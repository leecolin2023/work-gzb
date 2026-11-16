#!/usr/bin/env python3
"""Evidence-linked Decimal formula ledger; no eval, implicit zeros or model selection.

See references/calculation-models.md for schema and limitations.
"""
import argparse
import ast
import copy
import json
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext
from pathlib import Path


def number(value):
    if isinstance(value, bool) or value is None:
        raise ValueError("numeric value required; unknown is not zero")
    result = Decimal(str(value))
    if not result.is_finite():
        raise ValueError("finite value required")
    return result


def expression(text, values):
    """Small arithmetic language; decimal literals retain their source precision."""
    if not isinstance(text, str) or len(text) > 4000:
        raise ValueError("formula must be a string of at most 4000 characters")
    tree = ast.parse(text, mode="eval")
    if sum(1 for _ in ast.walk(tree)) > 500:
        raise ValueError("formula too complex")
    dependencies = set()

    def visit(node):
        if isinstance(node, ast.Constant) and type(node.value) in (int, float):
            return number(ast.get_source_segment(text, node))
        if isinstance(node, ast.Name):
            if node.id not in values:
                raise ValueError("missing or forward dependency: " + node.id)
            dependencies.add(node.id)
            return values[node.id]
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub)):
            v = visit(node.operand)
            return v if isinstance(node.op, ast.UAdd) else -v
        if isinstance(node, ast.BinOp):
            a, b = visit(node.left), visit(node.right)
            if isinstance(node.op, ast.Add):
                return a + b
            if isinstance(node.op, ast.Sub):
                return a - b
            if isinstance(node.op, ast.Mult):
                return a * b
            if isinstance(node.op, ast.Div):
                return a / b
            if isinstance(node.op, ast.Pow):
                if abs(b) > 1000 or a <= 0 and b != b.to_integral_value():
                    raise ValueError("unsupported power domain")
                return a ** b
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
            if node.func.id in ("min", "max") and node.args and not node.keywords:
                return (min if node.func.id == "min" else max)(visit(x) for x in node.args)
        raise ValueError("unsupported expression element: " + type(node).__name__)

    with localcontext() as context:
        context.prec = 40
        result = visit(tree.body)
    return number(result), sorted(dependencies)


def metadata(item):
    if not isinstance(item.get("unit"), str) or not item["unit"].strip():
        raise ValueError("unit required")
    refs = item.get("evidence")
    if not isinstance(refs, list) or not refs or not all(isinstance(x, str) and x.strip() for x in refs):
        raise ValueError("nonempty evidence list required")


def calculate(params):
    if params.get("model") != "formula_graph":
        raise ValueError("model must be formula_graph; do not route unknown models to legacy formulas")
    values, rows = {}, []
    inputs = params.get("inputs")
    nodes = params.get("nodes")
    if not isinstance(inputs, dict) or not inputs or not isinstance(nodes, list) or not nodes:
        raise ValueError("nonempty inputs and nodes required")
    for name, item in inputs.items():
        if not name.isidentifier() or name in ("min", "max"):
            raise ValueError("invalid input id: " + name)
        metadata(item)
        values[name] = number(item.get("value"))
    for node in nodes:
        name = node.get("id", "")
        if not name.isidentifier() or name in values or name in ("min", "max"):
            raise ValueError("invalid or duplicate node id: " + name)
        metadata(node)
        raw, dependencies = expression(node.get("formula"), values)
        rounding = node.get("round_to")
        value = raw
        if rounding is not None:
            unit = number(rounding)
            if unit <= 0:
                raise ValueError("round_to must be positive")
            value = (raw / unit).quantize(Decimal(1), rounding=ROUND_HALF_UP) * unit
        values[name] = value
        reported = number(node["reported"]) if node.get("reported") is not None else None
        if reported is not None and "tolerance" not in node:
            raise ValueError("explicit tolerance required for reported node: " + name)
        tolerance = number(node["tolerance"]) if "tolerance" in node else None
        if tolerance is not None and tolerance < 0:
            raise ValueError("tolerance cannot be negative")
        diff = None if reported is None else value - reported
        # Independently substitute the report's own intermediate values to detect
        # local arithmetic errors separately from upstream propagated differences.
        reported_inputs = dict(values)
        for prior in rows:
            if prior["reported"] is not None:
                reported_inputs[prior["id"]] = prior["reported"]
        local, local_error = None, None
        try:
            local, _ = expression(node["formula"], reported_inputs)
            if rounding is not None:
                local = (local / unit).quantize(Decimal(1), rounding=ROUND_HALF_UP) * unit
        except (ValueError, ArithmeticError) as error:
            local_error = type(error).__name__ + ": " + str(error)
        rows.append({"id": name, "label": node.get("label", name),
                     "formula": node["formula"], "dependencies": dependencies,
                     "substitution": {k: values[k] for k in dependencies},
                     "raw": raw, "recomputed": value, "unit": node["unit"],
                     "round_to": rounding, "reported": reported, "difference": diff,
                     "tolerance": tolerance,
                     "exceeds_tolerance": None if diff is None else abs(diff) > tolerance,
                     "using_reported_intermediates": local,
                     "local_error": local_error,
                     "local_difference": None if reported is None or local is None else local - reported,
                     "reported_formula_text": node.get("reported_formula_text"),
                     "reported_substitution_text": node.get("reported_substitution_text"),
                     "evidence": node["evidence"]})
    result_id = params.get("result")
    if result_id not in {n["id"] for n in nodes}:
        raise ValueError("result must name a calculated node")
    return {"values": values, "checks": rows, "result_id": result_id, "result": values[result_id]}


def verify(params, include_sensitivity=False):
    with localcontext() as context:
        context.prec = 40
        base = calculate(params)
        result = {"model": "formula_graph", "params": params, **base,
                  "scope": "Arithmetic only; units, evidence truth, formula applicability and business impact require review."}
        if include_sensitivity:
            scenarios = []
            for scenario in params.get("scenarios", []):
                if not scenario.get("name") or not scenario.get("reason"):
                    raise ValueError("scenario name and reason required")
                overrides = scenario.get("overrides")
                if not isinstance(overrides, dict) or not overrides:
                    raise ValueError("nonempty scenario overrides required")
                q = copy.deepcopy(params)
                for key, override in overrides.items():
                    if key not in q["inputs"]:
                        raise ValueError("scenario may override input values only: " + key)
                    # Distinguish supported evidence from illustrative stress assumptions.
                    if not isinstance(override, dict) or override.get("basis") not in ("evidence", "stress"):
                        raise ValueError("override basis must be evidence or stress")
                    if not isinstance(override.get("evidence"), list) or not override["evidence"]:
                        raise ValueError("scenario evidence or explicit stress rationale required")
                    q["inputs"][key]["value"] = str(number(override.get("value")))
                    q["inputs"][key]["evidence"] = override["evidence"]
                changed = calculate(q)
                scenarios.append({"name": scenario["name"], "reason": scenario["reason"],
                                  "overrides": overrides, "result": changed["result"],
                                  "delta": changed["result"] - base["result"],
                                  "values": changed["values"]})
            result["sensitivity"] = scenarios
        return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--params", required=True)
    parser.add_argument("--out")
    parser.add_argument("--sensitivity", action="store_true")
    args = parser.parse_args()
    try:
        params = json.loads(Path(args.params).read_text(encoding="utf-8"), parse_float=str)
        result = verify(params, args.sensitivity)
    except (ValueError, KeyError, TypeError, SyntaxError, InvalidOperation, ArithmeticError) as error:
        parser.exit(2, "Cannot verify: " + str(error) + "\n")
    text = json.dumps(result, ensure_ascii=False, indent=2,
                      default=lambda value: str(value) if isinstance(value, Decimal) else value)
    if args.out:
        dest = Path(args.out)
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_text(text + "\n", encoding="utf-8")
        print("written:", dest)
    else:
        print(text)


if __name__ == "__main__":
    main()
