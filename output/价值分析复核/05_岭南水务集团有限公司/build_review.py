"""Check the main deliverables and evidence references for one completed case."""
import json
import re
from pathlib import Path
from zipfile import ZipFile


def main():
    root = Path(__file__).resolve().parent
    case = root.name.split("_", 1)[0]
    read = lambda name: json.loads((root / name).read_text(encoding="utf-8"))
    review = read(f"review_data_{case}.json")
    assert len({row[0] for row in review["checks"]}) == 51
    evidence = {
        label: {x["id"] for x in read(f"evidence_{key}.json")["items"]}
        for label, key in (("VR", case), ("DD", "DD" + case))
    }
    for issue in review["issues"]:
        for kind, key in re.findall(r"(VR|DD):(P\d{4}|T\d{2}-R\d{2}-C\d{2})", issue[6]):
            assert key in evidence[kind], (kind, key)
    for pattern in ("*_价值分析复核意见.md", "*_价值分析报告叙事解读.md", "*_数据来源与定位索引.md"):
        assert len(list(root.glob(pattern))) == 1, pattern
    with ZipFile(next(root.glob("*_复核底稿.xlsx"))) as book:
        assert book.testzip() is None
    print(case, "51 checks, evidence links and workbook archive verified")


if __name__ == "__main__":
    main()
