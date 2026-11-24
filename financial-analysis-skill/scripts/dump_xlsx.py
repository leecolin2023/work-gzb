# -*- coding: utf-8 -*-
"""全量导出 xlsx：每张工作表逐行逐列输出 markdown 表格，公式与缓存值并列。"""
import openpyxl, sys, io, os

def fmt(v, f):
    if v is None: return ''
    if isinstance(v, str): return v.replace('|', '\\|')
    if isinstance(v, float):
        if v == int(v) and abs(v) < 1e15: return str(int(v))
        return ('%.10f' % v).rstrip('0').rstrip('.')
    return str(v)

def main(path, out):
    wbf = openpyxl.load_workbook(path, data_only=False)
    wbv = openpyxl.load_workbook(path, data_only=True)
    L = []
    L.append('# 原始数据全量导出（逐单元格，未做任何换算）\n')
    L.append('- 源文件：`%s`' % os.path.basename(path))
    L.append('- 工作表：%s\n' % ', '.join(wbf.sheetnames))
    for ws in wbf.worksheets:
        wv = wbv[ws.title]
        L.append('\n## 工作表：%s' % ws.title)
        L.append('- 已用区域：`%s`；行 %d，列 %d；合并单元格 %d 处：%s' % (
            ws.dimensions, ws.max_row, ws.max_column, len(ws.merged_cells.ranges),
            ', '.join(str(m) for m in sorted(ws.merged_cells.ranges, key=lambda x: (x.min_row, x.min_col))) or '无'))
        nf = sum(1 for r in ws.iter_rows() for c in r if isinstance(c.value, str) and c.value.startswith('='))
        L.append('- 公式单元格数：%d\n' % nf)
        cols = [openpyxl.utils.get_column_letter(c) for c in range(1, ws.max_column + 1)]
        L.append('| 行 | ' + ' | '.join(cols) + ' |')
        L.append('|---|' + '---|' * len(cols))
        for r in range(1, ws.max_row + 1):
            cells, has = [], False
            for c in range(1, ws.max_column + 1):
                cf = ws.cell(r, c).value; cv = wv.cell(r, c).value
                if cf is None and cv is None:
                    cells.append('')
                    continue
                has = True
                if isinstance(cf, str) and cf.startswith('='):
                    cells.append('`%s` → %s' % (cf, fmt(cv, None)))
                else:
                    cells.append(fmt(cv, None))
            if has:
                L.append('| %d | ' % r + ' | '.join(cells) + ' |')
    with io.open(out, 'w', encoding='utf-8') as f:
        f.write('\n'.join(L))
    print('written:', out, os.path.getsize(out), 'bytes')

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
