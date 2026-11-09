import ast
import io
import re
from datetime import datetime
from pathlib import Path
import xlsxwriter
from openpyxl import load_workbook

source = Path(__file__).with_name('app.py').read_text(encoding='utf-8')
compile(source, 'app.py', 'exec')
functions = {'build_installation_comparison_excel', 'build_saved_comparison_detail_excel'}
constants = {'INSTALLATION_ORDER_FIELDS', 'INSTALLATION_IDENTITY_FIELDS', 'INSTALLATION_QUANTITY_FIELDS', 'INSTALLATION_TARGET_PART_NUMBER'}
nodes = [n for n in ast.parse(source).body if
         isinstance(n, ast.FunctionDef) and n.name in functions or
         isinstance(n, ast.Assign) and any(isinstance(t, ast.Name) and t.id in constants for t in n.targets)]
ns = dict(io=io, re=re, datetime=datetime, xlsxwriter=xlsxwriter)
exec(compile(ast.Module(body=nodes, type_ignores=[]), 'app.py', 'exec'), ns)
row = {'日期': '2026-10-05', '訂單': '00001', '品名': '=test', '品號': '00002',
       '訂單數量': '5', '未交數量': '3', '已交數量': '2'}
changed = {**row, '新訂單數量': '5', '新未交數量': '1', '新已交數量': '4'}
comparison = {'新增': [row], '刪除': [row], '數量變更': [changed],
              '比較時間': '2026-10-15', '比較者': 'test', '前版檔名': 'old.xlsx', '新版本檔名': 'new.xlsx'}
empty = load_workbook(io.BytesIO(ns['build_saved_comparison_detail_excel']([comparison])))
assert all(s.max_row == 1 for s in list(empty)[1:])  # old/unmarked data excluded
comparison['統計明細'] = {'新增': [row], '刪除': [row], '數量變更': [changed]}
data = ns['build_saved_comparison_detail_excel']([comparison, comparison])
wb = load_workbook(io.BytesIO(data))
assert wb.sheetnames == ['比較摘要', '新增項目', '刪除項目', '數量變更']
for sheet in list(wb)[1:]:
    assert [c.value for c in sheet[1]] == ['日期', '訂單', '品名', '品號', '訂單數量', '未交數量', '已交數量']
    assert sheet.max_row == 3 and sheet.max_column == 7
    assert sheet['B2'].value == '00001' and sheet['D2'].value == '00002'
    assert sheet['C2'].data_type == 's'
    assert sheet['A2'].value == datetime(2026, 10, 5)
assert [wb['數量變更'].cell(2, col).value for col in [5, 6, 7]] == [5, 1, 4]
assert [wb['刪除項目'].cell(2, col).value for col in [5, 6, 7]] == [5, 3, 2]
original = load_workbook(io.BytesIO(ns['build_installation_comparison_excel'](comparison)))
assert original['新增項目'].max_column == 10
assert original['數量變更'].max_column == 13
assert '新已交數量' in [c.value for c in original['數量變更'][1]]
print('comparison detail export tests passed')
