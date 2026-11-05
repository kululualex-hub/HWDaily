"""Offline regression checks for stored baselines and missing worksheet."""
import ast
import io
import json
import re
from pathlib import Path
from unittest.mock import Mock

tree = ast.parse(Path(__file__).with_name('app.py').read_text(encoding='utf-8'))
names = {'load_installation_excel_versions', 'installation_excel_display_row',
         'normalize_order_number', 'parse_installation_excel'}
constants = {'INSTALLATION_ORDER_FIELDS', 'INSTALLATION_TARGET_PART_NUMBER'}
nodes = []
for node in tree.body:
    if isinstance(node, ast.FunctionDef) and node.name in names:
        node.decorator_list = []
        nodes.append(node)
    elif isinstance(node, ast.Assign) and any(isinstance(t, ast.Name) and t.id in constants for t in node.targets):
        nodes.append(node)
worksheet = Mock()
workbook = Mock(sheetnames=['已確認，未結案 (含合約訂單欄位)'])
ns = dict(json=json, re=re, io=io, load_workbook=lambda *a, **kw: workbook,
          get_installation_excel_worksheet=lambda: worksheet)
exec(compile(ast.Module(body=nodes, type_ignores=[]), 'app.py', 'exec'), ns)
def record(order, part, sheet):
    data = dict.fromkeys(ns['INSTALLATION_ORDER_FIELDS'], '')
    data.update(訂單=order, 品號=part)
    return {'訂單': 'stream', '品號': 'PM-0001-00', '版本ID': 'v1', '紀錄類型': '資料',
            '工作表': sheet, '資料JSON': json.dumps(data)}
worksheet.get_all_records.return_value = [
    record('A-0001', 'PM-0001-00', '已確認，未結案'),
    record('A-0002', 'PART', '已確認，未結案'),
    record('B-0001', 'PM-0001-00', '其他'),
    record('B-0002', 'PART', '已確認，未結案'),
]
version = ns['load_installation_excel_versions']('stream', 'PM-0001-00')[0]
assert len(version['資料']) == 2
assert version['工作表列'] == [2, 3, 4, 5]  # retain cleanup metadata
try:
    ns['parse_installation_excel'](b'')
    raise AssertionError('missing exact worksheet accepted')
except ValueError as error:
    assert '已確認，未結案' in str(error)
workbook.close.assert_called_once()
print('comparison sheet scope tests passed')
