import ast
import json
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

tree = ast.parse(Path(__file__).with_name('app.py').read_text(encoding='utf-8'))
nodes = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name in
         {'mark_installation_items_recorded', 'load_installation_excel_comparisons'}]
for node in nodes:
    node.decorator_list = []
sheet = Mock()
base = {'比較ID': 'CMP1', '紀錄類型': '新增項目', '處理狀態': '未紀錄', '資料JSON': json.dumps({'訂單': '001'})}
sheet.get_all_records.return_value = [base]
ns = dict(json=json, datetime=datetime, get_installation_comparison_worksheet=lambda: sheet,
          st=SimpleNamespace(session_state=SimpleNamespace(user_name='test', user_role='管理者')),
          load_unrecorded_installation_items=Mock(), ws_log=Mock())
exec(compile(ast.Module(body=nodes, type_ignores=[]), 'app.py', 'exec'), ns)
ns['load_installation_excel_comparisons'].clear = Mock()
mark = ns['mark_installation_items_recorded']
mark([2], inclusion_details={'工程名稱': 'test'})
included_payload = sheet.batch_update.call_args.args[0][0]['values'][0][0]
assert json.loads(included_payload)['_報告統計']['列入統計'] is True
mark([2])
excluded_payload = sheet.batch_update.call_args.args[0][0]['values'][0][0]
assert json.loads(excluded_payload)['_報告統計']['列入統計'] is False
sheet.get_all_records.return_value = [
    {**base, '處理狀態': '已記錄', '資料JSON': included_payload},
    {**base, '處理狀態': '已記錄', '資料JSON': excluded_payload},
    {**base, '處理狀態': '已記錄'},
    {**base, '資料JSON': included_payload},
]
result = ns['load_installation_excel_comparisons']()[0]
assert len(result['新增']) == 4
assert len(result['統計明細']['新增']) == 1
assert all('_報告統計' not in row for row in result['新增'])
print('comparison inclusion tests passed')
