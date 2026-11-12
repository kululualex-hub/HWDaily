"""Verify order progress includes exactly the exportable confirmed additions."""
import ast
import json
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

source = Path(__file__).with_name('app.py').read_text(encoding='utf-8')
tree = ast.parse(source)
compile(source, 'app.py', 'exec')
nodes = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name in
         {'mark_installation_items_recorded', 'load_installation_excel_comparisons'}]
for node in nodes:
    node.decorator_list = []
sheet = Mock()
base = {'比較ID': 'CMP1', '紀錄類型': '新增項目', '處理狀態': '已記錄',
        '新版本檔名': 'new.xlsx', '資料JSON': json.dumps({'訂單': '001', '品號': 'P1', '訂單數量': 3})}
ns = dict(json=json, datetime=datetime, get_installation_comparison_worksheet=lambda: sheet,
          st=SimpleNamespace(session_state=SimpleNamespace(user_name='test', user_role='管理者')),
          load_unrecorded_installation_items=Mock(), ws_log=Mock())
exec(compile(ast.Module(body=nodes, type_ignores=[]), 'app.py', 'exec'), ns)
ns['load_installation_excel_comparisons'].clear = Mock()
records = []
for category, included in [('新增項目', True), ('新增項目', False),
                           ('刪除項目', True), ('數量變更', True)]:
    record = {**base, '紀錄類型': category}
    sheet.get_all_records.return_value = [record]
    ns['mark_installation_items_recorded']([2], {'廠區': 'T8', '工程名稱': 'Test'} if included else None)
    payload = sheet.batch_update.call_args.args[0][0]['values'][0][0]
    records.append({**record, '資料JSON': payload})
# Existing confirmations without the former display flag must also appear.
old = {'訂單': 'OLD', '_報告統計': {'列入統計': True}}
records.append({**base, '資料JSON': json.dumps(old)})
records.append({**records[0], '處理狀態': '未紀錄'})
# Historical records without an explicit inclusion decision remain excluded.
records.append(base)
records.append({**base, '資料JSON': json.dumps({'訂單': 'FLAG_FALSE',
    '_報告統計': {'列入統計': True, '訂單進度顯示': False}})})
sheet.get_all_records.return_value = records
result = ns['load_installation_excel_comparisons']()[0]
assert len(result['已確認新增訂單']) == 3
row = result['已確認新增訂單'][0]
assert row['訂單'] == '001' and row['廠區'] == 'T8' and row['工程名稱'] == 'Test'
assert row['訂單數量'] == 3 and row['來源檔案'] == 'new.xlsx'
assert len(result['統計明細']['新增']) == 3
assert [r['訂單'] for r in result['已確認新增訂單']] == [
    r['訂單'] for r in result['統計明細']['新增']]
sheet.get_all_records.return_value = []
assert ns['load_installation_excel_comparisons']() == []
render = ast.unparse(next(n for n in tree.body if isinstance(n, ast.FunctionDef)
                         and n.name == 'render_order_progress_area'))
assert '已確認新增訂單' in render
assert 'build_comparison_order_progress' in render
print('confirmed comparison order tests passed')
