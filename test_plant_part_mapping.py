"""Validate persisted plant/part mappings and pending-item classification."""
import ast
import json
import re
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock
import pandas as pd

source = Path(__file__).with_name('app.py').read_text(encoding='utf-8')
tree = ast.parse(source)
compile(source, 'app.py', 'exec')
names = {'normalize_report_plant_name', 'plant_part_mapping_key',
         'build_plant_part_project_mappings', 'load_plant_part_project_mappings',
         'load_unrecorded_installation_items', 'mark_installation_items_recorded'}
nodes = [node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in names]
for node in nodes:
    node.decorator_list = []
sheet = Mock()
ns = dict(pd=pd, json=json, re=re, datetime=datetime,
          get_installation_comparison_worksheet=lambda: sheet,
          st=SimpleNamespace(session_state=SimpleNamespace(user_name='test', user_role='管理者')),
          ws_log=Mock(), load_installation_excel_comparisons=Mock(),
          installation_excel_display_row=lambda data: dict(data),
          INSTALLATION_IDENTITY_FIELDS=['客戶簡稱', '品號'],
          INSTALLATION_QUANTITY_FIELDS=['訂單數量', '未交數量', '已交數量'])
exec(compile(ast.Module(body=nodes, type_ignores=[]), 'app.py', 'exec'), ns)
ns['load_unrecorded_installation_items'].clear = Mock()
base = {'比較ID': 'C1', '紀錄類型': '新增項目', '處理狀態': '未紀錄',
        '資料JSON': json.dumps({'客戶簡稱': 'tsmc-14A', '品號': 'P1', '訂單': '001'})}
sheet.get_all_records.return_value = [base]
ns['mark_installation_items_recorded']([2], {'廠區': 'T14A', '工程名稱': '光學尺'})
payload = sheet.batch_update.call_args.args[0][0]['values'][0][0]
assert json.loads(payload)['_品號工程對照']['工程名稱'] == '光學尺'
saved = {**base, '處理狀態': '已記錄', '資料JSON': payload}
build = ns['build_plant_part_project_mappings']
key = ns['plant_part_mapping_key']
assert build([saved])[key('TSMC-14A', 'p1')]['工程名稱'] == '光學尺'
assert key('UMC 8C', ' P1 ') == ('umc 8c', 'p1')
other_payload = json.loads(payload)
other_payload['_品號工程對照'].update({'廠區': 'T15B', '工程名稱': '心電圖'})
other = {**saved, '資料JSON': json.dumps(other_payload)}
assert len(build([saved, other])) == 2
changed_payload = json.loads(payload)
changed_payload['_品號工程對照'].update({'工程名稱': '新工程', '更新時間': '2099-01-01 00:00:00'})
changed = {**saved, '資料JSON': json.dumps(changed_payload)}
assert build([changed, saved])[key('T14A', 'P1')]['工程名稱'] == '新工程'
ns['mark_installation_items_recorded']([2])
excluded = sheet.batch_update.call_args.args[0][0]['values'][0][0]
assert not build([{**saved, '資料JSON': excluded}])
assert not build([{**saved, '資料JSON': 'invalid'}, {**saved, '處理狀態': '未紀錄'}])
# Only pending items are classified; reading does not mark them included.
sheet.get_all_records.return_value = [saved, base, {**base, '資料JSON': json.dumps({'客戶簡稱':'T8','品號':'P1'})}]
sheet.batch_update.reset_mock()
pending = ns['load_unrecorded_installation_items']()
assert len(pending) == 2
assert pending[0]['配對工程'] == '光學尺' and pending[0]['紀錄狀態'] == '未紀錄'
assert pending[1]['工程配對狀態'] == '待配對'
sheet.batch_update.assert_not_called()
assert '工程分項（廠區＋工程名稱）' in source
assert '廠區＋品號與工程對照' in source
print('plant/part mapping persistence, isolation and classification tests passed')
