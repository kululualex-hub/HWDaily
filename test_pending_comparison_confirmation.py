"""Read-only mocked tests for delta preview and atomic manual confirmation."""
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
names = {'append_report_entry', 'report_integer', 'normalize_report_plant_name',
         'calculate_unrecorded_report_counts', 'confirm_pending_comparison'}
nodes = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name in names]
for n in nodes:
    n.decorator_list = []
comparison = Mock(id=11)
report = Mock(id=22)
report.col_values.return_value = ['建立時間', 'old']
state = SimpleNamespace(user_name='tester', user_role='管理者')
state.get = lambda key: getattr(state, key, None)
existing = {'年度':2026, '區域':'北', '廠區':'T8', '工程名稱':'Test',
            '建立時間':'2026-11-20 01:00:00', '訂單數量':10, '已出貨':4, '已安裝':2}
loader = Mock(return_value=pd.DataFrame([existing]))
headers = ['建立時間','建立者','區域','廠區','工程名稱','訂單數量','已出貨','未出貨',
           '已安裝','已出貨待安裝',*[f'{m}月份完成率' for m in range(4,13)],'年度']
ns = dict(pd=pd, re=re, json=json, datetime=datetime, st=SimpleNamespace(session_state=state),
          get_report_entry_worksheet=lambda:report, get_installation_comparison_worksheet=lambda:comparison,
          load_report_entry_records=loader, load_report_active_month=lambda:11,
          REPORT_ENTRY_HEADERS=headers, REPORT_MONTH_COLUMNS=headers[10:19], ws_log=Mock(),
          load_unrecorded_installation_items=Mock(), load_installation_excel_comparisons=Mock())
exec(compile(ast.Module(body=nodes, type_ignores=[]), 'app.py', 'exec'), ns)
raw = {'品號':'P1','訂單數量':3,'已交數量':0,'未交數量':3}
item = {'工作表列':2, '比較ID':'CMP1', '差異分類':'新增項目',
        '_原始差異資料':raw, **raw}
row = {'比較ID':'CMP1', '紀錄類型':'新增項目', '處理狀態':'未紀錄', '資料JSON':json.dumps(raw)}
comparison.get_all_records.return_value = [row]
calculate = ns['calculate_unrecorded_report_counts']
assert calculate(item, existing)['訂單變化'] == 3
assert calculate({**item,'差異分類':'刪除項目'},existing)['已出貨變化'] == 3
for old,new,expected in [(2,5,3),(5,2,0),(2,2,0)]:
    assert calculate({'差異分類':'數量變更','_原始差異資料':{'原已交數量':old,'新已交數量':new}},existing)['已出貨變化'] == expected
report.spreadsheet.batch_update.assert_not_called()
confirm = ns['confirm_pending_comparison']
confirm(item,'北','T8','Test',13,4,2,existing)
report.append_row.assert_not_called()
comparison.batch_update.assert_not_called()
requests = report.spreadsheet.batch_update.call_args.args[0]['requests']
assert len(requests) == 2
assert requests[0]['appendCells']['sheetId'] == 22
cells = requests[0]['appendCells']['rows'][0]['values']
assert cells[5]['userEnteredValue'] == {'numberValue':13}
assert cells[7]['userEnteredValue'] == {'formulaValue':'=F3-G3'}
update = requests[1]['updateCells']
assert update['range']['sheetId'] == 11 and update['range']['startRowIndex'] == 1
saved = update['rows'][0]['values']
assert saved[1]['userEnteredValue']['stringValue'] == '已記錄'
payload = json.loads(saved[0]['userEnteredValue']['stringValue'])
assert payload['_品號工程對照']['品號'] == 'P1'
assert payload['_報告統計']['列入統計'] is True
report.spreadsheet.batch_update.reset_mock()
comparison.get_all_records.return_value = [{**row,'處理狀態':'已記錄'}]
try:
    confirm(item,'北','T8','Test',13,4,2,existing)
    raise AssertionError('repeat accepted')
except ValueError:
    pass
report.spreadsheet.batch_update.assert_not_called()
comparison.get_all_records.return_value = [row]
loader.return_value = pd.DataFrame([{**existing,'訂單數量':12}])
try:
    confirm(item,'北','T8','Test',13,4,2,existing)
    raise AssertionError('stale report accepted')
except ValueError:
    pass
report.spreadsheet.batch_update.assert_not_called()
loader.return_value = pd.DataFrame([existing])
report.spreadsheet.batch_update.side_effect = RuntimeError('network failure')
try:
    confirm(item,'北','T8','Test',13,4,2,existing)
    raise AssertionError('failure hidden')
except RuntimeError:
    pass
comparison.batch_update.assert_not_called()
state.user_role = '工程師'
try:
    confirm(item,'北','T8','Test',13,4,2,existing)
    raise AssertionError('unauthorized confirmation')
except PermissionError:
    pass
assert 'scope="confirmation"' in source
assert '"訂單增加量": deltas["訂單變化"]' in source
assert '"已出貨增加量": deltas["已出貨變化"]' in source
print('pending comparison deltas, atomic confirmation and stale-state tests passed')
