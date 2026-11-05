"""Offline deletion tests: never connect to Google Sheets."""
import ast
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

class State(dict):
    __getattr__ = dict.__getitem__
    __setattr__ = dict.__setitem__

source = Path(__file__).with_name('app.py').read_text(encoding='utf-8')
compile(source, 'app.py', 'exec')
node = next(n for n in ast.parse(source).body
            if isinstance(n, ast.FunctionDef) and n.name == 'delete_installation_excel_comparison')
state = State(user_role='管理者', user_name='tester', installation_excel_last_result={'比較ID': 'A'})
sheet = Mock()
sheet.id = 123
sheet.get_all_records.return_value = [
    {'比較ID': 'A', '紀錄類型': '比較資訊', '前版ID': '20261005170016000000', '新版本ID': 'v2', '前版檔名': 'previous.xlsx'},
    {'比較ID': 'B'}, {'比較ID': 'A', '紀錄類型': '新增項目'},
    {'比較ID': 'A', '紀錄類型': '數量變更'}, {'比較ID': 'B'},
]
history, pending = Mock(), Mock()
versions = Mock()
version_sheet = Mock(id=456)
version_sheet.get_all_records.return_value = [
    {'版本ID': version, '訂單': '客戶訂單明細表', '品號': 'PM-0001-00',
     '上傳時間': version, '紀錄類型': '資料', '工作表': '已確認，未結案'}
    for version in ['20261005170016000000', 'v2']
]
parser = Mock(return_value={'資料': [{'工作表': '已確認，未結案', 'Excel列': 2,
                                      '資料JSON': '{}', '內容': 'test'}]})
ns = dict(st=SimpleNamespace(session_state=state),
          datetime=datetime, parse_installation_excel=parser,
          get_installation_comparison_worksheet=lambda: sheet,
          get_installation_excel_worksheet=lambda: version_sheet,
          load_installation_excel_versions=versions,
          INSTALLATION_TARGET_PART_NUMBER='PM-0001-00',
          load_installation_excel_comparisons=history,
          load_unrecorded_installation_items=pending)
exec(compile(ast.Module(body=[node], type_ignores=[]), 'app.py', 'exec'), ns)
delete = lambda identity: ns['delete_installation_excel_comparison'](identity, 'previous.xlsx', b'excel')
assert delete('A') == 3
requests = sheet.spreadsheet.batch_update.call_args.args[0]['requests']
assert [(r['updateCells']['range']['sheetId'], r['updateCells']['range']['startRowIndex'])
        for r in requests if 'updateCells' in r] == [(123, 1), (123, 3), (123, 4), (456, 1), (456, 2)]
assert requests[-1]['appendCells']['sheetId'] == 456
assert len(requests[-1]['appendCells']['rows']) == 2
assert requests[-1]['appendCells']['rows'][0]['values'][0]['userEnteredValue']['stringValue'] == '20261005170016000000'
assert state.installation_excel_last_result is None
assert state.installation_unrecorded_pending is None
history.clear.assert_called_once()
pending.clear.assert_called_once()
sheet.spreadsheet.batch_update.reset_mock()
assert delete('missing') == 0
sheet.spreadsheet.batch_update.assert_not_called()
state.user_role = '工程師'
try:
    delete('B')
    raise AssertionError('unauthorized deletion allowed')
except PermissionError:
    pass
state.user_role = '管理者'
sheet.spreadsheet.batch_update.side_effect = RuntimeError('network failure')
state.installation_excel_last_result = {'比較ID': 'A'}
try:
    delete('A')
    raise AssertionError('failure swallowed')
except RuntimeError:
    pass
assert state.installation_excel_last_result == {'比較ID': 'A'}
sheet.spreadsheet.batch_update.side_effect = None
sheet.spreadsheet.batch_update.reset_mock()
version_sheet.get_all_records.return_value = version_sheet.get_all_records.return_value[1:]
assert delete('A') == 3  # absent previous version is restored from attachment
sheet.spreadsheet.batch_update.reset_mock()
for name, content in [('wrong.xlsx', b'excel'), ('previous.xlsx', b''), ('previous.xlsx', b'x' * (10 * 1024 * 1024 + 1))]:
    try:
        ns['delete_installation_excel_comparison']('A', name, content)
        raise AssertionError('invalid attachment accepted')
    except ValueError:
        pass
sheet.spreadsheet.batch_update.assert_not_called()
parser.side_effect = ValueError('invalid workbook')
try:
    delete('A')
    raise AssertionError('invalid workbook accepted')
except ValueError:
    pass
parser.side_effect = None
sheet.spreadsheet.batch_update.assert_not_called()
sheet.get_all_records.return_value.append({'比較ID': 'C', '前版ID': 'v2', '新版本ID': 'v3'})
try:
    delete('A')
    raise AssertionError('older comparison accepted')
except ValueError:
    pass
sheet.spreadsheet.batch_update.assert_not_called()
print('comparison deletion tests passed')
