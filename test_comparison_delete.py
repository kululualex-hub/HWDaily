"""Offline deletion tests: never connect to Google Sheets."""
import ast
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
state = State(user_role='管理者', installation_excel_last_result={'比較ID': 'A'})
sheet = Mock()
sheet.id = 123
sheet.get_all_records.return_value = [
    {'比較ID': 'A', '紀錄類型': '比較資訊'},
    {'比較ID': 'B'}, {'比較ID': 'A', '紀錄類型': '新增項目'},
    {'比較ID': 'A', '紀錄類型': '數量變更'}, {'比較ID': 'B'},
]
history, pending = Mock(), Mock()
ns = dict(st=SimpleNamespace(session_state=state),
          get_installation_comparison_worksheet=lambda: sheet,
          load_installation_excel_comparisons=history,
          load_unrecorded_installation_items=pending)
exec(compile(ast.Module(body=[node], type_ignores=[]), 'app.py', 'exec'), ns)
delete = ns['delete_installation_excel_comparison']
assert delete('A') == 3
assert sheet.batch_clear.call_args.args[0] == ['A2:M2', 'A4:M4', 'A5:M5']
assert state.installation_excel_last_result is None
assert state.installation_unrecorded_pending is None
history.clear.assert_called_once()
pending.clear.assert_called_once()
sheet.batch_clear.reset_mock()
assert delete('missing') == 0
sheet.batch_clear.assert_not_called()
state.user_role = '工程師'
try:
    delete('B')
    raise AssertionError('unauthorized deletion allowed')
except PermissionError:
    pass
state.user_role = '管理者'
sheet.batch_clear.side_effect = RuntimeError('network failure')
state.installation_excel_last_result = {'比較ID': 'A'}
try:
    delete('A')
    raise AssertionError('failure swallowed')
except RuntimeError:
    pass
assert state.installation_excel_last_result == {'比較ID': 'A'}
print('comparison deletion tests passed')
