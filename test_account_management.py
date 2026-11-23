"""Account lifecycle tests using fake worksheets; never reads real credentials."""
import ast
import hashlib
import hmac
import secrets
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock
import gspread

source = Path(__file__).with_name('app.py').read_text(encoding='utf-8')
compile(source, 'app.py', 'exec')
tree = ast.parse(source)
names = {'account_password_hash', 'account_password_matches', 'read_managed_accounts', 'save_managed_account'}
nodes = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name in names]
sheet, log = Mock(), Mock()
state = {'user_id':'admin','user_role':'管理者'}
ns = dict(hashlib=hashlib, hmac=hmac, secrets=secrets, datetime=datetime, gspread=gspread,
          ws_accounts=sheet, ws_log=log, st=SimpleNamespace(session_state=state),
          ROLE_PERMISSION_DEFAULTS={'管理者':{},'工程師':{},'業務':{}})
exec(compile(ast.Module(body=nodes, type_ignores=[]), 'app.py', 'exec'), ns)
headers = ['帳號','密碼','姓名','權限','保留欄']
baseline = [headers, ['admin','old-secret','Admin','管理者','keep'],
            ['001','old-secret','Engineer','工程師','untouched']]
sheet.get_all_values.return_value = baseline
hash_password, verify = ns['account_password_hash'], ns['account_password_matches']
hashed = hash_password('test-password')
assert verify('test-password', hashed) and not verify('incorrect', hashed)
assert verify('old-secret','old-secret') and not verify('old','old-secret')
assert not verify('', '') and not verify('x','pbkdf2_sha256$bad')
assert hash_password('test-password') != hashed
save = ns['save_managed_account']
save('新增','002','New','業務','new-secret')
row = sheet.append_row.call_args.args[0]
assert row[0] == '002' and row[2:4] == ['New','業務']
assert row[1] != 'new-secret' and verify('new-secret', row[1])
assert sheet.append_row.call_args.kwargs['value_input_option'] == 'RAW'
assert 'new-secret' not in str(log.mock_calls) and hashed not in str(log.mock_calls)
sheet.reset_mock()
save('修改','001','Changed','業務','')
updates = sheet.batch_update.call_args.args[0]
assert all(update['range'] != 'B3' for update in updates)
assert all(update['range'] != 'E3' for update in updates)
sheet.reset_mock()
save('修改','001','Changed','工程師','another-secret')
updates = sheet.batch_update.call_args.args[0]
assert verify('another-secret', next(u for u in updates if u['range']=='B3')['values'][0][0])
save('刪除','001')
sheet.batch_clear.assert_called_once_with(['A3:E3'])
def rejects(action, target, name='', role='', password=''):
    sheet.reset_mock()
    try:
        save(action,target,name,role,password)
        raise AssertionError('invalid operation accepted')
    except (ValueError,PermissionError):
        pass
    sheet.batch_clear.assert_not_called()
    sheet.batch_update.assert_not_called()
    sheet.append_row.assert_not_called()
rejects('新增','ADMIN','x','工程師','new-secret')
rejects('新增','guest','x','工程師','new-secret')
rejects('新增','003','x','工程師','short')
rejects('修改','001','x','unknown')
rejects('刪除','missing')
rejects('刪除','admin')
rejects('修改','admin','Admin','工程師')
sheet.get_all_values.return_value = [*baseline, ['admin2','secret','Other','管理者','']]
rejects('刪除','admin')
state['user_role'] = '工程師'
rejects('刪除','001')
state['user_role'] = '管理者'
sheet.get_all_values.return_value = [headers, ['admin','secret','Admin','工程師','']]
rejects('新增','003','x','業務','new-secret')
assert 'account_credential_version != credential_version' in source
assert 'render_account_management()' in source
print('account management security and lifecycle tests passed')
