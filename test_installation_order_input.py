import ast
from pathlib import Path

source = Path(__file__).with_name('app.py').read_text(encoding='utf-8')
compile(source, 'app.py', 'exec')
node = next(n for n in ast.parse(source).body if isinstance(n, ast.FunctionDef)
            and n.name == 'installation_order_fields')
ns = {}
exec(compile(ast.Module(body=[node], type_ignores=[]), 'app.py', 'exec'), ns)
resolve = ns['installation_order_fields']
assert resolve(None, '  001-20261111-0001  ') == {'訂單': '001-20261111-0001', '業務紀錄ID': ''}
assert resolve(None, '') == {'訂單': '', '業務紀錄ID': ''}
assert resolve(None, None) == {'訂單': '', '業務紀錄ID': ''}
assert resolve({'訂單': ' EXISTING ', '業務紀錄ID': ' id1 '}, 'manual') == {'訂單': 'EXISTING', '業務紀錄ID': 'id1'}
print('installation order input tests passed')
initial_step = source[source.index('                matched_order_entries = matching_sales_orders('):]
assert 'identity_data["廠別"]' in initial_step[:200]
assert source.rindex('"訂單編號（選填）"') > source.index('f"已確認：{identity_data')
assert '**installation_order_fields(selected_order_record, manual_order_number)' in source
from datetime import datetime
from types import SimpleNamespace
import uuid
row_node = next(n for n in ast.parse(source).body if isinstance(n, ast.FunctionDef)
                and n.name == 'new_installation_row_values')
row_ns = dict(datetime=datetime, uuid=uuid,
              st=SimpleNamespace(session_state=SimpleNamespace(user_name='test', user_role='管理者')))
exec(compile(ast.Module(body=[row_node], type_ignores=[]), 'app.py', 'exec'), row_ns)
values = row_ns['new_installation_row_values']
old = {'訂單': 'OLD', '業務紀錄ID': 'id1'}
assert values({'訂單': 'NEW', '業務紀錄ID': ''}, old)[-2:] == ['NEW', '']
assert values({'訂單': '', '業務紀錄ID': ''}, old)[-2:] == ['', '']
assert values({}, old)[-2:] == ['OLD', 'id1']
print('history order persistence tests passed')
