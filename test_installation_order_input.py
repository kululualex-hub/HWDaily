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
