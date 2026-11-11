"""Exercise order matching and the actual history-filter loop without cloud writes."""
import ast
from pathlib import Path
import re
import pandas as pd

source = Path(__file__).with_name('app.py').read_text(encoding='utf-8')
tree = ast.parse(source)
compile(source, 'app.py', 'exec')
function = next(n for n in tree.body if isinstance(n, ast.FunctionDef)
                and n.name == 'installation_order_matches')
ns = {}
exec(compile(ast.Module(body=[function], type_ignores=[]), 'app.py', 'exec'), ns)
match = ns['installation_order_matches']
assert match({'訂單': '22R -20261111001-0001'}, '22r-20261111001')
assert match({'訂單': 'ABC.01'}, 'c.0')
assert not match({'訂單': 'ABCx01'}, 'c.0')
assert not match({}, '001')
assert not match({'訂單': None}, 'none')
assert match({}, '  ')
assert match({'訂單': '00123'}, '001')
loop = next(n for n in ast.walk(tree) if isinstance(n, ast.For)
            and isinstance(n.target, ast.Tuple)
            and ast.unparse(n.target) == '(record_index, record)'
            and 'installation_order_matches' in ast.unparse(n))
records = [
    {'訂單': '220 -20261111001-0001', '廠別': 'T8', '日期': '2026-11-11'},
    {'訂單': '220-20261111001-0001', '廠別': 'T3', '日期': '2026-11-11'},
    {'訂單': 'OTHER', '廠別': 'T8', '日期': '2026-11-11'},
    {'廠別': 'T8', '日期': '2026-11-11'},
]
ns.update(pd=pd, re=re, combined_results_records=records,
          dev_search_order='220-20261111001', dev_search_plant='T8',
          dev_search_case=[], dev_search_machine='（全部）',
          dev_search_installer='（全部）', dev_search_status='（全部）',
          dev_search_dates=[], filtered_record_indices=[])
code = compile(ast.Module(body=[loop], type_ignores=[]), 'app.py', 'exec')
exec(code, ns)
assert ns['filtered_record_indices'] == [0]
ns.update(dev_search_order='', filtered_record_indices=[])
exec(code, ns)
assert ns['filtered_record_indices'] == [0, 2, 3]
assert 'results_df.iloc[filtered_record_indices].copy()' in source
assert 'data=build_history_search_excel(display_results_df)' in source
print('order search and combined filtering tests passed')
