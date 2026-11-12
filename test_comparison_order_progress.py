import ast
from copy import deepcopy
from pathlib import Path
import re
import pandas as pd

source = Path(__file__).with_name('app.py').read_text(encoding='utf-8')
tree = ast.parse(source)
compile(source, 'app.py', 'exec')
nodes = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name in
         {'progress_order_without_prefix', 'normalize_progress_order', 'build_comparison_order_progress'}]
ns = dict(pd=pd, re=re)
exec(compile(ast.Module(body=nodes, type_ignores=[]), 'app.py', 'exec'), ns)
normalize, build = ns['normalize_progress_order'], ns['build_comparison_order_progress']
assert normalize(' 22r -20261112001-0001 ') == '20261112001'
assert normalize('220-20261112001-0002') == normalize('220-20261112001-0001')
assert normalize('20261112001') == normalize('220-20261112001-0001')
assert normalize('220-20261112001-123') == '20261112001-123'
assert normalize('ABC-123') == 'ABC-123'
for prefix in ['220', '22R', '230', '22A', 'SO', 'A01', '999', 'so']:
    assert normalize(f'{prefix}-20261112001-0001') == '20261112001'
assert normalize('20261112001-0001') == '20261112001'
assert normalize('20261112001') == '20261112001'
assert normalize(None) == ''
order = {'訂單': '220 -20261112001-0001', '品號': 'P1', '廠區': 'T8',
         '工程名稱': 'Test', '訂單數量': '2', '未交數量': 1, '已交數量': 1,
         '確認時間': '2026-11-12 01:00:00'}
record = {'訂單': '22R-20261112001-0001', '機台名稱': 'A01', '廠別': 'T8',
          '案件': 'Test', '狀態': '已完成', '日期': '2026-11-11'}
before = deepcopy(order)
result = build([order], [record, dict(record)])
assert len(result['進行中']) == 1
assert result['進行中'][0]['已安裝'] == 1
assert result['進行中'][0]['待安裝'] == 1
assert result['進行中'][0]['訂單數量'] == '2'
assert '未交數量' not in result['進行中'][0] and '已交數量' not in result['進行中'][0]
assert order == before
complete = build([order], [record, {**record, '機台名稱': 'A02'}])
assert len(complete['已完成']) == 1 and not complete['進行中']
assert complete['已完成'][0]['待安裝'] == 0
over = build([order], [{**record, '機台名稱': f'A{i}'} for i in range(3)])
assert over['已完成'][0]['已安裝'] == 3 and '提示' in over['已完成'][0]
assert build([order], [{**record, '狀態': '未完成'}])['進行中'][0]['已安裝'] == 0
assert build([order], [record, {**record, '日期': '2026-11-12', '狀態': '執行中'}])['進行中'][0]['已安裝'] == 0
unknown = build([order], [{**record, '訂單': '220-20261112999-0002'},
                          {**record, '訂單': ''}, {**record, '機台名稱': ''}])
assert len(unknown['無資料']) == 1 and unknown['進行中'][0]['已安裝'] == 0
assert unknown['無資料'][0]['訂單'] == '220-20261112999-0002'
for empty in ['', '  ', None]:
    assert not build([order], [{**record, '訂單': empty}])['無資料']
assert len(build([], [record])['無資料']) == 1
duplicate = build([order, {**order, '確認時間': '2026-11-12 02:00:00'}], [record])
assert len(duplicate['進行中']) == 1 and duplicate['進行中'][0]['已安裝'] == 1
ambiguous = build([order, {**order, '品號': 'P2'}], [record])
assert not ambiguous['無資料']
assert all('多個品項' in r['提示'] for r in ambiguous['進行中'])
assert sum(r['已安裝'] for r in ambiguous['進行中']) == 0
suffix_match = build([order], [{**record, '訂單': '22R-20261112001-0009'}])
assert suffix_match['進行中'][0]['已安裝'] == 1 and not suffix_match['無資料']
for prefix in ['230', '22A', 'SO', 'A01']:
    matched = build([order], [{**record, '訂單': f'{prefix}-20261112001-0009'}])
    assert matched['進行中'][0]['已安裝'] == 1 and not matched['無資料']
# Different suffixes for the same machine do not double-count installation.
suffix_duplicate = build([order], [record, {**record, '訂單': '20261112001-0009'}])
assert suffix_duplicate['進行中'][0]['已安裝'] == 1
# Distinct order lines must not be silently merged, even with the same part number.
multi_line = build([order, {**order, '訂單': '220-20261112001-0002'}], [record])
assert len(multi_line['進行中']) == 2 and not multi_line['無資料']
assert sum(r['已安裝'] for r in multi_line['進行中']) == 0
for quantity in ['', None, 'bad', 0, -1, 1.5]:
    invalid = build([{**order, '訂單數量': quantity}], [record])
    assert not invalid['已完成'] and invalid['進行中'][0]['待安裝'] is None
# Re-rendering is idempotent; no source quantity is decremented in storage.
assert build([order], [record]) == build([order], [record])
print('comparison order progress tests passed')

# Verify all three tabs render even without any legacy business orders.
from unittest.mock import MagicMock
ui = MagicMock()
ui.tabs.return_value = [MagicMock(), MagicMock(), MagicMock()]
ns.update(st=ui,
          load_installation_excel_comparisons=lambda: [{'已確認新增訂單': [order]}],
          load_new_installation_records=lambda: [record],
          legacy_records_for_new_interface=lambda: [],
          render_legacy_order_progress_area=MagicMock())
render_node = next(n for n in tree.body if isinstance(n, ast.FunctionDef)
                   and n.name == 'render_order_progress_area')
exec(compile(ast.Module(body=[render_node], type_ignores=[]), 'app.py', 'exec'), ns)
ns['render_order_progress_area']()
ui.tabs.assert_called_once_with(['進行中', '已完成', '無資料'])
ui.error.assert_not_called()
frame = ui.dataframe.call_args.args[0]
assert frame.iloc[0]['已安裝'] == 1 and frame.iloc[0]['待安裝'] == 1
assert '未交數量' not in frame.columns and '已交數量' not in frame.columns
print('order progress rendering tests passed')
