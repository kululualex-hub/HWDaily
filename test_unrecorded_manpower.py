import ast
from pathlib import Path

source = Path(__file__).with_name('app.py').read_text(encoding='utf-8')
compile(source, 'app.py', 'exec')
node = next(n for n in ast.parse(source).body if isinstance(n, ast.FunctionDef)
            and n.name == 'unrecorded_item_display_category')
ns = {}
exec(compile(ast.Module(body=[node], type_ignores=[]), 'app.py', 'exec'), ns)
categorize = ns['unrecorded_item_display_category']
items = []
for category in ['新增項目', '刪除項目', '數量變更']:
    for part in ['PM-0001-00', ' pm-0001-00 ', 'PM-0001-001', 'OTHER', None]:
        item = {'品號': part, '差異分類': category}
        before = dict(item)
        assert categorize(item) == ('人力' if part in ['PM-0001-00', ' pm-0001-00 '] else category)
        assert item == before
        items.append(item)
groups = {category: [i for i in items if categorize(i) == category]
          for category in ['新增項目', '刪除項目', '數量變更', '人力']}
assert sum(map(len, groups.values())) == len(items)
assert len(groups['人力']) == 6
print('unrecorded manpower grouping tests passed')
