"""Regression tests for TSMC plant aliases and non-TSMC preservation."""
import ast
import re
from pathlib import Path

source = Path(__file__).with_name('app.py').read_text(encoding='utf-8')
compile(source, 'app.py', 'exec')
nodes = [node for node in ast.parse(source).body if isinstance(node, ast.FunctionDef)
         and node.name in {'normalize_report_plant_name', 'plant_part_mapping_key'}]
ns = {'re': re}
exec(compile(ast.Module(body=nodes, type_ignores=[]), 'app.py', 'exec'), ns)
normalize = ns['normalize_report_plant_name']
for text in ['tsmc 12P7', 'tsmc-12P7', 'TSMC12P7', ' TSMC - 12p7 ', 'TSMC\u300012P7']:
    assert normalize(text) == 'T12P7', (text, normalize(text))
for text in ['tsmc 14A', 'tsmc-14A', 'TSMC14A']:
    assert normalize(text) == 'T14A'
for text in ['TSMC-08', 'tsmc 08', 'TSMC08']:
    assert normalize(text) == 'T8'
for text in ['TSMC-JP', 'tsmc JP', 'TSMCJP']:
    assert normalize(text) == 'JASM'
for text in ['UMC 8C', 'UMC 8A', 'VIS 2', 'T12P7', 'TSMX 12P7', 'TSMC', 'TSMCother']:
    assert normalize(text) == text, (text, normalize(text))
assert normalize('15B') == 'T15B'
assert normalize('TSMC-AP02') == 'AP02'  # Preserve existing hyphenated aliases.
assert normalize(None) == ''
key = ns['plant_part_mapping_key']
assert key('tsmc 12P7', 'P1') == key('TSMC-12P7', 'P1') == key('T12P7', 'P1')
assert key('UMC 8C', 'P1') == ('umc 8c', 'p1')
print('report plant normalization tests passed')
