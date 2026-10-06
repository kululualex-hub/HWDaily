"""Offline regression tests: no Streamlit startup or Google Sheets writes."""
import ast
import io
import re
import unittest
import zipfile
from datetime import datetime
from pathlib import Path
import xml.etree.ElementTree as ET

import pandas as pd
import xlsxwriter


source = Path(__file__).with_name('app.py').read_text(encoding='utf-8')
tree = ast.parse(source)
functions = {'build_report_statistics_excel', 'report_completion_rate_number', 'natural_plant_sort_key'}
constants = {'REPORT_AREA_ORDER', 'REPORT_COMPLETION_TARGETS', 'REPORT_COUNT_COLUMNS', 'REPORT_MONTH_COLUMNS'}
nodes = [node for node in tree.body if
         (isinstance(node, ast.FunctionDef) and node.name in functions) or
         (isinstance(node, ast.Assign) and any(isinstance(t, ast.Name) and t.id in constants for t in node.targets))]
ns = dict(pd=pd, io=io, re=re, datetime=datetime, xlsxwriter=xlsxwriter)
exec(compile(ast.Module(body=nodes, type_ignores=[]), 'app.py', 'exec'), ns)
build = ns['build_report_statistics_excel']
XML_NS = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
          's': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


class ReportExportTests(unittest.TestCase):
    def setUp(self):
        self.df = pd.DataFrame([
            {'年度': 2026, '區域': area, '廠區': f'T{i}', '工程名稱': f'工程{i}',
             '訂單數量': shipped + 2, '已出貨': shipped, '未出貨': 2,
             '已安裝': shipped * rate, '已出貨待安裝': shipped * (1-rate),
             '4月份完成率': rate, '5月份完成率': '', '6月份完成率': '待料'}
            for i, (area, shipped, rate) in enumerate([
                ('北', 10, .5), ('中', 30, 1), ('南', 20, .25), ('國外', 40, .75)])
        ])

    def charts(self, data):
        archive = zipfile.ZipFile(io.BytesIO(data))
        charts = [ET.fromstring(archive.read(name)) for name in archive.namelist()
                  if re.fullmatch(r'xl/charts/chart\d+\.xml', name)]
        return archive, [chart for chart in charts if chart.find('.//c:lineChart', XML_NS) is not None]

    def test_overall_and_custom_targets(self):
        targets = dict(zip(['全體', '北', '中', '南', '國外'], [.8, 0, .65, 1, None]))
        foreign_year = self.df.iloc[[0]].copy()
        foreign_year['年度'] = 2025
        archive, charts = self.charts(build(pd.concat([self.df, foreign_year]), 2026, targets))
        workbook = ET.fromstring(archive.read('xl/workbook.xml'))
        self.assertEqual([s.get('name') for s in workbook.findall('s:sheets/s:sheet', XML_NS)],
                         ['全體', '北', '中', '南', '國外'])
        self.assertEqual(len(charts), 5)
        for chart, target in zip(charts, targets.values()):
            series = chart.findall('.//c:lineChart/c:ser', XML_NS)
            self.assertEqual(len(series), 1 if target is None else 2)
            self.assertIsNotNone(series[0].find('c:dLbls', XML_NS))
            if target is not None:
                values = series[1].findall('c:val/c:numRef/c:numCache/c:pt/c:v', XML_NS)
                self.assertEqual([float(v.text) for v in values], [target] * 3)
        # Weighted across projects: (5 + 30 + 5 + 30) / 100 = 70%, not mean of regions.
        values = charts[0].findall('.//c:lineChart/c:ser/c:val/c:numRef/c:numCache', XML_NS)[0]
        points = values.findall('c:pt', XML_NS)
        self.assertEqual([(p.get('idx'), float(p.find('c:v', XML_NS).text)) for p in points], [('0', .7)])
        self.assertTrue(all('2026 年' in ''.join(chart.itertext()) for chart in charts))

    def test_defaults_disabled_and_no_data(self):
        _, charts = self.charts(build(self.df, 2026))
        self.assertEqual([len(c.findall('.//c:lineChart/c:ser', XML_NS)) for c in charts], [1, 2, 2, 2, 2])
        _, charts = self.charts(build(self.df, 2026, {s: None for s in ['全體', '北', '中', '南', '國外']}))
        self.assertTrue(all(len(c.findall('.//c:lineChart/c:ser', XML_NS)) == 1 for c in charts))
        _, charts = self.charts(build(self.df, 2024))
        self.assertEqual(charts, [])
        with self.assertRaises(ValueError):
            build(self.df, 2026, {'全體': 1.1})


if __name__ == '__main__':
    unittest.main()
