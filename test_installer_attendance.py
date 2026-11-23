"""Attendance uses the current roster and counts from the selected month's start."""
import ast
import calendar
import re
from datetime import date, datetime
from pathlib import Path
from types import SimpleNamespace

import pandas as pd

source = Path(__file__).with_name("app.py").read_text(encoding="utf-8")
compile(source, "app.py", "exec")
names = {"parse_installation_record_date", "split_installer_names",
         "build_monthly_attendance_summary", "build_installer_roster"}
nodes = [node for node in ast.parse(source).body if isinstance(node, ast.FunctionDef)
         and node.name in names]
ns = dict(pd=pd, date=date, datetime=datetime, re=re, calendar=calendar,
          holidays=SimpleNamespace(country_holidays=lambda *args, **kwargs: {date(2026, 11, 4): "測試假日"}),
          installers_list=["原有", "停用人員"])
exec(compile(ast.Module(body=nodes, type_ignores=[]), "app.py", "exec"), ns)
build = ns["build_monthly_attendance_summary"]
roster = ns["build_installer_roster"]([
    ["停用人員", "停用", "2026-11-03 09:00:00"],
    ["原有", "停用", "2026-11-03 09:00:00"],
    ["原有", "啟用", "2026-11-06 09:00:00"],
    ["Mary Lin", "啟用", "2026-11-06 10:00:00"],
])
active = [name for name, status in roster.items() if status == "啟用"]
records = [{"日期": "2026-11-03", "安裝人員": "停用人員、Mary Lin"},
           {"日期": "2026-11-04", "安裝人員": "Mary Lin"}]
def november(installations):
    frame, _ = build(2026, 11, active, installations, date(2026, 11, 6))
    return frame.set_index("姓名")

data = november(records)
assert "停用人員" not in data.index
assert data.loc["原有", "未出勤日期"] == "11/02、11/03、11/05、11/06"
assert data.loc["Mary Lin", "整月應出勤天數"] == data.loc["原有", "整月應出勤天數"]
assert data.loc["Mary Lin", "未出勤日期"] == "11/02、11/05、11/06"
assert november(records[1:]).loc["Mary Lin", "未出勤日期"] == "11/02、11/03、11/05、11/06"
prior, _ = build(2026, 10, active, [], date(2026, 11, 6))
assert set(prior["姓名"]) == {"原有", "Mary Lin"}
assert build(2026, 11, [], records, date(2026, 11, 6))[0].empty
assert 'name for name, status in load_installer_roster().items() if status == "啟用"' in source
assert 'if record["來源鍵"] not in converted_keys' in source
print("installer attendance tests passed")
