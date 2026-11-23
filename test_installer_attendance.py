"""Attendance follows roster dates and the updated installation records."""
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
         "build_monthly_attendance_summary", "parse_installer_attendance_events"}
nodes = [node for node in ast.parse(source).body if isinstance(node, ast.FunctionDef)
         and node.name in names]
ns = dict(pd=pd, date=date, datetime=datetime, re=re, calendar=calendar,
          holidays=SimpleNamespace(country_holidays=lambda *args, **kwargs: {date(2026, 11, 4): "測試假日"}),
          INSTALLER_SETTINGS_HEADERS=["姓名", "狀態", "更新時間", "更新者"])
exec(compile(ast.Module(body=nodes, type_ignores=[]), "app.py", "exec"), ns)
parse = ns["parse_installer_attendance_events"]
build = ns["build_monthly_attendance_summary"]
assert parse([[]]) == []
events = parse([ns["INSTALLER_SETTINGS_HEADERS"][:3],
                ["原有", "停用", "2026-11-03 09:00:00"],
                ["Mary Lin", "啟用", "2026-11-03 10:00:00"],
                ["原有", "啟用", "2026-11-06 12:00:00"],
                ["Mary Lin", "停用", "2026-11-06 12:00:00"]])
records = [{"日期": "2026-11-03", "安裝人員": "原有、Mary Lin"},
           {"日期": "2026-11-04", "安裝人員": "Mary Lin"}]
def november(installations):
    frame, _ = build(2026, 11, ["原有"], installations, date(2026, 11, 6), events)
    return frame.set_index("姓名")

data = november(records)
assert data.loc["原有", "未出勤日期"] == "11/02、11/06"
assert data.loc["原有", "有裝機紀錄天數"] == 0  # Inactive on day 3.
assert data.loc["Mary Lin", "整月應出勤天數"] == 2
assert data.loc["Mary Lin", "未出勤日期"] == "11/05"
assert november(records[1:]).loc["Mary Lin", "未出勤日期"] == "11/03、11/05"
prior, _ = build(2026, 10, ["原有"], [], date(2026, 11, 6), events)
assert list(prior["姓名"]) == ["原有"]
same_day = [("原有", "停用", date(2026, 11, 1)), ("原有", "啟用", date(2026, 11, 1))]
frame, _ = build(2026, 11, ["原有"], [], date(2026, 11, 2), same_day)
assert frame.iloc[0]["截至統計日應出勤天數"] == 1
assert 'if record["來源鍵"] not in converted_keys' in source
print("installer attendance tests passed")
