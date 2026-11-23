"""Roster lifecycle and history filtering without live Google Sheets access."""
import ast
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import gspread

source = Path(__file__).with_name("app.py").read_text(encoding="utf-8")
compile(source, "app.py", "exec")
names = {"build_installer_roster", "read_installer_roster", "load_installer_roster",
         "save_installer_status", "installation_editor_people"}
nodes = [n for n in ast.parse(source).body if isinstance(n, ast.FunctionDef) and n.name in names]
for node in nodes:
    node.decorator_list = []
sheet = Mock()
state = SimpleNamespace(user_role="管理者", user_id="admin")
state.get = lambda key: getattr(state, key)
ns = dict(sh=Mock(), st=SimpleNamespace(session_state=state), gspread=gspread,
          datetime=datetime, re=re, installers_list=["原有人員", "另一位"],
          TAIPEI_TZ=timezone(timedelta(hours=8)),
          INSTALLER_SETTINGS_SHEET="安裝人員設定",
          INSTALLER_SETTINGS_HEADERS=["姓名", "狀態", "更新時間", "更新者"])
exec(compile(ast.Module(body=nodes, type_ignores=[]), "app.py", "exec"), ns)
load, save = ns["load_installer_roster"], ns["save_installer_status"]
load.clear = Mock()
ns["sh"].worksheet.side_effect = gspread.exceptions.WorksheetNotFound
assert load() == {"原有人員": "啟用", "另一位": "啟用"}
ns["sh"].add_worksheet.assert_not_called()
ns["sh"].add_worksheet.return_value = sheet
sheet.get.return_value = [[]]  # Actual response from Google's blank tab.
sheet.row_values.return_value = []
save(" 新人 ", "啟用", new=True)
sheet.update.assert_called_once_with(range_name="A1:D1", values=[ns["INSTALLER_SETTINGS_HEADERS"]], raw=True)
assert sheet.append_row.call_args.args[0][:2] == ["新人", "啟用"]
assert sheet.append_row.call_args.kwargs["value_input_option"] == "RAW"
load.clear.assert_called_once()
ns["sh"].worksheet.side_effect = None
ns["sh"].worksheet.return_value = sheet
for blank in ([], [[]], [["", ""]]):
    sheet.get.return_value = blank
    assert load() == {"原有人員": "啟用", "另一位": "啟用"}
for invalid in ([["錯誤標題", "狀態"]], [[], ["新人", "啟用"]]):
    sheet.get.return_value = invalid
    try:
        load()
    except ValueError:
        pass
    else:
        raise AssertionError("unexpected roster schema accepted")
sheet.row_values.return_value = ns["INSTALLER_SETTINGS_HEADERS"]
sheet.get.return_value = [["姓名", "狀態"], ["新人", "啟用"], ["原有人員", "停用"]]
assert load() == {"原有人員": "停用", "另一位": "啟用", "新人": "啟用"}
save("原有人員", "啟用")
save("新人", "停用")
assert [c.args[0][:2] for c in sheet.append_row.call_args_list[-2:]] == [
    ["原有人員", "啟用"], ["新人", "停用"]
]
sheet.append_row.reset_mock()
for name, status, new in [("新人", "啟用", True), ("甲、乙", "啟用", True),
                          ("不存在", "停用", False), ("原有人員", "停用", False)]:
    try:
        save(name, status, new=new)
    except ValueError:
        pass
    else:
        raise AssertionError("invalid roster update accepted")
state.user_role = "工程師"
try:
    save("測試", "啟用", new=True)
except PermissionError:
    pass
else:
    raise AssertionError("non-admin update accepted")
sheet.append_row.assert_not_called()
people = ns["installation_editor_people"]
assert people("原有人員、停用人員\n原有人員", ["新人", "原有人員"]) == (
    ["原有人員", "停用人員", "新人"], ["原有人員", "停用人員"]
)
assert people("未指定", ["新人"]) == (["新人"], [])
assert people("Mary Lin；停用人員", []) == (["Mary Lin", "停用人員"], ["Mary Lin", "停用人員"])
assert 'edited_installers = "、".join(selected_installers)' in source
assert '"安裝人員": edited_installers or "未指定"' in source
assert "render_installer_management(installer_roster)" in source
assert 'for value in df_search["安裝人員"]' in source
print("installer roster tests passed")
