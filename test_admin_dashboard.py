"""Check dashboard routing and counts without using production Sheets."""
import ast
from datetime import datetime, timezone
import json
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import MagicMock
import pandas as pd

source = Path(__file__).with_name("app.py").read_text(encoding="utf-8")
tree = ast.parse(source)
compile(source, "app.py", "exec")
names = {"ADMIN_NAV_GROUPS", "ADMIN_PAGE_LABELS", "open_admin_page",
         "needs_installation_tracking", "render_admin_overview",
         "unrecorded_item_display_category", "render_admin_order_review",
         "select_pending_custom_installation_cases", "render_local_timetree_schedule",
         "render_morning_page"}
nodes = [n for n in tree.body if (isinstance(n, ast.FunctionDef) and n.name in names)
         or (isinstance(n, ast.Assign) and isinstance(n.targets[0], ast.Name)
             and n.targets[0].id in names)]


class State(dict):
    __setattr__ = dict.__setitem__

    def __getattr__(self, name):
        try:
            return self[name]
        except KeyError:
            raise AttributeError(name) from None


ui = MagicMock()
state = State(user_role="工程師", dev_active_tab="old")
ui.session_state = state
ui.columns.side_effect = lambda count: [MagicMock() for _ in range(count)]
new = [{"機台名稱": "A", "狀態": "已完成", "來源鍵": "OLD-1"},
       {"機台名稱": "B", "狀態": "未完成", "來源版本": "待確認"},
       {"機台名稱": "D", "狀態": "已完成", "來源版本": "待設定", "更新時間": "2026-11-25"}]
legacy = [{"機台名稱": "A", "狀態": "未完成", "來源鍵": "OLD-1"},
          {"機台名稱": "C", "狀態": "缺料", "來源鍵": "OLD-2"}]
comparison_ui, queue_ui = MagicMock(), MagicMock()
ns = dict(st=ui, pd=pd, load_new_installation_records=lambda: new,
          load_cloud_timetree_schedule=MagicMock(return_value=None),
          Path=Path, datetime=datetime, TAIPEI_TZ=timezone.utc, json=json,
          os=SimpleNamespace(environ={}),
          legacy_records_for_new_interface=lambda: legacy,
          CUSTOM_CASE_PENDING_SOURCE="待確認", CUSTOM_CASE_SETUP_SOURCE="待設定",
          load_unrecorded_installation_items=lambda: [
              {"品號": " pm-0001-00 ", "差異分類": "新增項目"},
              {"品號": "A", "差異分類": "刪除項目"},
              {"品號": "B", "差異分類": "數量變更"},
          ],
          render_installation_excel_version_area=comparison_ui,
          render_unrecorded_installation_items_panel=queue_ui,
          load_report_entry_records=lambda: pd.DataFrame())
exec(compile(ast.Module(body=nodes, type_ignores=[]), "app.py", "exec"), ns)
route = ns["open_admin_page"]
route("accounts")
assert "admin_page" not in state
ns["render_admin_order_review"]()
comparison_ui.assert_not_called()
queue_ui.assert_not_called()
state["user_role"] = "管理者"
route("invalid")
assert "admin_page" not in state
route("options")
assert state["admin_page"] == "options" and "dev_active_tab" not in state
route("installation")
assert state["admin_page"] == "installation"
route("custom_cases")
assert state["admin_page"] == "custom_cases"
assert [record["機台名稱"] for record in ns["select_pending_custom_installation_cases"](new)] == ["D", "B"]
ns["render_admin_overview"]()
assert [call.args for call in ui.metric.call_args_list] == [
    ("裝機待追蹤", 2), ("訂單差異待確認", 2), ("新版裝機的其他案件", 2)
]
destinations = {call.kwargs["args"][0] for call in ui.button.call_args_list}
assert destinations == {"tracking", "order_review", "custom_cases",
                        "add", "search", "options", "report"}
assert any(call.args == ("下拉式選項管理",) and call.kwargs["args"] == ("options",)
           for call in ui.button.call_args_list)
assert ns["ADMIN_PAGE_LABELS"]["attendance"] == "每月未出勤"
assert destinations <= ns["ADMIN_PAGE_LABELS"].keys()
assert source.count('render_unrecorded_installation_items_panel(pending_report_df, scope="confirmation")') == 1
ns["render_admin_order_review"]()
comparison_ui.assert_called_once_with()
queue_ui.assert_called_once()
assert queue_ui.call_args.kwargs == {"scope": "confirmation"}
assert source.count("    render_installation_excel_version_area()") == 1
ui.expander.reset_mock()
with TemporaryDirectory() as tmp:
    ns["os"].environ["LOCALAPPDATA"] = tmp
    path = Path(tmp) / "HWDaily-TimeTree-Test" / "daily.json"
    path.parent.mkdir()
    render = ns["render_local_timetree_schedule"]
    now = datetime.now(timezone.utc)
    data = {"checkedAt": now.isoformat(), "date": now.date().isoformat(),
            "status": "ok", "events": [{"time": "All-day", "title": "測試",
                                        "commentsStatus": "ok", "comments": ["測試文字留言"]}]}
    path.write_text(json.dumps(data), encoding="utf-8")
    render()
    ui.container.assert_called_once_with(border=True)
    assert [call.args for call in ui.text.call_args_list] == [
        ("工程｜測試",), ("測試文字留言",),
    ]
    ui.dataframe.assert_not_called()
    ui.expander.assert_called_once_with("查看留言（1 則）", expanded=False)
    ui.expander.reset_mock()
    unavailable = {**data["events"][0], "commentsStatus": "unavailable", "comments": []}
    path.write_text(json.dumps({**data, "events": [unavailable]}), encoding="utf-8")
    render()
    ui.caption.assert_called_with("留言尚未同步或無法讀取。")
    ui.expander.assert_not_called()
    ui.dataframe.reset_mock()
    path.write_text(json.dumps({**data, "date": "2000-01-01"}), encoding="utf-8")
    render()
    ui.warning.assert_called_with("今天尚未更新，不顯示前一天的行程。")
    path.write_text(json.dumps({**data, "status": "login_required"}), encoding="utf-8")
    render()
    ui.warning.assert_called_with("TimeTree 登入已失效，請在專用瀏覽器重新登入。")
    ui.dataframe.assert_not_called()
del ns["os"].environ["LOCALAPPDATA"]
cloud = ns["load_cloud_timetree_schedule"]
cloud.reset_mock()
cloud.return_value = data
ui.text.reset_mock()
state["user_role"] = "工程師"
render()
cloud.assert_not_called()
ui.text.assert_not_called()
state["user_role"] = "管理者"
render()
cloud.assert_called_once_with()
assert [call.args for call in ui.text.call_args_list] == [("工程｜測試",), ("測試文字留言",)]
cloud.side_effect = RuntimeError("private detail")
render()
ui.warning.assert_called_with("無法讀取 TimeTree 行程快照，請稍後再試。")

# Morning agenda is shown after search; only admins and guest sessions may read it.
cloud.side_effect = None
cloud.return_value = data
ui.button.return_value = False
state["tab1_search_active"] = False
morning = ns["render_morning_page"]
for role in ("工程師", "業務", "公用"):
    state["user_role"] = role
    ui.reset_mock()
    cloud.reset_mock()
    morning()
    cloud.assert_not_called()
    ui.text.assert_not_called()
    ui.subheader.assert_called_once_with("查詢晨會動態")
state["user_role"] = "管理者"
ui.reset_mock()
morning()
cloud.assert_called_once_with()
assert [call.args for call in ui.subheader.call_args_list] == [
    ("查詢晨會動態",), ("今日裝機行程",),
]
ui.expander.assert_called_once_with("查看留言（1 則）", expanded=False)
state["tab1_search_active"] = True
state["tab1_filtered_df"] = pd.DataFrame()
ui.reset_mock()
morning()
calls = [(call[0], call.args) for call in ui.method_calls]
assert calls.index(("info", ("📅 該日尚無裝機紀錄。",))) < calls.index(
    ("subheader", ("今日裝機行程",))
)
state.update(user_role="公用", user_id="guest", account_guest_session=True, logged_in=True)
cloud.reset_mock()
ui.reset_mock()
render()
cloud.assert_not_called()
morning()
cloud.assert_called_once_with()
assert [call.args for call in ui.text.call_args_list] == [("工程｜測試",), ("測試文字留言",)]
ui.expander.assert_called_once_with("查看留言（1 則）", expanded=False)
state["logged_in"] = False
cloud.reset_mock()
ui.reset_mock()
morning()
cloud.assert_not_called()
ui.text.assert_not_called()
ui.expander.assert_not_called()
print("admin dashboard navigation and count tests passed")
