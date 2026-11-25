"""Check dashboard routing and counts without using production Sheets."""
import ast
from pathlib import Path
from unittest.mock import MagicMock
import pandas as pd

source = Path(__file__).with_name("app.py").read_text(encoding="utf-8")
tree = ast.parse(source)
compile(source, "app.py", "exec")
names = {"ADMIN_NAV_GROUPS", "ADMIN_PAGE_LABELS", "open_admin_page",
         "needs_installation_tracking", "render_admin_overview",
         "unrecorded_item_display_category", "render_admin_order_review",
         "select_pending_custom_installation_cases"}
nodes = [n for n in tree.body if (isinstance(n, ast.FunctionDef) and n.name in names)
         or (isinstance(n, ast.Assign) and isinstance(n.targets[0], ast.Name)
             and n.targets[0].id in names)]


class State(dict):
    __setattr__ = dict.__setitem__


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
print("admin dashboard navigation and count tests passed")
