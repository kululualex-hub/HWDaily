"""Check the no-confirmation action without changing a live worksheet."""
import ast
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import gspread

source = Path(__file__).with_name("app.py").read_text(encoding="utf-8")
tree = ast.parse(source)
names = {"NEW_INSTALLATION_HEADERS", "CUSTOM_CASE_PENDING_SOURCE", "CUSTOM_CASE_SETUP_SOURCE",
         "CUSTOM_CASE_NOT_REQUIRED_SOURCE", "mark_custom_installation_confirmation_not_required",
         "select_pending_custom_installation_cases"}
nodes = [node for node in tree.body if (
    isinstance(node, ast.FunctionDef) and node.name in names
) or (
    isinstance(node, ast.Assign) and isinstance(node.targets[0], ast.Name)
    and node.targets[0].id in names
)]


class State(dict):
    __getattr__ = dict.__getitem__


state = State(user_role="管理者", user_name="測試", user_permissions={"installation_access": True})
sheet, loader = Mock(), Mock()
ns = dict(st=SimpleNamespace(session_state=state), gspread=gspread, datetime=datetime,
          get_new_installation_worksheet=lambda: sheet, load_new_installation_records=loader)
exec(compile(ast.Module(body=nodes, type_ignores=[]), "app.py", "exec"), ns)
mark = ns["mark_custom_installation_confirmation_not_required"]
select = ns["select_pending_custom_installation_cases"]
record = {"紀錄ID": "ID-1", "_new_sheet_row": 7, "案件": "原案件", "項目確認": "原確認項目"}
headers = ns["NEW_INSTALLATION_HEADERS"]
original = [""] * len(headers)
original[0] = record["紀錄ID"]
for source_name in ("CUSTOM_CASE_PENDING_SOURCE", "CUSTOM_CASE_SETUP_SOURCE"):
    original[headers.index("來源版本")] = ns[source_name]
    sheet.row_values.return_value = original
    mark(record)
    batch = sheet.batch_update.call_args
    assert batch.kwargs == {"value_input_option": "RAW"}
    changes = {item["range"]: item["values"][0][0] for item in batch.args[0]}
    assert set(changes) == {"C7", "N7", "O7"}
    assert changes["O7"] == ns["CUSTOM_CASE_NOT_REQUIRED_SOURCE"]
    assert changes["N7"] == "測試 (管理者)"
    assert not select([{**record, "來源版本": changes["O7"]}])
    assert select([{**record, "來源版本": ns[source_name]}])
assert loader.clear.call_count == 2
sheet.batch_update.reset_mock()
for identifier, status in [("ID-2", ns["CUSTOM_CASE_PENDING_SOURCE"]),
                           ("ID-1", ns["CUSTOM_CASE_NOT_REQUIRED_SOURCE"])]:
    original[0] = identifier
    original[headers.index("來源版本")] = status
    try:
        mark(record)
    except ValueError:
        pass
    else:
        raise AssertionError("stale record accepted")
state["user_permissions"] = {"installation_access": False}
try:
    mark(record)
except PermissionError:
    pass
else:
    raise AssertionError("unauthorized update accepted")
sheet.batch_update.assert_not_called()
print("custom case skip, preservation and permissions tests passed")
