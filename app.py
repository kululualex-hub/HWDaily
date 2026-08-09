import streamlit as st
import gspread
import pandas as pd
from datetime import datetime
import io
import xlsxwriter

# ==================== 1. 網頁基本與連線設定 ====================
st.set_page_config(page_title="鴻伍裝機日報系統", layout="wide")

# 初始化 Session State (記憶登入狀態與暫存資料)
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
if 'user_name' not in st.session_state:
    st.session_state.user_name = ""
if 'user_role' not in st.session_state:
    st.session_state.user_role = ""

if 'tab1_search_active' not in st.session_state:
    st.session_state.tab1_search_active = False
if 'tab1_filtered_df' not in st.session_state:
    st.session_state.tab1_filtered_df = pd.DataFrame()
    
if 'tab3_search_active' not in st.session_state:
    st.session_state.tab3_search_active = False
if 'tab3_filtered_df' not in st.session_state:
    st.session_state.tab3_filtered_df = pd.DataFrame()
if 'tab3_edit_requested' not in st.session_state:
    st.session_state.tab3_edit_requested = False
if 'tab3_edit_confirmed' not in st.session_state:
    st.session_state.tab3_edit_confirmed = False

if 'tab1_grid_key' not in st.session_state:
    st.session_state.tab1_grid_key = 0
if 'tab3_grid_key' not in st.session_state:
    st.session_state.tab3_grid_key = 0
if 'tab4_grid_key' not in st.session_state:
    st.session_state.tab4_grid_key = 0
if 'add_form_key' not in st.session_state:
    st.session_state.add_form_key = 0
if 'dev_add_form_key' not in st.session_state:
    st.session_state.dev_add_form_key = 0
if 'dev_add_preview' not in st.session_state:
    st.session_state.dev_add_preview = None
if 'dev_plant_options' not in st.session_state:
    st.session_state.dev_plant_options = []
if 'dev_case_options' not in st.session_state:
    st.session_state.dev_case_options = []
if 'dev_option_manager_key' not in st.session_state:
    st.session_state.dev_option_manager_key = 0
if 'dev_case_checklists' not in st.session_state:
    st.session_state.dev_case_checklists = {}
if 'dev_loaded_case' not in st.session_state:
    st.session_state.dev_loaded_case = None
if 'dev_checklist_key' not in st.session_state:
    st.session_state.dev_checklist_key = 0
if 'dev_test_records' not in st.session_state:
    st.session_state.dev_test_records = []
if 'dev_results_grid_key' not in st.session_state:
    st.session_state.dev_results_grid_key = 0
if 'dev_pending_preview' not in st.session_state:
    st.session_state.dev_pending_preview = None
if 'dev_reason_dialog_key' not in st.session_state:
    st.session_state.dev_reason_dialog_key = 0
if 'dev_flash_message' not in st.session_state:
    st.session_state.dev_flash_message = ""
if 'dev_pending_previous_record' not in st.session_state:
    st.session_state.dev_pending_previous_record = None
if 'dev_identity_draft' not in st.session_state:
    st.session_state.dev_identity_draft = None
if 'dev_previous_prefill' not in st.session_state:
    st.session_state.dev_previous_prefill = None
if 'dev_manual_only_initialized' not in st.session_state:
    st.session_state.dev_plant_options = []
    st.session_state.dev_case_options = []
    st.session_state.dev_case_checklists = {}
    st.session_state.dev_loaded_case = None
    st.session_state.dev_identity_draft = None
    st.session_state.dev_previous_prefill = None
    st.session_state.dev_pending_previous_record = None
    st.session_state.dev_add_preview = None
    st.session_state.dev_test_records = []
    st.session_state.dev_manual_only_initialized = True

# ==================== Google API 連線與快取優化 ====================
@st.cache_resource(ttl=3600)  # 快取 1 小時，避免頻繁呼叫 Google API 觸發限制
def init_google_sheets():
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            gc = gspread.service_account_from_dict(creds_dict)
        else:
            gc = gspread.service_account(filename='credentials.json')
        
        # 將 open 與 worksheet 的動作全部包在快取函數內，只執行一次
        sheet = gc.open("control table")
        ws_1 = sheet.worksheet("裝機人員進廠時間")
        ws_2 = sheet.worksheet("帳號管理")
        ws_3 = sheet.worksheet("修改紀錄")
        return sheet, ws_1, ws_2, ws_3
        
    except gspread.exceptions.WorksheetNotFound as e:
        st.error(f"⚠️ 系統錯誤：找不到必要的分頁 ({e})。請確認已建立『裝機人員進廠時間』、『帳號管理』與『修改紀錄』。")
        st.stop()
    except Exception as e:
        st.error(f"⚠️ Google API 連線異常，請稍後再試。詳細錯誤：{e}")
        st.stop()

# 取得快取後的試算表物件
sh, worksheet, ws_accounts, ws_log = init_google_sheets()
DEV_WORKSHEET_NAME = "開發區測試資料"


def sync_dev_data_to_google():
    """將開發區目前的手動輸入內容，以快照方式同步至獨立工作表。"""
    headers = [
        "同步批次",
        "紀錄類型",
        "操作時間",
        "裝機日期",
        "操作者",
        "廠別",
        "案件",
        "機台名稱",
        "項目確認",
        "安裝人員",
        "狀態",
        "未完成或缺貨原因",
        "Remark",
    ]
    sync_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    operator = f"{st.session_state.user_name} ({st.session_state.user_role})"
    rows = []

    for plant_name in st.session_state.dev_plant_options:
        rows.append([
            sync_time, "廠別選項", sync_time, "", operator,
            plant_name, "", "", "", "", "", "", "",
        ])

    for case_name in st.session_state.dev_case_options:
        checklist_items = st.session_state.dev_case_checklists.get(case_name, [])
        rows.append([
            sync_time, "案件設定", sync_time, "", operator,
            "", case_name, "", "\n".join(checklist_items), "", "", "", "",
        ])

    for record in st.session_state.dev_test_records:
        rows.append([
            sync_time,
            "裝機測試結果",
            record.get("建立時間", sync_time),
            record.get("日期", ""),
            operator,
            record.get("廠別", ""),
            record.get("案件", ""),
            record.get("機台名稱", ""),
            record.get("項目確認", ""),
            record.get("安裝人員", ""),
            record.get("狀態", ""),
            record.get("未完成或缺貨原因", ""),
            record.get("Remark", ""),
        ])

    if not rows:
        return 0

    try:
        dev_worksheet = sh.worksheet(DEV_WORKSHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        dev_worksheet = sh.add_worksheet(
            title=DEV_WORKSHEET_NAME,
            rows=1000,
            cols=len(headers),
        )

    existing_headers = dev_worksheet.row_values(1)
    if existing_headers and "裝機日期" not in existing_headers:
        dev_worksheet.insert_cols([["裝機日期"]], col=4)
        existing_headers = dev_worksheet.row_values(1)
    if existing_headers and "未完成或缺貨原因" not in existing_headers:
        reason_col = headers.index("未完成或缺貨原因") + 1
        dev_worksheet.insert_cols([["未完成或缺貨原因"]], col=reason_col)
        existing_headers = dev_worksheet.row_values(1)
    if not existing_headers:
        dev_worksheet.append_row(headers)
    elif existing_headers != headers:
        raise ValueError(f"「{DEV_WORKSHEET_NAME}」的欄位格式不符，請先確認標題列。")
    dev_worksheet.append_rows(rows, value_input_option="USER_ENTERED")
    return len(rows)


def load_latest_dev_data_from_google():
    """從開發區工作表載入最後一次同步快照。"""
    try:
        dev_worksheet = sh.worksheet(DEV_WORKSHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        return None

    saved_rows = dev_worksheet.get_all_records()
    if not saved_rows:
        return None

    latest_batch = next(
        (
            str(row.get("同步批次", "")).strip()
            for row in reversed(saved_rows)
            if str(row.get("同步批次", "")).strip()
        ),
        "",
    )
    if not latest_batch:
        return None

    batch_rows = [
        row for row in saved_rows
        if str(row.get("同步批次", "")).strip() == latest_batch
    ]

    plant_options = sorted({
        str(row.get("廠別", "")).strip()
        for row in batch_rows
        if row.get("紀錄類型") == "廠別選項" and str(row.get("廠別", "")).strip()
    })

    case_options = []
    case_checklists = {}
    test_records = []
    for row in batch_rows:
        record_type = str(row.get("紀錄類型", "")).strip()
        if record_type == "案件設定":
            case_name = str(row.get("案件", "")).strip()
            if case_name:
                case_options.append(case_name)
                case_checklists[case_name] = [
                    item.strip()
                    for item in str(row.get("項目確認", "")).splitlines()
                    if item.strip()
                ]
        elif record_type == "裝機測試結果":
            test_records.append({
                "建立時間": str(row.get("操作時間", "")).strip(),
                "日期": str(row.get("裝機日期", "")).strip(),
                "廠別": str(row.get("廠別", "")).strip(),
                "案件": str(row.get("案件", "")).strip(),
                "機台名稱": str(row.get("機台名稱", "")).strip(),
                "項目確認": str(row.get("項目確認", "")).strip(),
                "安裝人員": str(row.get("安裝人員", "")).strip(),
                "狀態": str(row.get("狀態", "")).strip(),
                "未完成或缺貨原因": str(row.get("未完成或缺貨原因", "")).strip(),
                "Remark": str(row.get("Remark", "")).strip(),
            })

    st.session_state.dev_plant_options = plant_options
    st.session_state.dev_case_options = sorted(set(case_options))
    st.session_state.dev_case_checklists = case_checklists
    st.session_state.dev_test_records = test_records
    st.session_state.dev_loaded_case = None
    st.session_state.dev_identity_draft = None
    st.session_state.dev_previous_prefill = None
    st.session_state.dev_pending_previous_record = None
    st.session_state.dev_add_preview = None
    st.session_state.dev_pending_preview = None
    st.session_state.dev_option_manager_key += 1
    st.session_state.dev_add_form_key += 1
    st.session_state.dev_checklist_key += 1

    return {
        "批次": latest_batch,
        "廠別": len(plant_options),
        "案件": len(case_options),
        "測試結果": len(test_records),
    }

# ===================================================================
# ==================== 2. 登入系統與權限驗證 ====================
if not st.session_state.logged_in:
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        # 🚀 透過自訂 CSS 容器包裹 st.image，利用 margin-left 完美對齊下方文字中心
        st.markdown(
            """
            <style>
            .centered-logo {
                display: flex;
                justify-content: center;
                margin-left: 20px; /* 根據視覺微調向右偏移 */
                margin-bottom: -10px;
            }
            </style>
            <div class="centered-logo">
            """,
            unsafe_allow_html=True
        )
        
        try:
            st.image("logo.jpg", width=300)
        except:
            st.warning("⚠️ 找不到 logo.jpg 圖片檔案。")
            
        st.markdown("</div>", unsafe_allow_html=True)
            
        st.markdown("<h2 style='text-align: center; margin-top: 5px;'>鴻伍裝機日報系統</h2>", unsafe_allow_html=True)
        st.markdown("### 🔐 系統登入")
        
        with st.form("login_form"):
            user_id = st.text_input("帳號 (ID)")
            user_pwd = st.text_input("密碼", type="password")
            submitted = st.form_submit_button("登入系統", use_container_width=True)
            
            if submitted:
                if not user_id or not user_pwd:
                    st.warning("請輸入帳號與密碼。")
                else:
                    accounts_data = ws_accounts.get_all_records()
                    df_acc = pd.DataFrame(accounts_data)
                    
                    if not df_acc.empty and '帳號' in df_acc.columns and '密碼' in df_acc.columns:
                        df_acc['帳號'] = df_acc['帳號'].astype(str)
                        df_acc['密碼'] = df_acc['密碼'].astype(str)
                        
                        match = df_acc[(df_acc['帳號'] == user_id) & (df_acc['密碼'] == user_pwd)]
                        if not match.empty:
                            st.session_state.logged_in = True
                            st.session_state.user_name = match.iloc[0].get('姓名', '未知使用者')
                            st.session_state.user_role = match.iloc[0].get('權限', '未設定權限')
                            st.rerun()
                        else:
                            st.error("帳號或密碼錯誤，請重新輸入。")
                    else:
                        st.error("『帳號管理』分頁中缺少必要欄位或無資料。")
        
        # 🚀 訪客快速登入按鈕
        if st.button("🚀 訪客快速登入 (公用權限)", use_container_width=True, type="secondary"):
            st.session_state.logged_in = True
            st.session_state.user_name = "訪客"
            st.session_state.user_role = "公用"
            st.rerun()

    st.stop() # 阻擋未登入者往下執行

# 權限定義
can_edit = st.session_state.user_role == "管理者"
can_add = st.session_state.user_role in ["管理者", "工程師", "業務", "RD"]

# 側邊欄狀態
with st.sidebar:
    try:
        st.image("logo.jpg", width=200)
    except:
        pass
    st.divider()
    st.header("👤 帳號資訊")
    st.markdown(f"**姓名：** {st.session_state.user_name}")
    st.markdown(f"**權限：** {st.session_state.user_role}")
    st.divider()
    if st.button("🚪 登出系統", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.user_name = ""
        st.session_state.user_role = ""
        st.rerun()

st.title("📊 鴻伍裝機日報系統 (Web 雲端版)")

# ==================== 3. 彈出視窗功能 (Dialog) ====================
def parse_checklist_summary(summary):
    items = [item.strip() for item in str(summary or "").split("、") if item.strip()]
    completed = [item.removeprefix("✅").strip() for item in items if item.startswith("✅")]
    incomplete = [item.removeprefix("❌").strip() for item in items if item.startswith("❌")]
    return completed, incomplete


def parse_checklist_definition(checklist_lines):
    """解析 [主分類]、[[子分類]] 與其確認項目。"""
    groups = []
    current_group = None
    current_subgroup = None

    def ensure_group():
        nonlocal current_group
        if current_group is None:
            current_group = {"name": "", "items": [], "subgroups": []}
            groups.append(current_group)
        return current_group

    for raw_line in checklist_lines:
        line = str(raw_line).strip()
        if not line:
            continue
        if line.startswith("[[") and line.endswith("]]") and len(line) > 4:
            parent_group = ensure_group()
            current_subgroup = {"name": line[2:-2].strip(), "items": []}
            parent_group["subgroups"].append(current_subgroup)
        elif line.startswith("[") and line.endswith("]") and len(line) > 2:
            current_group = {"name": line[1:-1].strip(), "items": [], "subgroups": []}
            groups.append(current_group)
            current_subgroup = None
        else:
            parent_group = ensure_group()
            if current_subgroup is not None:
                current_subgroup["items"].append(line)
            else:
                parent_group["items"].append(line)

    return [
        group for group in groups
        if group["items"] or any(subgroup["items"] for subgroup in group["subgroups"])
    ]


def format_checklist_progress(summary):
    completed, incomplete = parse_checklist_summary(summary)
    total = len(completed) + len(incomplete)
    return f"已施工 {len(completed)}/{total}" if total else "未設定"


def group_checklist_paths(items):
    """將 A / B / C 格式整理為主分類、子分類及項目。"""
    grouped = {}
    for raw_item in items:
        path_parts = [part.strip() for part in str(raw_item).split(" / ") if part.strip()]
        if len(path_parts) >= 3:
            category_name = path_parts[0]
            subcategory_name = path_parts[1]
            item_name = " / ".join(path_parts[2:])
        elif len(path_parts) == 2:
            category_name = path_parts[0]
            subcategory_name = ""
            item_name = path_parts[1]
        else:
            category_name = "其他項目"
            subcategory_name = ""
            item_name = path_parts[0] if path_parts else str(raw_item)

        grouped.setdefault(category_name, {}).setdefault(subcategory_name, []).append(item_name)
    return grouped


def render_grouped_checklist(items, icon):
    grouped_items = group_checklist_paths(items)
    for category_name, subcategories in grouped_items.items():
        category_count = sum(len(group_items) for group_items in subcategories.values())
        st.markdown(f"**📂 {category_name}（{category_count}）**")
        for subcategory_name, group_items in subcategories.items():
            compact_items = "　".join(f"{icon} {item_name}" for item_name in group_items)
            if subcategory_name:
                st.markdown(f"**↳ {subcategory_name}：** {compact_items}")
            elif category_name == "其他項目":
                for item_name in group_items:
                    st.markdown(f"- {icon} {item_name}")
            else:
                st.markdown(compact_items)


def get_checklist_export_labels(checklist_lines):
    """依主分類、子分類順序產生 Excel 欄位名稱。"""
    labels = []
    for category_group in parse_checklist_definition(checklist_lines):
        category_name = category_group["name"]
        for item_name in category_group["items"]:
            labels.append(f"{category_name} / {item_name}" if category_name else item_name)
        for subgroup in category_group["subgroups"]:
            for item_name in subgroup["items"]:
                path_parts = [category_name, subgroup["name"], item_name]
                labels.append(" / ".join(part for part in path_parts if part))
    return labels


def render_export_checklist_selector(checklist_lines, key_prefix):
    """依輸入時的主分類／子分類顯示 Excel 匯出項目選擇器。"""
    selected_labels = []
    item_index = 0
    for category_group in parse_checklist_definition(checklist_lines):
        category_name = category_group["name"]

        def render_group_items():
            nonlocal item_index
            for item_name in category_group["items"]:
                item_label = f"{category_name} / {item_name}" if category_name else item_name
                if st.checkbox(
                    item_name,
                    value=False,
                    key=f"{key_prefix}_{item_index}",
                ):
                    selected_labels.append(item_label)
                item_index += 1

            for subgroup in category_group["subgroups"]:
                st.markdown(f"**↳ {subgroup['name']}**")
                for item_name in subgroup["items"]:
                    path_parts = [category_name, subgroup["name"], item_name]
                    item_label = " / ".join(part for part in path_parts if part)
                    if st.checkbox(
                        item_name,
                        value=False,
                        key=f"{key_prefix}_{item_index}",
                    ):
                        selected_labels.append(item_label)
                    item_index += 1

        if category_name:
            with st.expander(f"📂 {category_name}", expanded=False):
                render_group_items()
        else:
            render_group_items()

    return selected_labels


def build_dev_excel_export(machine_records, selected_checklist_items):
    """依範例格式建立多機台確認項目 Excel。"""
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    sheet = workbook.add_worksheet("裝機確認")

    common_format = {
        "font_name": "新細明體",
        "font_size": 11,
        "border": 1,
        "align": "center",
        "valign": "vcenter",
    }
    header_format = workbook.add_format({**common_format, "text_wrap": True})
    cell_format = workbook.add_format(common_format)
    date_format = workbook.add_format({**common_format, "num_format": "m/d/yy"})
    remark_format = workbook.add_format({**common_format, "text_wrap": True})
    completed_format = workbook.add_format({**common_format, "bg_color": "#C6EFCE"})
    pending_format = workbook.add_format({**common_format, "bg_color": "#FFEB9C"})

    headers = ["日期", "廠別", "案件名稱", "機台名", *selected_checklist_items, "備註"]
    sheet.write_row(0, 0, headers, header_format)

    for row_index, record in enumerate(machine_records, start=1):
        date_value = pd.to_datetime(record.get("日期", ""), errors="coerce")
        if pd.notna(date_value):
            sheet.write_datetime(row_index, 0, date_value.to_pydatetime(), date_format)
        else:
            sheet.write(row_index, 0, str(record.get("日期", "")), cell_format)

        sheet.write(row_index, 1, record.get("廠別", ""), cell_format)
        sheet.write(row_index, 2, record.get("案件", ""), cell_format)
        sheet.write(row_index, 3, record.get("機台名稱", ""), cell_format)

        completed_items, _ = parse_checklist_summary(record.get("項目確認", ""))
        completed_set = set(completed_items)
        legacy_completed_names = {
            item.strip() for item in completed_items if " / " not in item
        }
        for item_offset, item_label in enumerate(selected_checklist_items, start=4):
            item_leaf_name = item_label.split(" / ")[-1].strip()
            is_completed = item_label in completed_set or item_leaf_name in legacy_completed_names
            status_format = completed_format if is_completed else pending_format
            sheet.write_blank(row_index, item_offset, None, status_format)

        sheet.write(
            row_index,
            4 + len(selected_checklist_items),
            record.get("Remark", ""),
            remark_format,
        )

    last_column = len(headers) - 1
    sheet.autofilter(0, 0, max(len(machine_records), 1), last_column)
    sheet.freeze_panes(1, 4)
    sheet.set_row(0, 30)
    sheet.set_column(0, 0, 12)
    sheet.set_column(1, 1, 12)
    sheet.set_column(2, 3, 18)
    if selected_checklist_items:
        sheet.set_column(4, 3 + len(selected_checklist_items), 22)
    sheet.set_column(4 + len(selected_checklist_items), last_column, 24)

    workbook.close()
    output.seek(0)
    return output.getvalue()


def reset_dialog_selection():
    reset_key = st.session_state.get("active_dialog_reset_key")
    if reset_key and reset_key in st.session_state:
        st.session_state[reset_key] += 1


@st.dialog("📝 詳細資料檢視", on_dismiss=reset_dialog_selection)
def show_details_dialog(row_data, reset_key):
    st.session_state.active_dialog_reset_key = reset_key
    st.markdown(f"### 🏭 廠別：{row_data.get('廠別', '')}")
    st.markdown(f"### 💻 機台名稱：{row_data.get('機台名稱', '')}")
    st.markdown(f"**📅 日期：** {row_data.get('日期', '')}")
    st.markdown(f"**📂 案件：** {row_data.get('案件', '')}")
    st.markdown(f"**👷 安裝人員：** {row_data.get('安裝人員', '')}")
    st.markdown(f"**🚦 狀態：** {row_data.get('狀態', '')}")
    if row_data.get('建立時間', ''):
        st.markdown(f"**🕒 建立時間：** {row_data.get('建立時間', '')}")
    if row_data.get('項目確認', ''):
        st.markdown("### ✅ 項目確認：")
        completed_items, incomplete_items = parse_checklist_summary(row_data.get('項目確認', ''))
        total_items = len(completed_items) + len(incomplete_items)
        if total_items:
            st.caption(
                f"已施工 {len(completed_items)} 項｜待施工 {len(incomplete_items)} 項"
            )
            if incomplete_items:
                with st.expander(f"❌ 待施工項目（{len(incomplete_items)}）", expanded=False):
                    render_grouped_checklist(incomplete_items, "❌")
            if completed_items:
                with st.expander(f"✅ 已施工項目（{len(completed_items)}）", expanded=False):
                    render_grouped_checklist(completed_items, "✅")
        else:
            st.write(row_data.get('項目確認', ''))
    if row_data.get('未完成或缺貨原因', ''):
        st.markdown("### ⚠️ 未完成原因：")
        st.write(row_data.get('未完成或缺貨原因', ''))
    st.markdown("---")
    st.markdown("### 📝 Remark (備忘內容)：")
    
    remark_text = str(row_data.get('Remark', '')).replace('\n', '<br>')
    if not remark_text.strip():
        remark_text = "（無備忘內容）"
        
    st.markdown(
        f"<div style='font-size: 1.3em; line-height: 1.6; background-color: #f0f2f6; padding: 15px; border-radius: 8px; color: #333; margin-bottom: 20px;'>{remark_text}</div>",
        unsafe_allow_html=True
    )
    
    if st.button("❌ 關閉視窗並取消選取", type="primary", use_container_width=True):
        st.session_state[reset_key] += 1  
        st.rerun()


def dismiss_dev_reason_dialog():
    st.session_state.dev_pending_preview = None


def prepare_dev_preview(record):
    """依狀態決定直接預覽，或先要求輸入原因。"""
    if record.get("狀態") == "未完成":
        st.session_state.dev_add_preview = None
        st.session_state.dev_pending_preview = record
    else:
        record["未完成或缺貨原因"] = ""
        st.session_state.dev_pending_preview = None
        st.session_state.dev_add_preview = record


@st.dialog("⚠️ 記錄原因", on_dismiss=dismiss_dev_reason_dialog)
def show_dev_reason_dialog():
    pending_record = st.session_state.dev_pending_preview
    if not pending_record:
        st.rerun()

    status = pending_record.get("狀態", "")
    st.warning(f"目前狀態為「{status}」，請記錄原因後再繼續。")
    reason = st.text_area(
        "原因 *",
        placeholder="請輸入未完成的原因",
        height=140,
        key=f"dev_reason_{st.session_state.dev_reason_dialog_key}",
    )

    confirm_col, cancel_col = st.columns(2)
    with confirm_col:
        if st.button("確認並產生預覽", type="primary", use_container_width=True):
            cleaned_reason = reason.strip()
            if not cleaned_reason:
                st.error("請輸入原因。")
            else:
                pending_record["未完成或缺貨原因"] = cleaned_reason
                st.session_state.dev_add_preview = pending_record
                st.session_state.dev_pending_preview = None
                st.session_state.dev_reason_dialog_key += 1
                st.rerun()
    with cancel_col:
        if st.button("取消", use_container_width=True):
            st.session_state.dev_pending_preview = None
            st.session_state.dev_reason_dialog_key += 1
            st.rerun()


def dismiss_previous_record_dialog():
    st.session_state.dev_pending_previous_record = None


@st.dialog("🔎 發現未完成紀錄", on_dismiss=dismiss_previous_record_dialog)
def show_previous_record_dialog():
    pending_data = st.session_state.dev_pending_previous_record
    if not pending_data:
        st.rerun()

    identity_data = dict(pending_data["基本資料"])
    previous_record = pending_data["上次資料"]
    st.warning("相同廠別、案件與機台名稱有未完成紀錄，是否帶入上次資料？")
    st.markdown(f"**廠別：** {identity_data.get('廠別', '')}")
    st.markdown(f"**案件：** {identity_data.get('案件', '')}")
    st.markdown(f"**機台名稱：** {identity_data.get('機台名稱', '')}")
    st.markdown(f"**上次狀態：** {previous_record.get('狀態', '')}")
    st.markdown("**上次確認項目：**")
    st.write(previous_record.get("項目確認", ""))
    st.markdown("**上次備註：**")
    st.write(previous_record.get("Remark", "") or "（無備註）")

    bring_col, keep_col = st.columns(2)
    with bring_col:
        if st.button("帶入上次資料", type="primary", use_container_width=True):
            st.session_state.dev_identity_draft = identity_data
            st.session_state.dev_loaded_case = identity_data["案件"]
            st.session_state.dev_previous_prefill = {
                "項目確認": previous_record.get("項目確認", ""),
                "狀態": previous_record.get("狀態", "未完成"),
                "Remark": previous_record.get("Remark", ""),
            }
            st.session_state.dev_pending_previous_record = None
            st.session_state.dev_add_form_key += 1
            st.session_state.dev_checklist_key += 1
            st.rerun()
    with keep_col:
        if st.button("不帶入，繼續輸入", use_container_width=True):
            st.session_state.dev_identity_draft = identity_data
            st.session_state.dev_loaded_case = identity_data["案件"]
            st.session_state.dev_previous_prefill = None
            st.session_state.dev_pending_previous_record = None
            st.session_state.dev_add_form_key += 1
            st.session_state.dev_checklist_key += 1
            st.rerun()

# ==================== 4. 建立功能分頁 ====================
installers_list = ["鍾博宇", "黃政欽", "張智偉", "林嬴燦", "吳建華", "何乙霆"]

tab_labels = ["🌅 晨會當日動態", "📝 新增裝機紀錄", "🔍 歷史搜尋與修改", "📌 待追蹤清單 (更新狀態)"]
if can_edit:
    tab_labels.append("🧪 開發測試區")

tabs = st.tabs(tab_labels)
tab1, tab2, tab3, tab4 = tabs[:4]
tab_dev = tabs[4] if can_edit else None

# ==================== 分頁 1：晨會當日動態 ====================
with tab1:
    st.subheader("查詢晨會動態")
    target_date = st.date_input("選擇日期", datetime.now(), key="morning_date")
    
    if st.button("🔍 查詢當日動態", key="btn_morning"):
        with st.spinner('讀取雲端資料中...'):
            data = worksheet.get_all_records()
            df = pd.DataFrame(data)
            
            if not df.empty:
                df['日期_temp'] = pd.to_datetime(df['日期'], format='mixed', errors='coerce').dt.date
                filtered_df = df[df['日期_temp'] == target_date].copy()
                
                if not filtered_df.empty:
                    filtered_df['日期'] = filtered_df['日期'].astype(str)
                    st.session_state.tab1_filtered_df = filtered_df
                    st.session_state.tab1_search_active = True
                else:
                    st.session_state.tab1_filtered_df = pd.DataFrame()
                    st.session_state.tab1_search_active = True
            else:
                st.warning("試算表中尚無任何資料。")
                st.session_state.tab1_search_active = False

    if st.session_state.tab1_search_active:
        filtered_df = st.session_state.tab1_filtered_df
        if not filtered_df.empty:
            st.success(f"找到 {len(filtered_df)} 筆紀錄 (🖱️ 提示：點擊任意列可彈出詳細資訊)")
            display_cols = ["廠別", "案件", "機台名稱", "安裝人員", "狀態", "Remark"]
            display_cols = [col for col in display_cols if col in filtered_df.columns]
            
            event = st.dataframe(
                filtered_df[display_cols], hide_index=True, use_container_width=True, 
                on_select="rerun", selection_mode="single-row", key=f"tab1_grid_{st.session_state.tab1_grid_key}"
            )
            if event.selection.rows:
                selected_idx = event.selection.rows[0]
                show_details_dialog(filtered_df.iloc[selected_idx], 'tab1_grid_key')
        else:
            st.info(f"📅 該日尚無裝機紀錄。")

# ==================== 分頁 2：新增裝機紀錄 ====================
with tab2:
    st.subheader("填寫裝機資訊")
    
    if st.session_state.user_role == "公用":
        st.warning("⚠️ 目前為「公用」身分，僅供檢視與查詢，無法新增裝機紀錄。")
    else:
        k_suffix = st.session_state.add_form_key
        col1, col2 = st.columns(2)
        
        with col1:
            input_date = st.date_input("裝機日期", datetime.now(), key=f"add_date_{k_suffix}")
            plant = st.text_input("廠別:", key=f"add_plant_{k_suffix}")
            case = st.text_input("案件:", key=f"add_case_{k_suffix}")
            machine = st.text_input("機台名稱:", key=f"add_machine_{k_suffix}")
            
        with col2:
            status = st.selectbox("狀態:", ["未完成", "缺料", "已完成"], key=f"add_status_{k_suffix}")
            installers = st.multiselect("安裝人員 (可複選):", installers_list, key=f"add_installers_{k_suffix}")
            remark = st.text_area("Remark (備忘):", height=130, key=f"add_remark_{k_suffix}")

        b_col1, b_col2 = st.columns(2)
        with b_col1:
            btn_submit = st.button("💾 新增紀錄", type="primary", key="btn_add")
        with b_col2:
            if st.button("🗑️ 清空欄位", key="btn_clear_form"):
                st.session_state.add_form_key += 1
                st.rerun()

        if btn_submit:
            if not plant or not machine:
                st.error("「廠別」與「機台名稱」為必填欄位！")
            else:
                with st.spinner('寫入雲端中...'):
                    installer_str = "\n".join(installers)
                    date_str = input_date.strftime("%Y-%m-%d")
                    
                    headers = worksheet.row_values(1)
                    new_row = [""] * len(headers)
                    def fill_col(col_name, val):
                        if col_name in headers:
                            new_row[headers.index(col_name)] = val

                    fill_col("日期", date_str)
                    fill_col("安裝人員", installer_str)
                    fill_col("廠別", plant)
                    fill_col("案件", case)
                    fill_col("機台名稱", machine)
                    fill_col("狀態", status)
                    fill_col("Remark", remark)
                    
                    worksheet.append_row(new_row)
                    
                    log_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    ws_log.append_row(
                        [log_time, f"{st.session_state.user_name} ({st.session_state.user_role})", f"新增機台: {machine} (廠別:{plant})", "", "建立新紀錄"],
                        table_range="A:E",
                    )
                    
                    st.success(f"✅ 成功將機台【{machine}】新增至雲端！")

# ==================== 分頁 3：歷史紀錄搜尋與修改 ====================
with tab3:
    st.subheader("🔍 進階條件篩選與修改")
    data = worksheet.get_all_records()
    df_search = pd.DataFrame(data)
    
    if not df_search.empty:
        df_search = df_search.fillna("")
        df_search['Sheet_Row'] = df_search.index + 2
        
        unique_plants = ["(全部)"] + sorted(list(set([str(x).strip() for x in df_search['廠別'] if str(x).strip()])))
        unique_installers = ["(全部)"] + installers_list
        
        st.markdown("##### 1. 設定搜尋條件 (設定完畢後請點擊下方搜尋按鈕)")
        col_s1, col_s2, col_s3, col_s4, col_s5 = st.columns(5)
        
        with col_s1: date_range = st.date_input("選擇日期區間:", [])
        with col_s2: search_plant = st.selectbox("廠別:", unique_plants)
            
        if search_plant != "(全部)":
            target_plant_df = df_search[df_search['廠別'].astype(str).str.strip() == search_plant]
            unique_machines = ["(全部)"] + sorted(list(set([str(x).strip() for x in target_plant_df['機台名稱'] if str(x).strip()])))
            is_machine_disabled = False
        else:
            unique_machines = ["(請先選擇廠別)"]
            is_machine_disabled, target_plant_df = True, pd.DataFrame()
            
        with col_s3: search_machine = st.selectbox("機台名稱 (廠別確定後解鎖):", unique_machines, disabled=is_machine_disabled)
            
        if search_plant != "(全部)":
            if search_machine not in ["(全部)", "(請先選擇廠別)"]:
                target_machine_df = target_plant_df[target_plant_df['機台名稱'].astype(str).str.strip() == search_machine]
                unique_cases = ["(全部)"] + sorted(list(set([str(x).strip() for x in target_machine_df['案件'] if str(x).strip()])))
            else:
                unique_cases = ["(全部)"] + sorted(list(set([str(x).strip() for x in target_plant_df['案件'] if str(x).strip()])))
            is_case_disabled = False
        else:
            unique_cases, is_case_disabled = ["(請先選擇廠別)"], True
            
        with col_s4: search_case = st.selectbox("案件 (廠別確定後解鎖):", unique_cases, disabled=is_case_disabled)
        with col_s5: search_installer = st.selectbox("安裝人員:", unique_installers)
            
        if st.button("🔍 開始搜尋", type="primary", key="btn_execute_search"):
            with st.spinner("搜尋中..."):
                filtered_df = df_search.copy()
                if len(date_range) == 2:
                    filtered_df['日期_temp'] = pd.to_datetime(filtered_df['日期'], format='mixed', errors='coerce').dt.date
                    filtered_df = filtered_df[(filtered_df['日期_temp'] >= date_range[0]) & (filtered_df['日期_temp'] <= date_range[1])].drop(columns=['日期_temp'])
                elif len(date_range) == 1:
                    filtered_df['日期_temp'] = pd.to_datetime(filtered_df['日期'], format='mixed', errors='coerce').dt.date
                    filtered_df = filtered_df[filtered_df['日期_temp'] == date_range[0]].drop(columns=['日期_temp'])
                    
                if search_plant != "(全部)": filtered_df = filtered_df[filtered_df['廠別'].astype(str).str.strip() == search_plant]
                if search_machine not in ["(全部)", "(請先選擇廠別)"]: filtered_df = filtered_df[filtered_df['機台名稱'].astype(str).str.strip() == search_machine]
                if search_case not in ["(全部)", "(請先選擇廠別)"]: filtered_df = filtered_df[filtered_df['案件'].astype(str).str.strip() == search_case]
                if search_installer != "(全部)": filtered_df = filtered_df[filtered_df['安裝人員'].astype(str).str.contains(search_installer)]
                
                st.session_state.tab3_filtered_df = filtered_df
                st.session_state.tab3_search_active = True
                st.session_state.tab3_edit_requested = False
                st.session_state.tab3_edit_confirmed = False
                st.rerun()
                
        st.divider()
        
        if st.session_state.tab3_search_active:
            filtered_df = st.session_state.tab3_filtered_df
            st.markdown(f"##### 2. 搜尋結果 (共計 <span style='color:red;'>{len(filtered_df)}</span> 筆)", unsafe_allow_html=True)
            
            if not filtered_df.empty:
                view_cols = ["日期", "廠別", "案件", "機台名稱", "安裝人員", "狀態", "Remark"]
                view_cols = [col for col in view_cols if col in filtered_df.columns]
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    filtered_df[view_cols].to_excel(writer, sheet_name='裝機搜尋結果', index=False)
                    workbook, worksheet_excel = writer.book, writer.sheets['裝機搜尋結果']
                    wrap_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})
                    default_format = workbook.add_format({'valign': 'top'})
                    for idx, col_name in enumerate(view_cols):
                        width = 45 if col_name == 'Remark' else (20 if col_name == '安裝人員' else 18)
                        worksheet_excel.set_column(idx, idx, width, wrap_format if col_name in ['Remark', '安裝人員'] else default_format)
                buffer.seek(0)
                
                st.download_button(label="📥 匯出搜尋結果為 Excel", data=buffer, file_name=f"鴻伍裝機搜尋結果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
                if not st.session_state.tab3_edit_confirmed:
                    event = st.dataframe(filtered_df[view_cols], hide_index=True, use_container_width=True, on_select="rerun", selection_mode="single-row", key=f"tab3_grid_{st.session_state.tab3_grid_key}")
                    if event.selection.rows:
                        show_details_dialog(filtered_df.iloc[event.selection.rows[0]], 'tab3_grid_key')
                    
                    if not st.session_state.tab3_edit_requested:
                        if can_edit:
                            if st.button("✏️ 開啟修改模式"):
                                st.session_state.tab3_edit_requested = True
                                st.rerun()
                        else:
                            st.info(f"💡 您的權限 ({st.session_state.user_role}) 僅供查詢與檢視，修改功能僅限「管理者」。")
                    else:
                        st.warning("⚠️ 即將進入修改，請確認")
                        c1, c2, c3 = st.columns([1, 1, 4])
                        with c1:
                            if st.button("✅ 確認修改", type="primary"):
                                st.session_state.tab3_edit_confirmed = True
                                st.session_state.tab3_edit_requested = False
                                st.rerun()
                        with c2:
                            if st.button("❌ 取消"):
                                st.session_state.tab3_edit_requested = False
                                st.rerun()
                else:
                    st.info("✏️ 編輯模式已開啟，請直接在下方表格修改內容。")
                    edited_df = st.data_editor(filtered_df[view_cols], hide_index=True, use_container_width=True, key="search_editor")
                    
                    if st.button("💾 儲存表格上的所有修改", type="primary"):
                        with st.spinner("正在批次同步更新並寫入日誌..."):
                            changed_cells = []
                            log_entries = []
                            headers = worksheet.row_values(1)
                            
                            for i in range(len(edited_df)):
                                orig_row = filtered_df[view_cols].iloc[i]
                                new_row = edited_df.iloc[i]
                                sheet_row_idx = filtered_df.iloc[i]['Sheet_Row']
                                machine_name = new_row['機台名稱']
                                
                                for col in view_cols:
                                    if str(orig_row[col]).strip() != str(new_row[col]).strip():
                                        if col in headers:
                                            col_idx = headers.index(col) + 1
                                            changed_cells.append(gspread.Cell(int(sheet_row_idx), col_idx, str(new_row[col])))
                                            log_entries.append([
                                                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                                f"{st.session_state.user_name} ({st.session_state.user_role})",
                                                f"{machine_name} (Row {sheet_row_idx}) - {col}",
                                                str(orig_row[col]),
                                                str(new_row[col])
                                            ])
                                            
                            if changed_cells:
                                worksheet.update_cells(changed_cells)
                                ws_log.append_rows(log_entries, table_range="A:E")
                                st.success(f"✅ 成功更新資料，並已記錄 {len(log_entries)} 筆修改日誌！")
                                st.session_state.tab3_search_active = False
                                st.session_state.tab3_edit_confirmed = False
                                st.rerun()
                            else:
                                st.info("沒有偵測到任何修改內容。")
            else:
                st.warning("⚠️ 找不到符合您設定條件的紀錄。")
    else:
        st.info("試算表中尚無資料。")

# ==================== 分頁 4：待追蹤清單與狀態更新 ====================
with tab4:
    st.subheader("📌 待追蹤機台與狀態更新")
    data = worksheet.get_all_records()
    df = pd.DataFrame(data)
    
    if not df.empty:
        df = df.fillna("")
        df['Sheet_Row'] = df.index + 2
        pending_df = df[(df['狀態'].astype(str).str.strip() != '已完成') & (df['狀態'].astype(str).str.strip() != '') & (df['機台名稱'].astype(str).str.strip() != '')]
        st.markdown(f"**目前待追蹤數量： <span style='color:red;'>{len(pending_df)}</span> 筆**", unsafe_allow_html=True)
        
        if not pending_df.empty:
            display_cols = ["日期", "廠別", "案件", "機台名稱", "安裝人員", "狀態", "Remark"]
            display_cols = [col for col in display_cols if col in pending_df.columns]
            
            event = st.dataframe(pending_df[display_cols], hide_index=True, use_container_width=True, on_select="rerun", selection_mode="single-row", key=f"tab4_grid_{st.session_state.tab4_grid_key}")
            if event.selection.rows:
                show_details_dialog(pending_df.iloc[event.selection.rows[0]], 'tab4_grid_key')
            
            st.divider()
            st.markdown("#### ✏️ 更新機台狀態")
            
            if not can_edit:
                st.info(f"💡 您的權限 ({st.session_state.user_role}) 僅供檢視，狀態更新功能僅限「管理者」操作。")
            else:
                options = pending_df['Sheet_Row'].tolist()
                upd_col1, upd_col2 = st.columns([2, 1])
                with upd_col1:
                    selected_item_row = st.selectbox(
                        "選擇要更新的機台：", options, 
                        format_func=lambda r: f"{pending_df.loc[pending_df['Sheet_Row'] == r, '日期'].values[0]} | {pending_df.loc[pending_df['Sheet_Row'] == r, '廠別'].values[0]} - {pending_df.loc[pending_df['Sheet_Row'] == r, '機台名稱'].values[0]} (目前: {pending_df.loc[pending_df['Sheet_Row'] == r, '狀態'].values[0]})"
                    )
                with upd_col2:
                    new_status = st.selectbox("修改為新狀態：", ["未完成", "缺料", "已完成"], index=2)
                    
                if st.button("送出狀態更新", type="primary", key="btn_update"):
                    with st.spinner("同步至雲端並記錄日誌中..."):
                        row_idx = int(selected_item_row)
                        headers = worksheet.row_values(1)
                        if "狀態" in headers and "Remark" in headers:
                            status_col_idx = headers.index("狀態") + 1
                            remark_col_idx = headers.index("Remark") + 1
                            
                            old_status = pending_df.loc[pending_df['Sheet_Row'] == row_idx, '狀態'].values[0]
                            target_machine = pending_df.loc[pending_df['Sheet_Row'] == row_idx, '機台名稱'].values[0]
                            
                            worksheet.update_cell(row_idx, status_col_idx, new_status)
                            
                            old_remark = worksheet.cell(row_idx, remark_col_idx).value or ""
                            today_str = datetime.now().strftime("%Y-%m-%d")
                            append_text = f"[{today_str} 更新狀態: {new_status}]"
                            new_remark = (str(old_remark).strip() + "\n" + append_text).strip()
                            worksheet.update_cell(row_idx, remark_col_idx, new_remark)
                            
                            log_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            ws_log.append_row(
                                [log_time, f"{st.session_state.user_name} ({st.session_state.user_role})", f"更新機台狀態: {target_machine}", old_status, new_status],
                                table_range="A:E",
                            )
                            
                            st.success(f"✅ 更新成功！該機台已標記為「{new_status}」。")
                            st.rerun()
                        else:
                            st.error("找不到「狀態」或「Remark」欄位。")
        else:
            st.success("🎉 所有機台皆已完工，沒有待追蹤項目。")

# ==================== 管理員專屬：開發測試區 ====================
if can_edit and tab_dev is not None:
    with tab_dev:
        st.subheader("🧪 開發測試區")
        st.warning("此區僅供管理員測試；平常異動保留在本次工作階段，只有手動同步時才會寫入開發區試算表。")
        if st.session_state.dev_flash_message:
            st.success(st.session_state.dev_flash_message)
            st.session_state.dev_flash_message = ""

        dev_form_tab, dev_options_tab, dev_results_tab, dev_excel_tab = st.tabs([
            "📝 新版新增裝機",
            "⚙️ 下拉選項管理",
            "📋 測試結果",
            "📥 Excel 匯出",
        ])

        with dev_form_tab:
            st.markdown("### 新版新增裝機紀錄（原型）")
            st.caption("請先完成廠別、案件與機台名稱識別，系統檢查舊紀錄後才會開放其餘欄位。")

            dev_key = st.session_state.dev_add_form_key

            if not st.session_state.dev_identity_draft:
                st.markdown("#### 1. 識別裝機資料")
                identity_col1, identity_col2, identity_col3 = st.columns(3)
                with identity_col1:
                    plant_choices = st.session_state.dev_plant_options or ["（請先至選項管理新增廠別）"]
                    identity_plant = st.selectbox(
                        "廠別 *",
                        plant_choices,
                        disabled=not st.session_state.dev_plant_options,
                        key=f"dev_identity_plant_{dev_key}",
                    )
                with identity_col2:
                    case_choices = st.session_state.dev_case_options or ["（請先至選項管理新增案件）"]
                    identity_case = st.selectbox(
                        "案件 *",
                        case_choices,
                        disabled=not st.session_state.dev_case_options,
                        key=f"dev_identity_case_{dev_key}",
                    )
                with identity_col3:
                    identity_machine = st.text_input(
                        "機台名稱 *",
                        placeholder="輸入機台名稱",
                        key=f"dev_identity_machine_{dev_key}",
                    )

                if st.button(
                    "檢查未完成紀錄並繼續",
                    type="primary",
                    use_container_width=True,
                    key=f"dev_check_identity_{dev_key}",
                ):
                    identity_missing = []
                    if not st.session_state.dev_plant_options:
                        identity_missing.append("廠別")
                    if not st.session_state.dev_case_options:
                        identity_missing.append("案件")
                    if not identity_machine.strip():
                        identity_missing.append("機台名稱")

                    if identity_missing:
                        st.error(f"請填寫必填欄位：{'、'.join(identity_missing)}")
                    else:
                        identity_data = {
                            "廠別": identity_plant,
                            "案件": identity_case,
                            "機台名稱": identity_machine.strip(),
                        }
                        target_identity = (
                            identity_plant.strip().casefold(),
                            identity_case.strip().casefold(),
                            identity_machine.strip().casefold(),
                        )
                        previous_unfinished = next(
                            (
                                record for record in reversed(st.session_state.dev_test_records)
                                if (
                                    str(record.get("廠別", "")).strip().casefold(),
                                    str(record.get("案件", "")).strip().casefold(),
                                    str(record.get("機台名稱", "")).strip().casefold(),
                                ) == target_identity
                                and str(record.get("狀態", "")).strip() == "未完成"
                            ),
                            None,
                        )

                        if previous_unfinished:
                            st.session_state.dev_pending_previous_record = {
                                "基本資料": identity_data,
                                "上次資料": dict(previous_unfinished),
                            }
                            show_previous_record_dialog()
                        else:
                            st.session_state.dev_identity_draft = identity_data
                            st.session_state.dev_loaded_case = identity_case
                            st.session_state.dev_previous_prefill = None
                            st.session_state.dev_pending_previous_record = None
                            st.session_state.dev_add_form_key += 1
                            st.session_state.dev_checklist_key += 1
                            st.rerun()
            else:
                identity_data = st.session_state.dev_identity_draft
                loaded_case = identity_data["案件"]
                previous_prefill = st.session_state.dev_previous_prefill or {}

                identity_info_col, identity_action_col = st.columns([4, 1])
                with identity_info_col:
                    st.success(
                        f"已確認：{identity_data['廠別']}／{loaded_case}／{identity_data['機台名稱']}"
                    )
                with identity_action_col:
                    if st.button(
                        "重新選擇",
                        use_container_width=True,
                        key=f"dev_reset_identity_{dev_key}",
                    ):
                        st.session_state.dev_identity_draft = None
                        st.session_state.dev_previous_prefill = None
                        st.session_state.dev_loaded_case = None
                        st.session_state.dev_add_preview = None
                        st.session_state.dev_pending_preview = None
                        st.session_state.dev_pending_previous_record = None
                        st.session_state.dev_add_form_key += 1
                        st.session_state.dev_checklist_key += 1
                        st.rerun()

                checklist_items = st.session_state.dev_case_checklists.get(loaded_case, [])
                checklist_key = st.session_state.dev_checklist_key
                previous_checklist = str(previous_prefill.get("項目確認", ""))
                status_options = ["未完成", "已完成"]
                previous_status = str(previous_prefill.get("狀態", ""))
                status_index = status_options.index(previous_status) if previous_status in status_options else 0

                with st.form(f"dev_add_installation_form_{dev_key}"):
                    st.markdown("#### 2. 裝機日期")
                    dev_date = st.date_input(
                        "裝機日期 *",
                        datetime.now(),
                        key=f"dev_date_{dev_key}",
                    )

                    st.markdown("#### 3. 項目確認")
                    checklist_results = {}
                    if not checklist_items:
                        st.info("此案件尚未設定確認項目，請至「下拉選項管理」新增。")
                    else:
                        checklist_groups = parse_checklist_definition(checklist_items)
                        item_index = 0
                        for category_group in checklist_groups:
                            category_name = category_group["name"]
                            if category_name:
                                with st.expander(f"📂 {category_name}", expanded=False):
                                    for item_name in category_group["items"]:
                                        item_label = f"{category_name} / {item_name}"
                                        checklist_results[item_label] = st.checkbox(
                                            item_name,
                                            value=(
                                                f"✅ {item_label}" in previous_checklist
                                                or f"✅ {item_name}" in previous_checklist
                                            ),
                                            key=f"dev_check_{dev_key}_{checklist_key}_{item_index}",
                                        )
                                        item_index += 1
                                    for subgroup in category_group["subgroups"]:
                                        st.markdown(f"**↳ {subgroup['name']}**")
                                        for item_name in subgroup["items"]:
                                            item_label = f"{category_name} / {subgroup['name']} / {item_name}"
                                            checklist_results[item_label] = st.checkbox(
                                                item_name,
                                                value=(
                                                    f"✅ {item_label}" in previous_checklist
                                                    or f"✅ {item_name}" in previous_checklist
                                                ),
                                                key=f"dev_check_{dev_key}_{checklist_key}_{item_index}",
                                            )
                                            item_index += 1
                            else:
                                for item_name in category_group["items"]:
                                    checklist_results[item_name] = st.checkbox(
                                        item_name,
                                        value=f"✅ {item_name}" in previous_checklist,
                                        key=f"dev_check_{dev_key}_{checklist_key}_{item_index}",
                                    )
                                    item_index += 1
                                for subgroup in category_group["subgroups"]:
                                    st.markdown(f"**↳ {subgroup['name']}**")
                                    for item_name in subgroup["items"]:
                                        item_label = f"{subgroup['name']} / {item_name}"
                                        checklist_results[item_label] = st.checkbox(
                                            item_name,
                                            value=(
                                                f"✅ {item_label}" in previous_checklist
                                                or f"✅ {item_name}" in previous_checklist
                                            ),
                                            key=f"dev_check_{dev_key}_{checklist_key}_{item_index}",
                                        )
                                        item_index += 1

                    st.markdown("#### 4. 執行資訊")
                    work_col1, work_col2 = st.columns(2)
                    with work_col1:
                        dev_status = st.selectbox(
                            "目前狀態",
                            status_options,
                            index=status_index,
                            key=f"dev_status_{dev_key}",
                        )
                    with work_col2:
                        dev_installers = st.multiselect(
                            "安裝人員",
                            installers_list,
                            key=f"dev_installers_{dev_key}",
                        )

                    st.markdown("#### 5. 備註")
                    dev_remark = st.text_area(
                        "Remark",
                        value=str(previous_prefill.get("Remark", "")),
                        placeholder="輸入進度、缺料項目或其他注意事項",
                        height=120,
                        key=f"dev_remark_{dev_key}",
                    )

                    preview_submitted = st.form_submit_button(
                        "產生送出預覽",
                        type="primary",
                        use_container_width=True,
                    )

                if preview_submitted:
                    checklist_summary = "、".join(
                        f"{'✅' if checked else '❌'} {item_name}"
                        for item_name, checked in checklist_results.items()
                    ) or "未設定確認項目"
                    current_preview = {
                        "日期": dev_date.strftime("%Y-%m-%d"),
                        "廠別": identity_data["廠別"],
                        "案件": loaded_case,
                        "機台名稱": identity_data["機台名稱"],
                        "項目確認": checklist_summary,
                        "安裝人員": "、".join(dev_installers) if dev_installers else "未指定",
                        "狀態": dev_status,
                        "Remark": dev_remark.strip(),
                    }
                    prepare_dev_preview(current_preview)
                    if st.session_state.dev_pending_preview:
                        show_dev_reason_dialog()

                if st.session_state.dev_pending_preview and not preview_submitted:
                    show_dev_reason_dialog()

                if st.session_state.dev_add_preview:
                    st.divider()
                    st.markdown("### 送出前預覽")
                    preview_df = pd.DataFrame([st.session_state.dev_add_preview]).rename(
                        columns={"未完成或缺貨原因": "未完成原因"}
                    )
                    st.dataframe(preview_df, hide_index=True, use_container_width=True)
                    st.success("請確認內容；確認後可加入測試結果清單。")

                    if st.button(
                        "確認加入測試結果",
                        type="primary",
                        use_container_width=True,
                        key="dev_confirm_test_record",
                    ):
                        test_record = dict(st.session_state.dev_add_preview)
                        test_record["建立時間"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        completed_count = 0
                        if test_record.get("狀態") == "已完成":
                            target_identity = (
                                str(test_record.get("廠別", "")).strip().casefold(),
                                str(test_record.get("案件", "")).strip().casefold(),
                                str(test_record.get("機台名稱", "")).strip().casefold(),
                            )
                            for existing_record in st.session_state.dev_test_records:
                                existing_identity = (
                                    str(existing_record.get("廠別", "")).strip().casefold(),
                                    str(existing_record.get("案件", "")).strip().casefold(),
                                    str(existing_record.get("機台名稱", "")).strip().casefold(),
                                )
                                if (
                                    existing_identity == target_identity
                                    and existing_record.get("狀態") != "已完成"
                                ):
                                    existing_record["狀態"] = "已完成"
                                    existing_record["未完成或缺貨原因"] = ""
                                    completed_count += 1

                        st.session_state.dev_test_records.append(test_record)
                        st.session_state.dev_add_preview = None
                        st.session_state.dev_previous_prefill = None
                        if completed_count:
                            st.success(
                                f"已加入測試結果，並將 {completed_count} 筆相同廠別、案件、機台的紀錄更新為已完成。"
                            )
                        else:
                            st.success("已加入測試結果，請至「📋 測試結果」分頁查看。")

                if st.button("清空原型表單", key="dev_clear_add_form"):
                    st.session_state.dev_add_form_key += 1
                    st.session_state.dev_checklist_key += 1
                    st.session_state.dev_loaded_case = None
                    st.session_state.dev_identity_draft = None
                    st.session_state.dev_previous_prefill = None
                    st.session_state.dev_add_preview = None
                    st.session_state.dev_pending_preview = None
                    st.session_state.dev_pending_previous_record = None
                    st.rerun()

        with dev_options_tab:
            st.markdown("### 下拉選項管理")
            st.caption("管理新版表單使用的廠別與案件選項；目前異動不會修改 Google Sheets。")

            def render_option_manager(title, state_key, widget_prefix):
                options = st.session_state[state_key]
                manager_key = st.session_state.dev_option_manager_key

                with st.container(border=True):
                    st.markdown(f"#### {title}")

                    add_col1, add_col2 = st.columns([3, 1])
                    with add_col1:
                        new_option = st.text_input(
                            f"新增{title}",
                            placeholder=f"輸入新的{title}名稱",
                            key=f"{widget_prefix}_new_{manager_key}",
                        )
                    with add_col2:
                        st.write("")
                        st.write("")
                        add_clicked = st.button(
                            "新增選項",
                            use_container_width=True,
                            key=f"{widget_prefix}_add_{manager_key}",
                        )

                    if add_clicked:
                        cleaned_name = new_option.strip()
                        if not cleaned_name:
                            st.error(f"請輸入{title}名稱。")
                        elif cleaned_name in options:
                            st.warning(f"此{title}已存在。")
                        else:
                            options.append(cleaned_name)
                            options.sort()
                            if state_key == "dev_case_options":
                                st.session_state.dev_case_checklists.setdefault(cleaned_name, [])
                            st.session_state.dev_option_manager_key += 1
                            st.session_state.dev_add_form_key += 1
                            st.rerun()

                    if options:
                        edit_col1, edit_col2 = st.columns(2)
                        with edit_col1:
                            selected_option = st.selectbox(
                                f"選擇要修改的{title}",
                                options,
                                key=f"{widget_prefix}_selected_{manager_key}",
                            )
                        with edit_col2:
                            renamed_option = st.text_input(
                                f"新的{title}名稱",
                                placeholder="輸入新名稱",
                                key=f"{widget_prefix}_rename_value_{manager_key}",
                            )

                        action_col1, action_col2 = st.columns(2)
                        with action_col1:
                            rename_clicked = st.button(
                                "儲存名稱修改",
                                use_container_width=True,
                                key=f"{widget_prefix}_rename_{manager_key}",
                            )
                        with action_col2:
                            delete_clicked = st.button(
                                "刪除選項",
                                use_container_width=True,
                                key=f"{widget_prefix}_delete_{manager_key}",
                            )

                        if rename_clicked:
                            cleaned_rename = renamed_option.strip()
                            if not cleaned_rename:
                                st.error(f"請輸入新的{title}名稱。")
                            elif cleaned_rename != selected_option and cleaned_rename in options:
                                st.warning(f"新的{title}名稱已存在。")
                            else:
                                option_index = options.index(selected_option)
                                options[option_index] = cleaned_rename
                                options.sort()
                                if state_key == "dev_case_options" and cleaned_rename != selected_option:
                                    existing_items = st.session_state.dev_case_checklists.pop(selected_option, [])
                                    st.session_state.dev_case_checklists[cleaned_rename] = existing_items
                                    if st.session_state.dev_loaded_case == selected_option:
                                        st.session_state.dev_loaded_case = cleaned_rename
                                st.session_state.dev_option_manager_key += 1
                                st.session_state.dev_add_form_key += 1
                                st.session_state.dev_add_preview = None
                                st.rerun()

                        if delete_clicked:
                            options.remove(selected_option)
                            if state_key == "dev_case_options":
                                st.session_state.dev_case_checklists.pop(selected_option, None)
                                if st.session_state.dev_loaded_case == selected_option:
                                    st.session_state.dev_loaded_case = None
                            st.session_state.dev_option_manager_key += 1
                            st.session_state.dev_add_form_key += 1
                            st.session_state.dev_add_preview = None
                            st.rerun()
                    else:
                        st.info(f"目前沒有{title}選項，請先新增。")

            render_option_manager("廠別", "dev_plant_options", "dev_plant")
            render_option_manager("案件", "dev_case_options", "dev_case")

            with st.container(border=True):
                st.markdown("#### 案件確認項目")
                st.caption("使用 [主分類] 與 [[子分類]] 建立階層；分類只供整理，不需勾選。")

                if st.session_state.dev_case_options:
                    checklist_manager_key = st.session_state.dev_option_manager_key
                    checklist_case = st.selectbox(
                        "選擇案件",
                        st.session_state.dev_case_options,
                        key=f"dev_checklist_case_{checklist_manager_key}",
                    )
                    checklist_case_index = st.session_state.dev_case_options.index(checklist_case)
                    existing_checklist = st.session_state.dev_case_checklists.get(checklist_case, [])
                    checklist_text = st.text_area(
                        "分類與確認項目",
                        value="\n".join(existing_checklist),
                        placeholder="例如：\n[電力系統]\n電源確認\n[[接地系統]]\n接地線確認\n接地阻抗確認",
                        height=180,
                        key=f"dev_checklist_text_{checklist_manager_key}_{checklist_case_index}",
                    )

                    if st.button(
                        "儲存案件確認項目",
                        type="primary",
                        use_container_width=True,
                        key=f"dev_save_checklist_{checklist_manager_key}",
                    ):
                        cleaned_items = []
                        for line in checklist_text.splitlines():
                            cleaned_item = line.strip()
                            if cleaned_item and cleaned_item not in cleaned_items:
                                cleaned_items.append(cleaned_item)

                        st.session_state.dev_case_checklists[checklist_case] = cleaned_items
                        st.session_state.dev_option_manager_key += 1
                        st.session_state.dev_checklist_key += 1
                        st.session_state.dev_add_preview = None
                        st.rerun()
                else:
                    st.info("請先新增案件，才能設定確認項目。")

        with dev_results_tab:
            st.markdown("### 新增裝機測試結果")
            st.caption("此處顯示開發區確認加入的資料；只有手動按下同步按鈕時才會寫入 Google Sheets。")

            if st.session_state.dev_test_records:
                results_df = pd.DataFrame(st.session_state.dev_test_records)
                preferred_columns = [
                    "建立時間",
                    "日期",
                    "廠別",
                    "案件",
                    "機台名稱",
                    "項目確認",
                    "安裝人員",
                    "狀態",
                    "未完成或缺貨原因",
                    "Remark",
                ]
                result_columns = [column for column in preferred_columns if column in results_df.columns]
                display_results_df = results_df[result_columns].copy()
                display_results_df = display_results_df.rename(
                    columns={"未完成或缺貨原因": "未完成原因"}
                )
                if "項目確認" in display_results_df.columns:
                    display_results_df["項目確認"] = display_results_df["項目確認"].apply(
                        format_checklist_progress
                    )
                results_event = st.dataframe(
                    display_results_df,
                    hide_index=True,
                    use_container_width=True,
                    on_select="rerun",
                    selection_mode="single-row",
                    key=f"dev_results_grid_{st.session_state.dev_results_grid_key}",
                )
                if (
                    results_event.selection.rows
                    and not st.session_state.dev_pending_preview
                    and not st.session_state.dev_pending_previous_record
                ):
                    selected_result_index = results_event.selection.rows[0]
                    show_details_dialog(
                        results_df.iloc[selected_result_index],
                        "dev_results_grid_key",
                    )
                st.info(f"目前共有 {len(results_df)} 筆測試資料。")

                if st.button(
                    "清空所有測試結果",
                    use_container_width=True,
                    key="dev_clear_test_records",
                ):
                    st.session_state.dev_test_records = []
                    st.rerun()
            else:
                st.info("目前沒有測試資料。請先在「📝 新版新增裝機」完成輸入並確認加入。")

            st.divider()
            st.markdown("### 同步開發資料")
            st.caption("同步會保存完整快照；載入則會用最近一次快照取代目前工作階段內容。")
            sync_col, load_col = st.columns(2)
            with sync_col:
                sync_clicked = st.button(
                    "儲存目前內容",
                    type="primary",
                    use_container_width=True,
                    key="dev_sync_to_google",
                )
            with load_col:
                load_clicked = st.button(
                    "載入上次保存內容",
                    use_container_width=True,
                    key="dev_load_from_google",
                )

            if sync_clicked:
                with st.spinner("正在保存開發區資料..."):
                    try:
                        synced_count = sync_dev_data_to_google()
                        if synced_count:
                            st.success(
                                f"保存完成！已將 {synced_count} 筆資料寫入「{DEV_WORKSHEET_NAME}」。"
                            )
                        else:
                            st.warning("目前沒有可保存的廠別、案件或測試結果。")
                    except Exception as e:
                        st.error(f"保存失敗，請稍後再試。詳細錯誤：{e}")

            if load_clicked:
                with st.spinner("正在載入最近一次保存內容..."):
                    try:
                        load_summary = load_latest_dev_data_from_google()
                        if load_summary:
                            st.session_state.dev_flash_message = (
                                f"已載入 {load_summary['批次']} 的保存內容："
                                f"{load_summary['廠別']} 個廠別、"
                                f"{load_summary['案件']} 個案件、"
                                f"{load_summary['測試結果']} 筆測試結果。"
                            )
                            st.rerun()
                        else:
                            st.warning("目前找不到可載入的開發區保存內容。")
                    except Exception as e:
                        st.error(f"載入失敗，請稍後再試。詳細錯誤：{e}")

        with dev_excel_tab:
            st.markdown("### 多機台確認項目 Excel")
            st.caption("依範例格式，每台機台一列；已施工使用綠色儲存格，待施工使用黃色儲存格。")

            if not st.session_state.dev_test_records:
                st.info("目前沒有可匯出的測試結果，請先新增資料或載入上次保存內容。")
            else:
                export_plants = sorted({
                    str(record.get("廠別", "")).strip()
                    for record in st.session_state.dev_test_records
                    if str(record.get("廠別", "")).strip()
                })
                export_plant = st.selectbox(
                    "選擇廠別",
                    export_plants,
                    key="dev_export_plant",
                )

                plant_records = [
                    record for record in st.session_state.dev_test_records
                    if str(record.get("廠別", "")).strip() == export_plant
                ]
                export_cases = sorted({
                    str(record.get("案件", "")).strip()
                    for record in plant_records
                    if str(record.get("案件", "")).strip()
                })
                export_case = st.selectbox(
                    "選擇案件",
                    export_cases,
                    key=f"dev_export_case_{export_plant}",
                )

                case_records = [
                    record for record in plant_records
                    if str(record.get("案件", "")).strip() == export_case
                ]
                latest_machine_records = {}
                for record in case_records:
                    machine_name = str(record.get("機台名稱", "")).strip()
                    if machine_name:
                        latest_machine_records[machine_name] = record

                machine_options = sorted(latest_machine_records)
                selected_machines = st.multiselect(
                    "選擇要輸出的機台（可複選）",
                    machine_options,
                    default=[],
                    key=f"dev_export_machines_{export_plant}_{export_case}",
                )

                checklist_definition = st.session_state.dev_case_checklists.get(export_case, [])
                checklist_options = get_checklist_export_labels(checklist_definition)
                st.markdown("#### 選擇要生成的確認項目")
                st.caption("選項預設全部不勾選，請依主分類與子分類展開挑選。")
                selected_export_items = render_export_checklist_selector(
                    checklist_definition,
                    key_prefix=f"dev_export_checklist_{export_plant}_{export_case}",
                )
                st.caption("🟩 已施工　🟨 待施工")

                if not checklist_options:
                    st.warning("此案件尚未建立確認項目，請先至「下拉選項管理」設定。")
                elif not selected_machines:
                    st.warning("請至少選擇一台機台。")
                elif not selected_export_items:
                    st.warning("請至少選擇一個確認項目。")
                else:
                    selected_machine_records = [
                        latest_machine_records[machine_name]
                        for machine_name in selected_machines
                    ]
                    excel_data = build_dev_excel_export(
                        selected_machine_records,
                        selected_export_items,
                    )
                    unsafe_filename = f"{export_plant}_{export_case}_裝機確認表_{datetime.now().strftime('%Y%m%d')}.xlsx"
                    safe_filename = "".join(
                        "_" if char in '\\/:*?\"<>|' else char
                        for char in unsafe_filename
                    )

                    st.success(
                        f"將輸出 {len(selected_machine_records)} 台機台、"
                        f"{len(selected_export_items)} 個確認項目。"
                    )
                    st.download_button(
                        "下載 Excel",
                        data=excel_data,
                        file_name=safe_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True,
                        key="dev_download_excel",
                    )
