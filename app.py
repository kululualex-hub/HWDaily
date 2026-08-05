import streamlit as st
import gspread
import pandas as pd
from datetime import datetime
import json
import io

# ==================== 1. 網頁基本與連線設定 ====================
st.set_page_config(page_title="鴻伍裝機日報系統", layout="wide")

# 初始化 Session State (記憶按鈕操作與暫存資料)
if 'user_role' not in st.session_state:
    st.session_state.user_role = "未登入"  # 可選: 管理者, RD, 業務, 工程師, 公用, 未登入
if 'user_name' not in st.session_state:
    st.session_state.user_name = "訪客"
    
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

# 用來強制重置表格勾選狀態的計數器
if 'tab1_grid_key' not in st.session_state:
    st.session_state.tab1_grid_key = 0
if 'tab3_grid_key' not in st.session_state:
    st.session_state.tab3_grid_key = 0
if 'tab4_grid_key' not in st.session_state:
    st.session_state.tab4_grid_key = 0

# 用來控制新增分頁欄位清空的計數器
if 'add_form_key' not in st.session_state:
    st.session_state.add_form_key = 0

@st.cache_resource
def get_sheets():
    # 智慧判斷環境：如果在雲端則讀取 st.secrets，若在本機則讀取 credentials.json 檔案
    if "gcp_service_account" in st.secrets:
        creds_dict = dict(st.secrets["gcp_service_account"])
        gc = gspread.service_account_from_dict(creds_dict)
    else:
        gc = gspread.service_account(filename='credentials.json')
        
    sh = gc.open("control table")
    worksheet_main = sh.worksheet("裝機人員進廠時間")
    
    # 嘗試取得或自動建立「修改紀錄」分頁
    try:
        worksheet_log = sh.worksheet("修改紀錄")
    except:
        worksheet_log = sh.add_worksheet(title="修改紀錄", rows=1000, cols=10)
        worksheet_log.append_row(["修改時間", "更改人員", "修改項目", "修改內容"])
        
    # 嘗試取得「帳號管理」分頁
    try:
        worksheet_accounts = sh.worksheet("帳號管理")
    except:
        worksheet_accounts = None
        
    return worksheet_main, worksheet_log, worksheet_accounts

worksheet, worksheet_log, worksheet_accounts = get_sheets()

# 記錄異動到「修改紀錄」分頁的輔助函式
def log_modification(editor_name, item_name, details):
    try:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        worksheet_log.append_row([now_str, editor_name, item_name, details])
    except Exception as e:
        st.error(f"寫入修改紀錄失敗: {e}")

# ==================== 2. 側邊欄：多重身分與訪客登入管理 ====================
with st.sidebar:
    st.header("🔐 身分與權限管理")
    
    current_role = st.session_state.user_role
    
    if current_role == "未登入":
        st.info("👋 歡迎！請選擇登入身分或使用訪客快速登入。")
        
        # 🚀 訪客快速登入按鈕 (公用訪問權)
        if st.button("🚀 訪客快速登入 (公用)", type="primary", use_container_width=True):
            st.session_state.user_role = "公用"
            st.session_state.user_name = "訪客(公用)"
            st.success("已以公用身分登入！")
            st.rerun()
            
        st.divider()
        st.markdown("##### 帳號密碼登入")
        input_id = st.text_input("帳號 (ID):")
        input_pwd = st.text_input("密碼:", type="password")
        
        if st.button("登入系統", use_container_width=True):
            success_login = False
            if worksheet_accounts:
                try:
                    acc_data = worksheet_accounts.get_all_records()
                    for row in acc_data:
                        r_user = str(row.get("帳號", "")).strip()
                        r_pass = str(row.get("密碼", "")).strip()
                        if r_user.lower() == input_id.strip().lower() and r_pass == input_pwd.strip():
                            st.session_state.user_role = str(row.get("權限", "工程師")).strip()
                            st.session_state.user_name = str(row.get("姓名", input_id)).strip()
                            success_login = True
                            break
                except Exception as e:
                    st.error(f"讀取帳號管理分頁失敗: {e}")
            
            # 若無設定帳號管理分頁，提供預設管理者密碼防呆 (預設 1234)
            if not success_login and input_id == "admin" and input_pwd == "1234":
                st.session_state.user_role = "管理者"
                st.session_state.user_name = "系統管理員"
                success_login = True
                
            if success_login:
                st.success(f"登入成功！身分：{st.session_state.user_role}")
                st.rerun()
            else:
                st.error("帳號或密碼錯誤！")
    else:
        st.success(f"👤 當前使用者：{st.session_state.user_name}\n\n🛡️ 身分權限：**{current_role}**")
        if current_role == "管理者":
            st.markdown("✨ **[管理者權限]** 您擁有本系統完整之查詢、新增、修改與狀態更新權限。")
        else:
            st.markdown(f"👁️ **[{current_role}權限]** 您可進行資料查詢、匯出與新增紀錄。修改資料僅限管理者。")
            
        if st.button("登出 / 切換身分", use_container_width=True):
            st.session_state.user_role = "未登入"
            st.session_state.user_name = "訪客"
            st.session_state.tab3_edit_requested = False
            st.session_state.tab3_edit_confirmed = False
            st.rerun()

st.title("📊 鴻伍裝機日報系統 (Web 雲端版)")

# ==================== 3. 彈出視窗功能 (Dialog) ====================
@st.dialog("📝 詳細資料檢視")
def show_details_dialog(row_data, reset_key):
    st.markdown(
        """
        <style>
        div[role="dialog"] button[aria-label="Close"] {
            display: none !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    st.markdown(f"### 🏭 廠別：{row_data.get('廠別', '')}")
    st.markdown(f"### 💻 機台名稱：{row_data.get('機台名稱', '')}")
    st.markdown(f"**📅 日期：** {row_data.get('日期', '')}")
    st.markdown(f"**📂 案件：** {row_data.get('案件', '')}")
    st.markdown(f"**👷 安裝人員：** {row_data.get('安裝人員', '')}")
    st.markdown(f"**🚦 狀態：** {row_data.get('狀態', '')}")
    st.markdown("---")
    st.markdown("### 📝 Remark (備忘內容)：")
    
    remark_text = str(row_data.get('Remark', '')).replace('\n', '<br>')
    if not remark_text.strip():
        remark_text = "（無備忘內容）"
        
    st.markdown(
        f"<div style='font-size: 1.3em; line-height: 1.6; background-color: #f0f2f6; padding: 15px; border-radius: 8px; color: #333; margin-bottom: 20px;'>"
        f"{remark_text}</div>",
        unsafe_allow_html=True
    )
    
    if st.button("❌ 關閉視窗並取消選取", type="primary", use_container_width=True):
        st.session_state[reset_key] += 1  
        st.rerun()

# ==================== 4. 建立四大功能分頁 ====================
installers_list = [
    "鍾博宇", "黃政欽", "張智偉", "林嬴燦", "吳建華", "何乙霆"
]

tab1, tab2, tab3, tab4 = st.tabs([
    "🌅 晨會當日動態", 
    "📝 新增裝機紀錄", 
    "🔍 歷史搜尋與修改", 
    "📌 待追蹤清單 (更新狀態)"
])

# ==================== 分頁 1：晨會當日動態 (完全開放) ====================
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
                filtered_df[display_cols], 
                hide_index=True,
                use_container_width=True, 
                on_select="rerun", 
                selection_mode="single-row",
                key=f"tab1_grid_{st.session_state.tab1_grid_key}"
            )
            if event.selection.rows:
                selected_idx = event.selection.rows[0]
                show_details_dialog(filtered_df.iloc[selected_idx], 'tab1_grid_key')
        else:
            st.info(f"📅 該日尚無裝機紀錄。")

# ==================== 分頁 2：新增裝機紀錄 (開放新增) ====================
with tab2:
    st.subheader("填寫裝機資訊")
    
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
        btn_clear = st.button("🗑️ 清空欄位", key="btn_clear_form")

    if btn_clear:
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
                        idx = headers.index(col_name)
                        new_row[idx] = val

                fill_col("日期", date_str)
                fill_col("安裝人員", installer_str)
                fill_col("廠別", plant)
                fill_col("案件", case)
                fill_col("機台名稱", machine)
                fill_col("狀態", status)
                fill_col("Remark", remark)
                
                worksheet.append_row(new_row)
                
                # 記錄到修改紀錄
                log_modification(
                    f"{st.session_state.user_name} ({st.session_state.user_role})", 
                    f"新增機台: {machine} ({plant})", 
                    f"日期: {date_str}, 狀態: {status}"
                )
                
                st.success(f"✅ 成功將機台【{machine}】新增至雲端！")

# ==================== 分頁 3：歷史紀錄搜尋與修改 (僅管理者可修改) ====================
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
        
        with col_s1:
            date_range = st.date_input("選擇日期區間:", [])
        
        with col_s2:
            search_plant = st.selectbox("廠別:", unique_plants)
            
        if search_plant != "(全部)":
            target_plant_df = df_search[df_search['廠別'].astype(str).str.strip() == search_plant]
            unique_machines = ["(全部)"] + sorted(list(set([str(x).strip() for x in target_plant_df['機台名稱'] if str(x).strip()])))
            is_machine_disabled = False
        else:
            unique_machines = ["(請先選擇廠別)"]
            is_machine_disabled = True
            target_plant_df = pd.DataFrame()
            
        with col_s3:
            search_machine = st.selectbox("機台名稱 (廠別確定後解鎖):", unique_machines, disabled=is_machine_disabled)
            
        if search_plant != "(全部)":
            if search_machine not in ["(全部)", "(請先選擇廠別)"]:
                target_machine_df = target_plant_df[target_plant_df['機台名稱'].astype(str).str.strip() == search_machine]
                unique_cases = ["(全部)"] + sorted(list(set([str(x).strip() for x in target_machine_df['案件'] if str(x).strip()])))
            else:
                unique_cases = ["(全部)"] + sorted(list(set([str(x).strip() for x in target_plant_df['案件'] if str(x).strip()])))
            is_case_disabled = False
        else:
            unique_cases = ["(請先選擇廠別)"]
            is_case_disabled = True
            
        with col_s4:
            search_case = st.selectbox("案件 (廠別確定後解鎖):", unique_cases, disabled=is_case_disabled)

        with col_s5:
            search_installer = st.selectbox("安裝人員:", unique_installers)
            
        if st.button("🔍 開始搜尋", type="primary", key="btn_execute_search"):
            with st.spinner("搜尋中..."):
                filtered_df = df_search.copy()
                
                if len(date_range) == 2:
                    filtered_df['日期_temp'] = pd.to_datetime(filtered_df['日期'], format='mixed', errors='coerce').dt.date
                    filtered_df = filtered_df[(filtered_df['日期_temp'] >= date_range[0]) & (filtered_df['日期_temp'] <= date_range[1])]
                    filtered_df = filtered_df.drop(columns=['日期_temp'])
                elif len(date_range) == 1:
                    filtered_df['日期_temp'] = pd.to_datetime(filtered_df['日期'], format='mixed', errors='coerce').dt.date
                    filtered_df = filtered_df[filtered_df['日期_temp'] == date_range[0]]
                    filtered_df = filtered_df.drop(columns=['日期_temp'])
                    
                if search_plant != "(全部)":
                    filtered_df = filtered_df[filtered_df['廠別'].astype(str).str.strip() == search_plant]
                if search_machine not in ["(全部)", "(請先選擇廠別)"]:
                    filtered_df = filtered_df[filtered_df['機台名稱'].astype(str).str.strip() == search_machine]
                if search_case not in ["(全部)", "(請先選擇廠別)"]:
                    filtered_df = filtered_df[filtered_df['案件'].astype(str).str.strip() == search_case]
                if search_installer != "(全部)":
                    filtered_df = filtered_df[filtered_df['安裝人員'].astype(str).str.contains(search_installer)]
                
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
                
                # 產生 Excel 檔案
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    filtered_df[view_cols].to_excel(writer, sheet_name='裝機搜尋結果', index=False)
                    workbook = writer.book
                    worksheet_excel = writer.sheets['裝機搜尋結果']
                    wrap_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})
                    default_format = workbook.add_format({'valign': 'top'})
                    
                    for idx, col_name in enumerate(view_cols):
                        if col_name in ['Remark', '安裝人員']:
                            width = 45 if col_name == 'Remark' else 20
                            worksheet_excel.set_column(idx, idx, width, wrap_format)
                        else:
                            worksheet_excel.set_column(idx, idx, 18, default_format)
                            
                buffer.seek(0)
                
                st.download_button(
                    label="📥 匯出搜尋結果為 Excel",
                    data=buffer,
                    file_name=f"鴻伍裝機搜尋結果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                if not st.session_state.tab3_edit_confirmed:
                    st.markdown("🖱️ 提示：點擊任意列可彈出詳細資訊")
                    event = st.dataframe(
                        filtered_df[view_cols], 
                        hide_index=True, 
                        use_container_width=True,
                        on_select="rerun",
                        selection_mode="single-row",
                        key=f"tab3_grid_{st.session_state.tab3_grid_key}"
                    )
                    
                    if event.selection.rows:
                        selected_idx = event.selection.rows[0]
                        show_details_dialog(filtered_df.iloc[selected_idx], 'tab3_grid_key')
                    
                    if not st.session_state.tab3_edit_requested:
                        # 權限控管：只有「管理者」才顯示修改按鈕
                        if st.session_state.user_role == "管理者":
                            if st.button("✏️ 開啟修改模式"):
                                st.session_state.tab3_edit_requested = True
                                st.rerun()
                        else:
                            st.info("💡 目前身分為唯讀/查詢，若需修改歷史資料，請透過側邊欄切換為「管理者」身分登入。")
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
                    
                    edited_df = st.data_editor(
                        filtered_df[view_cols],
                        hide_index=True,
                        use_container_width=True,
                        key="search_editor"
                    )
                    
                    if st.button("💾 儲存表格上的所有修改", type="primary"):
                        with st.spinner("正在批次同步更新至雲端 (請稍候)..."):
                            changed_rows = []
                            for i in range(len(edited_df)):
                                orig_row = filtered_df[view_cols].iloc[i]
                                new_row = edited_df.iloc[i]
                                
                                for col in view_cols:
                                    if str(orig_row[col]).strip() != str(new_row[col]).strip():
                                        changed_rows.append((filtered_df.iloc[i]['Sheet_Row'], new_row))
                                        break
                                        
                            if changed_rows:
                                headers = worksheet.row_values(1)
                                cells_to_update = []
                                
                                for sheet_idx, row_data in changed_rows:
                                    for col_name in view_cols:
                                        if col_name in headers:
                                            col_idx = headers.index(col_name) + 1
                                            cells_to_update.append(gspread.Cell(int(sheet_idx), col_idx, str(row_data[col_name])))
                                
                                if cells_to_update:
                                    worksheet.update_cells(cells_to_update)
                                    
                                # 記錄到修改紀錄
                                log_modification(
                                    f"{st.session_state.user_name} ({st.session_state.user_role})",
                                    "批次歷史修改",
                                    f"更新了 {len(changed_rows)} 筆資料列"
                                )
                                    
                                st.success(f"✅ 成功批次更新 {len(changed_rows)} 筆資料至雲端！")
                                st.session_state.tab3_search_active = False
                                st.session_state.tab3_edit_confirmed = False
                                st.rerun()
                            else:
                                st.info("沒有偵測到任何修改內容。")
            else:
                st.warning("⚠️ 找不到符合您設定條件的紀錄。")
    else:
        st.info("試算表中尚無資料。")

# ==================== 分頁 4：待追蹤清單與狀態更新 (僅管理者可更新) ====================
with tab4:
    st.subheader("📌 待追蹤機台與狀態更新")
    
    data = worksheet.get_all_records()
    df = pd.DataFrame(data)
    
    if not df.empty:
        df = df.fillna("")
        df['Sheet_Row'] = df.index + 2
        
        pending_df = df[
            (df['狀態'].astype(str).str.strip() != '已完成') & 
            (df['狀態'].astype(str).str.strip() != '') & 
            (df['機台名稱'].astype(str).str.strip() != '')
        ]
        
        st.markdown(f"**目前待追蹤數量： <span style='color:red;'>{len(pending_df)}</span> 筆** (🖱️ 提示：點擊下方表格任意列可彈出詳細資訊)", unsafe_allow_html=True)
        
        if not pending_df.empty:
            display_cols = ["日期", "廠別", "案件", "機台名稱", "安裝人員", "狀態", "Remark"]
            display_cols = [col for col in display_cols if col in pending_df.columns]
            
            event = st.dataframe(
                pending_df[display_cols], 
                hide_index=True,
                use_container_width=True,
                on_select="rerun",
                selection_mode="single-row",
                key=f"tab4_grid_{st.session_state.tab4_grid_key}"
            )
            
            if event.selection.rows:
                selected_idx = event.selection.rows[0]
                show_details_dialog(pending_df.iloc[selected_idx], 'tab4_grid_key')
            
            st.divider()
            
            st.markdown("#### ✏️ 更新機台狀態")
            
            # 權限控管：只有「管理者」才能操作狀態更新
            if st.session_state.user_role != "管理者":
                st.info("💡 狀態更新功能僅限「管理者」操作。若需更新狀態，請透過側邊欄切換身分登入。")
            else:
                options = pending_df['Sheet_Row'].tolist()
                
                upd_col1, upd_col2 = st.columns([2, 1])
                with upd_col1:
                    selected_item_row = st.selectbox(
                        "選擇要更新的機台：", 
                        options, 
                        format_func=lambda r: f"{pending_df.loc[pending_df['Sheet_Row'] == r, '日期'].values[0]} | {pending_df.loc[pending_df['Sheet_Row'] == r, '廠別'].values[0]} - {pending_df.loc[pending_df['Sheet_Row'] == r, '機台名稱'].values[0]} (目前: {pending_df.loc[pending_df['Sheet_Row'] == r, '狀態'].values[0]})"
                    )
                with upd_col2:
                    new_status = st.selectbox("修改為新狀態：", ["未完成", "缺料", "已完成"], index=2)
                    
                if st.button("送出狀態更新", type="primary", key="btn_update"):
                    with st.spinner("同步至雲端中..."):
                        row_idx = int(selected_item_row)
                        
                        headers = worksheet.row_values(1)
                        if "狀態" in headers and "Remark" in headers:
                            status_col_idx = headers.index("狀態") + 1
                            remark_col_idx = headers.index("Remark") + 1
                            
                            # 取得當前機台名稱供日誌紀錄使用
                            current_machine = worksheet.cell(row_idx, headers.index("機台名稱") + 1).value
                            current_plant = worksheet.cell(row_idx, headers.index("廠別") + 1).value
                            
                            worksheet.update_cell(row_idx, status_col_idx, new_status)
                            
                            old_remark = worksheet.cell(row_idx, remark_col_idx).value or ""
                            today_str = datetime.now().strftime("%Y-%m-%d")
                            append_text = f"[{today_str} 更新狀態: {new_status}]"
                            
                            new_remark = (str(old_remark).strip() + "\n" + append_text).strip()
                            worksheet.update_cell(row_idx, remark_col_idx, new_remark)
                            
                            # 記錄到修改紀錄分頁
                            log_modification(
                                f"{st.session_state.user_name} ({st.session_state.user_role})",
                                f"更新機台狀態: {current_machine} ({current_plant})",
                                f"狀態變更為: {new_status}"
                            )
                            
                            st.success(f"✅ 更新成功！該機台已標記為「{new_status}」。")
                            st.rerun()
                        else:
                            st.error("試算表中找不到「狀態」或「Remark」欄位，請檢查表頭名稱。")
        else:
            st.success("🎉 太棒了！目前所有機台皆已完工，沒有待追蹤項目。")