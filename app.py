import streamlit as st
import gspread
import pandas as pd
from datetime import datetime
import json

# ==================== 1. 網頁基本與連線設定 ====================
st.set_page_config(page_title="裝機進度日報表系統", layout="wide")

# 初始化 Session State (記憶按鈕操作與暫存資料)
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
def get_sheet():
    # 智慧判斷環境：如果在雲端則讀取 st.secrets，若在本機則讀取 credentials.json 檔案
    if "gcp_service_account" in st.secrets:
        creds_dict = dict(st.secrets["gcp_service_account"])
        gc = gspread.service_account_from_dict(creds_dict)
    else:
        gc = gspread.service_account(filename='credentials.json')
        
    sh = gc.open("control table")
    return sh.worksheet("裝機人員進廠時間")

worksheet = get_sheet()

st.title("📊 裝機進度日報表系統 (Web 雲端版)")

# ==================== 2. 彈出視窗功能 (Dialog) ====================
@st.dialog("📝 詳細資料檢視")
def show_details_dialog(row_data, reset_key):
    # 透過 CSS 強制隱藏彈出視窗右上角的原生 X 按鈕
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
    
    # 專屬關閉按鈕，點擊後會清除勾選並關閉視窗
    if st.button("❌ 關閉視窗並取消選取", type="primary", use_container_width=True):
        st.session_state[reset_key] += 1  
        st.rerun()

# ==================== 3. 建立四大功能分頁 ====================
installers_list = [
    "鍾博宇", "黃政欽", "張智偉", "林嬴燦", "吳建華", "何乙霆"
]

tab1, tab2, tab3, tab4 = st.tabs([
    "🌅 晨會當日動態", 
    "📝 新增裝機紀錄", 
    "🔍 歷史搜尋與修改", 
    "📌 待追蹤清單 (更新狀態)"
])

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
                filtered_df[display_cols], 
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

# ==================== 分頁 2：新增裝機紀錄 ====================
with tab2:
    st.subheader("填寫裝機資訊")
    
    # 動態產生用來重置輸入欄位的 Key 後綴
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

    # 把按鈕放在並排的欄位中：一個是新增紀錄，一個是清空欄位
    b_col1, b_col2 = st.columns(2)
    with b_col1:
        btn_submit = st.button("💾 新增紀錄", type="primary", key="btn_add")
    with b_col2:
        btn_clear = st.button("🗑️ 清空欄位", key="btn_clear_form")

    # 點擊清空按鈕時的動作
    if btn_clear:
        st.session_state.add_form_key += 1
        st.rerun()

    # 點擊新增紀錄時的動作
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
        
        st.markdown("##### 1. 設定搜尋條件 (設定完畢後請點擊下方搜尋按鈕)")
        col_s1, col_s2, col_s3, col_s4 = st.columns(4)
        
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
                
                st.session_state.tab3_filtered_df = filtered_df
                st.session_state.tab3_search_active = True
                st.session_state.tab3_edit_requested = False
                st.session_state.tab3_edit_confirmed = False
                st.rerun()
                
        st.divider()
        
        if st.session_state.tab3_search_active:
            st.markdown("##### 2. 搜尋結果")
            filtered_df = st.session_state.tab3_filtered_df
            
            if not filtered_df.empty:
                display_cols = ["日期", "廠別", "案件", "機台名稱", "安裝人員", "狀態", "Remark", "Sheet_Row"]
                display_cols = [col for col in display_cols if col in filtered_df.columns]
                
                if not st.session_state.tab3_edit_confirmed:
                    st.markdown("🖱️ 提示：點擊任意列可彈出詳細資訊")
                    event = st.dataframe(
                        filtered_df[display_cols], 
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
                        if st.button("✏️ 開啟修改模式"):
                            st.session_state.tab3_edit_requested = True
                            st.rerun()
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
                        filtered_df[display_cols],
                        hide_index=True,
                        disabled=["Sheet_Row"], 
                        use_container_width=True,
                        key="search_editor"
                    )
                    
                    if st.button("💾 儲存表格上的所有修改", type="primary"):
                        with st.spinner("正在比對修改並同步更新至雲端..."):
                            changed_rows = []
                            for i in range(len(edited_df)):
                                orig_row = filtered_df[display_cols].iloc[i]
                                new_row = edited_df.iloc[i]
                                
                                for col in display_cols:
                                    if str(orig_row[col]).strip() != str(new_row[col]).strip():
                                        changed_rows.append(new_row)
                                        break
                                        
                            if changed_rows:
                                headers = worksheet.row_values(1)
                                for row_data in changed_rows:
                                    sheet_idx = int(row_data['Sheet_Row'])
                                    for col_name in display_cols:
                                        if col_name != "Sheet_Row" and col_name in headers:
                                            col_idx = headers.index(col_name) + 1
                                            worksheet.update_cell(sheet_idx, col_idx, str(row_data[col_name]))
                                st.success(f"✅ 成功更新 {len(changed_rows)} 筆資料至雲端！")
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
            options = pending_df.apply(
                lambda x: f"列 {x['Sheet_Row']} | {x['廠別']} - {x['機台名稱']} (目前: {x['狀態']})", 
                axis=1
            ).tolist()
            
            upd_col1, upd_col2 = st.columns([2, 1])
            with upd_col1:
                selected_item = st.selectbox("選擇要更新的機台：", options)
            with upd_col2:
                new_status = st.selectbox("修改為新狀態：", ["未完成", "缺料", "已完成"], index=2)
                
            if st.button("送出狀態更新", type="primary", key="btn_update"):
                with st.spinner("同步至雲端中..."):
                    row_idx = int(selected_item.split(" | ")[0].replace("列 ", ""))
                    
                    headers = worksheet.row_values(1)
                    if "狀態" in headers and "Remark" in headers:
                        status_col_idx = headers.index("狀態") + 1
                        remark_col_idx = headers.index("Remark") + 1
                        
                        worksheet.update_cell(row_idx, status_col_idx, new_status)
                        
                        old_remark = worksheet.cell(row_idx, remark_col_idx).value or ""
                        today_str = datetime.now().strftime("%Y-%m-%d")
                        append_text = f"[{today_str} 更新狀態: {new_status}]"
                        
                        new_remark = (str(old_remark).strip() + "\n" + append_text).strip()
                        worksheet.update_cell(row_idx, remark_col_idx, new_remark)
                        
                        st.success(f"✅ 更新成功！該機台已標記為「{new_status}」。")
                        st.rerun()
                    else:
                        st.error("試算表中找不到「狀態」或「Remark」欄位，請檢查表頭名稱。")
        else:
            st.success("🎉 太棒了！目前所有機台皆已完工，沒有待追蹤項目。")