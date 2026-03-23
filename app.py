import streamlit as st
import pandas as pd
import io
from functions import parse_transcript, extract_text_from_uploaded_file

def clear_cache():
    """當上傳檔案變更時，清除先前的分析結果"""
    if 'analysis_results' in st.session_state:
        st.session_state['analysis_results'] = None
    if 'df_result' in st.session_state:
        st.session_state['df_result'] = None

# --- 介面設定 ---
st.set_page_config(page_title="立法院會議發言統計小幫手", layout="wide")
st.title("🏛️ 立法委員發言自動統計系統")
st.markdown("上傳 **委員名單 Excel** 與 **會議紀錄（txt/doc/docx）**，自動生成統計報表。")

# --- 側邊欄：上傳與設定 ---
with st.sidebar:
    st.header("1. 設定委員名單")
    
    # 下載範本按鈕
    template_df = pd.DataFrame({
        '委員會': ['內政', '外交國防', '經濟', '財政'],
        '黨籍': ['民進黨', '國民黨', '民眾黨', '無黨籍'],
        '姓名': ['OOO', 'OOO', 'OOO', 'OOO']
    })
    template_bytes = io.BytesIO()
    with pd.ExcelWriter(template_bytes, engine='xlsxwriter') as writer:
        template_df.to_excel(writer, sheet_name='委員名單', index=False)
    
    st.download_button(
        label="📋 下載範本名單（Excel）",
        data=template_bytes.getvalue(),
        file_name="委員名單_範本.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="下載包含『委員會』『黨籍』『姓名』三欄的空白範本，可直接修改後使用"
    )
    
    # 加入 on_change 機制：當檔案變更時，觸發 clear_cache 清除舊資料
    label_file = st.file_uploader("上傳委員名單 (xlsx)", type=['xlsx'], on_change=clear_cache)
    
    st.header("2. 載入逐字稿檔案")
    uploaded_files = st.file_uploader(
        "上傳逐字稿 (txt/doc/docx，可複選)",
        type=['txt', 'doc', 'docx'],
        accept_multiple_files=True,
        on_change=clear_cache
    )
        

# --- 主流程 ---
if label_file and uploaded_files:
    # 1. 讀取委員名單
    labels = {}
    try:
        df_config = pd.read_excel(label_file)

        # 欄位標準化：支援「姓名/政黨」與「委員會/黨籍/姓名」格式
        rename_map = {}
        if '黨籍' in df_config.columns and '政黨' not in df_config.columns:
            rename_map['黨籍'] = '政黨'
        if rename_map:
            df_config = df_config.rename(columns=rename_map)

        required_columns = {'姓名', '政黨'}
        if not required_columns.issubset(df_config.columns):
            raise ValueError("名單格式需包含『姓名』與『政黨』欄位；可選填『委員會』欄位。")

        # 清理空值並保留原始順序
        df_config = df_config.dropna(subset=['姓名', '政黨']).copy()
        df_config['姓名'] = df_config['姓名'].astype(str).str.strip()
        df_config['政黨'] = df_config['政黨'].astype(str).str.strip()
        if '委員會' in df_config.columns:
            df_config['委員會'] = df_config['委員會'].fillna('').astype(str).str.strip()

        # 可選擇委員會過濾
        filtered_config = df_config
        if '委員會' in df_config.columns:
            committees = [c for c in df_config['委員會'].unique().tolist() if c]
            selected_committees = st.sidebar.multiselect(
                "3. 篩選委員會（可複選）",
                options=committees,
                default=committees,
                on_change=clear_cache
            )
            if selected_committees:
                filtered_config = df_config[df_config['委員會'].isin(selected_committees)]
            else:
                filtered_config = df_config.iloc[0:0]

        labels = dict(zip(filtered_config['姓名'], filtered_config['政黨']))

        if not labels:
            st.warning("⚠️ 目前篩選條件下沒有可分析的委員，請調整委員會篩選。")
            st.stop()

        st.success(f"✅ 已載入 {len(labels)} 位委員名單")
        
        # 預覽一下名單給使用者看
        with st.expander("查看目前的委員名單"):
            st.dataframe(filtered_config)
            
    except Exception as e:
        st.error(f"讀取名單失敗: {e}")
        st.stop()

    # 初始化 Session State 用來存放結果
    if 'analysis_results' not in st.session_state:
        st.session_state['analysis_results'] = None
    if 'df_result' not in st.session_state:
        st.session_state['df_result'] = None

    # [修正] 建立一個固定的佔位區塊，避免進度條出現與消失造成下方元件位移而重置狀態
    status_holder = st.empty()

    # 2. 分析按鈕（僅負責計算同時存檔）
    if st.button("🚀 開始分析"):
        # 使用固定的 holder 來顯示進度條
        progress_bar = status_holder.progress(0)
        
        # 暫存
        processed_data = []
        all_row_data = []
        
        for i, uploaded_file in enumerate(uploaded_files):
            # 讀取並自動轉換文字內容（支援 txt/doc/docx 與多編碼）
            try:
                file_bytes = uploaded_file.getvalue()
                text, convert_note = extract_text_from_uploaded_file(uploaded_file.name, file_bytes)
            except ValueError as e:
                st.error(f"❌ 檔案轉換失敗: {uploaded_file.name}\n\n{str(e)}")
                st.stop()
            except Exception as e:
                st.error(f"❌ 處理檔案失敗: {uploaded_file.name}\n\n錯誤: {str(e)}")
                st.stop()
            
            # 解析
            info, stats = parse_transcript(text, labels)

            # 將顯示所需的資料存起來，而不是直接 st.write
            processed_data.append({
                'filename': uploaded_file.name,
                'preview': text, # 可選全文或前1000字供預覽[:1000]
                'info': info,
                'stats': stats,
                'index': i,
                'convert_note': convert_note
            })

            # 彙整 DataFrame 資料
            row_data = {
                ("基本資訊", "日期", ""): info['date'],
                ("基本資訊", "會議名稱", ""): info['name'],
                ("基本資訊", "主席", ""): info['chairman']
            }
            for name, party in labels.items():
                row_data[(party, name, "次數")] = stats.get(name, {}).get('count', 0)
                row_data[(party, name, "字數")] = stats.get(name, {}).get('words', 0)
            
            all_row_data.append(row_data)
            
            # 更新進度條
            progress_bar.progress((i + 1) / len(uploaded_files))
            
        # 分析完成後清空進度條，但空間還在 (因為 holder 是固定的)
        status_holder.empty()

        if not all_row_data:
            st.warning("⚠️ 未取得可輸出的分析資料，請檢查上傳檔案內容。")
            st.stop()

        # 轉成 DataFrame 並排序
        column_tuples = list(all_row_data[0].keys())
        df = pd.DataFrame(all_row_data, columns=column_tuples)
        df.columns = pd.MultiIndex.from_tuples(column_tuples)
        
        # [修正] 改用依照名單順序的排序邏輯
        name_index_map = {name: i for i, name in enumerate(labels.keys())}
        
        def list_based_sort_key(col_tuple):
            category, name, metric = col_tuple
            
            # 1. 基本資訊放最前面
            if category == "基本資訊":
            # 定義基本資訊的固定順序：日期 -> 會議名稱 -> 主席
                basic_info_order = {"日期": 0, "會議名稱": 1, "主席": 2}
                info_rank = basic_info_order.get(name, 999)
                return (-1, info_rank, str(metric))
            
            # 2. 依照 Excel 名單順序 (主要排序依據)
            idx = name_index_map.get(name, 9999)
            
            # 3. 同一個人的數據，次數先、字數後
            m_rank = 0 if metric == "次數" else 1
            
            return (0, idx, m_rank)

        # 應用排序
        sorted_cols = sorted(df.columns, key=list_based_sort_key)
        df = df[sorted_cols]

        # [關鍵步驟] 將計算結果存入 Session State
        st.session_state['analysis_results'] = processed_data
        st.session_state['df_result'] = df
        
        # 執行完畢後不需要在這裡顯示，讓下面的區塊負責顯示

        # 3. 顯示區塊 (只要 Session State 有資料就會顯示，不受按鈕影響)
    if st.session_state['analysis_results'] is not None and st.session_state['df_result'] is not None:
        results = st.session_state['analysis_results']
        df = st.session_state['df_result']

        for res in results:
            st.text(f"📄 逐字稿檔案：{res['filename']}")
            if res.get('convert_note'):
                st.caption(f"🔄 {res['convert_note']}")
            with st.expander("查看逐字稿", expanded=False):
                # 使用 container 固定高度並加入滑塊 (Scrollbar)
                with st.container(height=300):
                    st.code(res['preview'], language='text', line_numbers=True)
            
            info = res['info']
            stats = res['stats']
            i = res['index']

            # 詳細發言紀錄 (這裡的操作不會導致資料消失了)
            with st.expander(f"📝 查看 {info.get('name', '此會議')} 詳細發言紀錄"):
                # [修正] 依照 Excel 名單順序 (labels.keys()) 篩選有發言的委員
                active_speakers = [name for name in labels.keys() if name in stats and stats[name]['count'] > 0]
                chairman_name = info.get('chairman', '')
                
                if not active_speakers:
                    st.warning("⚠️ 此檔案似乎未偵測到任何委員發言。")
                else:
                    # 製作選單顯示名稱 (如果是主席就加上標記)
                    speaker_options = {
                        name: f"{name} {'(主席🔨)' if name == chairman_name else ''}" 
                        for name in active_speakers
                    } # 建立一個 map: 真實姓名 -> 顯示名稱

                    # key 必須設為獨一無二，這裡加上 index 區分不同檔案
                    selected_speaker_key = st.selectbox(
                        f"選擇委員 (共 {len(active_speakers)} 位有發言)", 
                        options=active_speakers, 
                        format_func=lambda x: speaker_options[x], # 使用 format_func 來自定義顯示
                        key=f"select_speaker_{i}"
                    )
                    
                    if selected_speaker_key:
                        selected_speaker = selected_speaker_key
                        data = stats[selected_speaker]
                        
                        st.markdown(f"**{selected_speaker} ({labels[selected_speaker]})** 發言`{data['count']}`次，共`{data['words']}`字")

                        # 使用 container 固定高度並加入滑塊，避免單一委員發言過多佔版面
                        with st.container(height=300):
                            for log in data['logs']:
                                st.markdown(log)
                                st.caption("-" * 30)
            st.divider()
        
        st.subheader("📊 分析結果預覽")
        st.dataframe(df)

        # 4. 下載按鈕
        # 將 DataFrame 寫入 Excel Bytes
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='統計結果')
        
        st.download_button(
            label="📥 下載 Excel 報表",
            data=output.getvalue(),
            file_name="legislator_stats.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif not label_file:
    st.info("👈 請先在左側上傳委員名單 (Excel 檔，需包含『姓名』『政黨』；可選『委員會』欄位)")