import streamlit as st
import random
from docx import Document
from docx.shared import RGBColor

# --- 核心邏輯區 ---

def parse_docx_red_font(uploaded_files):
    """
    解析 Word 檔，抓取紅色字體 (RGB 255,0,0) 作為答案。
    """
    all_questions = []
    
    for uploaded_file in uploaded_files:
        try:
            doc = Document(uploaded_file)
            
            for para in doc.paragraphs:
                text = para.text.strip()
                if not text:
                    continue # 跳過空行
                
                has_answer = False
                answer_parts = []
                
                # 遍歷段落中的每一個樣式片段 (Run)
                for run in para.runs:
                    # 檢查是否有設定字體顏色
                    font_color = run.font.color
                    
                    # 判斷是否為紅色 (RGB: 255, 0, 0)
                    # 注意：Word 的「標準紅色」通常就是這個數值
                    if font_color and font_color.rgb == RGBColor(255, 0, 0):
                        has_answer = True
                        answer_parts.append(run.text.strip())
                
                # 只有當這個段落有文字時才視為題目
                # (不管有沒有找到紅色答案，都先把題目收錄進來，如果是題目但沒標紅，會顯示無標記)
                if text:
                    all_questions.append({
                        "full_text": text,                # 完整題目文字
                        "has_answer": has_answer,         # 是否有偵測到答案
                        "answer_text": " ".join(answer_parts) if has_answer else "未偵測到紅色標記"
                    })
        except Exception as e:
            st.error(f"讀取檔案 {uploaded_file.name} 時發生錯誤: {e}")
            
    return all_questions

# --- 介面設定區 (UI) ---

st.set_page_config(page_title="紅色字體模考系統", page_icon="📝", layout="centered")

st.title("📝 自動模考產生器 (紅色字體版)")
st.markdown("""
**使用說明：**
1. 準備您的 Word 題庫。
2. 將 **正確答案的文字** 顏色設為 **「標準紅色」**。
3. 上傳檔案，系統會自動隨機抽出 25 題。
""")

# 初始化 Session State (用來記憶題目狀態，避免按按鈕時網頁重整導致題目跑掉)
if 'quiz_data' not in st.session_state:
    st.session_state.quiz_data = []
if 'show_answers' not in st.session_state:
    st.session_state.show_answers = {}
if 'exam_started' not in st.session_state:
    st.session_state.exam_started = False

# 1. 檔案上傳
uploaded_files = st.file_uploader("📂 請上傳 Word 檔案 (.docx)", type=['docx'], accept_multiple_files=True)

# 2. 生成按鈕與邏輯
if uploaded_files:
    if st.button("🚀 開始生成模考"):
        with st.spinner('正在讀取檔案並尋找紅色答案...'):
            all_questions = parse_docx_red_font(uploaded_files)
            
            if not all_questions:
                st.error("未讀取到任何文字，請確認檔案內容。")
            else:
                # 隨機抽題邏輯
                if len(all_questions) < 25:
                    st.warning(f"⚠️ 題庫總數僅有 {len(all_questions)} 題 (不足 25 題)，將全數顯示。")
                    st.session_state.quiz_data = all_questions
                else:
                    st.session_state.quiz_data = random.sample(all_questions, 25)
                    st.success(f"✅ 已從 {len(all_questions)} 題中隨機選出 25 題！")
                
                # 重置狀態
                st.session_state.show_answers = {i: False for i in range(len(st.session_state.quiz_data))}
                st.session_state.exam_started = True
                st.rerun()

# 3. 顯示考卷區
if st.session_state.exam_started and st.session_state.quiz_data:
    st.markdown("---")
    
    # 進度條或計數
    total_q = len(st.session_state.quiz_data)
    
    for idx, q in enumerate(st.session_state.quiz_data):
        # 每一題的區塊
        with st.container():
            st.markdown(f"#### 第 {idx + 1} 題 / 共 {total_q} 題")
            
            # 顯示題目內容 (預設顯示完整黑字，雖然答案是紅的，但為了模擬考試，這裡統一顯示黑色給使用者看比較好？
            # 為了方便，我們直接顯示文字。如果在 Word 裡原本就是紅字，Streamlit 預設會顯示成普通黑色文字，除非特別標註 Markdown 顏色。
            # 所以使用者看到題目時，是看不到顏色的，剛好符合考試需求！)
            st.info(q['full_text'])
            
            # 操作按鈕區
            col_btn, col_ans = st.columns([1, 4])
            
            with col_btn:
                # 每個按鈕需要獨立的 key
                if st.button(f"👁️ 看答案", key=f"btn_ans_{idx}"):
                    st.session_state.show_answers[idx] = not st.session_state.show_answers.get(idx, False)
                    # 這裡不使用 rerun，利用 Streamlit 的即時反應特性
            
            # 顯示答案邏輯
            if st.session_state.show_answers.get(idx, False):
                with col_ans:
                    if q['has_answer']:
                        st.success(f"**正確答案：** {q['answer_text']}")
                    else:
                        st.warning("⚠️ 此題未偵測到紅色字體")
            
            st.markdown("---")

    # 底部重置按鈕
    if st.button("🔄 重新抽題 / 上傳新檔"):
        st.session_state.quiz_data = []
        st.session_state.exam_started = False
        st.session_state.show_answers = {}
        st.rerun()

elif not uploaded_files:
    st.info("👋 請先在上方上傳檔案以開始使用。")