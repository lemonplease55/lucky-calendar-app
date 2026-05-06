# -*- coding: utf-8 -*-
import streamlit as st
import datetime
import pandas as pd
from io import BytesIO
import calendar
import sqlite3
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

# =========================
# 資料庫功能 (Backend Stats)
# =========================
DB_FILE = 'stats.db'

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS downloads 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                  timestamp DATETIME DEFAULT CURRENT_TIMESTAMP, 
                  filename TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS visits 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                  timestamp DATETIME DEFAULT CURRENT_TIMESTAMP)''')
    conn.commit()
    conn.close()

def log_download(filename):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("INSERT INTO downloads (filename) VALUES (?)", (filename,))
    conn.commit()
    conn.close()

def log_visit():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("INSERT INTO visits (timestamp) VALUES (CURRENT_TIMESTAMP)")
    conn.commit()
    conn.close()

def get_download_stats():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT timestamp, filename FROM downloads ORDER BY timestamp DESC", conn)
    conn.close()
    return df

def get_visit_stats():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT timestamp FROM visits ORDER BY timestamp DESC", conn)
    conn.close()
    return df

init_db()

# =========================
# Streamlit 設定
# =========================
st.set_page_config(page_title="樂覺製所生命靈數 | Numerology", layout="centered")

if 'has_visited' not in st.session_state:
    log_visit()
    st.session_state['has_visited'] = True

# =========================
# 公用數字處理
# =========================
def reduce_to_digit(n: int) -> int:
    while n > 9:
        n = sum(int(x) for x in str(n))
    return n

def sum_once(n: int) -> int:
    return sum(int(x) for x in str(n))

def format_layers(total: int) -> str:
    mid = sum_once(total)
    if mid > 9:
        return f"{total}/{mid}/{reduce_to_digit(mid)}"
    else:
        return f"{total}/{mid}"

# =========================
# 生命藍圖：階段數計算邏輯
# =========================
def calculate_blueprint_stages(birthday: datetime.date, hour: int, minute: int):
    # 拆解生日各項加總
    y_sum = sum(int(x) for x in str(birthday.year))
    m_sum = sum(int(x) for x in f"{birthday.month:02}")
    d_sum = sum(int(x) for x in f"{birthday.day:02}")
    h_sum = sum(int(x) for x in f"{hour:02}")
    min_sum = sum(int(x) for x in f"{minute:02}")

    # 累加邏輯 (遵循用戶指定)
    st_old = y_sum
    st_middle = st_old + m_sum
    st_young_adult = st_middle + d_sum
    st_teen = st_young_adult + h_sum
    st_child = st_teen + min_sum

    return [
        {"name": "老年階段", "age": "61歲以上", "val": format_layers(st_old)},
        {"name": "中年階段", "age": "41-60歲", "val": format_layers(st_middle)},
        {"name": "青年階段", "age": "21-40歲", "val": format_layers(st_young_adult)},
        {"name": "少年階段", "age": "11-20歲", "val": format_layers(st_teen)},
        {"name": "幼年階段", "age": "0-10歲", "val": format_layers(st_child)},
    ]

def calculate_life_path_number(birthday: datetime.date) -> tuple[int, int, str]:
    date_str = birthday.strftime("%Y%m%d")
    total_sum = sum(int(char) for char in date_str)
    final_num = reduce_to_digit(total_sum)
    process_str = f"{total_sum} → {final_num}"
    if total_sum != final_num and total_sum > 9:
        second_step = sum_once(total_sum)
        if second_step > 9 and second_step != final_num:
             process_str = f"{total_sum} → {second_step} → {final_num}"
    return final_num, total_sum, process_str

# =========================
# 流年與解說資料
# =========================
def life_year_number_for_date(birthday: datetime.date, query_date: datetime.date) -> int:
    cutoff = datetime.date(query_date.year, birthday.month, birthday.day)
    base_year = query_date.year - 1 if query_date < cutoff else query_date.year
    total = base_year + birthday.month + birthday.day
    return reduce_to_digit(sum_once(total))

def life_year_number_for_year(birthday: datetime.date, query_year: int) -> tuple[int, int]:
    before_total = (query_year - 1) + birthday.month + birthday.day
    after_total  = (query_year)     + birthday.month + birthday.day
    return reduce_to_digit(sum_once(before_total)), reduce_to_digit(sum_once(after_total))

def get_year_advice(n: int):
    advice = {
        1: ("自主與突破之年", "容易衝動、單打獨鬥", "設定清晰目標；決策前先蒐集意見。", "⭐⭐⭐⭐"),
        2: ("協作與關係之年", "過度迎合、忽略自我", "練習明確表達需求、建立健康邊界。", "⭐⭐⭐"),
        3: ("創意與表達之年", "分心、情緒起伏", "為創作與學習預留固定時段。", "⭐⭐⭐⭐"),
        4: ("穩定與基礎之年", "壓力感、僵化完美主義", "用『可持續的小步驟』築基礎。", "⭐⭐⭐"),
        5: ("變動與自由之年", "焦躁、衝動決策", "先設安全網再突破；用短衝測試。", "⭐⭐⭐⭐"),
        6: ("關懷與責任之年", "過度承擔、忽略自我", "把『照顧自己』寫進行程。", "⭐⭐⭐"),
        7: ("內省與學習之年", "孤立、鑽牛角尖", "安排獨處＋定期對談；冥想整理。", "⭐⭐⭐"),
        8: ("事業與財務之年", "過度追求成就、忽略健康", "設定績效與復原節奏並行。", "⭐⭐⭐⭐"),
        9: ("收尾與釋放之年", "抗拒結束、情緒回顧", "用感恩做結案；做斷捨離。", "⭐⭐⭐"),
    }
    return advice.get(n, ("年度主題", "—", "—", "⭐⭐⭐"))

lucky_map = {
    1: {"色": "🔴 紅色", "水晶": "紅瑪瑙", "小物": "原子筆"},
    2: {"色": "🟠 橙色", "水晶": "太陽石", "小物": "月亮吊飾"},
    3: {"色": "🟡 黃色", "水晶": "黃水晶", "小物": "紙膠帶"},
    4: {"色": "🟢 綠色", "水晶": "綠幽靈", "小物": "方形石頭"},
    5: {"色": "🔵 藍色", "水晶": "海藍寶", "小物": "交通票卡"},
    6: {"色": "🔷 靛色", "水晶": "青金石", "小物": "愛心吊飾"},
    7: {"色": "🟣 紫色", "水晶": "紫水晶", "小物": "書籤"},
    8: {"色": "💗 粉色", "水晶": "粉晶", "小物": "鋼筆"},
    9: {"色": "⚪ 白色", "水晶": "白水晶", "小物": "小香包"},
}

# [省略部分流日對照表以縮短代碼，請保留您原本的 flowing_day_guidance_map]
flowing_day_guidance_map = {"11/2": "與自己的內在靈性連結...", "12/3": "創意的想法..."} 

def get_flowing_day_guidance(fd_str): return flowing_day_guidance_map.get(fd_str, "")
def get_flowing_day_star(fd_str): return "🌟🌟🌟"

# =========================
# 匯出 Excel 功能 (維持不變)
# =========================
def style_excel(df: pd.DataFrame) -> BytesIO:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="流年月曆")
        workbook = writer.book
        worksheet = workbook["流年月曆"]
        # ... (樣式代碼)
    return output

# =========================
# Streamlit 介面 (核心變動區)
# =========================
st.title("🧭 樂覺製所生命靈數")
st.markdown("在數字之中，我們與自己不期而遇。\n**Be true, be you — 讓靈魂，自在呼吸。**")

# -------- 區塊 A：流年速算與階段藍圖 --------
st.subheader("🌟 生命靈數 & 流年速算 (Life Path & Yearly Flow)")

# 修改輸入區，加入時、分
col1, col2 = st.columns([1, 1])
with col1:
    birthday = st.date_input("請輸入生日 (Birthday)", value=datetime.date(1990, 1, 1))
    # 新增出生時分
    c_h, c_m = st.columns(2)
    with c_h:
        b_hour = st.number_input("出生時 (Hour)", 0, 23, 10)
    with c_m:
        b_min = st.number_input("出生分 (Min)", 0, 59, 0)
with col2:
    ref_date = st.date_input("查詢日期 (Query Date)", value=datetime.date.today())

if st.button("計算靈數與階段藍圖 (Calculate)"):
    # 1. 主命數
    life_num, life_sum, life_process = calculate_life_path_number(birthday)
    lucky_life = lucky_map.get(life_num, {})
    
    st.markdown("---")
    st.subheader(f"🔮 您的主命數：【 {life_num} 】號人")
    st.info(f"✨ **幸運色**：{lucky_life.get('色')} | **水晶**：{lucky_life.get('水晶')} | **小物**：{lucky_life.get('小物')}")

    # 2. 新增：生命藍圖五大階段 (關鍵新增)
    st.markdown("### 🗺️ 生命藍圖：五大階段數")
    stages = calculate_blueprint_stages(birthday, b_hour, b_min)
    
    # 用五個欄位顯示 (順序從幼年到老年，符合視覺邏輯)
    s_cols = st.columns(5)
    for i, s in enumerate(reversed(stages)):
        with s_cols[i]:
            st.markdown(f"**{s['name']}**")
            st.info(f"**{s['val']}**")
            st.caption(s['age'])
    st.markdown("---")

    # 3. 流年結果
    today_n = life_year_number_for_date(birthday, ref_date)
    before_n, after_n = life_year_number_for_year(birthday, ref_date.year)
    title, challenge, action, stars = get_year_advice(today_n)
    
    st.markdown(f"### 📊 流年結果：【 {today_n} 】")
    st.write(f"**主題**：{title} ({stars})")
    st.write(f"**建議行動**：{action}")
    
    with st.expander("查看流年詳細說明"):
        st.write(f"生日前流年：{before_n}")
        st.write(f"生日後流年：{after_n}")

# -------- 區塊 B：流年月曆產生器 (維持原本邏輯) --------
st.markdown("---")
st.subheader("📅 產生 1 個月份的『流年月曆』建議表")
target_month = st.selectbox("請選擇月份 (Select Month)", list(range(1, 13)), index=datetime.datetime.now().month - 1)

if st.button("🎉 產生日曆建議表 (Generate Excel)"):
    # ... (此處保留您原本產生日曆的程式碼)
    st.write("已產生報表。")

# =========================
# 後台管理區 (側邊欄)
# =========================
st.sidebar.markdown("---")
st.sidebar.subheader("🔒 管理員專區")
admin_password = st.sidebar.text_input("輸入密碼", type="password")
if admin_password == "admin123":
    st.sidebar.success("已登入")
    # ... (統計代碼)
