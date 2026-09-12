# -*- coding: utf-8 -*-
import streamlit as st
import datetime
import calendar
import pandas as pd
from io import BytesIO
from lunardate import LunarDate
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

# =========================
# 核心計算工具
# =========================
def reduce_to_digit(n):
    while n > 9:
        n = sum(int(c) for c in str(n))
    return n

def digit_sum(n):
    return sum(int(c) for c in str(n))

def format_layers(total):
    if total <= 9:
        return str(total)
    mid = digit_sum(total)
    if mid > 9:
        return f"{total}/{mid}/{reduce_to_digit(mid)}"
    return f"{total}/{mid}"

def digits_sum_str(s):
    return sum(int(c) for c in s)

# =========================
# 流年計算邏輯（修正版）
# 流年週期以「生日當天」為起點：
#   查詢日 >= 今年生日 → 流年年份 = 今年
#   查詢日 <  今年生日 → 流年年份 = 去年
# 計算：流年年份 + 生日月(02) + 生日日(02) → 逐位相加 → 化簡到個位數
# 例：生日 4/25，查詢日 2026/4/1
#   → 尚未過生日 → 流年年份 = 2025
#   → "20250425" → 2+0+2+5+0+4+2+5 = 20 → 2
# =========================
def _safe_cutoff(year, month, day):
    """處理 2/29 生日在非閏年的情況，退回 2/28"""
    try:
        return datetime.date(year, month, day)
    except ValueError:
        return datetime.date(year, month, day - 1)

def get_flow_year_base(birthday, query_date):
    cutoff = _safe_cutoff(query_date.year, birthday.month, birthday.day)
    return query_date.year if query_date >= cutoff else query_date.year - 1

def life_year_number(birthday, query_date):
    base_year = get_flow_year_base(birthday, query_date)
    total_str = f"{base_year}{birthday.month:02}{birthday.day:02}"
    return reduce_to_digit(digits_sum_str(total_str))

def flow_year_period(birthday, query_date):
    base_year = get_flow_year_base(birthday, query_date)
    next_year = base_year + 1
    start = _safe_cutoff(base_year, birthday.month, birthday.day)
    end   = _safe_cutoff(next_year, birthday.month, birthday.day) - datetime.timedelta(days=1)
    return start, end, base_year

# =========================
# 常量資料表
# =========================
YEAR_ADVICE = {
    1: ("自主與突破之年", "容易衝動、單打獨鬥", "設定清晰目標；蒐集意見、給自己緩衝時間。", "⭐⭐⭐⭐"),
    2: ("協作與關係之年", "過度迎合、忽略自我", "練習明確表達需求、建立健康邊界。", "⭐⭐⭐"),
    3: ("創意與表達之年", "分心、情緒起伏", "為創作與學習預留固定時段；公開練習。", "⭐⭐⭐⭐"),
    4: ("穩定與基礎之年", "壓力感、僵化完美主義", "用『可持續的小步驟』築基礎。", "⭐⭐⭐"),
    5: ("變動與自由之年", "焦躁、衝動決策", "先設安全網再突破；用短衝測試新方向。", "⭐⭐⭐⭐"),
    6: ("關懷與責任之年", "過度承擔、忽略自我", "把『照顧自己』寫進行程；清楚承諾。", "⭐⭐⭐"),
    7: ("內省與學習之年", "孤立、鑽牛角尖", "安排獨處＋定期對談；用寫作/冥想整理。", "⭐⭐⭐"),
    8: ("事業與財務之年", "過度追求成就、忽略健康情感", "設定績效與復原節奏並行；學會授權。", "⭐⭐⭐⭐"),
    9: ("收尾與釋放之年", "抗拒結束、情緒回顧", "用感恩做結案；做斷捨離，替新循環清出空間。", "⭐⭐⭐"),
}

LUCKY_MAP = {
    1: {"色": "🔴 紅色", "水晶": "紅瑪瑙、石榴石", "小物": "原子筆"},
    2: {"色": "🟠 橙色", "水晶": "太陽石、橙月光", "小物": "月亮吊飾"},
    3: {"色": "🟡 黃色", "水晶": "黃水晶、黃虎眼", "小物": "紙膠帶"},
    4: {"色": "🟢 綠色", "水晶": "綠幽靈、孔雀石", "小物": "方形石頭"},
    5: {"色": "🔵 藍色", "水晶": "海藍寶、藍紋瑪瑙", "小物": "交通票卡"},
    6: {"色": "🔷 靛色", "水晶": "青金石、蘇打石", "小物": "愛心吊飾"},
    7: {"色": "🟣 紫色", "水晶": "紫水晶", "小物": "書籤"},
    8: {"色": "💗 粉色", "水晶": "粉晶、草莓晶", "小物": "鋼筆"},
    9: {"色": "⚪ 白色", "水晶": "白水晶、白月光", "小物": "小香包"},
    0: {"色": "⚫️ 黑色", "水晶": "黑曜石", "小物": "護身符"},
}

UNKNOWN = "不知道"

# =========================
# 農曆換算
# =========================
HEAVENLY_STEMS   = "甲乙丙丁戊己庚辛壬癸"
EARTHLY_BRANCHES = "子丑寅卯辰巳午未申酉戌亥"
ZODIAC_ANIMALS   = "鼠牛虎兔龍蛇馬羊猴雞狗豬"
ZODIAC_EMOJIS    = ["🐭","🐂","🐯","🐰","🐲","🐍","🐴","🐑","🐵","🐔","🐶","🐷"]
LUNAR_MONTH_NAMES = ["正月","二月","三月","四月","五月","六月",
                     "七月","八月","九月","十月","冬月","臘月"]
LUNAR_DAY_NAMES = [
    "初一","初二","初三","初四","初五","初六","初七","初八","初九","初十",
    "十一","十二","十三","十四","十五","十六","十七","十八","十九","二十",
    "廿一","廿二","廿三","廿四","廿五","廿六","廿七","廿八","廿九","三十",
]

def solar_to_lunar(solar_date):
    ld = LunarDate.fromSolarDate(solar_date.year, solar_date.month, solar_date.day)
    idx    = (ld.year - 4) % 12
    stem   = HEAVENLY_STEMS[(ld.year - 4) % 10]
    branch = EARTHLY_BRANCHES[idx]
    month_label = ("閏" if ld.isLeapMonth else "") + LUNAR_MONTH_NAMES[ld.month - 1]
    return {
        "full":      f"{stem}{branch}年 {month_label}{LUNAR_DAY_NAMES[ld.day - 1]}",
        "ganzhi":    f"{stem}{branch}",
        "zodiac":    f"{ZODIAC_EMOJIS[idx]} {ZODIAC_ANIMALS[idx]}",
        "month_day": f"{month_label}{LUNAR_DAY_NAMES[ld.day - 1]}",
    }

# =========================
# 五階段藍圖
# =========================
def _current_stage_index(age):
    if age >= 61: return 0
    if age >= 41: return 1
    if age >= 21: return 2
    if age >= 11: return 3
    return 4

def calculate_blueprint_stages(birthday, hour_str, min_str, ref_date):
    y_sum     = digit_sum(birthday.year)
    month_sum = digit_sum(int(f"{birthday.month:02}"))
    day_sum   = digit_sum(int(f"{birthday.day:02}"))

    stage_old   = y_sum
    stage_mid   = stage_old + month_sum
    stage_young = stage_mid + day_sum

    if hour_str == UNKNOWN:
        teen_display  = "--"
        child_display = "--"
    else:
        hour_sum     = digit_sum(int(f"{int(hour_str):02}"))
        teen_display = format_layers(stage_young + hour_sum)
        if min_str == UNKNOWN:
            child_display = "--"
        else:
            minute_sum    = digit_sum(int(f"{int(min_str):02}"))
            child_display = format_layers(stage_young + hour_sum + minute_sum)

    age = ref_date.year - birthday.year - (
        (ref_date.month, ref_date.day) < (birthday.month, birthday.day)
    )
    active = _current_stage_index(age)

    return [
        ("老年階段", "61 歲以上",  format_layers(stage_old),   active == 0),
        ("中年階段", "41 – 60 歲", format_layers(stage_mid),   active == 1),
        ("青年階段", "21 – 40 歲", format_layers(stage_young), active == 2),
        ("少年階段", "11 – 20 歲", teen_display,               active == 3),
        ("幼年階段", "0 – 10 歲",  child_display,              active == 4),
    ]

# =========================
# 流日指引 & 星等
# =========================
FLOWING_DAY_GUIDANCE = {
    "11/2":    "與自己的內在靈性連結，打開心眼從心去看清楚背後的真相。",
    "12/3":    "創意的想法和能量正在湧現，用純粹且動聽的方式傳遞出來。",
    "13/4":    "讓想法不再只是想像，是時候設法落實到自己的現實生活中。",
    "14/5":    "轉化現有的狀態，從固有和凝滯的工作、關係中解脫。",
    "15/6":    "會特別渴望與某人深入交談、分享心事。",
    "16/7":    "整理內在與學習的好時機，感到精神渙散時，需要讓自己靜下來。",
    "17/8":    "會特別想處理與金錢、服務或管理相關的問題。",
    "18/9":    "在新階段來臨之前，先學會放下、告別與結束。",
    "19/10/1": "會發現自己比平時更容易接收到來自內在或外在的靈感。",
    "20/2":    "內在外在都將迎來翻轉式的改變，洞見更加清晰的真相。",
    "21/3":    "今天點子和想法會比平常要多，好好運用溝通和表達來創造。",
    "22/4":    "多任務、多變動的一天。保持耐心與行動力。",
    "23/5":    "是時候接收新的刺激和變動，考驗自己是否有足夠勇氣。",
    "24/6":    "關心自己身邊親近的家人朋友，承諾與責任是今天的主題。",
    "25/7":    "專注在自己的事情上，在這當中找回內在的平靜與和諧感。",
    "26/8":    "強化自信與擔當，適合接下責任、處理財務、設定下一步策略。",
    "27/9":    "透過真理看見真相，有意識地放下是今天的重點。",
    "28/10/1": "有強大顯化力與執行力的日子。保持務實、負責的態度。",
    "29/11/2": "透過傾聽和觀察，從更高智慧層次解讀事情。",
    "30/3":    "今天的主題是溝通與協調，運用創意來做包裝和行銷。",
    "31/4":    "創造中蘊含結構，靈感需要被規劃來落地。",
    "32/5":    "保持靈活和彈性，敞開心釋放和接收愛，有機會突破。",
    "33/6":    "用創意、好玩的方式去服務和關愛，釋放壓抑。",
    "34/7":    "今日會想獨處反思，注意情緒管控。",
    "35/8":    "推進與擴張的日子，結合創意與商業頭腦。",
    "36/9":    "在理想與現實之間取得平衡點，透過服務與奉獻幫助他人。",
    "37/10/1": "適時站出來為自己發聲，勇敢展現和展開新的行動。",
    "38/11/2": "運用累積的經驗協助夥伴家人，用風趣方式點出問題。",
    "39/12/3": "聲音和語言具有大能量，用話語去讚美自己和他人。",
    "40/4":    "以穩固為前提，更新現有的框架，建立新結構。",
    "41/5":    "穩定中尋求自由。突破常規，在變動中保持平衡。",
    "42/6":    "規矩紀律需與人際關係並重，考量感性層面。",
    "43/7":    "有強大的組織和分析能力，留意情緒控管與說話方式。",
    "44/8":    "具強大執行力與影響力，避免固執而忽略他人聲音。",
    "45/9":    "運用理性邏輯深入省思，成就自身智慧。",
    "46/10/1": "成為帶動者，展現組織合作能力，聚焦目標。",
    "47/11/2": "扮演穩定可靠的關鍵角色，在重要時刻協助他人。",
    "48/12/3": "在審慎評估下，做出富有創意的決策。",
    "49/13/4": "在穩定基礎下做出取捨，提升到更高境界。",
    "50/5":    "變動中隱藏機會，享受這美好的時刻。",
    "51/6":    "勇敢面對恐懼和創傷，與自己和解。",
    "52/7":    "從核心切入剖析，看見真相。適合獨處深思。",
    "53/8":    "有機會創造財富或經驗，保持開放。",
    "54/9":    "從漫無目的收斂聚焦，放下並感謝過往。",
    "55/10/1": "極度外放和自我展現，留意是否冒犯。保持專注。",
    "56/11/2": "跳脫二元對立的思維模式，平衡自由與承諾。",
    "57/12/3": "留意內在直覺，答案都在那裡。",
    "58/13/4": "在變動中整合出新流程和規則。",
    "59/14/5": "富有挑戰性的一天，過去所學將迎來轉化。",
}

FLOWING_DAY_STARS = {
    "11/2":"🌟🌟","12/3":"🌟🌟🌟🌟","13/4":"🌟🌟🌟🌟","14/5":"🌟🌟",
    "15/6":"🌟🌟🌟🌟","16/7":"🌟🌟🌟","17/8":"🌟🌟🌟🌟🌟","18/9":"🌟🌟",
    "19/10/1":"🌟🌟🌟🌟","20/2":"🌟🌟🌟","21/3":"🌟🌟🌟🌟","22/4":"🌟🌟🌟",
    "23/5":"🌟🌟🌟🌟","24/6":"🌟🌟🌟","25/7":"🌟🌟","26/8":"🌟🌟🌟🌟🌟",
    "27/9":"🌟🌟🌟","28/10/1":"🌟🌟🌟🌟🌟","29/11/2":"🌟🌟🌟","30/3":"🌟🌟🌟🌟",
    "31/4":"🌟🌟🌟🌟","32/5":"🌟🌟🌟🌟","33/6":"🌟🌟🌟","34/7":"🌟🌟",
    "35/8":"🌟🌟🌟🌟🌟","36/9":"🌟🌟🌟🌟","37/10/1":"🌟🌟🌟🌟🌟","38/11/2":"🌟🌟🌟",
    "39/12/3":"🌟🌟🌟🌟","40/4":"🌟🌟🌟","41/5":"🌟🌟🌟🌟","42/6":"🌟🌟🌟",
    "43/7":"🌟🌟🌟","44/8":"🌟🌟🌟🌟","45/9":"🌟🌟🌟","46/10/1":"🌟🌟🌟🌟",
    "47/11/2":"🌟🌟🌟","48/12/3":"🌟🌟🌟🌟","49/13/4":"🌟🌟🌟","50/5":"🌟🌟🌟🌟",
    "51/6":"🌟🌟","52/7":"🌟🌟🌟","53/8":"🌟🌟🌟🌟","54/9":"🌟🌟",
    "55/10/1":"🌟🌟🌟","56/11/2":"🌟🌟","57/12/3":"🌟🌟🌟🌟","58/13/4":"🌟🌟🌟",
    "59/14/5":"🌟🌟🌟🌟🌟",
}

def get_flowing_year_ref(query_date, bday):
    d = query_date.date() if hasattr(query_date, "date") else query_date
    cutoff = _safe_cutoff(d.year, bday.month, bday.day)
    return d.year if d >= cutoff else d.year - 1

def get_flowing_month_ref(query_date, birthday):
    d = query_date.date() if hasattr(query_date, "date") else query_date
    return (d.month - 1 if d.month > 1 else 12) if d.day < birthday.day else d.month

# =========================
# 匯出 Excel 樣式
# =========================
def style_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="流年月曆")
        ws = writer.book["流年月曆"]
        hf    = Font(size=12, bold=True, color="FFFFFF")
        hfill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        hal   = Alignment(horizontal="center", vertical="center")
        for idx, col in enumerate(df.columns):
            max_len = max((len(str(c)) for c in df[col]), default=15)
            ws.column_dimensions[chr(65 + idx)].width = max(15, min(int(max_len * 1.2), 100))
        for cell in ws[1]:
            cell.font = hf; cell.fill = hfill; cell.alignment = hal
        thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                      top=Side(style='thin'),  bottom=Side(style='thin'))
        for row in ws.iter_rows():
            for cell in row:
                cell.border = thin
                cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[row[0].row].height = 35
    return output

# =========================
# Streamlit 介面
# =========================
st.set_page_config(page_title="樂覺製所生命靈數", layout="centered")
st.title("🧭 樂覺製所生命靈數")
st.markdown("在數字之中，我們與自己不期而遇。")

# ── 輸入區 ──
st.subheader("🌟 生命靈數 & 階段藍圖速算")
col_in1, col_in2 = st.columns(2)

with col_in1:
    birthday = st.date_input(
        "請輸入生日 (Birthday)",
        value=datetime.date(1990, 1, 1),
        min_value=datetime.date(1900, 1, 1),
    )
    st.markdown("**出生時間 (Time)**")
    t_c1, t_c2 = st.columns(2)
    hour_options = [UNKNOWN] + [str(i) for i in range(24)]
    min_options  = [UNKNOWN] + [str(i) for i in range(60)]
    with t_c1:
        birth_hour = st.selectbox("時 (Hour)", options=hour_options, index=11)
    with t_c2:
        birth_min = st.selectbox("分 (Min)", options=min_options, index=1)

with col_in2:
    # ✅ 預設值為今天
    ref_date = st.date_input("查詢日期 (Query Date)", value=datetime.date.today())

if st.button("🔮 開始計算"):
    st.markdown("---")

    # ── 農曆換算 ──
    st.markdown("### 🌙 農曆換算 (Lunar Calendar)")
    lunar_bday  = solar_to_lunar(birthday)
    lunar_query = solar_to_lunar(ref_date)

    lc1, lc2 = st.columns(2)
    with lc1:
        st.markdown(
            f"<div style='background:#FFF8E7;padding:18px 20px;border-radius:12px;"
            f"border-left:5px solid #E6A817;'>"
            f"<p style='margin:0 0 6px;color:#888;font-size:0.85em;'>生日 — "
            f"{birthday.strftime('%Y/%m/%d')}</p>"
            f"<p style='margin:0;font-size:1.25em;font-weight:bold;color:#5A3E00;'>"
            f"{lunar_bday['full']}</p>"
            f"<p style='margin:6px 0 0;font-size:0.95em;color:#7A5C00;'>"
            f"生肖 {lunar_bday['zodiac']}</p></div>",
            unsafe_allow_html=True,
        )
    with lc2:
        st.markdown(
            f"<div style='background:#EEF4FF;padding:18px 20px;border-radius:12px;"
            f"border-left:5px solid #3366CC;'>"
            f"<p style='margin:0 0 6px;color:#888;font-size:0.85em;'>查詢日 — "
            f"{ref_date.strftime('%Y/%m/%d')}</p>"
            f"<p style='margin:0;font-size:1.25em;font-weight:bold;color:#1A337E;'>"
            f"{lunar_query['full']}</p>"
            f"<p style='margin:6px 0 0;font-size:0.95em;color:#2952A3;'>"
            f"生肖 {lunar_query['zodiac']}</p></div>",
            unsafe_allow_html=True,
        )

    st.markdown("---")

    # ── 五階段藍圖 ──
    st.markdown("### 🗺️ 生命藍圖五大階段 (Life Blueprint Stages)")
    stages = calculate_blueprint_stages(birthday, birth_hour, birth_min, ref_date)
    s_cols = st.columns(5)

    ACTIVE_STYLE = (
        "background-color:#1A337E;color:white;"
        "padding:25px 10px;border-radius:15px;text-align:center;margin:15px 0;"
    )
    INACTIVE_STYLE = (
        "background-color:#E8F0FE;color:#1A337E;"
        "padding:25px 10px;border-radius:15px;text-align:center;margin:15px 0;"
    )

    for col, (name, age_range, val, is_active) in zip(s_cols, stages):
        with col:
            label = f"{name} ◀ 目前" if is_active else name
            style = ACTIVE_STYLE if is_active else INACTIVE_STYLE
            st.markdown(
                f"<p style='text-align:center;font-weight:bold;margin-bottom:-10px;'>{label}</p>",
                unsafe_allow_html=True,
            )
            st.markdown(
                f"<div style='{style}'>"
                f"<span style='font-size:24px;font-weight:bold;'>{val}</span></div>",
                unsafe_allow_html=True,
            )
            st.markdown(
                f"<p style='text-align:center;color:gray;font-size:0.8em;'>{age_range}</p>",
                unsafe_allow_html=True,
            )

    st.markdown("---")

    # ── 流年結果 ──
    year_num = life_year_number(birthday, ref_date)
    title, challenge, action, stars = YEAR_ADVICE.get(
        year_num, ("年度主題", "—", "—", "⭐⭐⭐")
    )
    lucky = LUCKY_MAP.get(year_num, {})
    start_d, end_d, base_year = flow_year_period(birthday, ref_date)

    st.markdown(f"### 📊 查詢日期流年數：【 {year_num} 】")
    st.caption(f"目前流年週期：{start_d.strftime('%Y/%m/%d')} ～ {end_d.strftime('%Y/%m/%d')}")
    st.markdown(f"**年度主題 (Theme)**：{title}\n\n**運勢指數**：{stars}")
    st.markdown(f"**挑戰 (Challenge)**：{challenge}")
    st.markdown(f"**建議行動 (Action)**：{action}")
    if lucky:
        st.info(
            f"✨ **幸運色**：{lucky['色']} ｜ "
            f"**水晶**：{lucky['水晶']} ｜ "
            f"**小物**：{lucky['小物']}"
        )

# =========================
# 區塊 B：流年月曆 Excel 下載
# =========================
st.markdown("---")
st.subheader("📅 產生 1 個月份的『流年月曆』建議表 (Generate Monthly Calendar)")

target_month = st.selectbox(
    "請選擇月份 (Select Month)",
    list(range(1, 13)),
    index=datetime.datetime.now().month - 1
)

if st.button("🎉 產生日曆建議表 (Generate Excel)"):
    target_year = ref_date.year
    _, last_day = calendar.monthrange(target_year, target_month)
    days = pd.date_range(
        start=datetime.date(target_year, target_month, 1),
        end=datetime.date(target_year, target_month, last_day)
    )
    data = []
    for d in days:
        # 流日
        fd_total    = digits_sum_str(f"{birthday.year}{birthday.month:02}{d.day:02}")
        flowing_day = format_layers(fd_total)
        main_num    = reduce_to_digit(fd_total)
        lucky_d     = LUCKY_MAP.get(main_num, {})
        guidance    = FLOWING_DAY_GUIDANCE.get(flowing_day, "")
        day_stars   = FLOWING_DAY_STARS.get(flowing_day, "🌟🌟🌟")

        # 流年
        yr_ref      = get_flowing_year_ref(d, birthday)
        fy_total    = digits_sum_str(f"{yr_ref}{birthday.month:02}{birthday.day:02}")
        flowing_year = format_layers(fy_total)

        # 流月
        fm_ref      = get_flowing_month_ref(d, birthday)
        fm_total    = digits_sum_str(f"{birthday.year}{fm_ref:02}{birthday.day:02}")
        flowing_month = format_layers(fm_total)

        data.append({
            "日期 (Date)":      d.strftime("%Y-%m-%d"),
            "星期 (Day)":       d.strftime("%A"),
            "流年 (Year Num)":  flowing_year,
            "流月 (Month Num)": flowing_month,
            "流日 (Day Num)":   flowing_day,
            "運勢指數 (Stars)": day_stars,
            "指引 (Guidance)":  guidance,
            "幸運色 (Color)":   lucky_d.get("色", ""),
            "水晶 (Crystal)":   lucky_d.get("水晶", ""),
            "幸運小物 (Item)":  lucky_d.get("小物", ""),
        })

    df = pd.DataFrame(data)
    st.dataframe(df, use_container_width=True)

    file_name = f"LuckyCalendar_{target_year}_{str(target_month).zfill(2)}.xlsx"
    if not df.empty:
        output = style_excel(df)
        st.download_button(
            label="📥 點此下載 Excel (Download)",
            data=output.getvalue(),
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.warning("⚠️ 無法匯出 Excel：目前資料為空")

st.markdown("---")
st.caption("樂覺製所 © 2026 Numbertalk")
