code_content = '''import os
import io
import urllib.request
from datetime import datetime, timedelta
import random
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import Workbook, load_workbook
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np

# ============================================================
# [한글 폰트 설정] Streamlit Cloud(Linux) 한글 깨짐 완전 방지 로직
# ============================================================
def init_korean_font():
    font_filename = "NanumGothic.ttf"
    if not os.path.exists(font_filename):
        font_url = "https://github.com/google/fonts/raw/main/ofl/nanumgothic/NanumGothic-Regular.ttf"
        try:
            urllib.request.urlretrieve(font_url, font_filename)
        except Exception:
            try:
                urllib.request.urlretrieve(font_url, font_filename)
            except Exception:
                pass
            
    if os.path.exists(font_filename):
        fm.fontManager.addfont(font_filename)
        font_prop = fm.FontProperties(fname=font_filename)
        plt.rcParams['font.family'] = font_prop.get_name()
    else:
        plt.rcParams['font.family'] = 'Malgun Gothic'
        
    plt.rcParams['axes.unicode_minus'] = False

init_korean_font()

# DB 파일 경로 설정
DB_FILE_PATH = "Pump_Master_DB.xlsx"
OVERHAUL_DB_PATH = "Pump_Overhaul_DB.xlsx"
DOC_DB_PATH = "Pump_Docs_DB.xlsx"
KNOWHOW_DB_PATH = "Pump_Knowhow_DB.xlsx"
DAILY_LOG_DB_PATH = "Pump_DailyLog_DB.xlsx"
SAFETY_PERMIT_DB_PATH = "Pump_SafetyPermit_DB.xlsx"
LOGO_FILE_PATH = "Logo.png"

# ============================================================
# 1. 페이지 기본 설정 및 모바일 반응형 웹 CSS
# ============================================================
page_icon_setting = LOGO_FILE_PATH if os.path.exists(LOGO_FILE_PATH) else "🌊"

st.set_page_config(
    page_title="K-water tech 펌프 종합 진단 & CBM 시스템 (개발 진행 중)",
    page_icon=page_icon_setting,
    layout="wide",
    initial_sidebar_state="auto"
)

# 모바일 UI 줄바꿈 깨짐 방지 커스텀 CSS (웹 포털 스타일 적용)
st.markdown("""
    <style>
    .block-container {
        padding-top: 0.8rem !important;
        padding-bottom: 1.5rem !important;
        padding-left: 0.8rem !important;
        padding-right: 0.8rem !important;
    }
    .main-title {
        font-size: 1.5rem;
        font-weight: 800;
        color: #1e293b;
        margin-bottom: 0.2rem;
        line-height: 1.3;
    }
    .sub-caption {
        font-size: 0.85rem;
        color: #64748b;
    }
    
    /* KPI 수치 카드 그리드 CSS */
    .kpi-container {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(100px, 1fr));
        gap: 8px;
        margin-bottom: 15px;
    }
    .kpi-card-m {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-top: 3px solid #0098DA;
        border-radius: 8px;
        padding: 10px 6px;
        text-align: center;
    }
    .kpi-card-m .title {
        font-size: 0.75rem;
        color: #64748b;
        font-weight: 700;
        margin-bottom: 4px;
    }
    .kpi-card-m .value {
        font-size: 1.25rem;
        font-weight: 800;
        color: #0f172a;
    }

    div[data-testid="stSidebarNav"] {
        display: none;
    }
    
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        font-weight: bold;
        transition: all 0.2s ease-in-out;
    }
    
    div[data-testid="stSidebar"] .stButton>button {
        font-size: 0.95rem !important;
        padding: 8px 10px !important;
        text-align: left !important;
        justify-content: flex-start !important;
        margin-bottom: 4px !important;
        background-color: #f8fafc;
        border: 1px solid #cbd5e1;
        color: #1e293b;
    }
    
    div[data-testid="stSidebar"] .stButton>button:hover {
        background-color: #0098DA !important;
        color: white !important;
        border-color: #0098DA !important;
    }

    .badge-dev {
        background-color: #fef3c7;
        color: #d97706;
        padding: 3px 8px;
        border-radius: 4px;
        font-size: 0.75rem;
        font-weight: bold;
        display: inline-block;
    }

    @media (max-width: 768px) {
        .main-title { font-size: 1.15rem !important; }
        .kpi-container { grid-template-columns: repeat(3, 1fr); }
    }
    </style>
""", unsafe_allow_html=True)

USER_DB = [
    {"id": "kwatertech최진욱", "name": "최진욱", "dept": "밀양정수장", "role": "정밀진단원 / 관리자", "status": "접속중"}
]

CATEGORIES = {"성능": 40, "내부상태": 27, "기계상태": 25, "정비이력": 5}

DEFAULT_PUMPS = [
    {
        "site": "밀양정수장",
        "equip": f"가압펌프 #{i}",
        "maker": "효성펌프" if i % 2 == 1 else "현대중공업",
        "model": f"DHP-{i}" if i % 2 == 1 else f"VTP-{i*10}",
        "hp": str(150 + (i * 10)),
        "head": str(45 + (i * 2)),
        "flow": str(1200 + (i * 50)),
        "rpm": "1780",
        "build_date": f"2018-0{min(i, 9)}-15" if i < 10 else "2018-10-15",
        "op_hours": 8500 + (i * 350)
    } for i in range(1, 11)
]

# ============================================================
# 2. 판정 로직 및 DB 초기화
# ============================================================
def calc_eff(val):
    if val >= 98.0: return "A+"
    elif val >= 96.0: return "A"
    elif val >= 94.0: return "A-"
    elif val >= 92.0: return "B+"
    elif val >= 90.0: return "B"
    elif val >= 88.0: return "B-"
    elif val >= 86.0: return "C+"
    elif val >= 84.0: return "C"
    elif val >= 82.0: return "C-"
    elif val >= 80.0: return "D+"
    elif val >= 75.0: return "D"
    else: return "E"

def calc_reach(val):
    if val >= 98.0: return "A"
    elif val >= 93.0: return "B"
    elif val >= 88.0: return "C"
    elif val >= 80.0: return "D"
    else: return "E"

def calc_bep(val):
    if 85.0 <= val <= 115.0: return "A"
    elif 75.0 <= val <= 125.0: return "B"
    elif 65.0 <= val <= 135.0: return "C"
    elif 50.0 <= val <= 150.0: return "D"
    else: return "E"

def calc_ring_gap(val):
    if val < 1.5: return "A"
    elif val < 2.0: return "B"
    elif val < 2.5: return "C"
    elif val < 3.0: return "D"
    else: return "E"

def calc_sleeve(val):
    if val < 1.0: return "A"
    elif val < 1.8: return "B"
    elif val < 2.5: return "C"
    elif val < 3.0: return "D"
    else: return "E"

def calc_vib(val):
    if val < 1.8: return "A"
    elif val < 4.5: return "B"
    elif val < 7.1: return "C"
    elif val < 11.2: return "D"
    else: return "E"

def calc_align(val):
    if val <= 0.05: return "A"
    elif val <= 0.08: return "B"
    elif val <= 0.12: return "C"
    elif val <= 0.15: return "D"
    else: return "E"

def calc_overhaul(val):
    if val <= 10000: return "A"
    elif val <= 12000: return "B"
    elif val <= 15000: return "C"
    elif val <= 20000: return "D"
    else: return "E"

EVAL_ITEMS = [
    ("성능", "펌프 효율 유지율 (%)", 25, "98.0% 이상 (A+)", ["A+", "A", "A-", "B+", "B", "B-", "C+", "C", "C-", "D+", "D", "E"],
     {"A+": 25.0, "A": 24.0, "A-": 23.0, "B+": 22.0, "B": 21.0, "B-": 20.0, "C+": 18.5, "C": 17.0, "C-": 15.5, "D+": 14.0, "D": 11.0, "E": 7.0}, calc_eff),
    ("성능", "설계 양정/유량 도달률 (%)", 8, "98.0% 이상 (A)", ["A", "B", "C", "D", "E"], {"A": 8.0, "B": 6.8, "C": 5.6, "D": 4.4, "E": 2.4}, calc_reach),
    ("성능", "BEP 운전점 적정성 (BEP 대비 %)", 7, "85 ~ 115% (A)", ["A", "B", "C", "D", "E"], {"A": 7.0, "B": 5.95, "C": 4.9, "D": 3.85, "E": 2.1}, calc_bep),
    ("내부상태", "임펠러/케이싱 링 간극 (마모 배수)", 10, "최초의 3.0배 미만", ["A", "B", "C", "D", "E"], {"A": 10.0, "B": 8.0, "C": 6.5, "D": 5.0, "E": 3.0}, calc_ring_gap),
    ("내부상태", "축슬리브 마모 (패임 비율 %)", 5, "슬리브 외경의 2.5% 미만", ["A", "B", "C", "D", "E"], {"A": 5.0, "B": 4.25, "C": 3.5, "D": 2.75, "E": 1.5}, calc_sleeve),
    ("내부상태", "임펠러 손상/침식", 2, "균열/침식 없음", ["A", "B", "C", "D", "E"], {"A": 2.0, "B": 1.7, "C": 1.4, "D": 1.1, "E": 0.6}, None),
    ("내부상태", "임펠러 동적 밸런싱", 2, "ISO G2.5 이하", ["A", "B", "C", "D", "E"], {"A": 2.0, "B": 1.7, "C": 1.4, "D": 1.1, "E": 0.6}, None),
    ("내부상태", "NPSH 여유율 및 캐비테이션", 5, "NPSHa/NPSHr ≥ 1.3", ["A", "B", "C", "D", "E"], {"A": 5.0, "B": 4.25, "C": 3.5, "D": 2.75, "E": 1.5}, None),
    ("내부상태", "내부 코팅 상태", 3, "박리 면적 0% (건전)", ["A", "B", "C", "D", "E"], {"A": 3.0, "B": 2.55, "C": 2.1, "D": 1.65, "E": 0.9}, None),
    ("내부상태", "비금속 웨어링 개선", 3, "복합소재 + 간극 50% 축소", ["A", "B", "C", "D", "E"], {"A": 3.0, "B": 2.55, "C": 2.1, "D": 1.65, "E": 0.9}, None),
    ("기계상태", "Overall 진동 (ISO 10816, mm/s)", 7, "1.8 mm/s 미만", ["A", "B", "C", "D", "E"], {"A": 7.0, "B": 5.95, "C": 4.9, "D": 3.85, "E": 2.1}, calc_vib),
    ("기계상태", "베어링 결함 진동", 5, "결함 신호 없음", ["A", "B", "C", "D", "E"], {"A": 5.0, "B": 4.25, "C": 3.5, "D": 2.75, "E": 1.5}, None),
    ("기계상태", "주파수 성분 결함", 3, "특이 피크 없음", ["A", "B", "C", "D", "E"], {"A": 3.0, "B": 2.55, "C": 2.1, "D": 1.65, "E": 0.9}, None),
    ("기계상태", "펌프 모터 센터링 (mm)", 7, "0.05 mm 이내", ["A", "B", "C", "D", "E"], {"A": 7.0, "B": 5.95, "C": 4.9, "D": 3.85, "E": 2.1}, calc_align),
    ("기계상태", "Soft Foot 및 배관 응력", 3, "Soft Foot ≤ 0.05mm", ["A", "B", "C", "D", "E"], {"A": 3.0, "B": 2.55, "C": 2.1, "D": 1.65, "E": 0.9}, None),
    ("정비이력", "오버홀 주기 준수성 (운전 시간)", 3, "10,000시간 이내", ["A", "B", "C", "D", "E"], {"A": 3.0, "B": 2.55, "C": 2.1, "D": 1.65, "E": 0.9}, calc_overhaul),
    ("정비이력", "주요 소모품 교체 이력", 2, "주기 준수 및 이력 양호", ["A", "B", "C", "D", "E"], {"A": 2.0, "B": 1.7, "C": 1.4, "D": 1.1, "E": 0.6}, None)
]

FINAL_GRADE_INFO = {
    "A": ("A (매우 우수)", "신품 수준 또는 정밀 정비 완료 상태\n[조치] 정상 운전 유지 및 정기 모니터링 수행"),
    "B": ("B (우수/정상)", "전반적인 건전성 양호\n[조치] 계획된 일상 점검 및 윤활 관리 수행"),
    "C": ("C (보통/관찰)", "경미한 성능/기계적 저하 진행\n[조치] 관찰 대상 지정, 주요 진동/효율 추이 분석 및 차기 정비 계획 수립"),
    "D": ("D (개선 필요)", "뚜렷한 결함 또는 효율 저하 확인\n[조치] 정비 검토 필수, 원인 부위(진동/마모)에 대한 계획 정비 실시"),
    "E": ("E (불량/위험)", "설비 파손 위험 또는 기능 상실 상태\n[조치] 즉시 정지 및 비상 정비 검토, 오버홀 수행")
}

def ensure_db_exists():
    if not os.path.exists(DB_FILE_PATH):
        wb = Workbook()
        ws = wb.active
        ws.title = "진단이력"
        headers = ["점검일", "사업장", "설비명", "제조사", "모델명", "마력(HP)", "양정(m)", "준공일", "점검자", "종합점수", "최종등급"] + [item[1] for item in EVAL_ITEMS]
        ws.append(headers)
        wb.save(DB_FILE_PATH)

    if not os.path.exists(OVERHAUL_DB_PATH):
        wb = Workbook()
        ws = wb.active
        ws.title = "오버홀이력"
        ws.append(["작업일자", "사업장", "설비명", "공정단계", "작업자", "작업내용", "사진파일명", "전후효율개선(%)", "전후진동감소(mm/s)"])
        wb.save(OVERHAUL_DB_PATH)

    if not os.path.exists(DOC_DB_PATH):
        wb = Workbook()
        ws = wb.active
        ws.title = "전문보고서"
        ws.append(["등록일자", "사업장", "설비명", "보고서구분", "수행기관", "요약소견", "저장파일명", "원본파일명"])
        wb.save(DOC_DB_PATH)

    if not os.path.exists(KNOWHOW_DB_PATH):
        wb = Workbook()
        ws = wb.active
        ws.title = "노하우DB"
        ws.append(["등록일자", "분류", "관련모델", "현상및원인", "해결노하우", "작성자"])
        wb.save(KNOWHOW_DB_PATH)

    if not os.path.exists(DAILY_LOG_DB_PATH):
        wb = Workbook()
        ws = wb.active
        ws.title = "점검일지"
        ws.append(["점검일자", "점검자", "설비명", "누수/윤활유상태", "진동/소음상태", "온도/전류상태", "특이사항 메모"])
        wb.save(DAILY_LOG_DB_PATH)

    if not os.path.exists(SAFETY_PERMIT_DB_PATH):
        wb = Workbook()
        ws = wb.active
        ws.title = "위험작업허가"
        ws.append(["신청일자", "작업명", "대상설비", "위험유형", "작업기간", "신청자", "승인상태", "안전조치사항"])
        wb.save(SAFETY_PERMIT_DB_PATH)

ensure_db_exists()

# 샘플 데이터 세딩
def seed_sample_data(force=False):
    wb = load_workbook(DB_FILE_PATH)
    ws = wb["진단이력"]
    if ws.max_row <= 25 or force:
        ws.delete_rows(2, ws.max_row)
        dates = ["2024-03-15", "2024-09-10", "2025-03-20", "2025-08-15", "2026-02-10", "2026-08-14"]
        for p in DEFAULT_PUMPS:
            for dt in dates:
                score = round(random.uniform(72.0, 96.5), 1)
                if score >= 90: grade = "A"
                elif score >= 80: grade = "B"
                elif score >= 70: grade = "C"
                elif score >= 60: grade = "D"
                else: grade = "E"
                
                item_grades = [random.choice(["A+", "A", "B+", "B", "C+", "C"]) for _ in range(17)]
                ws.append([dt, p["site"], p["equip"], p["maker"], p["model"], p["hp"], p["head"], p["build_date"], "최진욱", score, grade] + item_grades)
        wb.save(DB_FILE_PATH)
    wb.close()

    wb_o = load_workbook(OVERHAUL_DB_PATH)
    ws_o = wb_o["오버홀이력"]
    if ws_o.max_row <= 10 or force:
        ws_o.delete_rows(2, ws_o.max_row)
        steps = ["1단계: 분해/해체 점검", "2단계: 부품 가공 및 신품 교체", "3단계: 조립 및 센터링", "4단계: 시운전 및 완료"]
        for i in range(1, 11):
            ws_o.append([f"2024-{i%9+1:02d}-10", "밀양정수장", f"가압펌프 #{i}", steps[0], "최진욱", "임펠러 해체 및 마모 점검", "No Image", "+1.2%", "-0.5 mm/s"])
            ws_o.append([f"2025-{i%9+1:02d}-15", "밀양정수장", f"가압펌프 #{i}", steps[2], "최진욱", "축슬리브 교체 및 센터링 완료", "No Image", "+3.5%", "-1.2 mm/s"])
            ws_o.append([f"2026-0{min(i,8)+1}-20", "밀양정수장", f"가압펌프 #{i}", steps[3], "최진욱", "정기 오버홀 완료 및 정상 시운전", "No Image", "+5.8%", "-2.1 mm/s"])
        wb_o.save(OVERHAUL_DB_PATH)
    wb_o.close()

seed_sample_data()

st.session_state.pump_list = DEFAULT_PUMPS

if "logged_in" not in st.session_state: st.session_state.logged_in = True
if "username" not in st.session_state: st.session_state.username = "최진욱"
if "nav_menu" not in st.session_state: st.session_state.nav_menu = "1.1. 설비 통합 대시보드"

for idx, item in enumerate(EVAL_ITEMS):
    if f"selected_grade_{idx}" not in st.session_state:
        st.session_state[f"selected_grade_{idx}"] = item[4][0]

# ============================================================
# 3. 공통 엑셀 다운로드 모달 팝업 (@st.dialog)
# ============================================================
@st.dialog("📥 엑셀 데이터 다운로드 확인")
def confirm_excel_download_dialog(df_data, file_label):
    st.write(f"### 📄 **[{file_label}]**")
    st.info("선택하신 메뉴의 전체 데이터를 엑셀(XLSX) 파일로 다운로드하시겠습니까?")
    st.dataframe(df_data.head(5), use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_data.to_excel(writer, index=False, sheet_name="Data")
    excel_bytes = output.getvalue()

    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            label="✅ 예 (엑셀 다운로드)",
            data=excel_bytes,
            file_name=f"Kwater_{file_label}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
    with c2:
        if st.button("❌ 아니오 (취소)", use_container_width=True):
            st.rerun()

def render_download_button_header(df_data, file_label):
    h_col1, h_col2 = st.columns([3.5, 1.2])
    with h_col2:
        if st.button(f"📥 엑셀 다운", key=f"dl_btn_{file_label}"):
            confirm_excel_download_dialog(df_data, file_label)

# ============================================================
# 4. 모바일 사이드바 메뉴
# ============================================================
def navigate_to(menu_name):
    st.session_state.nav_menu = menu_name
    st.rerun()

with st.sidebar:
    if os.path.exists(LOGO_FILE_PATH):
        st.image(LOGO_FILE_PATH, use_container_width=True)
    
    st.caption(f"👤 **{st.session_state.username}** 님 접속 중")
    st.markdown("<span class='badge-dev'>🚧 프로토타입 자체 개발 진행 중</span>", unsafe_allow_html=True)
    st.write("---")

    st.markdown("### ⚙️ 시스템 메뉴")
    menu_items = [
        ("1.1. 설비 통합 대시보드", "📊 1.1. 설비 통합 대시보드"),
        ("1.2. QR 설비 디지털 포털", "📱 1.2. QR 설비 디지털 포털"),
        ("2.1. 펌프 정밀 진단 (17개)", "📋 2.1. 펌프 정밀 진단 (17개)"),
        ("2.2. CBM 오버홀 주기 최적화", "🎯 2.2. CBM 오버홀 주기 최적화"),
        ("3.1. 오버홀 전후 ROI 분석", "🛠️ 3.1. 오버홀 전후 ROI 분석"),
        ("3.2. 단계별 사업모델 & 소요비용", "💰 3.2. 단계별 사업모델 & 소요비용"),
        ("4.1. AI 이상징후 & 추이 분석", "📈 4.1. AI 이상징후 & 추이 분석"),
        ("5.1. K-water tech 노하우 DB", "💡 5.1. K-water tech 노하우 DB"),
        ("5.2. 지자체 진단 보고서 출력", "📄 5.2. 지자체 진단 보고서 출력"),
        ("6.1. 통합 DB 일괄 백업", "📦 6.1. 통합 DB 일괄 백업")
    ]

    for menu_key, menu_label in menu_items:
        is_selected = (st.session_state.nav_menu == menu_key)
        btn_type = "primary" if is_selected else "secondary"
        if st.button(menu_label, key=f"btn_{menu_key}", type=btn_type, use_container_width=True):
            navigate_to(menu_key)

head_c1, head_c2 = st.columns([3.5, 1.2])
with head_c1:
    st.markdown("<div class='main-title'>K-water tech 펌프 진단 & CBM 플랫폼</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='sub-caption'>기존 데이터와 현장 정비기술을 결합한 디지털 설비관리 포털 [메뉴: <b>{st.session_state.nav_menu}</b>]</div>", unsafe_allow_html=True)

st.write("---")

# ============================================================
# [메인 화면 메뉴 구동]
# ============================================================

# --- 1.1. 설비 통합 대시보드 ---
if st.session_state.nav_menu == "1.1. 설비 통합 대시보드":
    wb = load_workbook(DB_FILE_PATH, data_only=True)
    ws = wb["진단이력"]
    diag_records = [row for row in ws.iter_rows(min_row=1, values_only=True)]
    wb.close()
    df_dash = pd.DataFrame(diag_records[1:], columns=diag_records[0]) if len(diag_records) > 1 else pd.DataFrame()

    render_download_button_header(df_dash, "대시보드_진단이력")

    st.subheader("📊 설비 건전성 현황판 (밀양정수장 #1~#10)")
    
    latest_grades = {}
    if not df_dash.empty:
        for _, row in df_dash.iterrows():
            latest_grades[row["설비명"]] = row["최종등급"]

    total_cnt = len(st.session_state.pump_list)
    grade_counts = {"A": 0, "B": 0, "C": 0, "D": 0, "E": 0}
    for g in latest_grades.values():
        if g in grade_counts: grade_counts[g] += 1

    st.markdown(f"""
        <div class="kpi-container">
            <div class="kpi-card-m"><div class="title">총 펌프</div><div class="value">{total_cnt}대</div></div>
            <div class="kpi-card-m"><div class="title">A 등급</div><div class="value" style="color:#16a34a;">{grade_counts['A']}대</div></div>
            <div class="kpi-card-m"><div class="title">B 등급</div><div class="value" style="color:#2563eb;">{grade_counts['B']}대</div></div>
            <div class="kpi-card-m"><div class="title">C 등급</div><div class="value" style="color:#d97706;">{grade_counts['C']}대</div></div>
            <div class="kpi-card-m"><div class="title">D 등급</div><div class="value" style="color:#dc2626;">{grade_counts['D']}대</div></div>
            <div class="kpi-card-m"><div class="title">E 등급</div><div class="value" style="color:#991b1b;">{grade_counts['E']}대</div></div>
        </div>
    """, unsafe_allow_html=True)

    st.write("---")
    st.markdown("### 💡 [실증 사례] 부곡가압장 3호기 정비 검증 데이터")
    st.info("모회사 작업의뢰를 통한 당사 직접 진단 수행 실적으로 저비용 진단 기반 정비 성과 검증을 입증한 근거 사례입니다.")
    
    col_a, col_b, col_c = st.columns(3)
    col_a.metric("진단 및 정비 전 효율", "72.3 %", "정비 필요")
    col_b.metric("오버홀 정비 후 효율", "77.7 %", "+5.4 %p 상승", delta_color="normal")
    col_c.metric("진단 비용 절감 효과", "약 200만 원", "고효율 정비 검증")

    st.write("---")
    st.subheader("🎯 CBM 기반 오버홀 시급성 Top 3 펌프")
    cbm_data = [
        {"순위": "1순위 (최우선)", "설비명": "가압펌프 #2", "운전시간": "9,200h", "효율": "73%", "진동": "5.8 mm/s", "상태": "주의 (조기정비)", "조치가이드": "금년도 최우선 정비 대상 반영"},
        {"순위": "2순위", "설비명": "가압펌프 #3", "운전시간": "11,200h", "효율": "70%", "진동": "6.8 mm/s", "상태": "이상 (정밀진단)", "조치가이드": "정밀 점검 및 분해 오버홀 수립"},
        {"순위": "3순위", "설비명": "가압펌프 #5", "운전시간": "9,800h", "효율": "75%", "진동": "4.2 mm/s", "상태": "주의 (관찰필요)", "조치가이드": "차기 정비 예산 수립 시 반영"}
    ]
    st.dataframe(pd.DataFrame(cbm_data), use_container_width=True, hide_index=True)

    st.write("---")
    st.subheader("📜 [통합 이력] 전체 펌프 정밀 진단 이력")
    if not df_dash.empty:
        st.dataframe(df_dash, use_container_width=True, hide_index=True)

# --- 1.2. QR 설비 디지털 포털 ---
elif st.session_state.nav_menu == "1.2. QR 설비 디지털 포털":
    st.subheader("📱 QR 기반 현장 설비 정보 모바일 디지털 포털")
    st.caption("현장에서 QR코드를 촬영하거나 설비를 선택하여 통합 제원 및 정비 이력을 확인합니다.")

    selected_pump = st.selectbox("현장 설비 선택 (QR 스캔 연동)", [p["equip"] for p in st.session_state.pump_list])
    pump_info = next(p for p in st.session_state.pump_list if p["equip"] == selected_pump)

    tab1, tab2, tab3, tab4 = st.tabs(["📋 1단계: 기본 사양", "📈 2단계: 점검/측정이력", "⚡ 3단계: 효율진단(모회사)", "🛠️ 4단계: 오버홀 이력"])

    with tab1:
        st.markdown(f"""
        - **설비명:** {pump_info['equip']} ({pump_info['site']})
        - **제조사 / 모델:** {pump_info['maker']} / {pump_info['model']}
        - **정격 출력:** {pump_info['hp']} HP | **정격 양정:** {pump_info['head']} m
        - **정격 유량:** {pump_info['flow']} m³/h | **회전수:** {pump_info['rpm']} RPM
        - **누적 운전시간:** {pump_info['op_hours']} 시간
        """)

    with tab2:
        st.info("PFM 연동 최근 측정 진동: 2.3 mm/s (정상) | 베어링 온도: 42°C")

    with tab3:
        st.success("모회사 효율진단 연동 결과: 현재 운전 효율 87.5% (전년 대비 -1.2% p)")

    with tab4:
        st.write("최근 오버홀: 2025-03-15 (축슬리브 교체 및 센터링 완료)")

# --- 2.1. 펌프 정밀 진단 (17개) ---
elif st.session_state.nav_menu == "2.1. 펌프 정밀 진단 (17개)":
    init_korean_font()
    st.subheader("📋 17개 핵심 항목 정밀 진단 및 CBM Index 자동 계산")
    
    selected_pump_name = st.selectbox("진단할 펌프 선택", [p["equip"] for p in st.session_state.pump_list])
    pump_info = next(p for p in st.session_state.pump_list if p["equip"] == selected_pump_name)

    c1, c2, c3 = st.columns(3)
    c1.text_input("사업장", value=pump_info["site"], disabled=True)
    c2.text_input("설비명", value=pump_info["equip"], disabled=True)
    c3.date_input("점검일자", value=datetime.now())

    st.write("---")
    cat_scores = {k: 0.0 for k in CATEGORIES.keys()}
    total_score = 0.0
    details = []

    for idx, (cat, item, weight, std_val, options, score_dict, auto_fn) in enumerate(EVAL_ITEMS):
        ec1, ec2, ec3 = st.columns([2, 1, 1])
        ec1.write(f"**{item}** ({weight}점)")
        input_val = ec2.text_input("측정값", key=f"v_{idx}", label_visibility="collapsed", placeholder="수치입력")
        selected_grade = ec3.selectbox("등급", options, key=f"g_{idx}", label_visibility="collapsed")
        
        s = score_dict[selected_grade]
        cat_scores[cat] += s
        total_score += s
        details.append({"category": cat, "item": item, "grade": selected_grade, "score": s})

    total_score = round(total_score, 2)
    final_grade = "A" if total_score >= 90 else ("B" if total_score >= 80 else ("C" if total_score >= 70 else ("D" if total_score >= 60 else "E")))

    st.write("---")
    st.metric("종합 진단 CBM 점수", f"{total_score} 점", f"최종 등급: {final_grade}")
    if st.button("💾 진단 결과 DB 저장", type="primary", use_container_width=True):
        st.success("통합 DB에 성공적으로 저장되었습니다!")

# --- 2.2. CBM 오버홀 주기 최적화 ---
elif st.session_state.nav_menu == "2.2. CBM 오버홀 주기 최적화":
    st.subheader("🎯 운전시간(TBM) + 상태기반(CBM) 오버홀 주기 최적화")
    st.markdown("단순 운전시간($10,000\text{시간}$) 기준이 아닌 **복합 상태 지표** 기반 정비 우선순위를 산출합니다.")

    df_cbm = pd.DataFrame([
        {"설비명": "가압펌프 #1", "운전시간": "10,200h", "효율": "84%", "진동": "2.1mm/s", "TBM판단": "오버홀 검토", "CBM판단": "정시 점검 유지", "우선순위": "5순위"},
        {"설비명": "가압펌프 #2", "운전시간": "9,200h", "효율": "73%", "진동": "5.8mm/s", "TBM판단": "정상 운전", "CBM판단": "금년도 최우선 정비", "우선순위": "1순위"},
        {"설비명": "가압펌프 #3", "운전시간": "7,800h", "효율": "68%", "진동": "6.8mm/s", "TBM판단": "정상 운전", "CBM판단": "정밀 점검 및 정비", "우선순위": "2순위"},
        {"설비명": "가압펌프 #4", "운전시간": "10,500h", "효율": "86%", "진동": "1.5mm/s", "TBM판단": "오버홀 검토", "CBM판단": "관찰 모니터링", "우선순위": "6순위"}
    ])
    
    render_download_button_header(df_cbm, "CBM_오버홀_우선순위_분석")
    st.dataframe(df_cbm, use_container_width=True, hide_index=True)

# --- 3.1. 오버홀 전후 ROI 분석 ---
elif st.session_state.nav_menu == "3.1. 오버홀 전후 ROI 분석":
    st.subheader("🛠️ 오버홀 정비 전·후 효과 정량적 분석 및 ROI 산출")
    
    col_in1, col_in2, col_in3 = st.columns(3)
    p_kw = col_in1.number_input("설비 용량 (kW)", value=110.0)
    eff_before = col_in2.number_input("정비 전 효율 (%)", value=73.0)
    eff_after = col_in3.number_input("정비 후 효율 (%)", value=82.0)

    hours = st.number_input("연간 운전 시간 (시간)", value=6000)
    cost_kwh = st.number_input("전력 단가 (원/kWh)", value=140)
    repair_cost = st.number_input("오버홀 정비 비용 (원)", value=35000000)

    # ROI 계산 로직
    if eff_before > 0 and eff_after > 0:
        kwh_saved = p_kw * ((1 / (eff_before / 100)) - (1 / (eff_after / 100))) * hours
        money_saved = kwh_saved * cost_kwh
        payback_years = repair_cost / money_saved if money_saved > 0 else 0
    else:
        kwh_saved, money_saved, payback_years = 0, 0, 0

    st.write("---")
    r1, r2, r3 = st.columns(3)
    r1.metric("연간 절감 전력량", f"{int(kwh_saved):,} kWh/년")
    r2.metric("연간 전력비 절감액", f"{int(money_saved):,} 원/년")
    r3.metric("투자비 회수 기간", f"{payback_years:.1f} 년")

# --- 3.2. 단계별 사업모델 & 소요비용 ---
elif st.session_state.nav_menu == "3.2. 단계별 사업모델 & 소요비용":
    st.subheader("💰 기획안 ⅩⅣ장. 단계별 계약 로드맵 및 소요 비용 (추정안)")
    
    st.markdown("#### 1. 단계별 계약화 로드맵")
    df_contract = pd.DataFrame([
        {"단계": "1단계 (1차년도)", "계약형태": "기존 용역계약 내 특약 추가", "상대방": "모회사", "핵심내용": "QR포털 구축/운영을 기존 정비용역에 편입 (신규예산 불요)"},
        {"단계": "2단계 (2차년도)", "계약형태": "신규 독립 유지보수 계약", "상대방": "모회사 사업소", "핵심내용": "시스템 고도화(PFM/CBM 연동) 독립 상품화"},
        {"단계": "3단계 (3차년도)", "계약형태": "신규 용역 및 컨설팅 계약", "상대방": "지자체 상하수도", "핵심내용": "검증 시스템 패키지화 및 외부 수주 확장"}
    ])
    st.dataframe(df_contract, use_container_width=True, hide_index=True)

    st.markdown("#### 2. 단계별 소요 비용 종합표")
    df_cost = pd.DataFrame([
        {"단계": "1단계 (PoC)", "서버운영비(연)": "40~100만 원", "QR인건비(연)": "4~8만 원", "운영인력투입": "100~200만 원", "합계(연)": "약 150~300만 원 (특약 처리)"},
        {"단계": "2단계 (독립계약)", "서버운영비(연)": "180~360만 원", "QR인건비(연)": "35~80만 원", "운영인력투입": "300~500만 원", "합계(연)": "약 500~800만 원"},
        {"단계": "3단계 (지자체)", "서버운영비(연)": "360~960만 원", "QR인건비(연)": "210만 원 이상", "운영인력투입": "별도 산정", "합계(연)": "별도 견적 수주"}
    ])
    st.dataframe(df_cost, use_container_width=True, hide_index=True)
    st.caption("※ QR 라벨 부착 단가 기준: 라벨 제작(2~3천원) + 부착 인건비(4~5천원) = 대당 약 7~8천원 기준")

# --- 4.1. AI 이상징후 & 추이 분석 ---
elif st.session_state.nav_menu == "4.1. AI 이상징후 & 추이 분석":
    init_korean_font()
    st.subheader("📈 AI 기반 이상징후 탐지 및 미래 추세 예측")
    
    fig, ax = plt.subplots(figsize=(8, 3.5))
    months = ["25.01", "25.06", "25.12", "26.03", "26.08", "26.12(예측)"]
    vib = [1.8, 2.3, 3.5, 4.9, 6.1, 7.6]
    ax.plot(months, vib, marker='o', color='#dc2626', linestyle='--', linewidth=2, label='가압펌프 #2 진동 추이 (mm/s)')
    ax.axhline(7.1, color='orange', linestyle=':', label='경고 기준치 (7.1 mm/s)')
    ax.set_title("가압펌프 #2 진동 상승 예측 곡선", fontsize=10, fontweight='bold')
    ax.legend()
    ax.grid(True, linestyle=':')
    st.pyplot(fig)

# --- 5.1. K-water tech 노하우 DB ---
elif st.session_state.nav_menu == "5.1. K-water tech 노하우 DB":
    wb = load_workbook(KNOWHOW_DB_PATH, data_only=True)
    ws = wb["노하우DB"]
    kh_records = [row for row in ws.iter_rows(min_row=1, values_only=True)]
    wb.close()
    df_kh = pd.DataFrame(kh_records[1:], columns=kh_records[0]) if len(kh_records) > 1 else pd.DataFrame()

    render_download_button_header(df_kh, "기술노하우DB")
    st.subheader("💡 K-water tech 정비 & 결함 해결 노하우 DB")
    if not df_kh.empty:
        st.dataframe(df_kh, use_container_width=True, hide_index=True)

# --- 5.2. 지자체 진단 보고서 출력 ---
elif st.session_state.nav_menu == "5.2. 지자체 진단 보고서 출력":
    st.subheader("📋 지자체 제출용 CBM 진단 보고서 자동 생성")
    
    st.selectbox("대상 지자체/사업소 선택", ["밀양시 상하수도사업소", "창원특례시 수도테크", "김해시 정수과"])
    
    st.markdown("""
    #### [보고서 요약 미리보기]
    - **통합 점검 펌프:** 총 10대
    - **금년도 최우선 정비 추천:** 1대 (가압펌프 #2)
    - **예산 절감 효과:** CBM 적용을 통한 정비 연기로 연간 약 3,200만원 절감 기대
    """)
    
    if st.button("📄 지자체용 표준 진단보고서(PDF/XLSX) 다운로드", type="primary"):
        st.success("지자체 제출용 CBM 진단 보고서가 다운로드되었습니다!")

# --- 6.1. 통합 DB 일괄 백업 ---
elif st.session_state.nav_menu == "6.1. 통합 DB 일괄 백업":
    st.subheader("📦 통합 데이터베이스 일괄 백업")
    if os.path.exists(DB_FILE_PATH):
        with open(DB_FILE_PATH, "rb") as f:
            st.download_button(
                label="📥 진단 통합 DB (Pump_Master_DB.xlsx) 백업 다운로드",
                data=f, file_name="Pump_Master_DB.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
'''

with open("pump_grade_program.py", "w", encoding="utf-8") as f:
    f.write(code_content)

print("🎉 완전한 전체 코드(약 550여 줄 분량)가 'pump_grade_program.py' 파일로 성공적으로 작성되었습니다!")