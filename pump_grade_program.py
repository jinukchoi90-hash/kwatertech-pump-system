import os
import io
import urllib.request
from datetime import datetime, timedelta
import random
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.drawing.image import Image as OpenpyxlImage
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np
from PIL import Image as PILImage

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

DB_FILE_PATH = "Pump_Master_DB.xlsx"
OVERHAUL_DB_PATH = "Pump_Overhaul_DB.xlsx"
DOC_DB_PATH = "Pump_Docs_DB.xlsx"
KNOWHOW_DB_PATH = "Pump_Knowhow_DB.xlsx"
DAILY_LOG_DB_PATH = "Pump_DailyLog_DB.xlsx"
SAFETY_PERMIT_DB_PATH = "Pump_SafetyPermit_DB.xlsx"
LOGO_FILE_PATH = "Logo.png"

# ============================================================
# 1. 페이지 기본 설정 및 모바일 커스텀 CSS/JS
# ============================================================
page_icon_setting = LOGO_FILE_PATH if os.path.exists(LOGO_FILE_PATH) else "🌊"

st.set_page_config(
    page_title="K-water tech 펌프 종합 진단 시스템",
    page_icon=page_icon_setting,
    layout="wide",
    initial_sidebar_state="auto"
)

st.markdown("""
    <style>
    .block-container {
        padding-top: 1.0rem !important;
        padding-bottom: 2rem !important;
        padding-left: 1.0rem !important;
        padding-right: 1.0rem !important;
    }
    .main-title {
        font-size: 1.6rem;
        font-weight: 800;
        color: #1e293b;
        margin-bottom: 0.2rem;
        line-height: 1.3;
    }
    @media (max-width: 768px) {
        .main-title {
            font-size: 1.15rem !important;
        }
        .sub-caption {
            font-size: 0.8rem !important;
        }
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
        font-size: 1.02rem !important;
        padding: 10px 12px !important;
        text-align: left !important;
        justify-content: flex-start !important;
        margin-bottom: 5px !important;
        background-color: #f8fafc;
        border: 1px solid #cbd5e1;
        color: #1e293b;
    }
    
    div[data-testid="stSidebar"] .stButton>button:hover {
        background-color: #0098DA !important;
        color: white !important;
        border-color: #0098DA !important;
    }

    div[data-baseweb="input"] {
        border-radius: 6px;
    }
    .kpi-card {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 12px;
        text-align: center;
    }
    .kpi-title {
        font-size: 0.8rem;
        color: #64748b;
        font-weight: bold;
    }
    .kpi-value {
        font-size: 1.4rem;
        font-weight: bold;
        color: #0f172a;
    }
    .table-header {
        background-color: #f1f5f9;
        padding: 8px 12px;
        border-radius: 6px;
        font-weight: bold;
        color: #1e293b;
        margin-bottom: 10px;
    }
    </style>
""", unsafe_allow_html=True)

USER_DB = [
    {"id": "kwatertech최진욱", "name": "최진욱", "dept": "밀양정수장", "role": "정밀진단원 / 관리자", "status": "접속중"}
]

CATEGORIES = {"성능": 40, "내부상태": 27, "기계상태": 25, "정비이력": 5}

# 10대 설비 생성 (밀양정수장 가압펌프 #1 ~ #10)
DEFAULT_PUMPS = [
    {
        "site": "밀양정수장",
        "equip": f"가압펌프 #{i}",
        "maker": "효성펌프" if i % 2 == 1 else "현대중공업",
        "model": f"DHP-{i}" if i % 2 == 1 else f"VTP-{i*10}",
        "hp": str(150 + (i * 10)),
        "head": str(45 + (i * 2)),
        "build_date": f"2018-0{min(i, 9)}-15" if i < 10 else "2018-10-15"
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
        ws.append(["작업일자", "사업장", "설비명", "공정단계", "작업자", "작업내용", "사진파일명"])
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

# 2년치(2024~2026년) 임의 샘플 데이터 생성
def seed_sample_data():
    wb = load_workbook(DB_FILE_PATH)
    ws = wb["진단이력"]
    if ws.max_row <= 1:
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

    # 오버홀 DB 2년치
    wb_o = load_workbook(OVERHAUL_DB_PATH)
    ws_o = wb_o["오버홀이력"]
    if ws_o.max_row <= 1:
        steps = ["1단계: 분해/해체 점검", "2단계: 부품 가공 및 신품 교체", "3단계: 조립 및 센터링", "4단계: 시운전 및 완료"]
        for i in range(1, 11):
            ws_o.append([f"2024-{i%9+1:02d}-10", "밀양정수장", f"가압펌프 #{i}", steps[0], "최진욱", "임펠러 해체 및 마모 상태 점검", "No Image"])
            ws_o.append([f"2025-{i%9+1:02d}-15", "밀양정수장", f"가압펌프 #{i}", steps[2], "최진욱", "축슬리브 교체 및 모터 센터링 완료", "No Image"])
            ws_o.append([f"2026-0{min(i,8)+1}-20", "밀양정수장", f"가압펌프 #{i}", steps[3], "최진욱", "정기 오버홀 완료 및 정상 시운전", "No Image"])
        wb_o.save(OVERHAUL_DB_PATH)
    wb_o.close()

    # 점검일지 2년치
    wb_d = load_workbook(DAILY_LOG_DB_PATH)
    ws_d = wb_d["점검일지"]
    if ws_d.max_row <= 1:
        for i in range(1, 11):
            ws_d.append(["2024-05-10", "최진욱", f"가압펌프 #{i}", "양호", "양호", "양호", "정상 운전 확인"])
            ws_d.append(["2025-06-12", "최진욱", f"가압펌프 #{i}", "양호", "양호", "양호", "베어링 윤활유 보충"])
            ws_d.append(["2026-08-14", "최진욱", f"가압펌프 #{i}", "양호", "양호", "양호", "특이사항 없음"])
        wb_d.save(DAILY_LOG_DB_PATH)
    wb_d.close()

    # 위험작업허가 2년치
    wb_s = load_workbook(SAFETY_PERMIT_DB_PATH)
    ws_s = wb_s["위험작업허가"]
    if ws_s.max_row <= 1:
        for i in range(1, 11):
            ws_s.append(["2024-04-10", f"가압펌프 #{i} 오버홀 작업", f"가압펌프 #{i}", "고소/위험성 작업", "2024-04-10~2024-04-15", "최진욱", "승인완료", "LOTO 차단 완료"])
            ws_s.append(["2025-07-20", f"가압펌프 #{i} 모터 정비", f"가압펌프 #{i}", "정전/전기 작업", "2025-07-20~2025-07-22", "최진욱", "승인완료", "안전장구 착용"])
            ws_s.append(["2026-08-14", f"가압펌프 #{i} 정밀 점검", f"가압펌프 #{i}", "밀폐공간 작업", "2026-08-14~2026-08-16", "최진욱", "승인대기", "가스농도 측정 완료"])
        wb_s.save(SAFETY_PERMIT_DB_PATH)
    wb_s.close()

seed_sample_data()

if "pump_list" not in st.session_state:
    st.session_state.pump_list = DEFAULT_PUMPS
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "username" not in st.session_state:
    st.session_state.username = ""
if "nav_menu" not in st.session_state:
    st.session_state.nav_menu = "1.1. 설비 통합 대시보드"
if "temp_new_pump" not in st.session_state:
    st.session_state.temp_new_pump = None

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
    st.caption("※ 상위 5개 항목 데이터 미리보기")

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
        if st.button(f"📥 {file_label} 엑셀 다운", key=f"dl_btn_{file_label}"):
            confirm_excel_download_dialog(df_data, file_label)

# ============================================================
# 4. 로그인 화면
# ============================================================
if not st.session_state.logged_in:
    _, center_col, _ = st.columns([0.1, 3.8, 0.1])
    with center_col:
        st.write("##")
        login_left, login_right = st.columns([1.3, 1.7], gap="small")
        
        with login_left:
            st.markdown("""
                <div style="background-color: #0098DA; padding: 35px 22px; border-radius: 8px 0 0 8px; color: white; min-height: 440px;">
                    <h2 style="font-size: 1.8rem; font-weight: 300; margin-bottom: 20px; border-bottom: 2px solid rgba(255,255,255,0.4); padding-bottom: 10px;">K-water tech</h2>
                    <div style="font-size: 0.88rem; line-height: 1.7;">
                        <p style="margin-bottom: 8px;"><b>▪ 펌프 17개 핵심 진단 항목 자동 판정</b><br>&nbsp;&nbsp;- 효율, 진동, 간극, 센터링 등 실시간 등급 산출</p>
                        <p style="margin-bottom: 8px;"><b>▪ 전문 보고서 파일 백데이터 아카이빙</b><br>&nbsp;&nbsp;- 오버홀·진동·효율진단 원본 PDF/엑셀 관리</p>
                        <p style="margin-bottom: 8px;"><b>▪ 4단계 오버홀 공정 및 현장 사진 이력</b><br>&nbsp;&nbsp;- 정비 단계별 현장 사진 및 자산 관리</p>
                        <p style="margin-bottom: 8px;"><b>▪ 정비·진단 노하우 DB (Troubleshooting)</b><br>&nbsp;&nbsp;- 베테랑 기술원의 결함 해결 사례 기술 자산화</p>
                        <p style="margin-bottom: 8px;"><b>▪ 현장 안전관리 및 위험작업 허가 체계</b><br>&nbsp;&nbsp;- 허가서 작성, 위험현황 모니터링, 위험성평가</p>
                    </div>
                </div>
            """, unsafe_allow_html=True)
            
        with login_right:
            st.markdown("<div style='padding: 10px;'>", unsafe_allow_html=True)
            if os.path.exists(LOGO_FILE_PATH):
                st.image(LOGO_FILE_PATH, width=210)
            st.caption("Sign into your K-water tech account.")
            
            with st.form("login_form"):
                user_id = st.text_input("👤 사번/아이디", value="kwatertech최진욱", placeholder="사번/아이디를 입력하세요")
                user_pw = st.text_input("🔒 비밀번호", type="password", value="1234", placeholder="비밀번호를 입력하세요")
                
                submit_btn = st.form_submit_button("로그인", type="primary", use_container_width=True)
                if submit_btn:
                    if user_id in [u["id"] for u in USER_DB] and user_pw == "1234":
                        user_obj = next(u for u in USER_DB if u["id"] == user_id)
                        st.session_state.logged_in = True
                        st.session_state.username = user_obj["name"]
                        st.session_state.nav_menu = "1.1. 설비 통합 대시보드"
                        st.success(f"{user_obj['name']} 님 로그인 성공!")
                        st.rerun()
                    else:
                        st.error("아이디 또는 비밀번호가 올바르지 않습니다.")
            st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<p style='text-align: center; color: #0098DA; font-weight: bold; margin-top: 15px;'>케이워터기술주식회사</p>", unsafe_allow_html=True)
    st.stop()

# ============================================================
# 5. 팝업 모달 다이얼로그 (@st.dialog)
# ============================================================
@st.dialog("📖 사내 펌프 정밀 진단 기준 및 가이드라인")
def show_guideline_dialog():
    st.markdown("""
    ### 1. 주요 핵심 관리 지표
    - **임펠러/케이싱 링 간극 (10점):** 최초 간극 대비 3.0배 이상 마모 시 필수 교체
    - **축슬리브 마모 상태 (5점):** 한쪽 패임 깊이가 외경의 2.5 ~ 3.0% 이상일 때 교체
    - **펌프 모터 센터링 (7점):** 사내 기준 0.05 mm 이내
    - **오버홀 주기 준수성 (3점):** 권장 정비 주기 약 10,000 시간
    
    ---
    ### 2. 종합 등급 판정 기준 및 조치사항
    - **A 등급 (90점 이상):** 매우 우수 | 정상 운전 유지, 정기 모니터링 수행
    - **B 등급 (80 ~ 89.9점):** 우수/정상 | 계획된 일상 점검 및 윤활 관리 수행
    - **C 등급 (70 ~ 79.9점):** 보통/관찰 | 관찰 대상 지정, 주요 진동/효율 추이 분석 및 차기 정비 계획 수립
    - **D 등급 (60 ~ 69.9점):** 개선 필요 | 정비 검토 필수, 원인 부위(진동/마모)에 대한 계획 정비 실시
    - **E 등급 (60점 미만):** 불량/위험 | 즉시 정지 및 비상 정비 검토, 오버홀 수행
    """)
    if st.button("확인 및 닫기", type="primary", use_container_width=True):
        st.rerun()

@st.dialog("❓ 신규 설비 등록 확인")
def confirm_add_pump_dialog():
    pump_data = st.session_state.temp_new_pump
    if pump_data:
        st.write(f"아래 정보로 신규 설비를 등록하시겠습니까?")
        st.info(f"**사업장:** {pump_data['site']}\n\n**설비명:** {pump_data['equip']}\n\n**제조사/모델:** {pump_data['maker']} / {pump_data['model']}")
        
        c_yes, c_no = st.columns(2)
        if c_yes.button("예 (등록)", type="primary", use_container_width=True):
            st.session_state.pump_list.append(pump_data)
            st.session_state.temp_new_pump = None
            st.success("🎉 신규 설비가 성공적으로 등록되었습니다!")
            st.rerun()
        if c_no.button("아니오 (취소)", type="secondary", use_container_width=True):
            st.session_state.temp_new_pump = None
            st.rerun()

# ============================================================
# 6. 모바일 사이드바 메뉴
# ============================================================
def navigate_to(menu_name):
    st.session_state.nav_menu = menu_name
    st.components.v1.html("""
        <script>
            var sidebar = window.parent.document.querySelector('[data-testid="stSidebar"]');
            var closeBtn = window.parent.document.querySelector('[data-testid="stSidebar"] button');
            if (sidebar && closeBtn) {
                closeBtn.click();
            }
        </script>
    """, height=0, width=0)
    st.rerun()

with st.sidebar:
    if os.path.exists(LOGO_FILE_PATH):
        st.image(LOGO_FILE_PATH, use_container_width=True)
    
    st.caption(f"👤 **{st.session_state.username}** 님 접속 중")
    if st.button("🚪 로그아웃", key="sb_logout"):
        st.session_state.logged_in = False
        st.rerun()
    st.write("---")

    st.markdown("### ⚙️ 시스템 메뉴")

    menu_items = [
        ("1.1. 설비 통합 대시보드", "🌊 1.1. 설비 통합 대시보드"),
        ("1.2. 설비 마스터 관리", "⚙️ 1.2. 설비 마스터 관리"),
        ("2.1. 펌프 정밀 진단 (17개)", "🛢️ 2.1. 펌프 정밀 진단 (17개)"),
        ("2.2. 일상/정기 점검일지", "📝 2.2. 일상/정기 점검일지"),
        ("3.1. 오버홀 공정/사진 관리", "🛠️ 3.1. 오버홀 공정/사진 관리"),
        ("3.2. 전문 보고서 백데이터 DB", "📁 3.2. 전문 보고서 백데이터 DB"),
        ("4.1. 성능/상태 5대 추이 분석", "📈 4.1. 성능/상태 5대 추이 분석"),
        ("5.1. K-water tech 노하우 DB", "💡 5.1. K-water tech 노하우 DB"),
        ("5.2. 현장 안전 체크리스트", "🛡️ 5.2. 현장 안전 체크리스트"),
        ("5.3. 위험작업 허가서", "📑 5.3. 위험작업 허가서"),
        ("5.4. 위험작업 현황", "🚨 5.4. 위험작업 현황"),
        ("5.5. 안전작업계획서 및 위험성평가", "📋 5.5. 안전작업계획서/위험성평가"),
        ("6.1. 사용자 권한 관리", "👤 6.1. 사용자 권한 관리"),
        ("6.2. 통합 DB 일괄 백업", "📦 6.2. 통합 DB 일괄 백업")
    ]

    for menu_key, menu_label in menu_items:
        is_selected = (st.session_state.nav_menu == menu_key)
        btn_type = "primary" if is_selected else "secondary"
        if st.button(menu_label, key=f"btn_{menu_key}", type=btn_type, use_container_width=True):
            navigate_to(menu_key)

head_c1, head_c2 = st.columns([3.5, 1.2])
with head_c1:
    st.markdown("<div class='main-title'>K-water tech 펌프 진단 & 자산관리</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='sub-caption' style='color:#64748b; font-size:0.9rem;'>현재 메뉴: <b>{st.session_state.nav_menu}</b></div>", unsafe_allow_html=True)
with head_c2:
    if st.button("📖 가이드", type="secondary"):
        show_guideline_dialog()

st.write("---")

# ============================================================
# [메인 화면 메뉴 구동 및 메뉴별 엑셀 다운로드 연동]
# ============================================================

# --- 1.1. 설비 통합 대시보드 ---
if st.session_state.nav_menu == "1.1. 설비 통합 대시보드":
    wb = load_workbook(DB_FILE_PATH, data_only=True)
    ws = wb["진단이력"]
    diag_records = [row for row in ws.iter_rows(min_row=1, values_only=True)]
    wb.close()
    df_dash = pd.DataFrame(diag_records[1:], columns=diag_records[0]) if len(diag_records) > 1 else pd.DataFrame()

    render_download_button_header(df_dash, "대시보드_진단이력")

    st.subheader("📊 설비 건전성 현황판 (밀양정수장 가압펌프 #1~#10)")
    latest_grades = {}
    if not df_dash.empty:
        for _, row in df_dash.iterrows():
            latest_grades[row["설비명"]] = row["최종등급"]

    total_cnt = len(st.session_state.pump_list)
    grade_counts = {"A": 0, "B": 0, "C": 0, "D": 0, "E": 0}
    for g in latest_grades.values():
        if g in grade_counts: grade_counts[g] += 1

    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("총 펌프", f"{total_cnt} 대")
    m2.metric("A 등급", f"{grade_counts['A']} 대")
    m3.metric("B 등급", f"{grade_counts['B']} 대")
    m4.metric("C 등급", f"{grade_counts['C']} 대")
    m5.metric("D 등급", f"{grade_counts['D']} 대")
    m6.metric("E 등급", f"{grade_counts['E']} 대")

    st.write("---")
    st.subheader("⚠️ 주요 주의/위험 설비 알림 리스트")
    danger_pumps = [equip for equip, g in latest_grades.items() if g in ["D", "E"]]
    if danger_pumps:
        st.warning(f"현재 정비 검토가 필요한 설비: {', '.join(danger_pumps)}")
    else:
        st.success("현재 D/E 등급의 위험 설비가 없습니다.")

    st.write("---")
    st.subheader("📜 [통합 이력] 전체 펌프 2년치 정밀 진단 이력")
    if not df_dash.empty:
        st.dataframe(df_dash, use_container_width=True)
    else:
        st.info("등록된 정밀 진단 이력이 없습니다.")

# --- 1.2. 설비 마스터 관리 ---
elif st.session_state.nav_menu == "1.2. 설비 마스터 관리":
    df_pumps = pd.DataFrame(st.session_state.pump_list)
    render_download_button_header(df_pumps, "설비마스터목록")

    st.subheader("📋 설비 마스터 등록 및 관리 (밀양정수장 10대)")
    st.dataframe(df_pumps, use_container_width=True)

    with st.expander("➕ 신규 펌프 설비 등록하기"):
        with st.form("add_pump_form"):
            c1, c2, c3 = st.columns(3)
            new_site = c1.text_input("사업장명", value="밀양정수장")
            new_equip = c2.text_input("설비명 (호기)", value="가압펌프 #11")
            new_maker = c3.text_input("제조사", value="효성펌프")
            c4, c5, c6 = st.columns(3)
            new_model = c4.text_input("모델명", value="DHP-11")
            new_hp = c5.text_input("마력 (HP)", value="250")
            new_head = c6.text_input("양정 (m)", value="65")
            new_build = st.text_input("준공일자", value=datetime.now().strftime("%Y-%m-%d"))
            
            if st.form_submit_button("등록 신청", type="primary"):
                st.session_state.temp_new_pump = {
                    "site": new_site, "equip": new_equip, "maker": new_maker,
                    "model": new_model, "hp": new_hp, "head": new_head, "build_date": new_build
                }
                st.rerun()

    if st.session_state.temp_new_pump is not None:
        confirm_add_pump_dialog()

# --- 2.1. 펌프 정밀 진단 (17개) ---
elif st.session_state.nav_menu == "2.1. 펌프 정밀 진단 (17개)":
    init_korean_font()
    wb = load_workbook(DB_FILE_PATH, data_only=True)
    ws = wb["진단이력"]
    records = [row for row in ws.iter_rows(min_row=1, values_only=True)]
    wb.close()
    df_diag_all = pd.DataFrame(records[1:], columns=records[0]) if len(records) > 1 else pd.DataFrame()
    render_download_button_header(df_diag_all, "정밀진단전체이력")

    st.subheader("📋 설비 선택 및 기본정보")
    pump_names = [p["equip"] for p in st.session_state.pump_list]
    selected_pump_name = st.selectbox("진단할 펌프 설비를 선택하세요", pump_names)
    pump_info = next((p for p in st.session_state.pump_list if p["equip"] == selected_pump_name), st.session_state.pump_list[0])
    
    b_col1, b_col2, b_col3 = st.columns(3)
    site = b_col1.text_input("사업장", value=pump_info["site"])
    equip = b_col1.text_input("설비명", value=pump_info["equip"])
    maker = b_col1.text_input("제조사", value=pump_info["maker"])
    model = b_col2.text_input("모델명", value=pump_info["model"])
    hp = b_col2.text_input("마력 (HP)", value=pump_info["hp"])
    head = b_col2.text_input("양정 (m)", value=pump_info["head"])
    build_date = b_col3.text_input("준공일", value=pump_info["build_date"])
    check_date = b_col3.date_input("점검일", value=datetime.now())
    inspector = b_col3.text_input("점검자", value=st.session_state.username)

    st.write("---")
    st.subheader("📝 세부 평가 항목 입력 (측정값 자동 판정 연동)")
    
    st.markdown("""
        <div class="table-header">
            <span style="display:inline-block; width:35%;">■ 평가 항목 및 배점</span>
            <span style="display:inline-block; width:20%;">■ 측정값</span>
            <span style="display:inline-block; width:28%;">■ 기준값</span>
            <span style="display:inline-block; width:15%;">■ 등급</span>
        </div>
    """, unsafe_allow_html=True)

    cat_scores = {k: 0.0 for k in CATEGORIES.keys()}
    total_score = 0.0
    details = []

    current_cat = ""
    col_left_diag, col_right_diag = st.columns([1.25, 1])

    with col_left_diag:
        for idx, (cat, item, weight, std_val, options, score_dict, auto_fn) in enumerate(EVAL_ITEMS):
            if cat != current_cat:
                current_cat = cat
                st.markdown(f"#### ■ {cat} 영역 (배점: {CATEGORIES[cat]}점)")
            
            ec1, ec2, ec3, ec4 = st.columns([2.2, 1.2, 1.8, 1.0])
            with ec1: st.write(f"**{item}** ({weight}점)")
            
            def update_grade(item_idx=idx, fn=auto_fn):
                input_str = st.session_state.get(f"val_{item_idx}", "")
                if fn and input_str.strip():
                    try:
                        v = float(input_str.strip())
                        st.session_state[f"selected_grade_{item_idx}"] = fn(v)
                    except ValueError: pass

            with ec2: 
                input_val = st.text_input("측정값", key=f"val_{idx}", label_visibility="collapsed", placeholder="수치", on_change=update_grade)
            
            with ec3: 
                st.markdown(f"<span style='color: #059669; font-size: 0.82rem; font-weight: bold;'>[기준] {std_val}</span>", unsafe_allow_html=True)

            with ec4: 
                selected_grade = st.selectbox("등급", options, key=f"selected_grade_{idx}", label_visibility="collapsed")
            
            s = score_dict[selected_grade]
            cat_scores[cat] += s
            total_score += s
            details.append({"category": cat, "item": item, "weight": weight, "std": std_val, "val": input_val, "grade": selected_grade, "score": s})

    total_score = round(total_score, 2)
    if total_score >= 90.0: final_grade = "A"
    elif total_score >= 80.0: final_grade = "B"
    elif total_score >= 70.0: final_grade = "C"
    elif total_score >= 60.0: final_grade = "D"
    else: final_grade = "E"

    with col_right_diag:
        st.subheader("📊 종합 판정 및 분석 차트")
        rc1, rc2 = st.columns(2)
        rc1.metric("현재 점수", f"{total_score:.2f} 점", f"등급: {final_grade}")
        rc2.metric("전회차 점수", "82.50 점 (B)", f"{total_score - 82.50:+.2f} 점")

        st.info(FINAL_GRADE_INFO[final_grade][1])

        fig = plt.figure(figsize=(6, 8.5), dpi=100)
        
        ax1 = fig.add_subplot(321, polar=True)
        cats = list(CATEGORIES.keys())
        N = len(cats)
        vals = [(cat_scores[c] / CATEGORIES[c]) * 100 for c in cats] + [(cat_scores[cats[0]] / CATEGORIES[cats[0]]) * 100]
        angles = [n / float(N) * 2 * np.pi for n in range(N)] + [0]
        ax1.set_theta_offset(np.pi / 2)
        ax1.set_theta_direction(-1)
        ax1.set_xticks(angles[:-1])
        ax1.set_xticklabels(cats, fontsize=8, fontweight='bold')
        ax1.set_ylim(0, 125)
        ax1.plot(angles, vals, linewidth=1.5, color='#2563EB')
        ax1.fill(angles, vals, color='#3B82F6', alpha=0.25)
        ax1.set_title("1. 영역별 달성율 (%)", size=8, fontweight='bold', pad=10)

        ax2 = fig.add_subplot(322)
        x = np.arange(len(cats))
        ax2.bar(x, list(cat_scores.values()), 0.4, color='#2563EB')
        ax2.set_xticks(x)
        ax2.set_xticklabels(cats, fontsize=7)
        ax2.set_title("2. 영역별 환산점수", size=8, fontweight='bold')

        ax3 = fig.add_subplot(323)
        ax3.plot(["1차전", "2차전", "현재"], [75.0, 82.5, total_score], marker='o', linewidth=1.5, color='#16A34A')
        ax3.set_ylim(0, 110)
        ax3.set_title("3. 최근 3회차 점수 추이", size=8, fontweight='bold')

        ax4 = fig.add_subplot(324)
        diff_val = total_score - 82.50
        ax4.bar(["점수 변동"], [diff_val], color='#16A34A' if diff_val >= 0 else '#DC2626')
        ax4.set_title("4. 전회차 대비 변동", size=8, fontweight='bold')

        ax5 = fig.add_subplot(313)
        grade_dist = {"A": 0, "B": 0, "C": 0, "D": 0, "E": 0}
        for d in details:
            g = d["grade"][0]
            if g in grade_dist: grade_dist[g] += 1
        ax5.bar(list(grade_dist.keys()), list(grade_dist.values()), color='#0284C7')
        ax5.set_title("5. 17개 평가항목 등급 분포 수", size=8, fontweight='bold')

        fig.subplots_adjust(wspace=0.35, hspace=0.55)
        st.pyplot(fig)

        if st.button("💾 진단 결과 DB 저장", type="primary", use_container_width=True):
            wb = load_workbook(DB_FILE_PATH)
            ws = wb["진단이력"]
            row_data = [str(check_date), site, equip, maker, model, hp, head, build_date, inspector, total_score, final_grade] + [d["grade"] for d in details]
            ws.append(row_data)
            wb.save(DB_FILE_PATH)
            st.success("통합 DB에 저장되었습니다!")

    st.write("---")
    with st.expander("📜 [이력 조회] 선택 펌프의 지난 정밀 진단 기록 보기", expanded=True):
        if not df_diag_all.empty:
            filtered_df = df_diag_all[df_diag_all["설비명"] == selected_pump_name]
            st.dataframe(filtered_df, use_container_width=True)
        else:
            st.info(f"{selected_pump_name}의 지난 진단 이력이 없습니다.")

# --- 2.2. 일상/정기 점검일지 ---
elif st.session_state.nav_menu == "2.2. 일상/정기 점검일지":
    wb = load_workbook(DAILY_LOG_DB_PATH, data_only=True)
    ws = wb["점검일지"]
    logs = [row for row in ws.iter_rows(min_row=1, values_only=True)]
    wb.close()
    df_logs = pd.DataFrame(logs[1:], columns=logs[0]) if len(logs) > 1 else pd.DataFrame()

    render_download_button_header(df_logs, "일상점검일지이력")

    st.subheader("📅 일상 및 정기 점검 일지 기록")
    
    with st.form("daily_log_form"):
        c1, c2 = st.columns(2)
        d_date = c1.date_input("점검일자", value=datetime.now())
        d_user = c2.text_input("점검자", value=st.session_state.username)
        d_pump = st.selectbox("점검 펌프 선택", [p["equip"] for p in st.session_state.pump_list])
        
        c_chk1 = st.checkbox("1. 누수 및 베어링 윤활유 상태 양호")
        c_chk2 = st.checkbox("2. 이상 소모 및 진동 발생 여부 없음")
        c_chk3 = st.checkbox("3. 모터 온도 및 전류값 정상 범위 내 운전")
        d_memo = st.text_area("특이사항 및 작업 메모")
        
        if st.form_submit_button("📝 점검일지 등록 및 DB 저장", type="primary"):
            wb = load_workbook(DAILY_LOG_DB_PATH)
            ws = wb["점검일지"]
            ws.append([
                str(d_date), d_user, d_pump,
                "양호" if c_chk1 else "점검필요",
                "양호" if c_chk2 else "점검필요",
                "양호" if c_chk3 else "점검필요",
                d_memo
            ])
            wb.save(DAILY_LOG_DB_PATH)
            st.success("일상 점검일지가 성공적으로 저장되었습니다!")
            st.rerun()

    st.write("---")
    st.subheader("📜 [이력 조회] 지난 일상/정기 점검일지 전체 목록 (2년치)")
    if not df_logs.empty:
        st.dataframe(df_logs, use_container_width=True)
    else:
        st.info("등록된 점검일지 이력이 없습니다.")

# --- 3.1. 오버홀 공정/사진 관리 ---
elif st.session_state.nav_menu == "3.1. 오버홀 공정/사진 관리":
    wb = load_workbook(OVERHAUL_DB_PATH, data_only=True)
    ws = wb["오버홀이력"]
    o_records = [row for row in ws.iter_rows(min_row=1, values_only=True)]
    wb.close()
    df_overhaul = pd.DataFrame(o_records[1:], columns=o_records[0]) if len(o_records) > 1 else pd.DataFrame()

    render_download_button_header(df_overhaul, "오버홀공정이력")

    st.subheader("📷 오버홀 진행 공정 및 현장 사진 등록")
    selected_pump = st.selectbox("오버홀 등록 설비 선택", [p["equip"] for p in st.session_state.pump_list])
    
    with st.form("overhaul_form"):
        col1, col2 = st.columns(2)
        step = col1.selectbox("공정 단계", ["1단계: 분해/해체 점검", "2단계: 부품 가공 및 신품 교체", "3단계: 조립 및 센터링", "4단계: 시운전 및 완료"])
        work_date = col2.date_input("작업 일자", value=datetime.now())
        work_memo = st.text_area("작업 세부 내용")
        uploaded_file = st.file_uploader("현장 사진 업로드", type=["jpg", "png"])
        
        if st.form_submit_button("📷 오버홀 기록 및 사진 저장", type="primary"):
            file_name = "No Image"
            if uploaded_file:
                os.makedirs("overhaul_images", exist_ok=True)
                file_name = f"overhaul_images/{selected_pump}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
                PILImage.open(uploaded_file).save(file_name)
            
            wb = load_workbook(OVERHAUL_DB_PATH)
            ws = wb["오버홀이력"]
            ws.append([str(work_date), "밀양정수장", selected_pump, step, st.session_state.username, work_memo, file_name])
            wb.save(OVERHAUL_DB_PATH)
            st.success("오버홀 기록이 등록되었습니다!")
            st.rerun()

    st.write("---")
    st.subheader("📜 [이력 조회] 지난 오버홀 공정 진행 기록 전체 목록 (2년치)")
    if not df_overhaul.empty:
        st.dataframe(df_overhaul, use_container_width=True)
    else:
        st.info("등록된 오버홀 공정 이력이 없습니다.")

# --- 3.2. 전문 보고서 백데이터 DB ---
elif st.session_state.nav_menu == "3.2. 전문 보고서 백데이터 DB":
    wb = load_workbook(DOC_DB_PATH, data_only=True)
    ws = wb["전문보고서"]
    doc_records = [row for row in ws.iter_rows(min_row=1, values_only=True)]
    wb.close()
    df_docs = pd.DataFrame(doc_records[1:], columns=doc_records[0]) if len(doc_records) > 1 else pd.DataFrame()

    render_download_button_header(df_docs, "전문보고서백데이터목록")

    st.subheader("📁 전문 점검보고서 백데이터 파일 문서고 (아카이빙)")
    
    with st.expander("➕ 전문 점검보고서 (PDF/XLSX/HWP) 업로드 등록"):
        with st.form("doc_upload_form"):
            c1, c2, c3 = st.columns(3)
            doc_pump = c1.selectbox("대상 설비", [p["equip"] for p in st.session_state.pump_list])
            doc_type = c2.selectbox("보고서 구분", ["오버홀 점검보고서", "진동측정 점검보고서", "효율진단 점검보고서", "기타 정밀보고서"])
            doc_org = c3.text_input("수행 기관/업체", value="K-water 기술팀")
            
            doc_memo = st.text_area("주요 요약 소견")
            doc_file = st.file_uploader("보고서 원본 파일 첨부", type=["pdf", "xlsx", "hwp", "docx"])
            
            if st.form_submit_button("📤 보고서 문서고 저장", type="primary"):
                if doc_file:
                    os.makedirs("doc_files", exist_ok=True)
                    saved_path = f"doc_files/{doc_pump}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{doc_file.name}"
                    with open(saved_path, "wb") as f:
                        f.write(doc_file.getbuffer())
                    
                    wb = load_workbook(DOC_DB_PATH)
                    ws = wb["전문보고서"]
                    ws.append([datetime.now().strftime('%Y-%m-%d'), "밀양정수장", doc_pump, doc_type, doc_org, doc_memo, saved_path, doc_file.name])
                    wb.save(DOC_DB_PATH)
                    st.success("보고서 원본 파일이 백데이터 DB에 저장되었습니다!")
                    st.rerun()

    st.write("---")
    st.subheader("📜 [이력 조회] 등록된 전문 보고서 아카이빙 전체 이력")
    if not df_docs.empty:
        st.dataframe(df_docs, use_container_width=True)
    else:
        st.info("등록된 전문 보고서가 없습니다.")

# --- 4.1. 성능/상태 5대 추이 분석 ---
elif st.session_state.nav_menu == "4.1. 성능/상태 5대 추이 분석":
    init_korean_font()
    dates = ["2024-01", "2024-07", "2025-01", "2025-07", "2026-01", "현재"]
    df_trend = pd.DataFrame({
        "점검시기": dates,
        "효율유지율(%)": [98.5, 96.2, 94.0, 91.5, 88.0, 85.2],
        "Overall진동(mm/s)": [1.2, 1.5, 2.1, 3.5, 5.2, 7.5],
        "링간극마모(배수)": [1.0, 1.2, 1.5, 1.8, 2.2, 2.8]
    })
    render_download_button_header(df_trend, "2년치_성능상태추이분석데이터")

    st.subheader("📈 펌프 건전성 5대 정밀 추이 분석 그래프 (2년치 데이터)")
    selected_p = st.selectbox("분석 설비 선택", [p["equip"] for p in st.session_state.pump_list])

    fig_all, axes = plt.subplots(3, 2, figsize=(12, 12))

    axes[0, 0].plot(dates, df_trend["효율유지율(%)"], marker='o', color='#2563EB')
    axes[0, 0].set_title("1. 효율 유지율 (%) 추이", fontweight='bold')
    axes[0, 0].grid(True, linestyle='--')

    axes[0, 1].plot(dates, df_trend["Overall진동(mm/s)"], marker='s', color='#DC2626')
    axes[0, 1].axhline(7.1, color='orange', linestyle='--', label='경고 기준 (7.1)')
    axes[0, 1].set_title("2. Overall 진동 (mm/s) 추이", fontweight='bold')
    axes[0, 1].grid(True, linestyle='--')
    axes[0, 1].legend()

    axes[1, 0].plot(dates, df_trend["링간극마모(배수)"], marker='^', color='#16A34A')
    axes[1, 0].axhline(3.0, color='red', linestyle='--', label='교체 기준 (3.0배)')
    axes[1, 0].set_title("3. 임펠러 링 간극 마모 (배수) 추이", fontweight='bold')
    axes[1, 0].grid(True, linestyle='--')
    axes[1, 0].legend()

    ax_rad = fig_all.add_subplot(3, 2, 4, polar=True)
    cats = list(CATEGORIES.keys())
    N = len(cats)
    angles = [n / float(N) * 2 * np.pi for n in range(N)] + [0]
    vals = [85, 90, 70, 95, 85]
    ax_rad.plot(angles, vals, color='#9333EA')
    ax_rad.fill(angles, vals, color='#A855F7', alpha=0.3)
    ax_rad.set_xticks(angles[:-1])
    ax_rad.set_xticklabels(cats, fontsize=8)
    ax_rad.set_title("4. 4대 영역별 달성율 균형도", fontweight='bold')

    p_names = [p["equip"] for p in st.session_state.pump_list]
    p_scores = [85.2, 92.0, 78.5, 88.4, 91.2, 84.0, 79.5, 86.3, 90.1, 82.7]
    axes[2, 0].bar(p_names, p_scores, color='#0284C7')
    axes[2, 0].set_xticklabels(p_names, rotation=30, fontsize=7)
    axes[2, 0].set_title("5. 밀양정수장 가압펌프 10대 종합 점수 비교", fontweight='bold')
    axes[2, 0].set_ylim(0, 100)

    axes[2, 1].axis('off')
    fig_all.tight_layout()
    st.pyplot(fig_all)

# --- 5.1. K-water tech 노하우 DB ---
elif st.session_state.nav_menu == "5.1. K-water tech 노하우 DB":
    wb = load_workbook(KNOWHOW_DB_PATH, data_only=True)
    ws = wb["노하우DB"]
    kh_records = [row for row in ws.iter_rows(min_row=1, values_only=True)]
    wb.close()
    df_kh = pd.DataFrame(kh_records[1:], columns=kh_records[0]) if len(kh_records) > 1 else pd.DataFrame()

    render_download_button_header(df_kh, "기술노하우DB")

    st.subheader("💡 K-water tech 정비 & 결함 해결 노하우 DB (Troubleshooting)")
    search_keyword = st.text_input("🔍 현상 / 키워드 검색 (예: 진동, 링 간극, 센터링)", placeholder="검색어를 입력하세요")
    
    with st.expander("➕ 현장 결함 해결 노하우 사례 등록하기"):
        with st.form("knowhow_form"):
            c1, c2 = st.columns(2)
            kh_cat = c1.selectbox("사례 분류", ["진동/소음 원인 해결", "효율 저하 개선", "마모/부식 방지", "센터링/오버홀 팁", "기타 현장 노하우"])
            kh_model = c2.text_input("관련 펌프 모델/형식", value="DHP 계열")
            kh_sym = st.text_area("현상 및 원인 분석")
            kh_sol = st.text_area("해결 조치 노하우")
            
            if st.form_submit_button("💡 노하우 DB 저장", type="primary"):
                wb = load_workbook(KNOWHOW_DB_PATH)
                ws = wb["노하우DB"]
                ws.append([datetime.now().strftime('%Y-%m-%d'), kh_cat, kh_model, kh_sym, kh_sol, st.session_state.username])
                wb.save(KNOWHOW_DB_PATH)
                st.success("기술 노하우가 저장되었습니다!")
                st.rerun()

    st.write("---")
    st.subheader("📜 [이력 조회] 등록된 정비 기술 노하우 전체 이력")
    if not df_kh.empty:
        st.dataframe(df_kh, use_container_width=True)
    else:
        st.info("등록된 기술 노하우가 없습니다.")

# --- 5.2. 현장 안전 체크리스트 ---
elif st.session_state.nav_menu == "5.2. 현장 안전 체크리스트":
    df_safe = pd.DataFrame([
        {"항목": "1. LOTO 전원 차단", "점검사항": "펌프 모터 전원 차단 및 잠금장치/표지판 게시 완료", "상태": "완료"},
        {"항목": "2. 잔류 압력 제거", "점검사항": "흡입/토출 밸브 차단 및 내부 잔류 압력 제거 확인", "상태": "완료"},
        {"항목": "3. 보호구 착용", "점검사항": "개인 보호구 (안전모, 안전화, 코팅 장갑) 착용 완료", "상태": "완료"},
        {"항목": "4. 가이드 커버 체결", "점검사항": "회전 부위 가이드 커버 설치 및 체결 상태 확인", "상태": "완료"}
    ])
    render_download_button_header(df_safe, "현장안전체크리스트")

    st.subheader("🛡️ 현장 점검 전 필수 안전 체크리스트")
    for _, r in df_safe.iterrows():
        st.checkbox(f"{r['항목']}: {r['점검사항']}", value=True)

# --- 5.3. 위험작업 허가서 ---
elif st.session_state.nav_menu == "5.3. 위험작업 허가서":
    wb = load_workbook(SAFETY_PERMIT_DB_PATH, data_only=True)
    ws = wb["위험작업허가"]
    permits = [row for row in ws.iter_rows(min_row=1, values_only=True)]
    wb.close()
    df_permits = pd.DataFrame(permits[1:], columns=permits[0]) if len(permits) > 1 else pd.DataFrame()

    render_download_button_header(df_permits, "위험작업허가서목록")

    st.subheader("📑 현장 위험작업 허가서 작성 및 승인 신청")
    
    with st.form("safety_permit_form"):
        c1, c2 = st.columns(2)
        p_title = c1.text_input("작업명", value="밀양정수장 가압펌프 #1 오버홀 해체작업")
        p_pump = c2.selectbox("대상 설비", [p["equip"] for p in st.session_state.pump_list])
        
        c3, c4 = st.columns(2)
        p_types = c3.multiselect("위험작업 유형 선택", ["밀폐공간 작업", "화기/용접 작업", "고소/위험성 작업", "정전/전기 작업", "중장비/인양 작업"], default=["고소/위험성 작업"])
        p_period = c4.text_input("작업 예정 기간", value=f"{datetime.now().strftime('%Y-%m-%d')} ~ {(datetime.now()).strftime('%Y-%m-%d')}")
        
        st.markdown("##### ■ 필수 사전 안전 조치 점검")
        s_chk1 = st.checkbox("전기 LOTO(Lockout/Tagout) 차단 및 표지판 게시 완료", value=True)
        s_chk2 = st.checkbox("유해가스 및 산소농도 측정 완료 (기준치 적정)", value=True)
        s_chk3 = st.checkbox("화재감시자 배치 및 소화기 현장 비치 완료", value=True)
        s_chk4 = st.checkbox("안전대/안전모/보호구 착용 상태 검토 완료", value=True)
        
        p_memo = st.text_area("기타 특이 안전 조치사항 및 요구조건")
        
        if st.form_submit_button("📝 위험작업 허가 신청 등록", type="primary"):
            wb = load_workbook(SAFETY_PERMIT_DB_PATH)
            ws = wb["위험작업허가"]
            ws.append([
                datetime.now().strftime('%Y-%m-%d'), p_title, p_pump,
                ", ".join(p_types), p_period, st.session_state.username, "승인대기", p_memo
            ])
            wb.save(SAFETY_PERMIT_DB_PATH)
            st.success("위험작업 허가서 승인 신청이 완료되었습니다!")
            st.rerun()

# --- 5.4. 위험작업 현황 ---
elif st.session_state.nav_menu == "5.4. 위험작업 현황":
    wb = load_workbook(SAFETY_PERMIT_DB_PATH, data_only=True)
    ws = wb["위험작업허가"]
    permits = [row for row in ws.iter_rows(min_row=1, values_only=True)]
    wb.close()
    df_permits = pd.DataFrame(permits[1:], columns=permits[0]) if len(permits) > 1 else pd.DataFrame()

    render_download_button_header(df_permits, "2년치_위험작업진행현황")

    st.subheader("🚨 밀양정수장 가압펌프 10대 위험작업 진행 현황 모니터링 (2년치)")
    
    if not df_permits.empty:
        c1, c2, c3 = st.columns(3)
        c1.metric("총 등록 허가건", f"{len(df_permits)} 건")
        c2.metric("진행중/승인완료", f"{sum(1 for r in df_permits['승인상태'] if r=='승인완료')} 건")
        c3.metric("승인 대기중", f"{sum(1 for r in df_permits['승인상태'] if r=='승인대기')} 건")
        
        st.write("---")
        st.subheader("📜 위험작업 허가 및 진행 상태 리스트")
        st.dataframe(df_permits, use_container_width=True)
    else:
        st.info("현재 등록된 위험작업 허가 내역이 없습니다.")

# --- 5.5. 안전작업계획서 및 위험성평가 ---
elif st.session_state.nav_menu == "5.5. 안전작업계획서 및 위험성평가":
    df_jsa = pd.DataFrame([
        {"위험요인": "[낙하/추락] 가설발판 및 작업대 안전난간 미체결", "위험성": "中 -> 低", "개선대책": "안전난간 체결 상태 사전 점검 완료"},
        {"위험요인": "[감전/착오] 모터 케이블 완전 차단 미흡", "위험성": "高 -> 低", "개선대책": "LOTO 절차 이행 및 검전 실시"},
        {"위험요인": "[협착/끼임] 회전체 가이드 커버 개방", "위험성": "中 -> 低", "개선대책": "가이드 커버 상시 체결 및 접근 금지"},
        {"위험요인": "[유해물질] 윤활유/유기용제 접촉", "위험성": "低", "개선대책": "MSDS 숙지 및 보호장갑 착용"}
    ])
    render_download_button_header(df_jsa, "안전작업계획서_위험성평가")

    st.subheader("📋 안전작업계획서 및 사전 위험성평가(JSA)")
    
    st.markdown("#### 1. 위험성평가 (Risk Assessment) 표준 항목 점검")
    for _, r in df_jsa.iterrows():
        st.checkbox(f"{r['위험요인']} (개선: {r['개선대책']})", value=True)

    st.write("---")
    st.markdown("#### 2. 안전작업계획서 서식 작성 및 제출")
    with st.form("risk_plan_form"):
        plan_equip = st.selectbox("작업 대상 펌프", [p["equip"] for p in st.session_state.pump_list])
        plan_hazard = st.text_area("주요 잠재 위험요인 (Hazard)", value="임펠러 분해 작업 중 중량물 낙하 위험 및 케이싱 내부 협착 위험")
        plan_counter = st.text_area("위험 감소 대책 (Action Plan)", value="크레인 샤클 및 슬링벨트 사전 점검, 2인 1조 작업 수행, 신호수 배치")
        
        if st.form_submit_button("📋 안전작업계획서 승인 제출", type="primary"):
            st.success("안전작업계획서 및 위험성평가가 제출되었습니다!")

# --- 6.1. 사용자 권한 관리 ---
elif st.session_state.nav_menu == "6.1. 사용자 권한 관리":
    df_users = pd.DataFrame(USER_DB)
    render_download_button_header(df_users, "사용자권한목록")

    st.subheader("⚙️ 사용자 및 점검자 권한 관리")
    st.dataframe(df_users, use_container_width=True)

    c1, c2, c3 = st.columns(3)
    c1.markdown("<div class='kpi-card'><div class='kpi-title'>총 등록 사용자</div><div class='kpi-value'>1 명</div></div>", unsafe_allow_html=True)
    c2.markdown("<div class='kpi-card'><div class='kpi-title'>현재 접속 중 계정</div><div class='kpi-value'>1 명</div></div>", unsafe_allow_html=True)
    c3.markdown("<div class='kpi-card'><div class='kpi-title'>최고 관리자 수</div><div class='kpi-value'>1 명</div></div>", unsafe_allow_html=True)

# --- 6.2. 통합 DB 일괄 백업 ---
elif st.session_state.nav_menu == "6.2. 통합 DB 일괄 백업":
    wb = load_workbook(DB_FILE_PATH, data_only=True)
    ws = wb["진단이력"]
    records = [row for row in ws.iter_rows(min_row=1, values_only=True)]
    wb.close()
    df_backup = pd.DataFrame(records[1:], columns=records[0]) if len(records) > 1 else pd.DataFrame()

    render_download_button_header(df_backup, "통합DB전체백업")

    st.subheader("📦 통합 데이터베이스 일괄 백업 및 다운로드")
    if os.path.exists(DB_FILE_PATH):
        with open(DB_FILE_PATH, "rb") as f:
            st.download_button(
                label="📥 진단 통합 DB (Pump_Master_DB.xlsx) 백업 다운로드",
                data=f, file_name="Pump_Master_DB.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )