# ============================================================
# K-water tech
# 펌프 설비관리 통합 플랫폼
# QR 기반 설비관리 + 정밀진단 + CBM + 오버홀 + AI + KPI
#
# Streamlit Single File Application
# ============================================================

import os
import io
import urllib.request
from datetime import datetime, date
import random
import math

import streamlit as st
import streamlit.components.v1 as components

import pandas as pd
import numpy as np
import openpyxl
from openpyxl import Workbook, load_workbook

import matplotlib.pyplot as plt
import matplotlib.font_manager as fm


# ============================================================
# 0. 기본 설정
# ============================================================

APP_TITLE = "K-water tech 설비관리 플랫폼"

DB_FILE_PATH = "Pump_Master_DB.xlsx"
OVERHAUL_DB_PATH = "Pump_Overhaul_DB.xlsx"
DOC_DB_PATH = "Pump_Docs_DB.xlsx"
KNOWHOW_DB_PATH = "Pump_Knowhow_DB.xlsx"
DAILY_LOG_DB_PATH = "Pump_DailyLog_DB.xlsx"
SAFETY_PERMIT_DB_PATH = "Pump_SafetyPermit_DB.xlsx"
KPI_DB_PATH = "Pump_KPI_DB.xlsx"

LOGO_FILE_PATH = "Logo.png"
SEED_FLAG_PATH = "seed_flag.txt"


# ============================================================
# 1. 한글 폰트
# ============================================================

def init_korean_font():

    font_filename = "NanumGothic.ttf"

    if not os.path.exists(font_filename):

        font_url = (
            "https://github.com/google/fonts/raw/main/"
            "ofl/nanumgothic/NanumGothic-Regular.ttf"
        )

        try:
            urllib.request.urlretrieve(font_url, font_filename)
        except Exception:
            pass

    if os.path.exists(font_filename):

        try:
            fm.fontManager.addfont(font_filename)

            font_prop = fm.FontProperties(
                fname=font_filename
            )

            plt.rcParams["font.family"] = font_prop.get_name()

        except Exception:
            plt.rcParams["font.family"] = "DejaVu Sans"

    else:

        plt.rcParams["font.family"] = "DejaVu Sans"

    plt.rcParams["axes.unicode_minus"] = False


init_korean_font()


# ============================================================
# 2. 페이지 설정
# ============================================================

st.set_page_config(

    page_title=APP_TITLE,

    page_icon="💧",

    layout="wide",

    initial_sidebar_state="expanded"
)


# ============================================================
# 2-1. HTML 렌더링 버그 수정 패치
#
# Streamlit의 st.markdown(unsafe_allow_html=True)은 내부적으로
# 마크다운 파서를 거치는데, 줄 앞에 공백이 4칸 이상 있으면
# 마크다운 문법상 "코드 블록"으로 인식되어 HTML 태그가
# 렌더링되지 않고 그대로 화면에 텍스트로 노출되는 문제가 있다.
#
# 이 프로젝트는 들여쓰기가 깊은 코드 스타일이라 여러 곳에서
# 이 문제가 발생하므로, st.markdown을 한 번만 감싸서
# unsafe_allow_html=True로 호출될 때 각 줄의 선행 공백을
# 자동으로 제거하도록 패치한다.
# ============================================================

_original_markdown = st.markdown


def _patched_markdown(body, *args, **kwargs):

    if kwargs.get("unsafe_allow_html") and isinstance(body, str):

        body = "\n".join(
            line.lstrip() for line in body.split("\n")
        )

    return _original_markdown(body, *args, **kwargs)


st.markdown = _patched_markdown


# ============================================================
# 3. 수자원 스타일 CSS
# ============================================================

st.markdown(
    """
<style>

/* ========================================================
   전체
======================================================== */

html,
body,
[data-testid="stAppViewContainer"] {

    background:
        linear-gradient(
            180deg,
            #f4f9fc 0%,
            #ffffff 55%
        );

}

.block-container {

    max-width: 1450px;

    padding-top: 1.0rem !important;

    padding-bottom: 2rem !important;

    padding-left: 1.2rem !important;

    padding-right: 1.2rem !important;

}


/* ========================================================
   사이드바
======================================================== */

section[data-testid="stSidebar"] {

    background:
        linear-gradient(
            180deg,
            #063b63 0%,
            #075985 45%,
            #087ea4 100%
        );

}

section[data-testid="stSidebar"] > div {

    background: transparent;

}

section[data-testid="stSidebar"] * {

    color: white !important;

}

section[data-testid="stSidebar"] .stButton > button {

    background: rgba(255,255,255,0.08);

    color: white !important;

    border: 1px solid rgba(255,255,255,0.12);

    border-radius: 9px;

    text-align: left;

    padding: 9px 12px;

    margin-bottom: 4px;

}

section[data-testid="stSidebar"] .stButton > button:hover {

    background: rgba(255,255,255,0.20);

    border-color: rgba(255,255,255,0.35);

}


/* ========================================================
   상단 헤더
======================================================== */

.top-header {

    background:
        linear-gradient(
            135deg,
            #063b63 0%,
            #087ea4 60%,
            #11a6c9 100%
        );

    border-radius: 18px;

    padding: 22px 26px;

    color: white;

    margin-bottom: 18px;

    box-shadow:
        0 8px 25px rgba(3, 65, 100, 0.12);

}

.top-title {

    font-size: 1.55rem;

    font-weight: 800;

    letter-spacing: -0.5px;

}

.top-sub {

    font-size: 0.85rem;

    opacity: 0.88;

    margin-top: 5px;

}


/* ========================================================
   섹션
======================================================== */

.section-title {

    font-size: 1.12rem;

    font-weight: 800;

    color: #0f3552;

    margin-top: 8px;

    margin-bottom: 10px;

}

.section-caption {

    font-size: 0.80rem;

    color: #64748b;

    margin-bottom: 12px;

}


/* ========================================================
   KPI
======================================================== */

.kpi-grid {

    display: grid;

    grid-template-columns:
        repeat(5, 1fr);

    gap: 12px;

    margin-bottom: 18px;

}

.kpi-card {

    background: white;

    border-radius: 14px;

    padding: 15px;

    border:
        1px solid #dce8ef;

    box-shadow:
        0 3px 12px rgba(10, 70, 100, 0.06);

}

.kpi-label {

    font-size: 0.74rem;

    color: #64748b;

    font-weight: 700;

}

.kpi-value {

    font-size: 1.45rem;

    font-weight: 800;

    color: #083b5c;

    margin-top: 4px;

}

.kpi-sub {

    font-size: 0.68rem;

    color: #94a3b8;

    margin-top: 3px;

}


/* ========================================================
   카드
======================================================== */

.platform-card {

    background: white;

    border: 1px solid #dce8ef;

    border-radius: 15px;

    padding: 18px;

    margin-bottom: 14px;

    box-shadow:
        0 3px 14px rgba(10, 70, 100, 0.055);

}

.card-title {

    color: #0f3552;

    font-weight: 800;

    font-size: 0.98rem;

    margin-bottom: 8px;

}


/* ========================================================
   상태
======================================================== */

.status-normal {

    display: inline-block;

    background: #e7f8f1;

    color: #087f5b;

    border-radius: 999px;

    padding: 4px 10px;

    font-size: 0.72rem;

    font-weight: 800;

}

.status-watch {

    display: inline-block;

    background: #fff7df;

    color: #a16207;

    border-radius: 999px;

    padding: 4px 10px;

    font-size: 0.72rem;

    font-weight: 800;

}

.status-danger {

    display: inline-block;

    background: #fff0f0;

    color: #c62828;

    border-radius: 999px;

    padding: 4px 10px;

    font-size: 0.72rem;

    font-weight: 800;

}


/* ========================================================
   메뉴
======================================================== */

.menu-caption {

    font-size: 0.68rem;

    opacity: 0.65;

    margin-top: 18px;

    margin-bottom: 4px;

    letter-spacing: 0.5px;

}


/* ========================================================
   버튼
======================================================== */

.stButton > button {

    border-radius: 9px;

    font-weight: 700;

    min-height: 38px;

}


/* ========================================================
   탭
======================================================== */

button[data-baseweb="tab"] {

    font-weight: 700;

}


/* ========================================================
   모바일
======================================================== */

@media (max-width: 900px) {

    .kpi-grid {

        grid-template-columns:
            repeat(2, 1fr);

    }

    .top-title {

        font-size: 1.2rem;

    }

}

@media (max-width: 600px) {

    .block-container {

        padding-left: 0.65rem !important;

        padding-right: 0.65rem !important;

    }

    .kpi-grid {

        grid-template-columns:
            repeat(2, 1fr);

        gap: 7px;

    }

    .kpi-card {

        padding: 11px;

    }

    .kpi-value {

        font-size: 1.1rem;

    }

}

</style>
""",
    unsafe_allow_html=True
)


# ============================================================
# 4. 설비 기본 DB
# ============================================================

DEFAULT_PUMPS = [

    {
        "site": "밀양정수장",
        "equip": "가압펌프 #1",
        "maker": "효성펌프",
        "model": "DHP-1",
        "hp": 160,
        "head": 47,
        "flow": 1250,
        "rpm": 1780,
        "build_date": "2018-01-15",
        "op_hours": 10200
    },

    {
        "site": "밀양정수장",
        "equip": "가압펌프 #2",
        "maker": "현대중공업",
        "model": "VTP-20",
        "hp": 170,
        "head": 49,
        "flow": 1300,
        "rpm": 1780,
        "build_date": "2018-02-15",
        "op_hours": 9200
    },

    {
        "site": "밀양정수장",
        "equip": "가압펌프 #3",
        "maker": "효성펌프",
        "model": "DHP-3",
        "hp": 180,
        "head": 51,
        "flow": 1350,
        "rpm": 1780,
        "build_date": "2018-03-15",
        "op_hours": 7800
    },

    {
        "site": "밀양정수장",
        "equip": "가압펌프 #4",
        "maker": "현대중공업",
        "model": "VTP-40",
        "hp": 190,
        "head": 53,
        "flow": 1400,
        "rpm": 1780,
        "build_date": "2018-04-15",
        "op_hours": 10500
    },

    {
        "site": "밀양정수장",
        "equip": "가압펌프 #5",
        "maker": "효성펌프",
        "model": "DHP-5",
        "hp": 200,
        "head": 55,
        "flow": 1450,
        "rpm": 1780,
        "build_date": "2018-05-15",
        "op_hours": 9800
    },

    {
        "site": "밀양정수장",
        "equip": "가압펌프 #6",
        "maker": "현대중공업",
        "model": "VTP-60",
        "hp": 210,
        "head": 57,
        "flow": 1500,
        "rpm": 1780,
        "build_date": "2018-06-15",
        "op_hours": 8700
    },

    {
        "site": "밀양정수장",
        "equip": "가압펌프 #7",
        "maker": "효성펌프",
        "model": "DHP-7",
        "hp": 220,
        "head": 59,
        "flow": 1550,
        "rpm": 1780,
        "build_date": "2018-07-15",
        "op_hours": 11400
    },

    {
        "site": "밀양정수장",
        "equip": "가압펌프 #8",
        "maker": "현대중공업",
        "model": "VTP-80",
        "hp": 230,
        "head": 61,
        "flow": 1600,
        "rpm": 1780,
        "build_date": "2018-08-15",
        "op_hours": 7200
    },

    {
        "site": "밀양정수장",
        "equip": "가압펌프 #9",
        "maker": "효성펌프",
        "model": "DHP-9",
        "hp": 240,
        "head": 63,
        "flow": 1650,
        "rpm": 1780,
        "build_date": "2018-09-15",
        "op_hours": 12300
    },

    {
        "site": "밀양정수장",
        "equip": "가압펌프 #10",
        "maker": "현대중공업",
        "model": "VTP-100",
        "hp": 250,
        "head": 65,
        "flow": 1700,
        "rpm": 1780,
        "build_date": "2018-10-15",
        "op_hours": 9100
    }

]


# ============================================================
# 5. 진단 기준
# ============================================================

CATEGORIES = {

    "성능": 40,

    "내부상태": 27,

    "기계상태": 25,

    "정비이력": 5

}


def calc_eff(val):

    if val >= 98:
        return "A+"

    if val >= 96:
        return "A"

    if val >= 94:
        return "A-"

    if val >= 92:
        return "B+"

    if val >= 90:
        return "B"

    if val >= 88:
        return "B-"

    if val >= 86:
        return "C+"

    if val >= 84:
        return "C"

    if val >= 82:
        return "C-"

    if val >= 80:
        return "D+"

    if val >= 75:
        return "D"

    return "E"


def calc_reach(val):

    if val >= 98:
        return "A"

    if val >= 93:
        return "B"

    if val >= 88:
        return "C"

    if val >= 80:
        return "D"

    return "E"


def calc_bep(val):

    if 85 <= val <= 115:
        return "A"

    if 75 <= val <= 125:
        return "B"

    if 65 <= val <= 135:
        return "C"

    if 50 <= val <= 150:
        return "D"

    return "E"


def calc_ring_gap(val):

    if val < 1.5:
        return "A"

    if val < 2:
        return "B"

    if val < 2.5:
        return "C"

    if val < 3:
        return "D"

    return "E"


def calc_sleeve(val):

    if val < 1:
        return "A"

    if val < 1.8:
        return "B"

    if val < 2.5:
        return "C"

    if val < 3:
        return "D"

    return "E"


def calc_vib(val):

    if val < 1.8:
        return "A"

    if val < 4.5:
        return "B"

    if val < 7.1:
        return "C"

    if val < 11.2:
        return "D"

    return "E"


def calc_align(val):

    if val <= 0.05:
        return "A"

    if val <= 0.08:
        return "B"

    if val <= 0.12:
        return "C"

    if val <= 0.15:
        return "D"

    return "E"


def calc_overhaul(val):

    if val <= 10000:
        return "A"

    if val <= 12000:
        return "B"

    if val <= 15000:
        return "C"

    if val <= 20000:
        return "D"

    return "E"


# ============================================================
# 6. 17개 핵심 진단항목
# ============================================================

EVAL_ITEMS = [

    (
        "성능",
        "펌프 효율 유지율 (%)",
        25,
        "98% 이상",
        ["A+", "A", "A-", "B+", "B", "B-", "C+", "C", "C-", "D+", "D", "E"],
        {
            "A+": 25,
            "A": 24,
            "A-": 23,
            "B+": 22,
            "B": 21,
            "B-": 20,
            "C+": 18.5,
            "C": 17,
            "C-": 15.5,
            "D+": 14,
            "D": 11,
            "E": 7
        },
        calc_eff
    ),

    (
        "성능",
        "설계 양정/유량 도달률 (%)",
        8,
        "98% 이상",
        ["A", "B", "C", "D", "E"],
        {
            "A": 8,
            "B": 6.8,
            "C": 5.6,
            "D": 4.4,
            "E": 2.4
        },
        calc_reach
    ),

    (
        "성능",
        "BEP 운전점 적정성 (%)",
        7,
        "85~115%",
        ["A", "B", "C", "D", "E"],
        {
            "A": 7,
            "B": 5.95,
            "C": 4.9,
            "D": 3.85,
            "E": 2.1
        },
        calc_bep
    ),

    (
        "내부상태",
        "임펠러/케이싱 링 간극",
        10,
        "1.5 mm 미만",
        ["A", "B", "C", "D", "E"],
        {
            "A": 10,
            "B": 8,
            "C": 6.5,
            "D": 5,
            "E": 3
        },
        calc_ring_gap
    ),

    (
        "내부상태",
        "축슬리브 마모",
        5,
        "1.0 mm 미만",
        ["A", "B", "C", "D", "E"],
        {
            "A": 5,
            "B": 4.25,
            "C": 3.5,
            "D": 2.75,
            "E": 1.5
        },
        calc_sleeve
    ),

    (
        "내부상태",
        "임펠러 손상/침식",
        2,
        "균열 및 침식 없음",
        ["A", "B", "C", "D", "E"],
        {
            "A": 2,
            "B": 1.7,
            "C": 1.4,
            "D": 1.1,
            "E": 0.6
        },
        None
    ),

    (
        "내부상태",
        "임펠러 동적 밸런싱",
        2,
        "ISO G2.5 이하",
        ["A", "B", "C", "D", "E"],
        {
            "A": 2,
            "B": 1.7,
            "C": 1.4,
            "D": 1.1,
            "E": 0.6
        },
        None
    ),

    (
        "내부상태",
        "NPSH 여유율/캐비테이션",
        5,
        "NPSHa/NPSHr ≥ 1.3",
        ["A", "B", "C", "D", "E"],
        {
            "A": 5,
            "B": 4.25,
            "C": 3.5,
            "D": 2.75,
            "E": 1.5
        },
        None
    ),

    (
        "내부상태",
        "내부 코팅 상태",
        3,
        "박리 없음",
        ["A", "B", "C", "D", "E"],
        {
            "A": 3,
            "B": 2.55,
            "C": 2.1,
            "D": 1.65,
            "E": 0.9
        },
        None
    ),

    (
        "내부상태",
        "비금속 웨어링 개선",
        3,
        "복합소재 적용 및 간극 개선",
        ["A", "B", "C", "D", "E"],
        {
            "A": 3,
            "B": 2.55,
            "C": 2.1,
            "D": 1.65,
            "E": 0.9
        },
        None
    ),

    (
        "기계상태",
        "Overall 진동 (mm/s)",
        7,
        "1.8 mm/s 미만",
        ["A", "B", "C", "D", "E"],
        {
            "A": 7,
            "B": 5.95,
            "C": 4.9,
            "D": 3.85,
            "E": 2.1
        },
        calc_vib
    ),

    (
        "기계상태",
        "베어링 결함 진동",
        5,
        "결함 신호 없음",
        ["A", "B", "C", "D", "E"],
        {
            "A": 5,
            "B": 4.25,
            "C": 3.5,
            "D": 2.75,
            "E": 1.5
        },
        None
    ),

    (
        "기계상태",
        "주파수 성분 결함",
        3,
        "특이 피크 없음",
        ["A", "B", "C", "D", "E"],
        {
            "A": 3,
            "B": 2.55,
            "C": 2.1,
            "D": 1.65,
            "E": 0.9
        },
        None
    ),

    (
        "기계상태",
        "펌프-모터 센터링",
        7,
        "0.05 mm 이하",
        ["A", "B", "C", "D", "E"],
        {
            "A": 7,
            "B": 5.95,
            "C": 4.9,
            "D": 3.85,
            "E": 2.1
        },
        calc_align
    ),

    (
        "기계상태",
        "Soft Foot 및 배관 응력",
        3,
        "0.05 mm 이하",
        ["A", "B", "C", "D", "E"],
        {
            "A": 3,
            "B": 2.55,
            "C": 2.1,
            "D": 1.65,
            "E": 0.9
        },
        None
    ),

    (
        "정비이력",
        "오버홀 주기",
        3,
        "10,000시간 이내",
        ["A", "B", "C", "D", "E"],
        {
            "A": 3,
            "B": 2.55,
            "C": 2.1,
            "D": 1.65,
            "E": 0.9
        },
        calc_overhaul
    ),

    (
        "정비이력",
        "주요 소모품 교체이력",
        2,
        "주기 및 이력 양호",
        ["A", "B", "C", "D", "E"],
        {
            "A": 2,
            "B": 1.7,
            "C": 1.4,
            "D": 1.1,
            "E": 0.6
        },
        None
    )

]


# ============================================================
# 7. 종합등급
# ============================================================

def get_final_grade(score):

    if score >= 90:
        return "A"

    if score >= 80:
        return "B"

    if score >= 70:
        return "C"

    if score >= 60:
        return "D"

    return "E"


def get_grade_text(grade):

    data = {

        "A": "매우 우수 · 정상운전",

        "B": "우수 · 계획관리",

        "C": "관찰 · 추이관리",

        "D": "개선 필요 · 정비검토",

        "E": "위험 · 즉시정비"

    }

    return data.get(
        grade,
        "판정 필요"
    )


# ============================================================
# 8. Excel DB 생성
# ============================================================

def ensure_excel_file(
    path,
    sheet,
    headers
):

    if not os.path.exists(path):

        wb = Workbook()

        ws = wb.active

        ws.title = sheet

        ws.append(headers)

        wb.save(path)

        wb.close()


def ensure_db_exists():

    ensure_excel_file(

        DB_FILE_PATH,

        "진단이력",

        [
            "점검일",
            "사업장",
            "설비명",
            "제조사",
            "모델명",
            "마력",
            "양정",
            "준공일",
            "점검자",
            "종합점수",
            "최종등급"
        ]
        +
        [
            item[1]
            for item in EVAL_ITEMS
        ]

    )

    ensure_excel_file(

        OVERHAUL_DB_PATH,

        "오버홀이력",

        [
            "작업일자",
            "사업장",
            "설비명",
            "공정단계",
            "작업자",
            "작업내용",
            "사진파일명",
            "전후효율",
            "전후진동"
        ]

    )

    ensure_excel_file(

        DOC_DB_PATH,

        "전문보고서",

        [
            "등록일자",
            "사업장",
            "설비명",
            "보고서구분",
            "수행기관",
            "요약소견",
            "저장파일명"
        ]

    )

    ensure_excel_file(

        KNOWHOW_DB_PATH,

        "노하우DB",

        [
            "등록일자",
            "분류",
            "관련모델",
            "현상및원인",
            "해결노하우",
            "작성자"
        ]

    )

    ensure_excel_file(

        DAILY_LOG_DB_PATH,

        "점검일지",

        [
            "점검일자",
            "점검자",
            "설비명",
            "누수",
            "진동",
            "온도",
            "전류",
            "특이사항"
        ]

    )

    ensure_excel_file(

        SAFETY_PERMIT_DB_PATH,

        "위험작업허가",

        [
            "신청일자",
            "작업명",
            "대상설비",
            "위험유형",
            "작업기간",
            "신청자",
            "승인상태",
            "안전조치"
        ]

    )

    ensure_excel_file(

        KPI_DB_PATH,

        "KPI실적",

        [
            "등록일자",
            "지표",
            "목표",
            "실적",
            "단위",
            "비고"
        ]

    )


ensure_db_exists()


# ============================================================
# 9. 샘플 데이터
# ============================================================

def seed_sample_data():

    if os.path.exists(SEED_FLAG_PATH):

        return

    wb = load_workbook(DB_FILE_PATH)

    ws = wb["진단이력"]

    dates = [

        "2024-03-15",
        "2024-09-10",
        "2025-03-20",
        "2025-08-15",
        "2026-02-10",
        "2026-08-15"

    ]

    for pump in DEFAULT_PUMPS:

        for dt in dates:

            score = round(
                random.uniform(
                    72,
                    96
                ),
                1
            )

            grade = get_final_grade(
                score
            )

            grades = [

                random.choice(
                    [
                        "A",
                        "B",
                        "B",
                        "C"
                    ]
                )

                for _ in EVAL_ITEMS
            ]

            ws.append(

                [
                    dt,
                    pump["site"],
                    pump["equip"],
                    pump["maker"],
                    pump["model"],
                    pump["hp"],
                    pump["head"],
                    pump["build_date"],
                    "최진욱",
                    score,
                    grade
                ]
                +
                grades

            )

    wb.save(DB_FILE_PATH)

    wb.close()

    with open(
        SEED_FLAG_PATH,
        "w",
        encoding="utf-8"
    ) as f:

        f.write(
            datetime.now().isoformat()
        )


seed_sample_data()


# ============================================================
# 10. 데이터 읽기
# ============================================================

def read_excel(
    path,
    sheet
):

    if not os.path.exists(path):

        return pd.DataFrame()

    try:

        return pd.read_excel(
            path,
            sheet_name=sheet
        )

    except Exception:

        return pd.DataFrame()


df_history = read_excel(
    DB_FILE_PATH,
    "진단이력"
)


# ============================================================
# 11. 상태 계산
# ============================================================

def pump_status(
    pump
):

    name = pump["equip"]

    score = 85

    vibration = 2.2

    efficiency = 88

    hours = pump["op_hours"]

    if name == "가압펌프 #2":

        score = 68

        vibration = 5.8

        efficiency = 73

    elif name == "가압펌프 #3":

        score = 64

        vibration = 6.8

        efficiency = 68

    elif name == "가압펌프 #5":

        score = 74

        vibration = 4.2

        efficiency = 75

    elif hours > 11000:

        score = 77

        vibration = 3.8

        efficiency = 81

    elif hours > 10000:

        score = 82

        vibration = 2.8

        efficiency = 84

    grade = get_final_grade(
        score
    )

    if grade in ["A", "B"]:

        status = "정상"

    elif grade == "C":

        status = "관찰"

    else:

        status = "정비검토"

    return {

        "점수": score,

        "등급": grade,

        "상태": status,

        "진동": vibration,

        "효율": efficiency

    }


# ============================================================
# 12. QR 비용 계산
# ============================================================

QR_UNIT_COST = {

    "QR 라벨 제작": 2500,

    "표면 세척/준비": 1000,

    "QR 부착 작업": 3000,

    "설비 등록 및 검수": 1500

}


def calculate_qr_cost(
    count
):

    unit = sum(
        QR_UNIT_COST.values()
    )

    return unit, unit * count


# ============================================================
# 13. 사이드바
# ============================================================

if "page" not in st.session_state:

    st.session_state.page = "홈"


with st.sidebar:

    st.markdown(
        """
        <div style="
        text-align:center;
        padding:8px 0 18px 0;
        ">
            <div style="
            font-size:2rem;
            ">
            💧
            </div>
            <div style="
            font-size:1.05rem;
            font-weight:800;
            ">
            K-water tech
            </div>
            <div style="
            font-size:0.72rem;
            opacity:0.7;
            ">
            설비관리 통합 플랫폼
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        "<div class='menu-caption'>MAIN</div>",
        unsafe_allow_html=True
    )

    main_menus = [

        ("홈", "🏠 설비관리 홈"),

        ("설비", "🏭 설비 관리"),

        ("QR", "📱 QR 설비 포털"),

        ("진단", "🔍 정밀 진단"),

        ("CBM", "🎯 CBM 정비판단"),

        ("오버홀", "🛠️ 오버홀 관리")

    ]

    for key, label in main_menus:

        if st.button(
            label,
            key=f"menu_{key}",
            use_container_width=True
        ):

            st.session_state.page = key

            st.rerun()

    st.markdown(
        "<div class='menu-caption'>ANALYSIS</div>",
        unsafe_allow_html=True
    )

    analysis_menus = [

        ("AI", "📈 AI 이상징후"),

        ("ROI", "💰 정비효과·ROI"),

        ("KPI", "📊 성과관리")

    ]

    for key, label in analysis_menus:

        if st.button(
            label,
            key=f"menu_{key}",
            use_container_width=True
        ):

            st.session_state.page = key

            st.rerun()

    st.markdown(
        "<div class='menu-caption'>KNOWLEDGE</div>",
        unsafe_allow_html=True
    )

    knowledge_menus = [

        ("노하우", "💡 기술 노하우"),

        ("보고서", "📄 진단 보고서"),

        ("백업", "💾 데이터 관리")

    ]

    for key, label in knowledge_menus:

        if st.button(
            label,
            key=f"menu_{key}",
            use_container_width=True
        ):

            st.session_state.page = key

            st.rerun()

    st.markdown("---")

    st.caption(
        "최진욱 · 정밀진단원 / 관리자"
    )

    st.caption(
        "밀양정수장"
    )


# ============================================================
# 14. 공통 헤더
# ============================================================

st.markdown(
    f"""
    <div class="top-header">

        <div class="top-title">
            💧 K-water tech 설비관리 통합 플랫폼
        </div>

        <div class="top-sub">
            QR 기반 설비정보 · 상태진단 · CBM 정비판단 ·
            오버홀 · 이력관리 · 데이터 기반 설비관리
        </div>

    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# 15. 홈
# ============================================================

if st.session_state.page == "홈":

    st.markdown(
        """
        <div class="section-title">
        설비관리 현황
        </div>

        <div class="section-caption">
        현장의 설비 상태를 한 화면에서 확인하고
        정비 우선순위를 판단합니다.
        </div>
        """,
        unsafe_allow_html=True
    )

    total = len(
        DEFAULT_PUMPS
    )

    normal = 0
    watch = 0
    repair = 0

    for pump in DEFAULT_PUMPS:

        result = pump_status(
            pump
        )

        if result["상태"] == "정상":

            normal += 1

        elif result["상태"] == "관찰":

            watch += 1

        else:

            repair += 1

    st.markdown(
        f"""
        <div class="kpi-grid">

            <div class="kpi-card">
                <div class="kpi-label">관리 설비</div>
                <div class="kpi-value">{total}대</div>
                <div class="kpi-sub">등록 설비</div>
            </div>

            <div class="kpi-card">
                <div class="kpi-label">정상</div>
                <div class="kpi-value">{normal}대</div>
                <div class="kpi-sub">정상 운전</div>
            </div>

            <div class="kpi-card">
                <div class="kpi-label">관찰</div>
                <div class="kpi-value">{watch}대</div>
                <div class="kpi-sub">추이관리</div>
            </div>

            <div class="kpi-card">
                <div class="kpi-label">정비 검토</div>
                <div class="kpi-value">{repair}대</div>
                <div class="kpi-sub">CBM 우선관리</div>
            </div>

            <div class="kpi-card">
                <div class="kpi-label">QR 관리</div>
                <div class="kpi-value">{total}개</div>
                <div class="kpi-sub">설비별 1개 기준</div>
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(
        [1.45, 1]
    )

    with col1:

        st.markdown(
            """
            <div class="platform-card">

            <div class="card-title">
            🏭 설비 건전성 현황
            </div>

            """
            ,
            unsafe_allow_html=True
        )

        rows = []

        for pump in DEFAULT_PUMPS:

            result = pump_status(
                pump
            )

            rows.append(

                {
                    "설비": pump["equip"],
                    "운전시간": f'{pump["op_hours"]:,} h',
                    "효율": f'{result["효율"]:.1f}%',
                    "진동": f'{result["진동"]:.1f} mm/s',
                    "CBM": result["점수"],
                    "등급": result["등급"],
                    "판정": result["상태"]
                }

            )

        st.dataframe(
            pd.DataFrame(rows),
            use_container_width=True,
            hide_index=True
        )

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )

    with col2:

        st.markdown(
            """
            <div class="platform-card">

            <div class="card-title">
            🎯 CBM 정비 우선순위
            </div>

            """,
            unsafe_allow_html=True
        )

        ranking = []

        for pump in DEFAULT_PUMPS:

            result = pump_status(
                pump
            )

            ranking.append(

                [
                    pump["equip"],
                    result["점수"],
                    result["등급"],
                    pump["op_hours"]
                ]

            )

        ranking = sorted(
            ranking,
            key=lambda x: x[1]
        )

        ranking_df = pd.DataFrame(

            ranking[:5],

            columns=[
                "설비",
                "CBM Score",
                "등급",
                "운전시간"
            ]

        )

        st.dataframe(
            ranking_df,
            use_container_width=True,
            hide_index=True
        )

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )

    st.markdown(
        """
        <div class="platform-card">

        <div class="card-title">
        🔄 플랫폼 운영 흐름
        </div>

        <div style="
        display:flex;
        flex-wrap:wrap;
        gap:8px;
        align-items:center;
        font-size:0.82rem;
        color:#164e63;
        font-weight:700;
        ">

        <span>📱 QR 접속</span>
        →
        <span>🏭 설비정보</span>
        →
        <span>🔍 상태진단</span>
        →
        <span>🎯 CBM Score</span>
        →
        <span>📌 정비 Ranking</span>
        →
        <span>🛠️ 오버홀</span>
        →
        <span>📈 효과검증</span>
        →
        <span>📚 이력축적</span>

        </div>

        </div>
        """,
        unsafe_allow_html=True
    )


# ============================================================
# 16. 설비 관리
# ============================================================

elif st.session_state.page == "설비":

    st.markdown(
        """
        <div class="section-title">
        🏭 설비 관리
        </div>

        <div class="section-caption">
        관리 중인 펌프의 기본 제원과 운전상태를 통합 관리합니다.
        </div>
        """,
        unsafe_allow_html=True
    )

    selected = st.selectbox(

        "설비 선택",

        [
            p["equip"]
            for p in DEFAULT_PUMPS
        ]

    )

    pump = next(
        p for p in DEFAULT_PUMPS
        if p["equip"] == selected
    )

    result = pump_status(
        pump
    )

    a, b, c, d = st.columns(4)

    a.metric(
        "CBM Score",
        f'{result["점수"]}점'
    )

    b.metric(
        "상태",
        result["상태"]
    )

    c.metric(
        "현재 효율",
        f'{result["효율"]:.1f}%'
    )

    d.metric(
        "진동",
        f'{result["진동"]:.1f} mm/s'
    )

    st.write("")

    tab1, tab2, tab3 = st.tabs(

        [
            "설비 기본정보",
            "운전정보",
            "정비이력"
        ]

    )

    with tab1:

        c1, c2 = st.columns(2)

        c1.text_input(
            "사업장",
            pump["site"],
            disabled=True
        )

        c2.text_input(
            "설비명",
            pump["equip"],
            disabled=True
        )

        c1.text_input(
            "제조사",
            pump["maker"],
            disabled=True
        )

        c2.text_input(
            "모델",
            pump["model"],
            disabled=True
        )

        c1.number_input(
            "정격출력(HP)",
            value=pump["hp"],
            disabled=True
        )

        c2.number_input(
            "정격양정(m)",
            value=pump["head"],
            disabled=True
        )

        c1.number_input(
            "정격유량(m³/h)",
            value=pump["flow"],
            disabled=True
        )

        c2.number_input(
            "회전수(RPM)",
            value=pump["rpm"],
            disabled=True
        )

    with tab2:

        st.dataframe(

            pd.DataFrame(
                [
                    [
                        "누적 운전시간",
                        f'{pump["op_hours"]:,} h'
                    ],
                    [
                        "최근 진동",
                        f'{result["진동"]:.1f} mm/s'
                    ],
                    [
                        "현재 효율",
                        f'{result["효율"]:.1f}%'
                    ],
                    [
                        "CBM 상태",
                        result["상태"]
                    ]
                ],
                columns=[
                    "항목",
                    "현재값"
                ]
            ),

            use_container_width=True,

            hide_index=True

        )

    with tab3:

        st.info(
            "최근 오버홀 : "
            "2025-03-15 / 축슬리브 교체 / "
            "센터링 완료"
        )

        st.info(
            "다음 정비판단 : CBM Score 및 "
            "운전시간을 함께 반영"
        )


# ============================================================
# 17. QR 설비 포털
# ============================================================

elif st.session_state.page == "QR":

    st.markdown(
        """
        <div class="section-title">
        📱 QR 설비 디지털 포털
        </div>

        <div class="section-caption">
        펌프 옆 QR코드를 스마트폰으로 스캔하면
        해당 설비의 정보·진단·정비이력을 바로 확인합니다.
        </div>
        """,
        unsafe_allow_html=True
    )

    selected = st.selectbox(

        "설비 선택",

        [
            p["equip"]
            for p in DEFAULT_PUMPS
        ]

    )

    pump = next(
        p for p in DEFAULT_PUMPS
        if p["equip"] == selected
    )

    result = pump_status(
        pump
    )

    c1, c2 = st.columns(
        [1, 2]
    )

    with c1:

        st.markdown(
            """
            <div style="
            background:white;
            border:1px solid #dce8ef;
            border-radius:16px;
            padding:25px;
            text-align:center;
            ">
            <div style="
            font-size:5rem;
            ">
            ▦
            </div>

            <b>설비 QR CODE</b>

            <div style="
            margin-top:8px;
            color:#64748b;
            font-size:0.75rem;
            ">
            PUMP-MLY-001
            </div>

            </div>
            """,
            unsafe_allow_html=True
        )

    with c2:

        st.markdown(
            f"""
            <div class="platform-card">

            <div class="card-title">
            {pump["equip"]} 디지털 설비카드
            </div>

            <b>사업장</b> : {pump["site"]}<br>
            <b>제조사</b> : {pump["maker"]}<br>
            <b>모델</b> : {pump["model"]}<br>
            <b>출력</b> : {pump["hp"]} HP<br>
            <b>운전시간</b> : {pump["op_hours"]:,} h<br>
            <b>현재 효율</b> : {result["효율"]:.1f}%<br>
            <b>현재 진동</b> : {result["진동"]:.1f} mm/s<br>
            <b>CBM Score</b> : {result["점수"]}점<br>
            <b>상태</b> : {result["상태"]}

            </div>
            """,
            unsafe_allow_html=True
        )

    st.write("")

    t1, t2, t3, t4, t5 = st.tabs(

        [
            "설비정보",
            "진단이력",
            "효율진단",
            "오버홀",
            "작업기록"
        ]

    )

    with t1:

        st.success(
            "QR 접속 성공 · 설비정보 확인 가능"
        )

    with t2:

        st.info(
            "최근 정밀진단 : 2026-08-15"
        )

        st.info(
            f"현재 CBM 등급 : {result['등급']}"
        )

    with t3:

        st.info(
            "효율진단은 설비 종합상태 판단을 위한 "
            "입력 데이터 중 하나로 활용됩니다."
        )

        st.metric(
            "최근 효율",
            f'{result["효율"]:.1f}%'
        )

    with t4:

        st.info(
            "최근 오버홀 : 2025-03-15"
        )

    with t5:

        st.text_area(
            "현장 작업 메모",
            placeholder="점검 및 작업 내용을 입력하세요."
        )

        if st.button(
            "작업기록 저장",
            type="primary"
        ):

            st.success(
                "현장 작업기록이 저장되었습니다."
            )


# ============================================================
# 18. 정밀진단
# ============================================================

elif st.session_state.page == "진단":

    st.markdown(
        """
        <div class="section-title">
        🔍 펌프 정밀 상태진단
        </div>

        <div class="section-caption">
        17개 핵심 진단항목을 기반으로 설비 상태를 점수화하고
        CBM 정비판단에 활용합니다.
        </div>
        """,
        unsafe_allow_html=True
    )

    selected = st.selectbox(

        "진단 대상",

        [
            p["equip"]
            for p in DEFAULT_PUMPS
        ]

    )

    pump = next(
        p for p in DEFAULT_PUMPS
        if p["equip"] == selected
    )

    st.info(
        f'{pump["site"]} / '
        f'{pump["equip"]} / '
        f'{pump["maker"]} {pump["model"]}'
    )

    total_score = 0

    category_score = {
        "성능": 0,
        "내부상태": 0,
        "기계상태": 0,
        "정비이력": 0
    }

    details = []

    for idx, item in enumerate(
        EVAL_ITEMS
    ):

        category = item[0]

        name = item[1]

        weight = item[2]

        standard = item[3]

        options = item[4]

        score_map = item[5]

        auto_fn = item[6]

        with st.expander(
            f"{idx + 1}. {name} · {weight}점",
            expanded=False
        ):

            c1, c2 = st.columns(
                [2, 1]
            )

            with c1:

                st.caption(
                    f"판정기준 : {standard}"
                )

            with c2:

                grade = st.selectbox(

                    "판정",

                    options,

                    key=f"grade_{idx}"

                )

            score = score_map[
                grade
            ]

            category_score[
                category
            ] += score

            total_score += score

            details.append(
                {
                    "항목": name,
                    "등급": grade,
                    "점수": score
                }
            )

    total_score = round(
        total_score,
        2
    )

    final_grade = get_final_grade(
        total_score
    )

    st.write("---")

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "종합점수",
        f"{total_score}점"
    )

    c2.metric(
        "최종등급",
        final_grade
    )

    c3.metric(
        "판정",
        get_grade_text(
            final_grade
        )
    )

    st.markdown(
        "### 분야별 상태"
    )

    cat_df = pd.DataFrame(

        [
            [
                key,
                round(
                    value,
                    1
                ),
                CATEGORIES[key]
            ]

            for key, value
            in category_score.items()
        ],

        columns=[
            "분야",
            "획득점수",
            "배점"
        ]

    )

    st.dataframe(
        cat_df,
        use_container_width=True,
        hide_index=True
    )

    if st.button(
        "💾 진단결과 저장",
        type="primary",
        use_container_width=True
    ):

        wb = load_workbook(
            DB_FILE_PATH
        )

        ws = wb[
            "진단이력"
        ]

        row = [

            datetime.now().strftime(
                "%Y-%m-%d"
            ),

            pump["site"],

            pump["equip"],

            pump["maker"],

            pump["model"],

            pump["hp"],

            pump["head"],

            pump["build_date"],

            "최진욱",

            total_score,

            final_grade

        ] + [

            x["등급"]
            for x in details
        ]

        ws.append(row)

        wb.save(
            DB_FILE_PATH
        )

        wb.close()

        st.success(
            f"{pump['equip']} "
            f"진단결과가 저장되었습니다."
        )


# ============================================================
# 19. CBM
# ============================================================

elif st.session_state.page == "CBM":

    st.markdown(
        """
        <div class="section-title">
        🎯 CBM 정비판단
        </div>

        <div class="section-caption">
        운전시간만으로 오버홀을 결정하지 않고
        성능·진동·내부상태·정비이력을 종합하여
        정비 우선순위를 판단합니다.
        </div>
        """,
        unsafe_allow_html=True
    )

    ranking = []

    for pump in DEFAULT_PUMPS:

        result = pump_status(
            pump
        )

        risk = round(
            100 - result["점수"],
            1
        )

        ranking.append(

            {
                "우선순위": 0,
                "설비": pump["equip"],
                "CBM Score": result["점수"],
                "Risk": risk,
                "운전시간": pump["op_hours"],
                "효율": result["효율"],
                "진동": result["진동"],
                "등급": result["등급"],
                "정비판단": result["상태"]
            }

        )

    ranking = sorted(
        ranking,
        key=lambda x: x["Risk"],
        reverse=True
    )

    for i, row in enumerate(
        ranking
    ):

        row["우선순위"] = i + 1

    df_rank = pd.DataFrame(
        ranking
    )

    st.dataframe(
        df_rank,
        use_container_width=True,
        hide_index=True
    )

    st.write("")

    top = ranking[0]

    st.warning(
        f"최우선 정비 대상 : "
        f"{top['설비']} / "
        f"CBM Score {top['CBM Score']}점 / "
        f"Risk {top['Risk']}%"
    )

    st.markdown(
        """
        <div class="platform-card">

        <div class="card-title">
        CBM 판단 로직
        </div>

        <b>① 성능상태</b>
        → 효율·양정·유량·BEP

        <br>

        <b>② 내부상태</b>
        → 링간극·축슬리브·임펠러·코팅

        <br>

        <b>③ 기계상태</b>
        → 진동·베어링·센터링·배관응력

        <br>

        <b>④ 정비이력</b>
        → 운전시간·소모품·오버홀 이력

        <br><br>

        <b>⑤ 종합 CBM Score</b>
        → 정비 우선순위 Ranking

        </div>
        """,
        unsafe_allow_html=True
    )


# ============================================================
# 20. 오버홀 관리
# ============================================================

elif st.session_state.page == "오버홀":

    st.markdown(
        """
        <div class="section-title">
        🛠️ 오버홀 관리
        </div>

        <div class="section-caption">
        진단 결과를 실제 정비활동과 연결하고
        오버홀 전·후 상태변화를 데이터로 축적합니다.
        </div>
        """,
        unsafe_allow_html=True
    )

    selected = st.selectbox(

        "작업 대상",

        [
            p["equip"]
            for p in DEFAULT_PUMPS
        ]

    )

    pump = next(
        p for p in DEFAULT_PUMPS
        if p["equip"] == selected
    )

    tab1, tab2, tab3 = st.tabs(

        [
            "작업계획",
            "작업기록",
            "전후 효과"
        ]

    )

    with tab1:

        st.write(
            f"대상설비 : **{pump['equip']}**"
        )

        steps = [

            "분해 및 상태확인",

            "부품 측정 및 마모판정",

            "가공·교체",

            "조립 및 축정렬",

            "시운전",

            "최종 성능확인"

        ]

        for i, step in enumerate(
            steps
        ):

            st.checkbox(
                f"{i + 1}. {step}",
                key=f"oh_{i}"
            )

    with tab2:

        work = st.text_area(
            "작업내용",
            height=150,
            placeholder=
            "분해점검, 마모측정, "
            "부품교체, 센터링 등"
        )

        if st.button(
            "오버홀 작업기록 저장",
            type="primary"
        ):

            wb = load_workbook(
                OVERHAUL_DB_PATH
            )

            ws = wb[
                "오버홀이력"
            ]

            ws.append(

                [
                    datetime.now().strftime(
                        "%Y-%m-%d"
                    ),

                    pump["site"],

                    pump["equip"],

                    "오버홀",

                    "최진욱",

                    work,

                    "",

                    "",

                    ""

                ]

            )

            wb.save(
                OVERHAUL_DB_PATH
            )

            wb.close()

            st.success(
                "오버홀 작업기록이 저장되었습니다."
            )

    with tab3:

        c1, c2 = st.columns(2)

        before_eff = c1.number_input(
            "정비 전 효율 (%)",
            value=72.3
        )

        after_eff = c2.number_input(
            "정비 후 효율 (%)",
            value=77.7
        )

        before_vib = c1.number_input(
            "정비 전 진동",
            value=5.8
        )

        after_vib = c2.number_input(
            "정비 후 진동",
            value=3.7
        )

        st.metric(
            "효율 개선",
            f"{after_eff-before_eff:+.1f}%p"
        )

        st.metric(
            "진동 감소",
            f"{before_vib-after_vib:+.1f} mm/s"
        )


# ============================================================
# 21. AI 이상징후
# ============================================================

elif st.session_state.page == "AI":

    st.markdown(
        """
        <div class="section-title">
        📈 AI 이상징후 및 추세분석
        </div>

        <div class="section-caption">
        축적된 진동·효율·운전시간·정비이력을 기반으로
        설비 상태변화를 시각화하고 향후 정비시점을 판단합니다.
        </div>
        """,
        unsafe_allow_html=True
    )

    selected = st.selectbox(

        "분석 대상",

        [
            p["equip"]
            for p in DEFAULT_PUMPS
        ]

    )

    pump = next(
        p for p in DEFAULT_PUMPS
        if p["equip"] == selected
    )

    result = pump_status(
        pump
    )

    months = [

        "2025.01",
        "2025.06",
        "2025.12",
        "2026.03",
        "2026.08",
        "2026.12"

    ]

    base = result["진동"]

    vibration = [

        max(
            1.2,
            base - 2.5
        ),

        max(
            1.4,
            base - 2
        ),

        max(
            1.6,
            base - 1.4
        ),

        max(
            1.8,
            base - 0.8
        ),

        base,

        base + 1.2

    ]

    fig, ax = plt.subplots(
        figsize=(9, 4)
    )

    ax.plot(
        months,
        vibration,
        marker="o",
        linewidth=2
    )

    ax.axhline(
        4.5,
        linestyle="--",
        label="관찰 기준"
    )

    ax.axhline(
        7.1,
        linestyle=":",
        label="주의 기준"
    )

    ax.set_ylabel(
        "진동 (mm/s)"
    )

    ax.set_title(
        f"{selected} 진동 추세"
    )

    ax.grid(
        alpha=0.2
    )

    ax.legend()

    st.pyplot(
        fig
    )

    if result["진동"] >= 7.1:

        st.error(
            "고위험 상태 · 정비검토 필요"
        )

    elif result["진동"] >= 4.5:

        st.warning(
            "주의 상태 · 추이관찰 및 정밀진단 권고"
        )

    else:

        st.success(
            "현재 진동상태 양호"
        )


# ============================================================
# 22. ROI
# ============================================================

elif st.session_state.page == "ROI":

    st.markdown(
        """
        <div class="section-title">
        💰 정비효과 및 ROI 분석
        </div>

        <div class="section-caption">
        오버홀 전·후 효율 변화와 실제 부하율을 반영하여
        에너지 절감 및 투자회수 효과를 계산합니다.
        </div>
        """,
        unsafe_allow_html=True
    )

    c1, c2 = st.columns(2)

    rated_kw = c1.number_input(
        "모터 정격출력(kW)",
        value=110.0
    )

    load = c2.number_input(
        "평균 부하율(%)",
        value=80.0
    )

    c1, c2 = st.columns(2)

    before = c1.number_input(
        "정비 전 효율(%)",
        value=73.0
    )

    after = c2.number_input(
        "정비 후 효율(%)",
        value=82.0
    )

    hours = st.number_input(
        "연간 운전시간",
        value=6000
    )

    price = st.number_input(
        "전력단가(원/kWh)",
        value=140
    )

    repair = st.number_input(
        "오버홀 비용(원)",
        value=35_000_000
    )

    effective_kw = (

        rated_kw
        *
        load
        /
        100

    )

    saved_kwh = (

        effective_kw
        *
        (
            1 / (before / 100)
            -
            1 / (after / 100)
        )
        *
        hours

    )

    saved_money = (
        saved_kwh
        *
        price
    )

    payback = (

        repair / saved_money

        if saved_money > 0

        else 0

    )

    a, b, c = st.columns(3)

    a.metric(
        "연간 절감전력",
        f"{saved_kwh:,.0f} kWh"
    )

    b.metric(
        "연간 절감액",
        f"{saved_money:,.0f} 원"
    )

    c.metric(
        "투자회수",
        f"{payback:.1f}년"
    )


# ============================================================
# 23. KPI
# ============================================================

elif st.session_state.page == "KPI":

    st.markdown(
        """
        <div class="section-title">
        📊 플랫폼 성과관리
        </div>

        <div class="section-caption">
        PoC와 사업화 과정에서 시스템의 실제 성과를 관리합니다.
        </div>
        """,
        unsafe_allow_html=True
    )

    kpis = [

        [
            "CBM 판정-분해점검 일치율",
            80,
            "%"
        ],

        [
            "TBM 대비 정비비용 절감률",
            10,
            "%"
        ],

        [
            "현장 QR 활용",
            50,
            "건/월"
        ],

        [
            "현장기술자 만족도",
            4,
            "점/5점"
        ],

        [
            "오버홀 전후 효과검증",
            90,
            "%"
        ]

    ]

    for name, target, unit in kpis:

        value = st.number_input(
            name,
            min_value=0.0,
            value=0.0,
            key=f"kpi_{name}"
        )

        st.progress(

            min(
                value / target
                if target
                else 0,
                1
            )

        )

        st.caption(
            f"목표 : {target} {unit}"
        )


# ============================================================
# 24. 기술 노하우
# ============================================================

elif st.session_state.page == "노하우":

    st.markdown(
        """
        <div class="section-title">
        💡 기술 노하우 DB
        </div>

        <div class="section-caption">
        현장에서 축적된 정비경험과 결함원인·조치방법을
        조직의 기술자산으로 관리합니다.
        </div>
        """,
        unsafe_allow_html=True
    )

    wb = load_workbook(
        KNOWHOW_DB_PATH,
        data_only=True
    )

    ws = wb[
        "노하우DB"
    ]

    data = list(
        ws.values
    )

    wb.close()

    if len(data) > 1:

        df = pd.DataFrame(
            data[1:],
            columns=data[0]
        )

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True
        )

    st.write("")

    with st.expander(
        "➕ 새로운 기술 노하우 등록"
    ):

        category = st.selectbox(
            "분류",
            [
                "펌프",
                "모터",
                "진동",
                "축정렬",
                "오버홀",
                "누수",
                "기타"
            ]
        )

        model = st.text_input(
            "관련 모델"
        )

        phenomenon = st.text_area(
            "현상 및 원인"
        )

        solution = st.text_area(
            "해결 노하우"
        )

        if st.button(
            "노하우 저장",
            type="primary"
        ):

            wb = load_workbook(
                KNOWHOW_DB_PATH
            )

            ws = wb[
                "노하우DB"
            ]

            ws.append(

                [
                    datetime.now().strftime(
                        "%Y-%m-%d"
                    ),

                    category,

                    model,

                    phenomenon,

                    solution,

                    "최진욱"

                ]

            )

            wb.save(
                KNOWHOW_DB_PATH
            )

            wb.close()

            st.success(
                "기술 노하우가 등록되었습니다."
            )


# ============================================================
# 25. 보고서
# ============================================================

elif st.session_state.page == "보고서":

    st.markdown(
        """
        <div class="section-title">
        📄 진단 보고서
        </div>

        <div class="section-caption">
        축적된 진단 및 정비 데이터를 기반으로
        사업소·지자체 제출용 보고서 자료를 구성합니다.
        </div>
        """,
        unsafe_allow_html=True
    )

    site = st.selectbox(
        "대상 사업장",
        [
            "밀양정수장",
            "밀양시 상하수도사업소",
            "창원특례시",
            "김해시"
        ]
    )

    pump_count = st.number_input(
        "진단 설비 수",
        min_value=1,
        value=10
    )

    priority = st.selectbox(
        "최우선 정비 대상",
        [
            "가압펌프 #2",
            "가압펌프 #3",
            "가압펌프 #5"
        ]
    )

    st.markdown(
        f"""
        <div class="platform-card">

        <div class="card-title">
        보고서 미리보기
        </div>

        <b>사업장</b> : {site}<br>
        <b>진단설비</b> : {pump_count}대<br>
        <b>최우선 정비대상</b> : {priority}<br>
        <b>진단방식</b> : QR 기반 설비정보 + 정밀상태진단 + CBM<br>
        <b>정비방식</b> : 상태기반 정비(CBM)<br>
        <b>효과검증</b> : 오버홀 전후 데이터 비교

        </div>
        """,
        unsafe_allow_html=True
    )

    report_text = f"""
K-water tech 설비관리 진단보고서

사업장 : {site}

진단설비 : {pump_count}대

최우선 정비대상 : {priority}

진단체계 :
QR 기반 설비정보 관리
+
17개 핵심 정밀진단
+
CBM 상태평가
+
정비 우선순위 산정
+
오버홀 효과검증

본 시스템은 운전시간만을 기준으로 하는
기존 TBM 방식에서 벗어나 설비의 실제 상태를
기반으로 정비시점을 판단하는 것을 목적으로 한다.
"""

    st.download_button(

        "📥 보고서 자료 다운로드",

        data=report_text,

        file_name=
        f"설비진단보고서_{site}.txt",

        mime="text/plain",

        type="primary"

    )


# ============================================================
# 26. 데이터 관리 / QR 비용
# ============================================================

elif st.session_state.page == "백업":

    st.markdown(
        """
        <div class="section-title">
        💾 데이터 및 사업관리
        </div>

        <div class="section-caption">
        설비 DB 백업과 QR 구축·부착 비용,
        플랫폼 운영비용을 관리합니다.
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        "### 📱 QR 구축비용"
    )

    qr_count = st.number_input(
        "QR 부착 설비 수",
        min_value=1,
        value=10
    )

    unit_cost, total_cost = calculate_qr_cost(
        qr_count
    )

    c1, c2 = st.columns(2)

    c1.metric(
        "QR 1개당 구축비",
        f"{unit_cost:,}원"
    )

    c2.metric(
        f"총 {qr_count}개 구축비",
        f"{total_cost:,}원"
    )

    qr_df = pd.DataFrame(

        [

            [
                "QR 라벨 제작",
                QR_UNIT_COST["QR 라벨 제작"]
            ],

            [
                "표면 세척/준비",
                QR_UNIT_COST["표면 세척/준비"]
            ],

            [
                "QR 부착 작업",
                QR_UNIT_COST["QR 부착 작업"]
            ],

            [
                "설비 등록 및 검수",
                QR_UNIT_COST["설비 등록 및 검수"]
            ]

        ],

        columns=[
            "항목",
            "원가"
        ]

    )

    st.dataframe(
        qr_df,
        use_container_width=True,
        hide_index=True
    )

    st.info(
        "권장 현장 부착시간 : "
        "설비 1대당 약 10~15분"
    )

    st.markdown(
        "### 🏷️ QR 부착 방법"
    )

    st.markdown(
        """
        **① 설치 위치 선정**

        펌프 본체 또는 제어반 등
        작업자가 쉽게 확인할 수 있는 위치 선정

        **② 표면 세척**

        먼지·유분 제거

        **③ QR 라벨 부착**

        평탄한 면에 부착 후
        모서리 및 접착면 압착

        **④ 설비번호 확인**

        QR과 설비번호가 일치하는지 확인

        **⑤ 스마트폰 테스트**

        실제 현장에서 스캔하여
        해당 설비 페이지 연결 확인

        **⑥ DB 등록 완료**

        설비 Master DB와 QR ID 연결
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        "### 💰 플랫폼 사업화 예상비용"
    )

    cost_df = pd.DataFrame(

        [

            [
                "1단계 PoC",
                "자체개발",
                "약 150~300만원",
                "기존 용역 내 실증"
            ],

            [
                "2단계 플랫폼화",
                "고도화",
                "약 500~800만원/년",
                "유지보수·서버·운영"
            ],

            [
                "3단계 외부사업",
                "지자체",
                "별도 견적",
                "설비 수량별 산정"
            ]

        ],

        columns=[
            "단계",
            "구축방식",
            "예상비용",
            "사업모델"
        ]

    )

    st.dataframe(
        cost_df,
        use_container_width=True,
        hide_index=True
    )

    st.write("")

    st.markdown(
        "### 📦 DB 백업"
    )

    for path, label in [

        (
            DB_FILE_PATH,
            "진단 DB"
        ),

        (
            OVERHAUL_DB_PATH,
            "오버홀 DB"
        ),

        (
            KNOWHOW_DB_PATH,
            "노하우 DB"
        )

    ]:

        if os.path.exists(path):

            with open(
                path,
                "rb"
            ) as f:

                st.download_button(

                    f"📥 {label} 다운로드",

                    data=f.read(),

                    file_name=os.path.basename(
                        path
                    ),

                    key=f"download_{path}"

                )


# ============================================================
# 27. QR 비용 별도 상세 계산
# ============================================================

st.write("")


# ============================================================
# 28. 하단 플랫폼 정보
# ============================================================

st.markdown(
    """
    <div style="
    margin-top:35px;
    padding:14px;
    border-top:1px solid #dce8ef;
    text-align:center;
    color:#94a3b8;
    font-size:0.7rem;
    ">

    K-water tech · 설비관리 통합 플랫폼

    <br>

    QR 기반 설비정보 · 상태진단 · CBM · 오버홀 ·
    데이터 기반 유지관리

    </div>
    """,
    unsafe_allow_html=True
)