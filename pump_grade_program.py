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
import urllib.parse
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

import qrcode
from filelock import FileLock


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
EQUIP_DB_PATH = "Pump_Equipment_DB.xlsx"

LOGO_FILE_PATH = "Logo.png"
SEED_FLAG_PATH = "seed_flag.txt"

PHOTO_DIR = "overhaul_photos"
AUTH_PIN = "2580"


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

    padding-top: 3.5rem !important;

    padding-bottom: 2rem !important;

    padding-left: 1.2rem !important;

    padding-right: 1.2rem !important;

}


/* ========================================================
   Streamlit 기본 상단 툴바가 커스텀 헤더 위에
   겹쳐서 그래픽 상단이 잘리는 것을 방지
======================================================== */

header[data-testid="stHeader"] {

    background: transparent;

}

.top-header {

    margin-top: 0.4rem;

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

        font-size: 1.2rem;

    }

    /* ====================================================
       실외(햇빛 아래) 스마트폰 사용을 고려해
       글자를 더 굵고 크게, 명암을 더 진하게
    ==================================================== */

    body, p, span, div, label {

        font-size: 1.02rem !important;

    }

    .section-title {

        font-size: 1.25rem !important;

    }

    .card-title {

        font-size: 1.05rem !important;

    }

    .kpi-label {

        color: #334155 !important;

        font-weight: 800 !important;

    }

    .platform-card,
    .kpi-card {

        border-width: 1.5px !important;

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
# 6-1. 자동판정 항목별 입력 설정
#
# EVAL_ITEMS 중 auto_fn(계산함수)이 지정된 항목은
# 등급을 직접 고르는 대신 측정값을 입력하면
# 자동으로 등급이 산출된다.
# (min, max, 기본값, step, 단위표시)
# ============================================================

AUTO_INPUT_CONFIG = {

    "펌프 효율 유지율 (%)": (
        0.0, 110.0, 95.0, 0.1, "%"
    ),

    "설계 양정/유량 도달률 (%)": (
        0.0, 150.0, 96.0, 0.1, "%"
    ),

    "BEP 운전점 적정성 (%)": (
        0.0, 200.0, 100.0, 0.1, "%"
    ),

    "임펠러/케이싱 링 간극": (
        0.0, 10.0, 1.2, 0.1, "mm"
    ),

    "축슬리브 마모": (
        0.0, 10.0, 0.8, 0.1, "mm"
    ),

    "Overall 진동 (mm/s)": (
        0.0, 20.0, 2.0, 0.1, "mm/s"
    ),

    "펌프-모터 센터링": (
        0.0, 1.0, 0.05, 0.01, "mm"
    ),

    "오버홀 주기": (
        0.0, 30000.0, 8000.0, 100.0, "시간"
    )

}


# ============================================================
# 6-2. 전문용어 사전
#
# 신입/관리부서 등 비전문가도 볼 수 있도록
# 항목명 옆에 간단한 설명을 붙인다. (도움말 아이콘)
# ============================================================

GLOSSARY = {

    "펌프 효율 유지율 (%)":
        "설계(신품) 효율 대비 현재 효율의 비율. "
        "100%에 가까울수록 신품 상태에 가깝다.",

    "설계 양정/유량 도달률 (%)":
        "설계상 목표한 양정(압력)·유량을 실제로 얼마나 "
        "달성하고 있는지의 비율.",

    "BEP 운전점 적정성 (%)":
        "BEP(Best Efficiency Point, 최고효율점) 대비 "
        "현재 운전점의 위치. 100%에 가까울수록 "
        "펌프가 설계상 가장 효율적인 지점에서 운전 중.",

    "임펠러/케이싱 링 간극":
        "회전체(임펠러)와 고정체(케이싱) 사이의 틈. "
        "마모되어 간극이 커지면 내부 누설이 늘어 효율이 떨어진다.",

    "축슬리브 마모":
        "축을 보호하는 슬리브(축 보호관)의 마모량. "
        "패킹·메커니컬씰과 맞닿는 부분이라 마모되면 누수로 이어진다.",

    "NPSH 여유율/캐비테이션":
        "NPSH(Net Positive Suction Head, 유효흡입양정). "
        "여유가 부족하면 공동현상(캐비테이션)이 발생해 "
        "임펠러가 손상될 수 있다.",

    "Overall 진동 (mm/s)":
        "펌프 전체적인 진동의 크기(속도 실효값, RMS). "
        "베어링·축정렬·불균형 등 여러 원인을 종합적으로 반영한다.",

    "펌프-모터 센터링":
        "펌프축과 모터축의 중심을 맞추는 작업(축정렬)의 오차. "
        "오차가 크면 진동·베어링 조기마모의 원인이 된다.",

    "Soft Foot 및 배관 응력":
        "Soft Foot: 설비 받침(Foot) 중 한 곳이 뜨거나 "
        "완전히 밀착되지 않은 상태. "
        "배관 응력: 배관이 설비를 당기거나 미는 힘으로, "
        "둘 다 진동·정렬불량의 숨은 원인이 된다.",

    "오버홀 주기":
        "직전 오버홀(분해정비) 이후 누적 운전시간."

}


def term_help(name):

    return GLOSSARY.get(
        name,
        None
    )


# ============================================================
# 6-3. 설비별 맞춤 진동 기준
#
# 원래는 모든 설비에 똑같은 진동 기준(1.8/4.5/7.1/11.2 mm/s)을
# 적용했다. 실제로는 펌프 마력·모델마다 정상범위가 다를 수 있어,
# 설비 마스터에 "기준진동" 값이 등록되어 있으면 그 값을 기준으로
# 등급을 재계산하고, 없으면 기존 공통 기준을 그대로 쓴다.
# ============================================================

def get_effective_auto_fn(
    name,
    default_fn,
    pump
):

    if (

        name == "Overall 진동 (mm/s)"

        and
        pump.get("기준진동")

    ):

        limit = pump["기준진동"]

        def custom_vib(
            val,
            limit=limit
        ):

            if val < limit * 0.6:
                return "A"

            if val < limit:
                return "B"

            if val < limit * 1.4:
                return "C"

            if val < limit * 1.8:
                return "D"

            return "E"

        return custom_vib

    return default_fn


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

def get_lock(
    path
):

    # 여러 사람이 거의 동시에 저장할 때 파일이 깨지는 것을
    # 막기 위한 파일 잠금. 실제 다중 사용자 동시편집을 완벽히
    # 해결하지는 못하지만(마지막 저장이 우선), 최소한
    # 두 저장 작업이 겹쳐서 엑셀 파일 자체가 손상되는 것은 막아준다.

    return FileLock(
        f"{path}.lock",
        timeout=10
    )


def safe_append_row(
    path,
    sheet,
    row
):

    with get_lock(path):

        wb = load_workbook(
            path
        )

        ws = wb[
            sheet
        ]

        ws.append(
            row
        )

        wb.save(
            path
        )

        wb.close()


def migrate_sheet_headers(
    path,
    sheet,
    required_headers
):

    # 이미 배포되어 있는 엑셀 파일에 컬럼이 부족하면
    # (예전 버전으로 만들어진 파일) 뒤쪽에 새 컬럼을 추가한다.
    # 기존 행 데이터는 그대로 두고 새 컬럼만 빈 값으로 늘어난다.

    with get_lock(path):

        wb = load_workbook(
            path
        )

        ws = wb[
            sheet
        ]

        existing = [

            cell.value

            for cell in ws[1]

        ]

        changed = False

        for header in required_headers:

            if header not in existing:

                ws.cell(

                    row=1,

                    column=len(existing) + 1,

                    value=header

                )

                existing.append(
                    header
                )

                changed = True

        if changed:

            wb.save(
                path
            )

        wb.close()


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
        +
        [
            "효율측정값(%)",
            "진동측정값(mm/s)",
            "온도측정값(°C)"
        ]

    )

    migrate_sheet_headers(

        DB_FILE_PATH,

        "진단이력",

        [
            "효율측정값(%)",
            "진동측정값(mm/s)",
            "온도측정값(°C)"
        ]

    )

    ensure_excel_file(

        EQUIP_DB_PATH,

        "설비마스터",

        [
            "사업장",
            "설비명",
            "제조사",
            "모델명",
            "정격출력(HP)",
            "정격양정(m)",
            "정격유량(m3/h)",
            "회전수(RPM)",
            "준공일",
            "누적운전시간",
            "기준진동(mm/s)",
            "기준효율(%)"
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
# 10-1. 설비 마스터 (추가·삭제 가능한 설비 목록)
#
# 예전에는 설비 10대가 코드(DEFAULT_PUMPS)에 고정되어 있어서
# 현장에서 설비를 추가·삭제할 수 없었다.
# 이제 설비 목록을 엑셀 DB(EQUIP_DB_PATH)로 관리하고,
# 비어 있으면 기존 10대로 최초 1회 시딩한다.
# ============================================================

def seed_equipment_if_empty():

    df_equip = read_excel(
        EQUIP_DB_PATH,
        "설비마스터"
    )

    if not df_equip.empty:

        return

    with get_lock(EQUIP_DB_PATH):

        wb = load_workbook(
            EQUIP_DB_PATH
        )

        ws = wb[
            "설비마스터"
        ]

        for pump in DEFAULT_PUMPS:

            ws.append(

                [
                    pump["site"],
                    pump["equip"],
                    pump["maker"],
                    pump["model"],
                    pump["hp"],
                    pump["head"],
                    pump["flow"],
                    pump["rpm"],
                    pump["build_date"],
                    pump["op_hours"],
                    None,
                    None
                ]

            )

        wb.save(
            EQUIP_DB_PATH
        )

        wb.close()


seed_equipment_if_empty()


def get_all_pumps():

    df_equip = read_excel(
        EQUIP_DB_PATH,
        "설비마스터"
    )

    if df_equip.empty:

        return list(
            DEFAULT_PUMPS
        )

    pumps = []

    for _, row in df_equip.iterrows():

        pumps.append(

            {
                "site": row["사업장"],
                "equip": row["설비명"],
                "maker": row["제조사"],
                "model": row["모델명"],
                "hp": row["정격출력(HP)"],
                "head": row["정격양정(m)"],
                "flow": row["정격유량(m3/h)"],
                "rpm": row["회전수(RPM)"],
                "build_date": str(
                    row["준공일"]
                )[:10],
                "op_hours": row["누적운전시간"],

                "기준진동": (

                    float(row["기준진동(mm/s)"])

                    if pd.notna(
                        row.get("기준진동(mm/s)")
                    )

                    else None

                ),

                "기준효율": (

                    float(row["기준효율(%)"])

                    if pd.notna(
                        row.get("기준효율(%)")
                    )

                    else None

                )

            }

        )

    return pumps


def add_equipment(new_pump):

    safe_append_row(

        EQUIP_DB_PATH,

        "설비마스터",

        [
            new_pump["site"],
            new_pump["equip"],
            new_pump["maker"],
            new_pump["model"],
            new_pump["hp"],
            new_pump["head"],
            new_pump["flow"],
            new_pump["rpm"],
            new_pump["build_date"],
            new_pump["op_hours"],
            new_pump.get("기준진동"),
            new_pump.get("기준효율")
        ]

    )


def delete_equipment(equip_name):

    with get_lock(EQUIP_DB_PATH):

        wb = load_workbook(
            EQUIP_DB_PATH
        )

        ws = wb[
            "설비마스터"
        ]

        rows = list(
            ws.iter_rows(
                min_row=2
            )
        )

        for r in rows:

            if r[1].value == equip_name:

                ws.delete_rows(
                    r[0].row
                )

                break

        wb.save(
            EQUIP_DB_PATH
        )

        wb.close()


ALL_PUMPS = get_all_pumps()


# ============================================================
# 11. 상태 계산
# ============================================================

def pump_status(
    pump
):

    name = pump["equip"]

    hours = pump["op_hours"]

    # ------------------------------------------------------
    # 1순위: 실제로 저장된 정밀진단 이력이 있으면 그 값을 사용한다.
    # (예전 버전은 설비명 문자열에 따라 값을 하드코딩해서
    #  정밀진단에서 아무리 값을 입력해도 다른 화면에 반영이
    #  안 되는 문제가 있었다. 이제는 최근 진단 결과를 읽어온다.)
    # ------------------------------------------------------

    latest = None

    if (

        not df_history.empty

        and
        "설비명" in df_history.columns

    ):

        pump_rows = df_history[

            df_history["설비명"] == name

        ]

        if not pump_rows.empty:

            latest = pump_rows.iloc[
                -1
            ]

    if (

        latest is not None

        and
        "효율측정값(%)" in latest.index

        and
        pd.notna(
            latest["효율측정값(%)"]
        )

    ):

        score = float(
            latest["종합점수"]
        )

        efficiency = float(
            latest["효율측정값(%)"]
        )

        vibration = float(
            latest["진동측정값(mm/s)"]
        ) if pd.notna(
            latest.get("진동측정값(mm/s)")
        ) else 2.2

        temperature = float(
            latest["온도측정값(°C)"]
        ) if pd.notna(
            latest.get("온도측정값(°C)")
        ) else round(
            43 + vibration * 2.3,
            1
        )

    else:

        # ------------------------------------------------------
        # 2순위: 아직 정밀진단 이력이 없는 설비는
        # 데모용 기본값(과거 방식)을 그대로 사용한다.
        # ------------------------------------------------------

        score = 85

        vibration = 2.2

        efficiency = 88

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

        temperature = round(
            43 + vibration * 2.3,
            1
        )

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

        "효율": efficiency,

        "온도": temperature,

        "실측이력있음": latest is not None

    }


# ============================================================
# 11-1. 진동 추세 차트 (AI 진단 공용)
#
# AI 이상징후 페이지와 QR 포털의 AI진단 탭이
# 동일한 그래프를 공유하므로 함수로 분리한다.
# ============================================================

def build_vibration_trend_fig(
    pump,
    result,
    figsize=(9, 4)
):

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
        figsize=figsize
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
        f"{pump['equip']} 진동 추세"
    )

    ax.grid(
        alpha=0.2
    )

    ax.legend()

    return fig


def build_trend_fig(

    pump,

    label,

    unit,

    months,

    values,

    watch_threshold=None,

    danger_threshold=None,

    figsize=(6.2, 3.4)

):

    # 진동 외 다른 지표(효율·온도·CBM Score 등)에도
    # 재사용하는 범용 추세 차트 함수.

    fig, ax = plt.subplots(
        figsize=figsize
    )

    ax.plot(
        months,
        values,
        marker="o",
        linewidth=2
    )

    if watch_threshold is not None:

        ax.axhline(
            watch_threshold,
            linestyle="--",
            label="관찰 기준"
        )

    if danger_threshold is not None:

        ax.axhline(
            danger_threshold,
            linestyle=":",
            label="주의 기준"
        )

    ax.set_ylabel(
        f"{label} ({unit})"
    )

    ax.set_title(
        f"{pump['equip']} {label} 추세"
    )

    ax.grid(
        alpha=0.2
    )

    if watch_threshold is not None or danger_threshold is not None:

        ax.legend()

    return fig


TREND_MONTHS = [

    "2025.01",
    "2025.06",
    "2025.12",
    "2026.03",
    "2026.08",
    "2026.12"

]


def build_efficiency_trend_fig(
    pump,
    result,
    figsize=(6.2, 3.4)
):

    # 효율은 막대그래프로 표현

    base = result["효율"]

    values = [

        min(100, base + 8),

        min(100, base + 6),

        min(100, base + 4),

        min(100, base + 2),

        base,

        max(0, base - 3)

    ]

    colors = [

        "#e03131" if v < 70
        else "#f08c00" if v < 80
        else "#087ea4"

        for v in values

    ]

    fig, ax = plt.subplots(
        figsize=figsize
    )

    ax.bar(

        TREND_MONTHS,

        values,

        color=colors,

        width=0.55

    )

    ax.axhline(
        80,
        linestyle="--",
        color="#a16207",
        label="관찰 기준"
    )

    ax.axhline(
        70,
        linestyle=":",
        color="#c62828",
        label="주의 기준"
    )

    ax.set_ylabel(
        "효율 (%)"
    )

    ax.set_title(
        f"{pump['equip']} 효율 추세"
    )

    ax.set_ylim(
        0,
        105
    )

    ax.grid(
        alpha=0.2,
        axis="y"
    )

    ax.legend()

    return fig


def build_temperature_trend_fig(
    pump,
    result,
    figsize=(6.2, 3.4)
):

    # 온도는 영역(면적)그래프로 표현

    base = result["온도"]

    values = [

        max(35, base - 6),

        max(37, base - 4.5),

        max(39, base - 3),

        max(40, base - 1.5),

        base,

        base + 2

    ]

    fig, ax = plt.subplots(
        figsize=figsize
    )

    ax.fill_between(

        TREND_MONTHS,

        values,

        color="#ff922b",

        alpha=0.35

    )

    ax.plot(

        TREND_MONTHS,

        values,

        marker="o",

        linewidth=2,

        color="#e8590c"

    )

    ax.axhline(
        50,
        linestyle="--",
        color="#a16207",
        label="관찰 기준"
    )

    ax.axhline(
        55,
        linestyle=":",
        color="#c62828",
        label="주의 기준"
    )

    ax.set_ylabel(
        "온도 (°C)"
    )

    ax.set_title(
        f"{pump['equip']} 온도 추세"
    )

    ax.grid(
        alpha=0.2
    )

    ax.legend()

    return fig


def build_score_gauge_fig(
    pump,
    result,
    figsize=(6.2, 2.3)
):

    # CBM Score는 게이지(구간별 색상 막대 + 현재값 표시)로 표현

    score = result["점수"]

    zones = [

        (0, 60, "#ffc9c9"),

        (60, 70, "#ffd8a8"),

        (70, 80, "#fff3bf"),

        (80, 90, "#d3f9d8"),

        (90, 100, "#b2f2bb")

    ]

    fig, ax = plt.subplots(
        figsize=figsize
    )

    for start, end, color in zones:

        ax.barh(

            0,

            end - start,

            left=start,

            color=color,

            height=0.6,

            edgecolor="white"

        )

    ax.barh(

        0,

        2,

        left=max(
            0,
            min(score, 98)
        ),

        color="#083b5c",

        height=0.9

    )

    ax.set_xlim(
        0,
        100
    )

    ax.set_yticks(
        []
    )

    ax.set_xlabel(
        "CBM Score (0~100)"
    )

    ax.set_title(
        f"{pump['equip']} 현재 CBM Score : {score}점 ({result['등급']}등급)"
    )

    return fig


def build_op_hours_trend_fig(
    pump,
    figsize=(9, 3.2)
):

    # 누적 운전시간은 영역(면적)그래프로 표현

    current_hours = pump["op_hours"]

    values = [

        max(0, current_hours - 5000),

        max(0, current_hours - 4000),

        max(0, current_hours - 3000),

        max(0, current_hours - 2000),

        max(0, current_hours - 1000),

        current_hours

    ]

    fig, ax = plt.subplots(
        figsize=figsize
    )

    ax.fill_between(

        TREND_MONTHS,

        values,

        color="#087ea4",

        alpha=0.30

    )

    ax.plot(

        TREND_MONTHS,

        values,

        marker="o",

        linewidth=2,

        color="#063b63"

    )

    ax.set_ylabel(
        "누적 운전시간 (h)"
    )

    ax.set_title(
        f"{pump['equip']} 누적 운전시간 추세"
    )

    ax.grid(
        alpha=0.2
    )

    return fig


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
# 12-1. 로그인 게이트 · 보기 전용 모드
#
# URL만 알면 누구나 데이터를 수정할 수 있던 문제에 대한
# 최소한의 보완. 정식 계정 시스템은 아니고
# 단순 PIN 게이트 수준이라는 점을 감안해서 사용할 것.
# ============================================================

if "authenticated" not in st.session_state:

    st.session_state.authenticated = False

if "read_only" not in st.session_state:

    st.session_state.read_only = False


def is_read_only():

    return st.session_state.read_only


if not st.session_state.authenticated:

    st.markdown(
        """
        <div class="top-header">
        <div class="top-title">
        💧 K-water tech 설비관리 플랫폼
        </div>
        <div class="top-sub">
        접속하려면 PIN 번호를 입력하세요.
        </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    pin_input = st.text_input(

        "PIN 번호",

        type="password",

        key="pin_input"

    )

    col_a, col_b = st.columns(2)

    with col_a:

        if st.button(
            "접속",
            type="primary",
            use_container_width=True
        ):

            if pin_input == AUTH_PIN:

                st.session_state.authenticated = True

                st.rerun()

            else:

                st.error(
                    "PIN 번호가 올바르지 않습니다."
                )

    with col_b:

        if st.button(
            "👀 보기 전용으로 둘러보기",
            use_container_width=True
        ):

            st.session_state.authenticated = True

            st.session_state.read_only = True

            st.rerun()

    st.stop()


# ============================================================
# 13. 사이드바
# ============================================================

if "page" not in st.session_state:

    # QR 스캔으로 들어온 경우 URL의 ?page=QR&equip=... 를
    # 그대로 반영해서 해당 설비 화면으로 바로 이동시킨다.

    st.session_state.page = st.query_params.get(
        "page",
        "홈"
    )


def go_to_page(
    page_key
):

    # 버튼의 on_click 콜백으로 사용.
    # 상태 변경 직후 Streamlit이 자동으로 rerun을 실행하므로
    # 별도의 st.rerun() 호출이 필요 없고,
    # 모바일 인앱 브라우저(카카오톡 등)에서도
    # 클릭 반응이 더 안정적이다.

    st.session_state.page = page_key


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

        st.button(
            label,
            key=f"menu_{key}",
            use_container_width=True,
            on_click=go_to_page,
            args=(key,)
        )

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

        st.button(
            label,
            key=f"menu_{key}",
            use_container_width=True,
            on_click=go_to_page,
            args=(key,)
        )

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

        st.button(
            label,
            key=f"menu_{key}",
            use_container_width=True,
            on_click=go_to_page,
            args=(key,)
        )

    st.markdown("---")

    st.caption(
        "최진욱 · 정밀진단원 / 관리자"
    )

    st.caption(
        "밀양정수장"
    )

    st.toggle(

        "🔒 보기 전용 모드",

        key="read_only",

        help="켜면 모든 저장·삭제 버튼이 비활성화됩니다. "
             "외부인에게 시연할 때 실수로 데이터가 "
             "바뀌는 것을 막아줍니다."

    )

    _risk_count = sum(

        1

        for p in ALL_PUMPS

        if pump_status(p)["상태"] == "정비검토"

    )

    if _risk_count > 0:

        st.error(
            f"🚨 정비검토 필요 설비 {_risk_count}대"
        )


# ============================================================
# 14. 공통 헤더 및 상단 이동 메뉴
#
# - 큰 배너(K-water tech ... 소개문구)는 "홈" 화면에서만 보여준다.
#   다른 화면에서는 각 화면 자체의 section-title/caption이
#   맨 위에 바로 보이도록 한다. (요청 6)
#
# - 사이드바(왼쪽 서랍 메뉴)는 모바일 브라우저에서 버튼을 눌러도
#   서랍이 자동으로 닫히지 않는 Streamlit 자체의 한계가 있어,
#   대신 본문 맨 위에 항상 떠 있는 "메뉴 이동" 드롭다운을 추가한다.
#   드롭다운은 선택하면 자동으로 닫히는 표준 UI라 이 문제가 없다.
#   (요청 3)
# ============================================================

ALL_MENUS = (
    main_menus
    +
    analysis_menus
    +
    knowledge_menus
)

MENU_LABEL_BY_KEY = dict(
    ALL_MENUS
)

MENU_KEY_BY_LABEL = {

    label: key

    for key, label in ALL_MENUS

}

_current_page = st.session_state.page

_current_label = MENU_LABEL_BY_KEY.get(
    _current_page,
    "🏠 설비관리 홈"
)

_nav_labels = [
    label
    for key, label in ALL_MENUS
]


def _on_top_nav_change():

    chosen_label = st.session_state[
        f"top_nav_{_current_page}"
    ]

    st.session_state.page = MENU_KEY_BY_LABEL[
        chosen_label
    ]


st.selectbox(

    "메뉴 이동",

    _nav_labels,

    index=_nav_labels.index(
        _current_label
    ),

    key=f"top_nav_{_current_page}",

    on_change=_on_top_nav_change,

    label_visibility="collapsed"

)

if st.session_state.page == "홈":

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
        ALL_PUMPS
    )

    normal = 0
    watch = 0
    repair = 0

    for pump in ALL_PUMPS:

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

        for pump in ALL_PUMPS:

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

        for pump in ALL_PUMPS:

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
        관리 중인 펌프의 기본 제원과 운전상태를 한눈에 확인합니다.
        </div>
        """,
        unsafe_allow_html=True
    )

    selected = st.selectbox(

        "설비 선택",

        [
            p["equip"]
            for p in ALL_PUMPS
        ]

    )

    pump = next(
        p for p in ALL_PUMPS
        if p["equip"] == selected
    )

    result = pump_status(
        pump
    )

    status_class = {

        "정상": "status-normal",

        "관찰": "status-watch",

        "정비검토": "status-danger"

    }.get(
        result["상태"],
        "status-watch"
    )

    st.markdown(
        f"""
        <div class="platform-card">

        <div class="card-title">
        {pump['equip']} 종합 현황
        </div>

        <span class="{status_class}">
        {result['상태']}
        </span>

        </div>
        """,
        unsafe_allow_html=True
    )

    a, b, c, d = st.columns(4)

    a.metric(
        "CBM Score",
        f'{result["점수"]}점'
    )

    b.metric(
        "현재 효율",
        f'{result["효율"]:.1f}%'
    )

    c.metric(
        "진동",
        f'{result["진동"]:.1f} mm/s'
    )

    d.metric(
        "온도",
        f'{result["온도"]:.1f}°C'
    )

    st.pyplot(

        build_score_gauge_fig(

            pump,

            result

        )

    )

    st.write("")

    tab1, tab2, tab3 = st.tabs(

        [
            "기본정보",
            "운전상태",
            "정비이력"
        ]

    )

    with tab1:

        st.dataframe(

            pd.DataFrame(

                [

                    ["사업장", pump["site"]],

                    ["설비명", pump["equip"]],

                    ["제조사", pump["maker"]],

                    ["모델명", pump["model"]],

                    ["정격출력(HP)", f'{pump["hp"]}'],

                    ["정격양정(m)", f'{pump["head"]}'],

                    ["정격유량(m³/h)", f'{pump["flow"]}'],

                    ["회전수(RPM)", f'{pump["rpm"]}'],

                    ["준공일", pump["build_date"]]

                ],

                columns=[
                    "항목",
                    "값"
                ]

            ),

            use_container_width=True,

            hide_index=True

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
                        "현재 효율",
                        f'{result["효율"]:.1f}%'
                    ],
                    [
                        "현재 진동",
                        f'{result["진동"]:.1f} mm/s'
                    ],
                    [
                        "현재 온도",
                        f'{result["온도"]:.1f}°C'
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

        df_overhaul = read_excel(
            OVERHAUL_DB_PATH,
            "오버홀이력"
        )

        pump_overhaul = df_overhaul[
            df_overhaul["설비명"] == pump["equip"]
        ] if not df_overhaul.empty else pd.DataFrame()

        if not pump_overhaul.empty:

            st.dataframe(

                pump_overhaul[
                    [
                        "작업일자",
                        "공정단계",
                        "작업내용"
                    ]
                ].tail(5),

                use_container_width=True,

                hide_index=True

            )

        else:

            st.info(
                "저장된 오버홀·작업 이력이 없습니다. "
                "QR 포털의 '이력' 탭에서 현장 작업기록을 추가할 수 있습니다."
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
        해당 설비의 제원·도면·이력·AI진단을
        한 화면에서 바로 확인합니다.
        </div>
        """,
        unsafe_allow_html=True
    )

    _pump_names = [
        p["equip"]
        for p in ALL_PUMPS
    ]

    _query_equip = st.query_params.get(
        "equip"
    )

    _default_index = (

        _pump_names.index(_query_equip)

        if _query_equip in _pump_names

        else 0

    )

    selected = st.selectbox(

        "설비 선택",

        _pump_names,

        index=_default_index

    )

    pump = next(
        p for p in ALL_PUMPS
        if p["equip"] == selected
    )

    result = pump_status(
        pump
    )

    equip_no = ALL_PUMPS.index(
        pump
    ) + 1

    qr_id = f"PUMP-MLY-{equip_no:03d}"

    # 실제로 스캔 가능한 QR 이미지 생성.
    # 배포 주소를 모르는 로컬 환경에서는 예시 도메인으로
    # 대체되며, 실제 배포 후에는 REAL APP URL로 바꿔주면 된다.

    APP_BASE_URL = "https://kwatertech-pump.streamlit.app"

    qr_target_url = (

        f"{APP_BASE_URL}/?page=QR&equip="

        +
        urllib.parse.quote(
            selected
        )

    )

    qr_img = qrcode.make(
        qr_target_url
    )

    qr_buffer = io.BytesIO()

    qr_img.save(
        qr_buffer,
        format="PNG"
    )

    c1, c2 = st.columns(
        [1, 2]
    )

    with c1:

        st.image(

            qr_buffer.getvalue(),

            use_container_width=True,

            caption=qr_id

        )

        st.caption(
            "실제 스캔 가능한 QR입니다. "
            "스캔하면 이 설비 페이지로 바로 연결됩니다."
        )

    with c2:

        st.markdown(
            f"""
            <div class="platform-card">

            <div class="card-title">
            {pump["equip"]} 디지털 이력카드
            </div>

            <b>사업장</b> : {pump["site"]}<br>
            <b>제조사</b> : {pump["maker"]}<br>
            <b>모델</b> : {pump["model"]}<br>
            <b>출력</b> : {pump["hp"]} HP<br>
            <b>운전시간</b> : {pump["op_hours"]:,} h<br>
            <b>현재 효율</b> : {result["효율"]:.1f}%<br>
            <b>현재 진동</b> : {result["진동"]:.1f} mm/s<br>
            <b>현재 온도</b> : {result["온도"]:.1f}°C<br>
            <b>CBM Score</b> : {result["점수"]}점<br>
            <b>상태</b> : {result["상태"]}


            </div>
            """,
            unsafe_allow_html=True
        )

    st.write("")

    t1, t2, t3, t4 = st.tabs(

        [
            "제원",
            "도면",
            "이력",
            "AI진단"
        ]

    )

    with t1:

        st.dataframe(

            pd.DataFrame(

                [

                    ["사업장", pump["site"]],

                    ["설비명", pump["equip"]],

                    ["제조사", pump["maker"]],

                    ["모델명", pump["model"]],

                    ["정격출력(HP)", f'{pump["hp"]}'],

                    ["정격양정(m)", f'{pump["head"]}'],

                    ["정격유량(m³/h)", f'{pump["flow"]}'],

                    ["회전수(RPM)", f'{pump["rpm"]}'],

                    ["준공일", pump["build_date"]],

                    [
                        "누적 운전시간",
                        f'{pump["op_hours"]:,} h'
                    ]

                ],

                columns=[
                    "항목",
                    "값"
                ]

            ),

            use_container_width=True,

            hide_index=True

        )

    with t2:

        if "drawings" not in st.session_state:

            st.session_state.drawings = {}

        drawing_file = st.file_uploader(

            "설비 도면 업로드 (배관도·조립도 등)",

            type=["png", "jpg", "jpeg"],

            key=f"drawing_upload_{pump['equip']}"

        )

        if drawing_file is not None:

            st.session_state.drawings[
                pump["equip"]
            ] = drawing_file.getvalue()

        if pump["equip"] in st.session_state.drawings:

            st.image(

                st.session_state.drawings[
                    pump["equip"]
                ],

                use_container_width=True,

                caption=f"{pump['equip']} 도면"

            )

            st.caption(
                "스마트폰에서는 두 손가락으로 확대(핀치 줌)해서 볼 수 있습니다."
            )

        else:

            st.info(
                "등록된 도면이 없습니다. "
                "배관도·조립도 이미지를 업로드하면 "
                "현장에서 바로 확대해 볼 수 있습니다."
            )

    with t3:

        st.markdown(
            "##### 최근 정밀진단 이력"
        )

        pump_history = df_history[
            df_history["설비명"] == pump["equip"]
        ] if not df_history.empty else pd.DataFrame()

        if not pump_history.empty:

            st.dataframe(

                pump_history[
                    [
                        "점검일",
                        "종합점수",
                        "최종등급"
                    ]
                ].tail(5),

                use_container_width=True,

                hide_index=True

            )

            if len(pump_history) > 5:

                with st.expander(
                    f"전체 진단이력 펼쳐보기 (총 {len(pump_history)}건)"
                ):

                    st.dataframe(

                        pump_history[
                            [
                                "점검일",
                                "종합점수",
                                "최종등급"
                            ]
                        ],

                        use_container_width=True,

                        hide_index=True

                    )

        else:

            st.info(
                "저장된 정밀진단 이력이 없습니다."
            )

        st.markdown(
            "##### 오버홀·현장 작업 이력"
        )

        df_overhaul = read_excel(
            OVERHAUL_DB_PATH,
            "오버홀이력"
        )

        pump_overhaul = df_overhaul[
            df_overhaul["설비명"] == pump["equip"]
        ] if not df_overhaul.empty else pd.DataFrame()

        if not pump_overhaul.empty:

            st.dataframe(

                pump_overhaul[
                    [
                        "작업일자",
                        "공정단계",
                        "작업내용"
                    ]
                ].tail(5),

                use_container_width=True,

                hide_index=True

            )

            if len(pump_overhaul) > 5:

                with st.expander(
                    f"전체 작업이력 펼쳐보기 (총 {len(pump_overhaul)}건)"
                ):

                    st.dataframe(

                        pump_overhaul[
                            [
                                "작업일자",
                                "공정단계",
                                "작업내용"
                            ]
                        ],

                        use_container_width=True,

                        hide_index=True

                    )

            photo_rows = pump_overhaul[

                pump_overhaul["사진파일명"].astype(str) != ""

            ] if "사진파일명" in pump_overhaul.columns else pd.DataFrame()

            for _, prow in photo_rows.tail(3).iterrows():

                photo_path = os.path.join(
                    PHOTO_DIR,
                    str(prow["사진파일명"])
                )

                if os.path.exists(photo_path):

                    st.image(

                        photo_path,

                        caption=f'{prow["작업일자"]} 작업사진',

                        use_container_width=True

                    )

        else:

            st.info(
                "저장된 오버홀·작업 이력이 없습니다."
            )

        st.write("")

        work = st.text_area(

            "현장 작업 메모 추가",

            placeholder="점검 및 작업 내용을 입력하세요.",

            key=f"qr_note_{pump['equip']}",

            disabled=is_read_only()

        )

        if is_read_only():

            st.info(
                "🔒 보기 전용 모드에서는 저장할 수 없습니다."
            )

        elif st.button(

            "작업기록 저장",

            type="primary",

            key=f"qr_save_{pump['equip']}"

        ):

            safe_append_row(

                OVERHAUL_DB_PATH,

                "오버홀이력",

                [
                    datetime.now().strftime(
                        "%Y-%m-%d"
                    ),

                    pump["site"],

                    pump["equip"],

                    "현장메모",

                    "최진욱",

                    work,

                    "",

                    "",

                    ""

                ]

            )

            st.success(
                "현장 작업기록이 저장되었습니다."
            )

    with t4:

        status_class = {

            "정상": "status-normal",

            "관찰": "status-watch",

            "정비검토": "status-danger"

        }.get(
            result["상태"],
            "status-watch"
        )

        st.markdown(
            f"""
            <span class="{status_class}">
            {result['상태']}
            </span>
            """,
            unsafe_allow_html=True
        )

        a, b, c = st.columns(3)

        a.metric(
            "CBM Score",
            f'{result["점수"]}점'
        )

        b.metric(
            "진동(RMS)",
            f'{result["진동"]:.1f} mm/s'
        )

        c.metric(
            "온도",
            f'{result["온도"]:.1f}°C'
        )

        fig = build_vibration_trend_fig(

            pump,

            result,

            figsize=(7, 3)

        )

        st.pyplot(
            fig
        )

        if result["진동"] >= 7.1:

            st.error(
                "CBM 예측 : 고위험 상태 · 정비검토 필요"
            )

        elif result["진동"] >= 4.5:

            st.warning(
                "CBM 예측 : 주의 상태 · 추이관찰 및 정밀진단 권고"
            )

        else:

            st.success(
                "CBM 예측 : 현재 상태 양호"
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
            for p in ALL_PUMPS
        ]

    )

    pump = next(
        p for p in ALL_PUMPS
        if p["equip"] == selected
    )

    st.info(
        f'{pump["site"]} / '
        f'{pump["equip"]} / '
        f'{pump["maker"]} {pump["model"]}'
    )

    if pump.get("기준진동") or pump.get("기준효율"):

        st.caption(
            f"⚙️ 이 설비는 맞춤 기준 적용 중 · "
            f"기준진동 {pump.get('기준진동') or '공통기준'} mm/s · "
            f"기준효율 {pump.get('기준효율') or '공통기준'} %"
        )

    total_score = 0

    category_score = {
        "성능": 0,
        "내부상태": 0,
        "기계상태": 0,
        "정비이력": 0
    }

    details = []

    auto_raw_values = {}

    def render_eval_item(
        idx,
        item
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

            st.caption(
                f"판정기준 : {standard}"
            )

            glossary_text = term_help(
                name
            )

            if glossary_text:

                st.caption(
                    f"ℹ️ {glossary_text}"
                )

            if auto_fn is not None:

                min_v, max_v, default_v, step_v, unit = AUTO_INPUT_CONFIG[
                    name
                ]

                raw_value = st.number_input(

                    f"측정값 입력 ({unit})",

                    min_value=float(min_v),

                    max_value=float(max_v),

                    value=float(default_v),

                    step=float(step_v),

                    key=f"val_{idx}",

                    disabled=is_read_only()

                )

                auto_raw_values[
                    name
                ] = raw_value

                effective_fn = get_effective_auto_fn(
                    name,
                    auto_fn,
                    pump
                )

                grade = effective_fn(
                    raw_value
                )

                st.metric(
                    "자동판정",
                    grade
                )

            else:

                grade = st.selectbox(

                    "판정",

                    options,

                    key=f"grade_{idx}",

                    disabled=is_read_only()

                )

        score = score_map[
            grade
        ]

        return category, name, grade, score

    category_order = [
        "성능",
        "내부상태",
        "기계상태",
        "정비이력"
    ]

    category_tabs = st.tabs(

        [
            f"{cat} · {CATEGORIES[cat]}점"
            for cat in category_order
        ]

    )

    for tab, cat in zip(
        category_tabs,
        category_order
    ):

        with tab:

            cat_items = [

                (idx, item)

                for idx, item in enumerate(EVAL_ITEMS)

                if item[0] == cat

            ]

            for i in range(
                0,
                len(cat_items),
                2
            ):

                pair = cat_items[i:i + 2]

                cols = st.columns(
                    len(pair)
                )

                for col, (idx, item) in zip(
                    cols,
                    pair
                ):

                    with col:

                        category, name, grade, score = render_eval_item(
                            idx,
                            item
                        )

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

    temp_measured = st.number_input(

        "측정 온도 (°C) — EVAL_ITEMS에는 없지만 "
        "AI 이상징후 추세에 함께 반영됩니다",

        min_value=0.0,

        max_value=120.0,

        value=45.0,

        step=0.5,

        key="diag_temp",

        disabled=is_read_only()

    )

    if is_read_only():

        st.info(
            "🔒 보기 전용 모드입니다. "
            "저장하려면 사이드바에서 보기 전용 모드를 해제하세요."
        )

    else:

        if st.button(
            "💾 진단결과 저장",
            type="primary",
            use_container_width=True
        ):

            save_time = datetime.now().strftime(
                "%Y-%m-%d"
            )

            row = [

                save_time,

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

            ] + [

                auto_raw_values.get(
                    "펌프 효율 유지율 (%)",
                    ""
                ),

                auto_raw_values.get(
                    "Overall 진동 (mm/s)",
                    ""
                ),

                temp_measured

            ]

            safe_append_row(

                DB_FILE_PATH,

                "진단이력",

                row

            )

            st.success(
                f"{pump['equip']} "
                f"진단결과가 저장되었습니다."
            )

            # 저장 직후 방금 저장한 값이 실제로
            # 반영됐는지 바로 눈으로 확인할 수 있도록
            # 파일을 다시 읽어서 보여준다.

            df_just_saved = read_excel(
                DB_FILE_PATH,
                "진단이력"
            )

            st.markdown(
                "###### ✅ 방금 저장된 기록"
            )

            st.dataframe(

                df_just_saved[
                    df_just_saved["설비명"] == pump["equip"]
                ].tail(1),

                use_container_width=True,

                hide_index=True

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

    for pump in ALL_PUMPS:

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
            for p in ALL_PUMPS
        ]

    )

    pump = next(
        p for p in ALL_PUMPS
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
            "부품교체, 센터링 등",
            disabled=is_read_only()
        )

        work_photo = st.file_uploader(

            "작업 전후 사진 첨부 (선택)",

            type=["png", "jpg", "jpeg"],

            key="overhaul_photo_upload",

            disabled=is_read_only()

        )

        if is_read_only():

            st.info(
                "🔒 보기 전용 모드에서는 저장할 수 없습니다."
            )

        elif st.button(
            "오버홀 작업기록 저장",
            type="primary"
        ):

            photo_filename = ""

            if work_photo is not None:

                os.makedirs(
                    PHOTO_DIR,
                    exist_ok=True
                )

                photo_filename = (

                    f"{pump['equip']}_"

                    +
                    datetime.now().strftime(
                        "%Y%m%d_%H%M%S"
                    )

                    +
                    "_"
                    +
                    work_photo.name

                )

                with open(

                    os.path.join(
                        PHOTO_DIR,
                        photo_filename
                    ),

                    "wb"

                ) as f:

                    f.write(
                        work_photo.getbuffer()
                    )

            safe_append_row(

                OVERHAUL_DB_PATH,

                "오버홀이력",

                [
                    datetime.now().strftime(
                        "%Y-%m-%d"
                    ),

                    pump["site"],

                    pump["equip"],

                    "오버홀",

                    "최진욱",

                    work,

                    photo_filename,

                    "",

                    ""

                ]

            )

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
            for p in ALL_PUMPS
        ]

    )

    pump = next(
        p for p in ALL_PUMPS
        if p["equip"] == selected
    )

    result = pump_status(
        pump
    )

    g1, g2 = st.columns(2)

    with g1:

        st.pyplot(

            build_vibration_trend_fig(

                pump,

                result,

                figsize=(6.2, 3.4)

            )

        )

    with g2:

        st.pyplot(

            build_efficiency_trend_fig(

                pump,

                result

            )

        )

    g3, g4 = st.columns(2)

    with g3:

        st.pyplot(

            build_temperature_trend_fig(

                pump,

                result

            )

        )

    with g4:

        st.pyplot(

            build_score_gauge_fig(

                pump,

                result

            )

        )

    st.pyplot(

        build_op_hours_trend_fig(

            pump

        )

    )

    st.write("")

    if (

        result["진동"] >= 7.1

        or
        result["효율"] <= 70

        or
        result["온도"] >= 55

        or
        result["점수"] < 60

    ):

        st.error(
            "고위험 상태 · 정비검토 필요"
        )

    elif (

        result["진동"] >= 4.5

        or
        result["효율"] <= 80

        or
        result["온도"] >= 50

        or
        result["점수"] < 80

    ):

        st.warning(
            "주의 상태 · 추이관찰 및 정밀진단 권고"
        )

    else:

        st.success(
            "현재 상태 양호 · 모든 지표 정상범위"
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

    # 결과 카드를 입력값보다 먼저 화면에 보여주기 위해
    # 자리(placeholder)를 미리 만들어 둔다.
    # 실제 값은 아래에서 입력을 받은 뒤 이 자리에 채워 넣는다.

    result_slot = st.container()

    with st.expander(
        "🔧 계산 조건 입력",
        expanded=True
    ):

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

        c1, c2 = st.columns(2)

        hours = c1.number_input(
            "연간 운전시간",
            value=6000
        )

        price = c2.number_input(
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

    with result_slot:

        st.markdown(
            f"""
            <div class="kpi-grid" style="
            grid-template-columns:
            repeat(3, 1fr);
            ">

                <div class="kpi-card">
                    <div class="kpi-label">연간 절감전력</div>
                    <div class="kpi-value">{saved_kwh:,.0f}</div>
                    <div class="kpi-sub">kWh / 년</div>
                </div>

                <div class="kpi-card">
                    <div class="kpi-label">연간 절감액</div>
                    <div class="kpi-value">{saved_money:,.0f}</div>
                    <div class="kpi-sub">원 / 년</div>
                </div>

                <div class="kpi-card">
                    <div class="kpi-label">투자회수 기간</div>
                    <div class="kpi-value">{payback:.1f}년</div>
                    <div class="kpi-sub">오버홀 비용 대비</div>
                </div>

            </div>
            """,
            unsafe_allow_html=True
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
            "설비정보 조회시간 단축률",
            50,
            "%"
        ],

        [
            "QR 스캔 성공률",
            95,
            "%"
        ],

        [
            "설비정보 정확성",
            98,
            "%"
        ],

        [
            "현장 작업자 수용성",
            70,
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

    df_knowhow = read_excel(
        KNOWHOW_DB_PATH,
        "노하우DB"
    )

    if not df_knowhow.empty:

        fc1, fc2 = st.columns(
            [2, 1]
        )

        search_word = fc1.text_input(

            "🔍 키워드 검색 (현상·원인·해결노하우·모델명)",

            key="knowhow_search"

        )

        category_options = ["전체"] + sorted(

            df_knowhow["분류"].dropna().unique().tolist()

        )

        category_filter = fc2.selectbox(

            "분류 필터",

            category_options,

            key="knowhow_category_filter"

        )

        filtered = df_knowhow.copy()

        if category_filter != "전체":

            filtered = filtered[

                filtered["분류"] == category_filter

            ]

        if search_word:

            mask = (

                filtered["관련모델"].astype(str).str.contains(search_word, case=False, na=False)

                |
                filtered["현상및원인"].astype(str).str.contains(search_word, case=False, na=False)

                |
                filtered["해결노하우"].astype(str).str.contains(search_word, case=False, na=False)

            )

            filtered = filtered[mask]

        st.caption(
            f"총 {len(df_knowhow)}건 중 {len(filtered)}건 표시"
        )

        st.dataframe(
            filtered,
            use_container_width=True,
            hide_index=True
        )

        if not is_read_only() and not filtered.empty:

            with st.expander(
                "🗑️ 노하우 삭제"
            ):

                delete_options = [

                    f'{i} · {row["등록일자"]} · {row["관련모델"]} · {str(row["현상및원인"])[:20]}'

                    for i, row in filtered.iterrows()

                ]

                delete_choice = st.selectbox(

                    "삭제할 항목 선택",

                    delete_options,

                    key="knowhow_delete_select"

                )

                if st.button(
                    "선택 항목 삭제",
                    key="knowhow_delete_btn"
                ):

                    del_idx = int(
                        delete_choice.split(" · ")[0]
                    )

                    with get_lock(KNOWHOW_DB_PATH):

                        wb = load_workbook(
                            KNOWHOW_DB_PATH
                        )

                        ws = wb[
                            "노하우DB"
                        ]

                        # 엑셀 행 번호 = 데이터프레임 인덱스 + 2
                        # (1행은 헤더이고, iterrows 인덱스는 0부터 시작)

                        ws.delete_rows(
                            del_idx + 2
                        )

                        wb.save(
                            KNOWHOW_DB_PATH
                        )

                        wb.close()

                    st.success(
                        "삭제되었습니다. 페이지를 새로고침하면 반영됩니다."
                    )

    else:

        st.info(
            "등록된 노하우가 아직 없습니다."
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
            ],
            disabled=is_read_only()
        )

        model = st.text_input(
            "관련 모델",
            disabled=is_read_only()
        )

        phenomenon = st.text_area(
            "현상 및 원인",
            disabled=is_read_only()
        )

        solution = st.text_area(
            "해결 노하우",
            disabled=is_read_only()
        )

        if is_read_only():

            st.info(
                "🔒 보기 전용 모드에서는 등록할 수 없습니다."
            )

        elif st.button(
            "노하우 저장",
            type="primary"
        ):

            safe_append_row(

                KNOWHOW_DB_PATH,

                "노하우DB",

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

    st.info(
        "현재 PFM 연동 단계 : 1단계 (수동 입력형) · "
        "독자 DB 구축·운영 중 — "
        "2단계(엑셀 Batch 연계) 및 "
        "3단계(API 연계)는 향후 추진 예정"
    )

    st.markdown(
        "### ⚙️ 설비 마스터 관리"
    )

    st.caption(
        "설비 목록이 코드에 고정되어 있지 않고 여기서 "
        "추가·삭제할 수 있습니다. 설비별로 진동·효율 기준을 "
        "다르게 두고 싶으면 기준값도 함께 입력하세요 "
        "(비워두면 공통 기준 적용)."
    )

    with st.expander(
        "➕ 새 설비 추가"
    ):

        nc1, nc2 = st.columns(2)

        new_site = nc1.text_input(
            "사업장",
            "밀양정수장",
            key="new_equip_site"
        )

        new_equip = nc2.text_input(
            "설비명",
            key="new_equip_name"
        )

        new_maker = nc1.text_input(
            "제조사",
            key="new_equip_maker"
        )

        new_model = nc2.text_input(
            "모델명",
            key="new_equip_model"
        )

        new_hp = nc1.number_input(
            "정격출력(HP)",
            value=150,
            key="new_equip_hp"
        )

        new_head = nc2.number_input(
            "정격양정(m)",
            value=45,
            key="new_equip_head"
        )

        new_flow = nc1.number_input(
            "정격유량(m³/h)",
            value=1200,
            key="new_equip_flow"
        )

        new_rpm = nc2.number_input(
            "회전수(RPM)",
            value=1780,
            key="new_equip_rpm"
        )

        new_build = nc1.text_input(
            "준공일 (YYYY-MM-DD)",
            "2024-01-01",
            key="new_equip_build"
        )

        new_hours = nc2.number_input(
            "누적 운전시간",
            value=0,
            key="new_equip_hours"
        )

        nc1, nc2 = st.columns(2)

        new_vib_limit = nc1.number_input(
            "기준진동(mm/s) — 선택사항, 0이면 공통기준",
            value=0.0,
            key="new_equip_vib"
        )

        new_eff_target = nc2.number_input(
            "기준효율(%) — 선택사항, 0이면 공통기준",
            value=0.0,
            key="new_equip_eff"
        )

        if is_read_only():

            st.info(
                "🔒 보기 전용 모드에서는 설비를 추가할 수 없습니다."
            )

        elif st.button(
            "설비 추가",
            type="primary",
            key="add_equip_btn"
        ):

            if not new_equip:

                st.error(
                    "설비명을 입력해주세요."
                )

            else:

                add_equipment(

                    {
                        "site": new_site,
                        "equip": new_equip,
                        "maker": new_maker,
                        "model": new_model,
                        "hp": new_hp,
                        "head": new_head,
                        "flow": new_flow,
                        "rpm": new_rpm,
                        "build_date": new_build,
                        "op_hours": new_hours,
                        "기준진동": new_vib_limit if new_vib_limit > 0 else None,
                        "기준효율": new_eff_target if new_eff_target > 0 else None
                    }

                )

                st.success(
                    f"{new_equip} 설비가 추가되었습니다. "
                    "메뉴를 다시 열면 목록에 반영됩니다."
                )

    with st.expander(
        "🗑️ 설비 삭제"
    ):

        del_target = st.selectbox(

            "삭제할 설비",

            [
                p["equip"]
                for p in ALL_PUMPS
            ],

            key="del_equip_select"

        )

        if is_read_only():

            st.info(
                "🔒 보기 전용 모드에서는 삭제할 수 없습니다."
            )

        elif st.button(
            "선택한 설비 삭제",
            key="del_equip_btn"
        ):

            delete_equipment(
                del_target
            )

            st.success(
                f"{del_target} 설비가 삭제되었습니다. "
                "정밀진단·오버홀 이력은 그대로 남아있습니다."
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

    st.warning(
        "⚠️ 이 데이터는 서버 로컬 파일로 저장됩니다. "
        "배포 환경이 재시작·재배포되면 초기화될 수 있으니 "
        "정기적으로 아래에서 전체 백업을 받아두는 것을 권장합니다."
    )

    _backup_files = [

        (DB_FILE_PATH, "진단 DB"),
        (OVERHAUL_DB_PATH, "오버홀 DB"),
        (KNOWHOW_DB_PATH, "노하우 DB"),
        (EQUIP_DB_PATH, "설비마스터 DB"),
        (KPI_DB_PATH, "KPI DB")

    ]

    import zipfile

    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(
        zip_buffer,
        "w",
        zipfile.ZIP_DEFLATED
    ) as zf:

        for path, label in _backup_files:

            if os.path.exists(path):

                zf.write(
                    path,
                    arcname=os.path.basename(path)
                )

    st.download_button(

        "📥 전체 DB 한번에 백업 (zip)",

        data=zip_buffer.getvalue(),

        file_name=

        f"kwatertech_backup_"

        +
        datetime.now().strftime("%Y%m%d_%H%M")

        +
        ".zip",

        mime="application/zip",

        type="primary",

        use_container_width=True

    )

    for path, label in _backup_files:

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
