# -*- coding: utf-8 -*-
"""
충청북도교육청 지방공무원 인사발령 정리 프로그램

[필요한 라이브러리 설치]
    pip install pandas openpyxl customtkinter tkinterdnd2 pywin32

[exe 만들 때 권장 옵션]
    pyinstaller --onefile --noconsole --collect-all tkinterdnd2 --collect-all customtkinter 인사발령정리.py

[실행 조건]
    - exe(또는 py)와 같은 폴더에 '기관현황.xlsx' 파일이 있어야 합니다.
    - 아래아한글(HWP)이 설치된 PC에서 실행해야 합니다.
"""
import pandas as pd
import re
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.pagebreak import Break
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES
import win32com.client as win32
import win32com.client
import os
import sys
import time
import traceback
import threading
import ctypes
import json
import webbrowser
import urllib.request
import urllib.parse
from datetime import datetime

# ==================================================
# ★ 업데이트 확인 설정 (버전을 올릴 때는 아래 current_version만 고치세요)
# ==================================================
current_version = "2.1.0"
version_info = "https://raw.githubusercontent.com/tworemind-cpu/cb_personnel_appointment/main/version.json"

# --- ★ 실행 파일(.exe) 기준 현재 폴더 경로 찾기 ---
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

AGENCY_FILE = os.path.join(BASE_DIR, '기관현황.xlsx')
SETTINGS_FILE = os.path.join(BASE_DIR, 'config.json')


class AppError(Exception):
    """사용자에게 팝업으로 알려줄 오류 (제목, 내용)"""
    def __init__(self, title, message):
        super().__init__(message)
        self.title = title


# --- ★ 업데이트 확인 함수 ---
def parse_version(ver_str):
    """'1.2.3' 같은 문자열을 (1, 2, 3) 숫자 묶음으로 바꿉니다. (크기 비교용)"""
    nums = re.findall(r'\d+', str(ver_str))
    return tuple(int(n) for n in nums) if nums else (0,)

def check_for_update(parent=None):
    """
    깃허브의 version.json을 읽어 새 버전이 있으면 안내 팝업을 띄웁니다.
    - 인터넷이 안 되거나 주소가 틀려도 프로그램은 멈추지 않고 그냥 넘어갑니다.
    - '예'를 누르면 다운로드 페이지를 열고 True(프로그램 종료 필요)를 돌려줍니다.
    """
    try:
        safe_url = urllib.parse.quote(version_info, safe=":/?=&%")
        req = urllib.request.Request(safe_url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=5) as res:
            data = json.loads(res.read().decode("utf-8"))

        latest_version = str(data.get("version", "")).strip()
        download_url = str(data.get("download_url", "")).strip()
        notes = str(data.get("notes", "")).strip()

        if not latest_version:
            return False

        if parse_version(latest_version) > parse_version(current_version):
            msg = (f"새 버전이 나왔습니다.\n\n"
                   f"현재 버전: {current_version}\n"
                   f"최신 버전: {latest_version}\n")
            if notes:
                msg += f"\n[변경 내용]\n{notes}\n"
            msg += "\n지금 다운로드 페이지를 여시겠습니까?\n(예를 누르면 프로그램이 종료됩니다.)"

            if messagebox.askyesno("업데이트 안내", msg, parent=parent):
                if download_url:
                    webbrowser.open(download_url)
                return True
    except Exception:
        # 인터넷 연결 없음, 주소 오류, 파일 형식 오류 등 → 조용히 넘어감
        pass
    return False


# --- ★ 직급 비교용 함수 ---
def get_rank_level(rank_str):
    if not rank_str: return -1
    r = str(rank_str)
    if '부이사관' in r: return 7
    if '서기관' in r: return 6
    if '사무관' in r or '연구관' in r: return 5
    if '주사보' in r: return 3
    if '주사' in r or '연구사' in r: return 4
    if '서기보시보' in r: return 0
    if '서기보' in r: return 1
    if '서기' in r: return 2
    return -1

# --- ★ 직급 및 부서 분리 함수 ---
def parse_rank_dept(raw_list):
    if not raw_list: return "", ""
    rank_keywords = ['부이사관', '서기관', '사무관', '주사', '서기', '연구관', '연구사', '시보']
    rank_len = 1
    if len(raw_list) >= 2:
        if any(k in raw_list[1] for k in rank_keywords):
            rank_len = 2
            
    rank_str = " ".join(raw_list[:rank_len])
    dept_str = " ".join(raw_list[rank_len:])
    return rank_str, dept_str

# --- ★ 직급 명칭 간소화 함수 ---
def simplify_rank_str(s, is_fixed_term=False):
    if not s: return ""
    
    if '부이사관' in s or '3급' in s: return "3급"
    if '서기관' in s or '4급' in s: return "4급"

    s = str(s).replace("지방", "")
    s = s.replace("교육행정", "행정")
    
    s = s.replace("사무관", "5급")
    s = s.replace("주사보", "7급")
    s = s.replace("주사", "6급")
    
    s = s.replace("서기보시보", "9급 시보") 
    s = s.replace("서기보", "9급")
    s = s.replace("서기", "8급")
    
    s = s.strip()
    
    if is_fixed_term:
        match = re.search(r'(\d급)', s)
        if match:
            idx = match.start()
            if idx > 0 and s[idx-1] != ' ':
                s = s[:idx] + ' ' + s[idx:]
    return s

# --- ★ 전의 직렬 판별 함수 ---
def get_series_and_group(rank_str):
    clean_str = str(rank_str).replace(" ", "")
    m_series = re.match(r'^([가-힣]+)', clean_str)
    series = m_series.group(1) if m_series else "기타"
    series = re.sub(r'\d급.*', '', series)
    
    if any(s in series for s in ['조리', '시설관리', '운전', '사무운영']):
        return 2, series
    elif any(s in series for s in ['기록연구', '기록']): 
        return 3, series
    elif any(s in series for s in ['행정', '전산', '사서', '공업', '보건', '식품위생', '시설']):
        return 1, series
    else:
        return 1, series

# ==================================================
# ★ 기관현황.xlsx 읽기 (지역 목록 + 선택한 지역의 기관/부서 순서)
# ==================================================
# 아래 3개는 '선택한 지역'에 따라 apply_region()이 채워 넣습니다. (정렬/판정 함수들이 함께 사용)
dept_order_map = {}
sub_dept_order_map = {}
local_agencies_for_sort = []

ignore_keywords = [
    '유치원', '초등학교', '중학교', '고등학교', '특수학교', '대안학교',
    '교육지원청', '직속기관', '소속기관', '관내학교', '기관명', '부서명',
    '행정과', '교육과', '총무과', '재정과', '행정국', '교육국', '기획국',
    '체육건강과', '학교지원센터', '학교지원과', '학생지원과',
    '유초등교육과', '중등교육과', '창의특수교육과', '인성시민과', '교원인사과',
    '체육건강안전과', '노사정책과', '예산과', '정책기획과', '미래교육추진단',
    '공보관', '감사관', '단설유치원', '병설유치원',
    '행정실', '행정실장', '교무실', '사무관', '서기관', '부이사관',
    '주사', '주사보', '서기', '서기보', '서기보시보', '주무관',
    '지방교육행정', '지방사서', '지방전산', '지방시설',
    '지방공업', '지방보건', '지방식품위생', '지방간호', '본청', '도교육청',
    '교육도서관', '학생수련원', '단재교육연수원', '해양교육원', '유아교육진흥원', '특수교육원'
]

# 엑셀에 적지 않아도 항상 포함되는 특수 기관명 (신규임용 표시용)
ALWAYS_INCLUDED_AGENCIES = ['신규임용']


def read_agency_table():
    """기관현황.xlsx를 읽어서 (표 데이터, 지역 목록)을 돌려줍니다.
       A열(2행부터) = 지역명 / B열(2행부터) = 부서명 / C열(2행부터) = 기관(학교)명  (모두 직제순)"""
    if not os.path.exists(AGENCY_FILE):
        raise AppError("파일 없음", "기관현황.xlsx 파일이 존재하지 않습니다.\n프로그램(exe)과 같은 폴더에 넣어주세요.")
    try:
        with pd.ExcelFile(AGENCY_FILE) as xl:
            sheet = '기관현황' if '기관현황' in xl.sheet_names else xl.sheet_names[0]
            df_agency = xl.parse(sheet, header=None)
    except PermissionError:
        raise AppError("파일 열림 오류", "'기관현황.xlsx' 파일을 읽을 수 없습니다.\n엑셀에서 열려 있다면 저장 후 닫고 다시 시도해주세요.")
    except Exception as e:
        raise AppError("엑셀 읽기 오류", f"'기관현황.xlsx' 분석 중 오류가 발생했습니다.\n{e}")

    if df_agency.shape[1] < 3:
        raise AppError("엑셀 형식 오류", "기관현황.xlsx의 A열(지역명), B열(부서명), C열(기관·학교명) 형식이 아닙니다.")

    def clean(v):
        if pd.isna(v):
            return ''
        t = str(v).strip()
        return '' if t.lower() in ('nan', 'none', '<na>') else t

    rows = []
    cur_region = ''
    for a, b, c in zip(df_agency.iloc[1:, 0].tolist(), df_agency.iloc[1:, 1].tolist(), df_agency.iloc[1:, 2].tolist()):
        a, b, c = clean(a), clean(b), clean(c)
        if a:
            cur_region = a          # 지역명이 비어 있는 줄은 바로 위 지역을 이어받음(병합 셀 대비)
        if cur_region and (b or c):
            rows.append((cur_region, b, c))
    df_data = pd.DataFrame(rows, columns=['지역', '부서', '기관'])

    regions = []
    for r in df_data['지역'].tolist():
        if r not in regions:
            regions.append(r)
    if not regions:
        raise AppError("엑셀 형식 오류", "기관현황.xlsx의 A열(2행부터)에서 지역명을 찾지 못했습니다.")
    return df_data, regions


def apply_region(df_data, region):
    """선택한 지역의 기관·부서 순서(직제순)를 판정/정렬용 자료에 반영합니다."""
    df_local = df_data[df_data['지역'] == region]
    raw_agencies = [x for x in df_local['기관'].tolist() if x]
    raw_sub_depts = [x for x in df_local['부서'].tolist() if x]

    if not raw_agencies:
        raise AppError("기관 없음", f"기관현황.xlsx에 '{region}' 지역의 기관(C열)이 없습니다.")

    raw_agencies.extend(ALWAYS_INCLUDED_AGENCIES)

    local_agencies_for_sort[:] = [a for a in raw_agencies if len(a) > 1 and a not in ignore_keywords]
    dept_order_map.clear()
    dept_order_map.update({agency: i for i, agency in enumerate(local_agencies_for_sort)})
    sub_dept_order_map.clear()
    sub_dept_order_map.update({dept: i for i, dept in enumerate(raw_sub_depts) if len(dept) > 1})


def get_dept_order(dept_str):
    clean_dept = re.sub(r'[^가-힣a-zA-Z0-9]', '', str(dept_str))
    best_idx = 999
    for agency, idx in dept_order_map.items():
        clean_agency = re.sub(r'[^가-힣a-zA-Z0-9]', '', agency)
        if clean_agency in clean_dept and idx < best_idx:
            best_idx = idx
            
    sub_idx = 999
    if "교육지원청" in clean_dept:
        for sub_dept, idx in sub_dept_order_map.items():
            clean_sub_dept = re.sub(r'[^가-힣a-zA-Z0-9]', '', sub_dept)
            if clean_sub_dept in clean_dept and idx < sub_idx:
                sub_idx = idx
    return (best_idx, sub_idx)

series_map = {
    '행정': 1, '전산': 2, '사서': 3, '공업': 4, '보건': 5, '식품위생': 6, '시설': 7,
    '조리': 8, '시설관리': 9, '운전': 10, '사무운영': 11, '기록연구': 12, '기록': 12
}

def get_gubun_priority(gubun):
    if gubun in ['승진', '전보', '전입', '전출', '휴직', '복직', '파견', '파견복귀', '신규임용']: return 1
    elif gubun in ['전문관지정', '전문관 지정']: return 3
    elif gubun in ['명예퇴직', '정년퇴직', '의원면직', '퇴직']: return 4
    return 99

def parse_rank_for_sort(rank_str):
    if '연구관' in str(rank_str) or '연구사' in str(rank_str):
        r_order = 10
    else:
        m_rank = re.search(r'(\d)급', str(rank_str))
        r_order = int(m_rank.group(1)) if m_rank else 99
    
    clean_str = str(rank_str).replace(" ", "")
    m_series = re.match(r'^([가-힣]+)', clean_str)
    series = m_series.group(1) if m_series else ""
    series = re.sub(r'\d급.*', '', series)
    s_order = series_map.get(series, 99)
    
    return r_order, s_order

# --- ★ 직급 -> 직렬 -> 발령구분 -> 부서 -> 성명 정렬 ---
def sort_key_in(row):
    g_order = get_gubun_priority(row['구분'])
    r_order, s_order = parse_rank_for_sort(row['직급'])
    d_order = get_dept_order(row['임용 부서'])
    return (r_order, s_order, g_order, d_order, str(row['성명']))

def sort_key_out(row):
    gubun = row['구분_']
    if gubun == '복직': g_order = 96
    elif gubun == '파견복귀': g_order = 97 
    elif gubun in ['명예퇴직', '정년퇴직', '의원면직', '퇴직']: g_order = 98
    else: g_order = get_gubun_priority(gubun)
        
    r_order, s_order = parse_rank_for_sort(row['직급_'])
    d_order = get_dept_order(row['현임 부서_'])
    return (r_order, s_order, g_order, d_order, str(row['성명_']))


# ==================================================
# ★ 한글(hwp) 파일에서 텍스트 추출
# ==================================================
def auto_allow_popup():
    """한글 보안 팝업창 자동 허용 (Alt+N)"""
    VK_MENU = 0x12
    VK_N = 0x4E
    KEYEVENTF_KEYUP = 0x0002

    try:
        time.sleep(2.5)
        ctypes.windll.user32.keybd_event(VK_MENU, 0, 0, 0)
        time.sleep(0.1)
        ctypes.windll.user32.keybd_event(VK_N, 0, 0, 0)
        time.sleep(0.1)
        ctypes.windll.user32.keybd_event(VK_N, 0, KEYEVENTF_KEYUP, 0)
        time.sleep(0.1)
        ctypes.windll.user32.keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0)
    except Exception as e:
        print(f"팝업 제어 예외 발생: {e}")
    finally:
        ctypes.windll.user32.keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0)


def extract_hwp_text(hwp_file_path):
    temp_txt_path = os.path.join(BASE_DIR, "temp_발령서_임시추출.txt")
    hwp = None
    try:
        try:
            hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        except AttributeError:
            import shutil
            from importlib import reload
            cache_path = win32.gencache.GetGeneratePath()
            shutil.rmtree(cache_path, ignore_errors=True)
            reload(win32com.client.gencache)
            hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")

        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
        hwp.XHwpWindows.Item(0).Visible = True

        threading.Thread(target=auto_allow_popup, daemon=True).start()

        hwp.Open(hwp_file_path)
        hwp.SaveAs(temp_txt_path, "TEXT")
    except Exception as e:
        raise AppError("한글 파일 오류", f"한글 파일 처리에 실패했습니다.\n{e}")
    finally:
        try:
            hwp.Quit()
        except Exception:
            pass
        ctypes.windll.user32.keybd_event(0x12, 0, 0x0002, 0)

    try:
        try:
            with open(temp_txt_path, 'r', encoding='utf-16') as f:
                text_list = [line.strip() for line in f.readlines() if line.strip()]
        except Exception:
            with open(temp_txt_path, 'r', encoding='cp949') as f:
                text_list = [line.strip() for line in f.readlines() if line.strip()]
        if os.path.exists(temp_txt_path):
            os.remove(temp_txt_path)
    except Exception as e:
        raise AppError("텍스트 변환 오류", f"텍스트 변환 중 오류가 발생했습니다.\n{e}")
    return text_list


# ==================================================
# ★ 발령 내용 분석 → '전의' 시트가 담긴 엑셀(워크북) 만들기
#    (아래 본문은 기존 코드의 분류·정렬 로직을 그대로 가져왔습니다)
# ==================================================
def build_report(text_list):
    chunks = []
    current_chunk = []
    recording = False
    date_pattern = re.compile(r'^\d{4}\.\s*\d{1,2}\.\s*\d{1,2}\.?$')

    for line in text_list:
        if not recording:
            if line.isdigit() and 1 <= len(line) <= 4:
                recording = True
                current_chunk = [line]
        else:
            if line.isdigit() and 1 <= len(line) <= 4 and len(current_chunk) >= 5:
                if current_chunk[0].isdigit() and int(line) > int(current_chunk[0]):
                    chunks.append(current_chunk)
                    current_chunk = [line]
                    continue
            current_chunk.append(line)
            if date_pattern.match(line): 
                chunks.append(current_chunk)
                recording = False
                current_chunk = []

    if current_chunk:
        chunks.append(current_chunk)

    _now = datetime.now()
    first_date_str = f"{_now.year}. {_now.month}. {_now.day}."
    for chunk in chunks:
        if chunk and date_pattern.match(chunk[-1]):
            raw_date = chunk[-1]
            parts = [p.strip() for p in raw_date.split('.') if p.strip()]
            if len(parts) >= 3:
                first_date_str = f"{parts[0]}. {int(parts[1])}. {int(parts[2])}."
            else:
                first_date_str = raw_date
            break

    def check_is_local(text, agencies):
        words = text.split()
        for agency in agencies:
            clean_agency = re.sub(r'[^가-힣a-zA-Z0-9]', '', str(agency))
            if len(clean_agency) < 2: continue
            for word in words:
                clean_word = re.sub(r'[^가-힣a-zA-Z0-9]', '', word)
                if clean_agency == clean_word:
                    return True
                if clean_word.startswith(clean_agency):
                    remainder = clean_word[len(clean_agency):]
                    if remainder in ["", "학교", "등학교", "파견", "겸임", "지원", "근무"]:
                        return True
        return False

    categories = {
        "5급 이상": {"in": [], "out": []},
        "6급": {"in": [], "out": []},
        "7급 이하": {"in": [], "out": []},
        "임기제": {"in": [], "out": []}
    }
    
    jeoneuian_dict = {
        "1_5": {"in": [], "out": []}, "1_6": {"in": [], "out": []}, "1_7": {"in": [], "out": []},
        "2_5": {"in": [], "out": []}, "2_6": {"in": [], "out": []}, "2_7": {"in": [], "out": []},
        "기록연구사": {"in": [], "out": []}, 
        "임기제": {"in": [], "out": []}    
    }

    for chunk in chunks:
        if len(chunk) < 5: continue
        
        is_skip_worker = False
        is_specialist_dismiss = False
        is_specialist_appoint = False
        is_gyeomim = False 
        
        for text in chunk:
            cleaned_text = text.replace(" ", "")
            if "필수실무요원" in cleaned_text: is_skip_worker = True
            if "전문관" in cleaned_text and "면함" in cleaned_text: is_specialist_dismiss = True
            if "전문관" in cleaned_text and "명함" in cleaned_text: is_specialist_appoint = True
            if "겸임" in cleaned_text: is_gyeomim = True 
        
        if is_skip_worker or (is_specialist_dismiss and not is_specialist_appoint) or is_gyeomim: continue

        name = chunk[1]
        
        new_hire_idx = -1
        for i, text in enumerate(chunk):
            if "신규임용" in text.replace(" ", ""):
                new_hire_idx = i
                break
                
        if new_hire_idx != -1:
            split_idx = new_hire_idx
        else:
            jibang_indices = [i for i, text in enumerate(chunk) if text.startswith('지방') and '지방공무원법' not in text]
            if len(jibang_indices) >= 2: split_idx = jibang_indices[-1]
            elif len(jibang_indices) == 1: 
                if jibang_indices[0] == 2: split_idx = len(chunk) - 3 if len(chunk) >= 6 else 3
                else: split_idx = jibang_indices[0]
            else: split_idx = -2
            
        to_raw = chunk[2:split_idx]
        from_raw = chunk[split_idx:-1] if date_pattern.match(chunk[-1]) else chunk[split_idx:]
            
        to_rank, to_dept = parse_rank_dept(to_raw)
        from_rank, from_dept = parse_rank_dept(from_raw)
        
        to_full = " ".join(to_raw)
        from_full = " ".join(from_raw)
        
        if "근무를" in to_full:
            if "임함" in to_full:
                dept_match = re.search(r'임함\s*(.+?)\s*근무를\s*명함', to_full)
            else:
                dept_match = re.search(r'(.+?)\s*근무를\s*명함', to_full)
                
            if dept_match: 
                to_dept = dept_match.group(1).strip()
            
        rank_general_match = re.search(r'(.*?)\s*에\s*임함', to_full)
        promo_match = re.search(r'(부이사관|사무관|서기관|주사보|주사|서기보시보|서기보|서기)에\s*임함', to_full)
        
        if promo_match:
            to_rank = promo_match.group(1)
        elif rank_general_match:
            to_rank = rank_general_match.group(1).strip()

        is_fixed_term_person = "임기제" in to_full.replace(" ", "") or "임기제" in from_full.replace(" ", "")

        if not to_rank.strip() or (get_rank_level(to_rank) == -1 and not is_fixed_term_person) or (is_fixed_term_person and "임기제" not in to_rank):
            to_rank = from_rank
            
        if not to_dept.strip():
            to_dept = from_dept

        level_from = get_rank_level(from_rank)
        level_to = get_rank_level(to_rank)
        max_level = max(level_from, level_to)
        
        cat_key = "7급 이하"
        if is_fixed_term_person: cat_key = "임기제"
        elif max_level >= 5: cat_key = "5급 이상"
        elif max_level == 4: cat_key = "6급"
        
        is_promotion = (level_to > level_from) and (level_from != -1)
        if '대우' in to_full: is_promotion = False
            
        is_new_hire = "신규임용" in from_full.replace(" ", "") or "신규임용" in to_full.replace(" ", "")
        is_retirement = "퇴직" in to_full
        is_dispatch = "파견" in to_full
        is_retire_dispatch = is_retirement and is_dispatch 
        is_transfer_out_prov = "전출을명함" in to_full.replace(" ", "")
        is_leave_start = "휴직" in to_full and "명함" in to_full and "복직" not in to_full

        is_transfer_in_prov = False
        if "전입" in to_full.replace(" ", ""):
            is_transfer_in_prov = True
        else:
            other_provinces = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종', '경기', '강원', '충남', '충청남도', '전북', '전라북도', '전남', '전라남도', '경북', '경상북도', '경남', '경상남도', '제주', '교육부']
            for prov in other_provinces:
                if prov in from_full and ("교육청" in from_full or "교육부" in from_full):
                    is_transfer_in_prov = True
                    break
        
        out_gubun = "전출"
        out_rank = to_rank 
        out_note = to_full
        
        if is_retire_dispatch:
            out_gubun = "파견"
            out_note = "퇴직준비교육 파견"
        elif is_retirement:
            if "명예퇴직" in to_full.replace(" ", ""): out_gubun = "명예퇴직"
            elif "정년퇴직" in to_full.replace(" ", ""): out_gubun = "정년퇴직"
            else: out_gubun = "퇴직"
            promo_match2 = re.search(r'(지방[가-힣\s]+?)(부이사관|사무관|서기관|주사|주사보|서기|서기보)에\s*임함', to_full)
            if promo_match2: 
                base_title = from_rank.split()[0] if from_rank else "지방공무원"
                new_rank_title = promo_match2.group(2)
                out_rank = f"{base_title} {new_rank_title}"
            out_note = ""
        elif is_dispatch:
            out_gubun = "파견"
            cleaned_dispatch = re.sub(r'.*?(?:의하여|따라|의거|규정에|호에|항에|법에|근거하여)\s*', '', to_full)
            dispatch_match = re.search(r'(.+?파견)', cleaned_dispatch)
            if dispatch_match: 
                out_note = dispatch_match.group(1).replace("충청북도", "").strip()
            else: 
                out_note = "파견"
            to_dept = out_note 
        elif is_transfer_out_prov:
            out_gubun = "전출"
            transfer_match = re.search(r'(.*?)\s*전출을\s*명함', to_full)
            if transfer_match: out_note = transfer_match.group(1).strip()
            else: out_note = "타시도 전출"
        elif is_leave_start:
            out_gubun = "휴직"
            law_match = re.search(r'(제\d+조(?:\s*제\d+항)?(?:\s*제\d+호)?)', to_full)
            if law_match: out_note = law_match.group(1).replace(" ", "")
            else: out_note = "휴직"
        else:
            out_note = to_dept

        if "파견중" in from_full.replace(" ", ""):
            if not is_retirement:
                out_gubun = "파견복귀"

        if "휴직" in from_full.replace(" ", "") and out_gubun == "전출": out_gubun = "복직"

        in_note = from_dept
        in_rank = to_rank
        in_gubun = "전입"
        
        if is_new_hire:
            in_note = "" 
            in_gubun = "신규임용"
            in_rank = to_rank
        elif "파견" in from_full: in_gubun = "파견복귀"
        elif "휴직" in from_full.replace(" ", ""): in_gubun = "복직"
        
        if is_promotion:
            if in_gubun == "전입": in_gubun = "승진"
            if out_gubun == "전출": out_gubun = "승진"
        elif not is_new_hire and not is_retirement and not is_dispatch and not is_transfer_out_prov and not is_transfer_in_prov and not is_leave_start:
            if level_from != -1 and level_from == level_to and from_dept.replace(" ", "") != to_dept.replace(" ", ""):
                if in_gubun == "전입": in_gubun = "전보"
                if out_gubun == "전출": out_gubun = "전보"

        is_specialist = "전문관" in to_full.replace(" ", "") and "명함" in to_full.replace(" ", "")
        if is_specialist:
            specialist_match = re.search(r'(전문관\s*\([^)]+\))', to_full)
            specialist_str = specialist_match.group(1).strip() if specialist_match else "전문관"
            in_note = specialist_str
            if out_gubun == "전출" or out_gubun == "전보": out_note = f"{to_dept} {specialist_str}"
            else: out_note = specialist_str
            in_gubun = "전문관지정"
            out_gubun = "전문관지정"

        is_in_local = check_is_local(to_full, local_agencies_for_sort)
        is_out_local = check_is_local(from_full, local_agencies_for_sort)
        if is_new_hire: is_out_local = False
            
        in_rank = simplify_rank_str(in_rank, is_fixed_term_person)
        out_rank = simplify_rank_str(out_rank, is_fixed_term_person)
        
        grp_in_num, _ = get_series_and_group(in_rank)
        grp_out_num, _ = get_series_and_group(out_rank)

        if max_level >= 5: sub_grp_str = "5"
        elif max_level == 4: sub_grp_str = "6"
        else: sub_grp_str = "7"

        if is_fixed_term_person:
            in_dict_key = "임기제"
            out_dict_key = "임기제"
        else:
            if grp_in_num == 3: in_dict_key = "기록연구사"
            else: in_dict_key = f"{grp_in_num}_{sub_grp_str}"
            
            if grp_out_num == 3: out_dict_key = "기록연구사"
            else: out_dict_key = f"{grp_out_num}_{sub_grp_str}"

        in_note = str(in_note).replace("근무를 명함", "").replace("근무를명함", "").replace("((", "(").replace("))", ")").strip()
        out_note = str(out_note).replace("근무를 명함", "").replace("근무를명함", "").replace("((", "(").replace("))", ")").strip()
        to_dept = str(to_dept).replace("근무를 명함", "").replace("근무를명함", "").replace("((", "(").replace("))", ")").strip()

        if is_in_local and not is_out_local:
            in_dict = {"구분": in_gubun, "임용 부서": to_dept, "직급": in_rank, "성명": name, "비고": in_note}
            categories[cat_key]["in"].append(in_dict) 
            if in_gubun != "전문관지정": 
                jeoneuian_dict[in_dict_key]["in"].append(in_dict)
            
        elif is_out_local and not is_in_local:
            out_dict = {"구분_": out_gubun, "현임 부서_": from_dept.replace("(퇴직준비교육 파견중)", "").strip(), "직급_": out_rank, "성명_": name, "비고_": out_note}
            categories[cat_key]["out"].append(out_dict)
            if out_gubun != "전문관지정": 
                jeoneuian_dict[out_dict_key]["out"].append(out_dict)
            
        elif is_in_local and is_out_local:
            gubun_str = "자체전보" 
            if is_promotion: gubun_str = "승진"
            if is_new_hire: gubun_str = "신규임용" 
            elif is_dispatch: gubun_str = "파견"
            elif "파견" in from_full: gubun_str = "파견복귀"
            elif "휴직" in from_full.replace(" ", ""): gubun_str = "복직"
            elif is_leave_start: gubun_str = "휴직"
            elif is_specialist: gubun_str = "전문관지정"
                
            in_dict = {"구분": gubun_str, "임용 부서": to_dept, "직급": in_rank, "성명": name, "비고": in_note}
            out_dict = {"구분_": gubun_str, "현임 부서_": from_dept.replace("(퇴직준비교육 파견중)", "").strip(), "직급_": out_rank, "성명_": name, "비고_": out_note}
            
            categories[cat_key]["in"].append(in_dict)
            categories[cat_key]["out"].append(out_dict)
            
            if gubun_str != "전문관지정":
                jeoneuian_dict[in_dict_key]["in"].append(in_dict)
                jeoneuian_dict[out_dict_key]["out"].append(out_dict)

    for key in jeoneuian_dict:
        jeoneuian_dict[key]["in"].sort(key=sort_key_in)
        jeoneuian_dict[key]["out"].sort(key=sort_key_out)

    def write_sheet_data(wb, sheet_name, main_title, sections_data):
        # ★ [수정] 전입자와 전출자가 모두 없는 그룹은 표시하지 않음
        #   (한쪽이라도 있으면 그 그룹은 그대로 표시하고, 빈 쪽만 '해당 없음'으로 표시)
        non_empty_sections = [s for s in sections_data if s[2] or s[3]]
        if non_empty_sections:
            sections_data = non_empty_sections
        else:
            # 전 그룹이 비어 있는 경우에도 제목/양식이 깨지지 않도록 첫 그룹 하나만 남김
            sections_data = sections_data[:1]

        if sheet_name in wb.sheetnames:
            sheet_idx = wb.sheetnames.index(sheet_name)
            wb.remove(wb[sheet_name])
            worksheet = wb.create_sheet(sheet_name, sheet_idx)
        else:
            worksheet = wb.create_sheet(sheet_name)
            
        current_row = 4 
        header_rows = []
        data_ranges = []
        is_first_section = True
        
        MAX_ROWS_PER_PAGE = 38
        
        total_rows_needed = 3
        for _, _, in_list, out_list in sections_data:
            len_in = len(in_list) if in_list else 1
            len_out = len(out_list) if out_list else 1
            total_rows_needed += (max(len_in, len_out) + 3)
            
        is_total_fits_one_page = (total_rows_needed <= MAX_ROWS_PER_PAGE)
        
        current_page_rows = 3
        
        for idx, (title_in, title_out, in_list, out_list) in enumerate(sections_data):
            if not in_list:
                in_list.append({"구분": "", "임용 부서": "해당 없음", "직급": "", "성명": "", "비고": ""})
            if not out_list:
                out_list.append({"구분_": "", "현임 부서_": "해당 없음", "직급_": "", "성명_": "", "비고_": ""})
                
            max_len = max(len(in_list), len(out_list))
            section_rows = max_len + 3
            
            is_first_after_group1 = False
            if idx > 0:
                prev_title = sections_data[idx-1][0]
                if "1그룹" in prev_title and "1그룹" not in title_in:
                    is_first_after_group1 = True
            
            force_break_for_group2 = is_first_after_group1 and not is_total_fits_one_page
            
            # 1. 페이지 나누기 판단
            if force_break_for_group2 or (current_page_rows + section_rows > MAX_ROWS_PER_PAGE):
                if current_row > 4 and current_page_rows > 3:
                    worksheet.row_breaks.append(Break(id=current_row - 1))
                    current_page_rows = 3
                    
            # 2. 이번에 쓴 그룹의 물량을 현재 페이지에 누적
            current_page_rows += section_rows
            
            # 3. ★ 거대 그룹 자동 페이지 넘김 시 잔여 줄 수 추적
            while current_page_rows > MAX_ROWS_PER_PAGE:
                current_page_rows -= (MAX_ROWS_PER_PAGE - 3)
            
            in_padded = in_list + [{"구분":"", "임용 부서":"", "직급":"", "성명":"", "비고":""}] * (max_len - len(in_list))
            out_padded = out_list + [{"구분_":"", "현임 부서_":"", "직급_":"", "성명_":"", "비고_":""}] * (max_len - len(out_list))

            df_in = pd.DataFrame(in_padded, columns=["구분", "임용 부서", "직급", "성명", "비고"])
            df_empty = pd.DataFrame([{" ": ""}] * max_len)
            df_out = pd.DataFrame(out_padded, columns=["구분_", "현임 부서_", "직급_", "성명_", "비고_"])
            df_out.columns = ["구분", "현임 부서", "직급", "성명", "비고"]
            df_final = pd.concat([df_in, df_empty, df_out], axis=1)

            headers = ["구분", "임용 부서", "직급", "성명", "비고", " ", "구분", "현임 부서", "직급", "성명", "비고"]
            for c_idx, val in enumerate(headers, start=2): 
                worksheet.cell(row=current_row + 1, column=c_idx, value=val)

            for r_idx, row_data in enumerate(df_final.values, start=current_row + 2):
                for c_idx, val in enumerate(row_data, start=2):
                    if pd.isna(val):
                        val = ""
                    worksheet.cell(row=r_idx, column=c_idx, value=val)
            
            if is_first_section:
                main_title_row = 2
                worksheet.merge_cells(f'B{main_title_row}:L{main_title_row}')
                title_cell = worksheet[f'B{main_title_row}']
                title_cell.value = main_title
                title_cell.font = Font(bold=True, size=20)
                title_cell.alignment = Alignment(horizontal='center', vertical='center')
                worksheet.row_dimensions[main_title_row].height = 30
                is_first_section = False
            
            worksheet.merge_cells(f'B{current_row}:F{current_row}')
            title_cell_in = worksheet[f'B{current_row}']
            title_cell_in.value = title_in
            title_cell_in.font = Font(bold=True, size=13)
            
            worksheet.merge_cells(f'H{current_row}:L{current_row}')
            title_cell_out = worksheet[f'H{current_row}']
            title_cell_out.value = title_out
            title_cell_out.font = Font(bold=True, size=13)
            
            header_row = current_row + 1
            header_rows.append(header_row)
            data_end_row = header_row + max_len
            data_ranges.append((header_row, data_end_row))
            
            current_row = data_end_row + 2
            
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        header_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        center_align = Alignment(horizontal='center', vertical='center')

        data_columns = list(range(2, 7)) + list(range(8, 13))
        
        for start_r, end_r in data_ranges:
            for row in range(start_r, end_r + 1):
                for col in data_columns:
                    cell = worksheet.cell(row=row, column=col)
                    cell.border = thin_border
                    if row in header_rows:
                        cell.fill = header_fill
                        cell.font = Font(bold=True)
                        cell.alignment = center_align

        title_strings = [s[0] for s in sections_data] + [s[1] for s in sections_data]
        for col in data_columns:
            max_length = 0
            col_letter = get_column_letter(col)
            for cell in worksheet[col_letter]:
                if cell.row == 2: continue
                if cell.value and str(cell.value) in title_strings: continue
                if cell.value:
                    val_str = str(cell.value)
                    length = sum(1.8 if ord(c) > 127 else 1.1 for c in val_str)
                    if length > max_length: max_length = length
            adjusted_width = max(max_length + 2, 8)
            worksheet.column_dimensions[col_letter].width = adjusted_width
            
        worksheet.column_dimensions['G'].width = 3
        
        worksheet.page_setup.orientation = worksheet.ORIENTATION_LANDSCAPE
        worksheet.page_setup.paperSize = worksheet.PAPERSIZE_A4 
        worksheet.page_margins = PageMargins(left=0.25, right=0.25, top=0.75, bottom=0.75, header=0.3, footer=0.3) 
        
        worksheet.sheet_properties.pageSetUpPr.fitToPage = True
        worksheet.page_setup.fitToWidth = 1  
        worksheet.page_setup.fitToHeight = 0 
        
        worksheet.print_title_rows = '1:3'   

    jeoneuian_sections = [
        ("[ 1그룹 5급 이상 전입 ]", "[ 1그룹 5급 이상 전출 ]", jeoneuian_dict["1_5"]["in"], jeoneuian_dict["1_5"]["out"]),
        ("[ 1그룹 6급 전입 ]", "[ 1그룹 6급 전출 ]", jeoneuian_dict["1_6"]["in"], jeoneuian_dict["1_6"]["out"]),
        ("[ 1그룹 7급 이하 전입 ]", "[ 1그룹 7급 이하 전출 ]", jeoneuian_dict["1_7"]["in"], jeoneuian_dict["1_7"]["out"]),
        ("[ 2그룹 5급 이상 전입 ]", "[ 2그룹 5급 이상 전출 ]", jeoneuian_dict["2_5"]["in"], jeoneuian_dict["2_5"]["out"]),
        ("[ 2그룹 6급 전입 ]", "[ 2그룹 6급 전출 ]", jeoneuian_dict["2_6"]["in"], jeoneuian_dict["2_6"]["out"]),
        ("[ 2그룹 7급 이하 전입 ]", "[ 2그룹 7급 이하 전출 ]", jeoneuian_dict["2_7"]["in"], jeoneuian_dict["2_7"]["out"]),
        ("[ 기록연구사 전입 ]", "[ 기록연구사 전출 ]", jeoneuian_dict["기록연구사"]["in"], jeoneuian_dict["기록연구사"]["out"]),
        ("[ 임기제 전입 ]", "[ 임기제 전출 ]", jeoneuian_dict["임기제"]["in"], jeoneuian_dict["임기제"]["out"])
    ]

    wb = openpyxl.Workbook()
    wb.remove(wb.active)   # 새 파일의 기본 빈 시트 제거

    write_sheet_data(wb, '전의', f"{first_date_str} 자 지방공무원 인사발령 전의 기초자료", jeoneuian_sections)

    wb.active = 0
    return wb, first_date_str


# ==================================================
# ★ 설정 저장/불러오기 (현재 상태 기억하기)
# ==================================================
def load_settings():
    try:
        with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}

def save_settings(data):
    try:
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ==================================================
# ★ 화면(GUI)
# ==================================================
FONT = "맑은 고딕"
C_BG = "#F3F5F9"
C_CARD = "#FFFFFF"
C_BORDER = "#D5DBE6"
C_ACCENT = "#2563EB"
C_ACCENT_HOVER = "#1D4ED8"
C_TEXT = "#1F2937"
C_SUB = "#6B7280"
C_SOFT = "#E8EEFB"
C_SOFT_HOVER = "#D6E1F8"


class App(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, regions):
        super().__init__()
        self.TkdndVersion = TkinterDnD._require(self)

        self.title(f"충청북도교육청 지방공무원 인사발령 정리 v{current_version}")
        self.geometry("780x640")
        self.resizable(False, False)
        self.configure(fg_color=C_BG)

        self.regions = list(regions)
        self.hwp_path = ""
        self._loading = True

        saved = load_settings()
        remember = bool(saved.get("remember", False))
        region0 = saved.get("region", "") if remember else ""
        if region0 not in self.regions:
            region0 = self.regions[0]
        save_mode0 = saved.get("save_mode", "same") if remember else "same"
        if save_mode0 not in ("same", "ask"):
            save_mode0 = "same"

        self.var_region = tk.StringVar(value=region0)
        self.var_remember = tk.BooleanVar(value=remember)
        self.var_save_mode = tk.StringVar(value=save_mode0)

        self.build_ui()
        self.register_drop_targets(self)

        self._loading = False
        for v in (self.var_region, self.var_remember, self.var_save_mode):
            v.trace_add("write", lambda *a: self.persist())

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(600, self.check_update_once)

    # ---------- 화면 만들기 ----------
    def build_ui(self):
        self.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(self, text="충청북도교육청 지방공무원 인사발령 정리", font=ctk.CTkFont(FONT, 24, "bold"), text_color=C_TEXT)
        title.grid(row=0, column=0, padx=28, pady=(22, 0), sticky="w")
        sub = ctk.CTkLabel(self, text="충청북도교육청 인사발령서에서 선택한 지역의 전입·전출을 정리해 엑셀로 저장합니다.",
                           font=ctk.CTkFont(FONT, 13), text_color=C_SUB)
        sub.grid(row=1, column=0, padx=28, pady=(2, 14), sticky="w")

        # ----- 가운데: 왼쪽(인사발령서) + 오른쪽(지역/기관현황) -----
        mid = ctk.CTkFrame(self, fg_color="transparent")
        mid.grid(row=2, column=0, padx=28, sticky="ew")
        mid.grid_columnconfigure(0, weight=3, uniform="mid")
        mid.grid_columnconfigure(1, weight=2, uniform="mid")

        # 왼쪽 카드: 드롭 영역
        drop = ctk.CTkFrame(mid, fg_color=C_CARD, corner_radius=14, border_width=2, border_color=C_ACCENT)
        drop.grid(row=0, column=0, padx=(0, 10), sticky="nsew")
        ctk.CTkLabel(drop, text="📄", font=ctk.CTkFont(FONT, 38)).pack(pady=(22, 0))
        ctk.CTkLabel(drop, text="충청북도교육청 인사발령서를\n여기로 끌어다 놓으세요",
                     font=ctk.CTkFont(FONT, 15, "bold"), text_color=C_TEXT).pack(pady=(4, 2))
        ctk.CTkLabel(drop, text="(hwp, hwpx)", font=ctk.CTkFont(FONT, 12), text_color=C_SUB).pack()
        ctk.CTkButton(drop, text="인사발령서 선택", width=170, height=36, corner_radius=10,
                      font=ctk.CTkFont(FONT, 14, "bold"), fg_color=C_ACCENT, hover_color=C_ACCENT_HOVER,
                      command=self.pick_file).pack(pady=(14, 8))
        self.lbl_file = ctk.CTkLabel(drop, text="선택된 파일 없음", font=ctk.CTkFont(FONT, 12),
                                     text_color=C_SUB, wraplength=330)
        self.lbl_file.pack(pady=(0, 18), padx=12)

        # 오른쪽 카드: 지역 + 기관현황 열기
        side = ctk.CTkFrame(mid, fg_color=C_CARD, corner_radius=14, border_width=1, border_color=C_BORDER)
        side.grid(row=0, column=1, padx=(10, 0), sticky="nsew")
        ctk.CTkLabel(side, text="지역 선택", font=ctk.CTkFont(FONT, 14, "bold"),
                     text_color=C_TEXT).pack(anchor="w", padx=18, pady=(20, 6))
        row = ctk.CTkFrame(side, fg_color="transparent")
        row.pack(fill="x", padx=18)
        self.menu_region = ctk.CTkOptionMenu(row, values=self.regions, variable=self.var_region, height=36,
                                             font=ctk.CTkFont(FONT, 14), dropdown_font=ctk.CTkFont(FONT, 14),
                                             fg_color=C_SOFT, button_color=C_ACCENT, button_hover_color=C_ACCENT_HOVER,
                                             text_color=C_TEXT)
        self.menu_region.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(row, text="↻", width=36, height=36, corner_radius=8, font=ctk.CTkFont(FONT, 16, "bold"),
                      fg_color=C_SOFT, hover_color=C_SOFT_HOVER, text_color=C_ACCENT,
                      command=self.refresh_regions).pack(side="left", padx=(6, 0))
        ctk.CTkLabel(side, text="↻ : 기관현황.xlsx를 고친 뒤\n지역 목록 새로고침", justify="left",
                     font=ctk.CTkFont(FONT, 11), text_color=C_SUB).pack(anchor="w", padx=18, pady=(6, 0))
        ctk.CTkButton(side, text="기관현황 열기", height=36, corner_radius=10, font=ctk.CTkFont(FONT, 14, "bold"),
                      fg_color=C_SOFT, hover_color=C_SOFT_HOVER, text_color=C_ACCENT,
                      command=self.open_agency_file).pack(fill="x", padx=18, pady=(18, 20))

        # ----- 옵션 -----
        opt = ctk.CTkFrame(self, fg_color=C_CARD, corner_radius=14, border_width=1, border_color=C_BORDER)
        opt.grid(row=3, column=0, padx=28, pady=(14, 0), sticky="ew")
        ctk.CTkLabel(opt, text="옵션", font=ctk.CTkFont(FONT, 14, "bold"),
                     text_color=C_TEXT).grid(row=0, column=0, columnspan=3, padx=18, pady=(14, 6), sticky="w")
        ctk.CTkCheckBox(opt, text="현재 상태 기억하기 (지역과 옵션값을 저장했다가 다음 실행 때 불러옵니다)",
                        variable=self.var_remember, font=ctk.CTkFont(FONT, 13), text_color=C_TEXT,
                        fg_color=C_ACCENT, hover_color=C_ACCENT_HOVER).grid(row=1, column=0, columnspan=3, padx=18, pady=(0, 8), sticky="w")
        ctk.CTkLabel(opt, text="결과물 저장 위치", font=ctk.CTkFont(FONT, 13, "bold"),
                     text_color=C_TEXT).grid(row=2, column=0, columnspan=3, padx=18, pady=(4, 4), sticky="w")
        ctk.CTkRadioButton(opt, text="충청북도교육청 인사발령서가 있는 폴더", variable=self.var_save_mode, value="same",
                           font=ctk.CTkFont(FONT, 13), text_color=C_TEXT,
                           fg_color=C_ACCENT, hover_color=C_ACCENT_HOVER).grid(row=3, column=0, columnspan=3, padx=18, pady=3, sticky="w")
        ctk.CTkRadioButton(opt, text="직접 선택하여 저장 ([정리 시작]을 누르면 저장할 폴더를 묻습니다)",
                           variable=self.var_save_mode, value="ask",
                           font=ctk.CTkFont(FONT, 13), text_color=C_TEXT,
                           fg_color=C_ACCENT, hover_color=C_ACCENT_HOVER).grid(row=4, column=0, columnspan=3, padx=18, pady=(3, 14), sticky="w")

        # ----- 실행 -----
        self.btn_run = ctk.CTkButton(self, text="정리 시작", height=48, corner_radius=12,
                                     font=ctk.CTkFont(FONT, 17, "bold"), fg_color=C_ACCENT,
                                     hover_color=C_ACCENT_HOVER, command=self.start)
        self.btn_run.grid(row=4, column=0, padx=28, pady=(16, 6), sticky="ew")
        self.lbl_status = ctk.CTkLabel(self, text="인사발령서를 선택하고 지역을 고른 뒤 [정리 시작]을 눌러주세요.",
                                       font=ctk.CTkFont(FONT, 12), text_color=C_SUB)
        self.lbl_status.grid(row=5, column=0, padx=28, pady=(0, 10))

    # ---------- 드래그 앤 드롭 ----------
    def register_drop_targets(self, widget):
        """창 안의 어느 곳에 놓아도 받을 수 있도록 모든 위젯에 등록합니다."""
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind('<<Drop>>', self.on_drop)
        except Exception:
            pass
        for child in widget.winfo_children():
            self.register_drop_targets(child)

    def on_drop(self, event):
        try:
            files = self.tk.splitlist(event.data)
        except Exception:
            files = [event.data]
        for f in files:
            if str(f).lower().endswith(('.hwp', '.hwpx')):
                self.set_hwp(str(f))
                return
        messagebox.showwarning("파일 형식 확인", "한글 파일(hwp, hwpx)만 사용할 수 있습니다.", parent=self)

    # ---------- 버튼 동작 ----------
    def set_hwp(self, path):
        self.hwp_path = path
        self.lbl_file.configure(text=os.path.basename(path), text_color=C_TEXT)
        self.set_status("인사발령서가 선택되었습니다. 지역을 확인하고 [정리 시작]을 눌러주세요.")

    def pick_file(self):
        path = filedialog.askopenfilename(
            parent=self, title="충청북도교육청 인사발령서 파일 선택",
            filetypes=[("한글 파일", "*.hwp *.hwpx"), ("모든 파일", "*.*")])
        if path:
            self.set_hwp(path)

    def open_agency_file(self):
        if not os.path.exists(AGENCY_FILE):
            messagebox.showerror("파일 없음", "기관현황.xlsx 파일이 존재하지 않습니다.\n프로그램(exe)과 같은 폴더에 넣어주세요.", parent=self)
            return
        try:
            os.startfile(AGENCY_FILE)
        except Exception as e:
            messagebox.showerror("열기 실패", f"기관현황.xlsx를 여는 데 실패했습니다.\n{e}", parent=self)

    def refresh_regions(self):
        try:
            _, regions = read_agency_table()
        except AppError as e:
            messagebox.showerror(e.title, str(e), parent=self)
            return False
        self.regions = regions
        self.menu_region.configure(values=regions)
        if self.var_region.get() not in regions:
            self.var_region.set(regions[0])
        return True

    def set_status(self, text):
        self.lbl_status.configure(text=text)
        self.update()

    def persist(self):
        if self._loading:
            return
        if self.var_remember.get():
            save_settings({
                "remember": True,
                "region": self.var_region.get(),
                "save_mode": self.var_save_mode.get(),
            })
        else:
            save_settings({"remember": False})

    def on_close(self):
        self.persist()
        self.destroy()

    def check_update_once(self):
        if check_for_update(self):
            self.on_close()

    # ---------- 정리 실행 ----------
    def start(self):
        if not self.hwp_path or not os.path.exists(self.hwp_path):
            messagebox.showwarning("인사발령서 선택", "충청북도교육청 인사발령서 파일을 먼저 선택해주세요.", parent=self)
            return

        region = self.var_region.get()
        self.btn_run.configure(state="disabled", text="처리 중입니다...")
        try:
            self.set_status("기관현황.xlsx를 읽고 있습니다...")
            df_data, regions = read_agency_table()
            if region not in regions:
                raise AppError("지역 확인", f"기관현황.xlsx에 '{region}' 지역이 없습니다.\n지역을 다시 선택해주세요.")
            apply_region(df_data, region)

            if self.var_save_mode.get() == "ask":
                out_dir = filedialog.askdirectory(
                    parent=self, title="결과물을 저장할 폴더 선택",
                    initialdir=os.path.dirname(self.hwp_path))
                if not out_dir:
                    self.set_status("저장할 폴더를 선택하지 않아 작업을 취소했습니다.")
                    return
            else:
                out_dir = os.path.dirname(self.hwp_path)

            self.set_status("한글 파일에서 데이터를 추출하고 있습니다... (한글 창이 잠시 열립니다)")
            text_list = extract_hwp_text(self.hwp_path)

            self.set_status("인사발령을 분류·정렬하고 있습니다...")
            wb, date_str = build_report(text_list)

            out_path = os.path.join(out_dir, f"{date_str} 인사발령 전의 기초자료.xlsx")
            if os.path.exists(out_path):
                if not messagebox.askyesno("덮어쓰기 확인", f"같은 이름의 파일이 이미 있습니다.\n덮어쓰시겠습니까?\n\n{out_path}", parent=self):
                    self.set_status("저장을 취소했습니다.")
                    return
            try:
                wb.save(out_path)
            except PermissionError:
                raise AppError("파일 열림 오류", f"'{os.path.basename(out_path)}' 파일이 이미 열려 있어 저장할 수 없습니다.\n엑셀을 완전히 종료하고 다시 실행해주세요.")
            finally:
                wb.close()

            self.set_status(f"완료! '{os.path.basename(out_path)}' 파일을 저장했습니다.")
            try:
                os.startfile(out_path)
            except Exception as e:
                messagebox.showwarning("파일 열기 실패", f"저장은 완료되었지만 자동으로 여는 데 실패했습니다.\n{out_path}\n{e}", parent=self)

        except AppError as e:
            self.set_status("오류가 발생했습니다.")
            messagebox.showerror(e.title, str(e), parent=self)
        except Exception as e:
            err = traceback.format_exc()
            print(err)
            self.set_status("오류가 발생했습니다.")
            messagebox.showerror("프로그램 오류", f"데이터 분석 또는 엑셀 생성 중 오류가 발생했습니다.\n\n원인: {e}", parent=self)
        finally:
            self.btn_run.configure(state="normal", text="정리 시작")


# ==================================================
# ★ 시작
# ==================================================
def main():
    try:
        _, regions = read_agency_table()
    except AppError as e:
        r = tk.Tk()
        r.withdraw()
        if "존재하지 않습니다" in str(e):
            messagebox.showerror("파일 없음", "기관현황.xlsx 파일이 없어 프로그램을 종료합니다.\n\n프로그램(exe)과 같은 폴더에 기관현황.xlsx를 넣고 다시 실행해주세요.")
        else:
            messagebox.showerror(e.title, f"{e}\n\n프로그램을 종료합니다.")
        sys.exit()

    ctk.set_appearance_mode("light")
    app = App(regions)
    app.mainloop()


if __name__ == "__main__":
    main()
