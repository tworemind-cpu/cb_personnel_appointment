import pandas as pd
import re
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.pagebreak import Break
import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client as win32
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

# ==================================================
# ★ 업데이트 확인 설정 (버전을 올릴 때는 아래 current_version만 고치세요)
# ==================================================
current_version = "1.7.0"
version_info = "https://raw.githubusercontent.com/tworemind-cpu/personnel_appointment/main/version.json"

# --- ★ 실행 파일(.exe) 기준 현재 폴더 경로 찾기 ---
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# --- ★ 업데이트 확인 함수 ---
def parse_version(ver_str):
    """'1.2.3' 같은 문자열을 (1, 2, 3) 숫자 묶음으로 바꿉니다. (크기 비교용)"""
    nums = re.findall(r'\d+', str(ver_str))
    return tuple(int(n) for n in nums) if nums else (0,)

def check_for_update():
    """
    깃허브의 version.json을 읽어 새 버전이 있으면 안내 팝업을 띄웁니다.
    - 인터넷이 안 되거나 주소가 틀려도 프로그램은 멈추지 않고 그냥 넘어갑니다.
    - '예'를 누르면 다운로드 페이지를 열고 프로그램을 종료합니다.
    """
    try:
        # 주소에 한글이 들어 있어도 읽을 수 있도록 변환
        safe_url = urllib.parse.quote(version_info, safe=":/?=&%")
        req = urllib.request.Request(safe_url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=5) as res:
            data = json.loads(res.read().decode("utf-8"))

        latest_version = str(data.get("version", "")).strip()
        download_url = str(data.get("download_url", "")).strip()
        notes = str(data.get("notes", "")).strip()

        if not latest_version:
            return

        if parse_version(latest_version) > parse_version(current_version):
            msg = (f"새 버전이 나왔습니다.\n\n"
                   f"현재 버전: {current_version}\n"
                   f"최신 버전: {latest_version}\n")
            if notes:
                msg += f"\n[변경 내용]\n{notes}\n"
            msg += "\n지금 다운로드 페이지를 여시겠습니까?\n(예를 누르면 프로그램이 종료됩니다.)"

            if messagebox.askyesno("업데이트 안내", msg):
                if download_url:
                    webbrowser.open(download_url)
                sys.exit()
    except SystemExit:
        raise
    except Exception:
        # 인터넷 연결 없음, 주소 오류, 파일 형식 오류 등 → 조용히 넘어감
        pass

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

# --- 0. UI 설정 및 파일 확인 ---
root = tk.Tk()
root.withdraw()
root.attributes('-topmost', True)

# --- ★ 새 버전 확인 (프로그램 시작 직후 1회) ---
check_for_update()

agency_path = os.path.join(BASE_DIR, '인사발령정리.xlsx')
if not os.path.exists(agency_path):
    messagebox.showerror("파일 누락 안내", f"현재 폴더에 [ 인사발령정리.xlsx ] 파일이 없습니다.\n해당 파일 내에 '기관현황' 시트가 필요합니다.")
    sys.exit()

print("==================================================")
print("지방공무원 인사발령 정리 자동화 프로그램")
print("==================================================\n")

print("화면에 뜬 팝업창에서 '도교육청 인사발령서(hwp, hwpx)' 파일을 선택해주세요...")
hwp_file_path = filedialog.askopenfilename(
    title="도교육청 인사발령서 파일 선택",
    filetypes=[("한글 파일", "*.hwp *.hwpx"), ("모든 파일", "*.*")]
)

if not hwp_file_path:
    print("\n파일 선택이 취소되었습니다.")
    time.sleep(2)
    sys.exit()

print(f"\n[선택된 파일]: {os.path.basename(hwp_file_path)}")
print("한글 파일에서 데이터를 추출하고 있습니다...")

temp_txt_path = os.path.join(BASE_DIR, "temp_발령서_임시추출.txt")
text_list = []

# ★ 한글 보안 팝업창 자동 허용 함수
def auto_allow_popup():
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
    print(f"\n[오류] 한글 파일 처리 실패: {e}")
    time.sleep(3)
    sys.exit()
finally:
    try: hwp.Quit()
    except: pass
    ctypes.windll.user32.keybd_event(0x12, 0, 0x0002, 0)

try:
    try:
        with open(temp_txt_path, 'r', encoding='utf-16') as f:
            text_list = [line.strip() for line in f.readlines() if line.strip()]
    except:
        with open(temp_txt_path, 'r', encoding='cp949') as f:
            text_list = [line.strip() for line in f.readlines() if line.strip()]
    if os.path.exists(temp_txt_path):
        os.remove(temp_txt_path)
except Exception as e:
    print(f"텍스트 변환 오류: {e}")
    sys.exit()

print("인사발령정리 엑셀의 '기관현황' 시트와 대조를 시작합니다...")
try:
    df_agency = pd.read_excel(agency_path, sheet_name='기관현황', header=None)
except PermissionError:
    messagebox.showerror("파일 열림 오류", f"'{os.path.basename(agency_path)}' 파일이 이미 엑셀에서 열려 있습니다.\n\n안전한 데이터 처리를 위해 열려있는 엑셀 파일을 완전히 종료한 후 파이썬을 다시 실행해주세요.")
    sys.exit()
except Exception as e:
    messagebox.showerror("엑셀 읽기 오류", f"'기관현황' 시트 분석 중 오류가 발생했습니다.\n{e}")
    sys.exit()
    
try:
    base_region = str(df_agency.iloc[2, 1]).strip()
    
    try:
        special_agency = str(df_agency.iloc[1, 3]).strip()
    except IndexError:
        special_agency = ""
    
    df_data = df_agency.iloc[2:1000, 2:5].copy()
    df_data.columns = ['지역', '임용기관', '부서']
    
    df_data['지역'] = df_data['지역'].astype(str).str.strip()
    df_data['임용기관'] = df_data['임용기관'].astype(str).str.strip()
    df_data['부서'] = df_data['부서'].astype(str).str.strip()
    
    df_local = df_data[df_data['지역'] == base_region]
    
    raw_agencies = df_local['임용기관'].replace('nan', '').replace('None', '').dropna().tolist()
    raw_sub_depts = df_local['부서'].replace('nan', '').replace('None', '').dropna().tolist()
    
    if special_agency and special_agency.lower() not in ['nan', 'none']:
        raw_agencies.append(special_agency)
    
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
    
    local_agencies_for_sort = [a for a in raw_agencies if len(a) > 1 and a not in ignore_keywords]
    dept_order_map = {agency: i for i, agency in enumerate(local_agencies_for_sort)}
    sub_dept_order_map = {dept: i for i, dept in enumerate(raw_sub_depts) if len(dept) > 1}
        
except Exception as e:
    messagebox.showerror("엑셀 읽기 오류", f"'기관현황' 데이터 처리 중 오류가 발생했습니다.\n{e}")
    sys.exit()

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

try:
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

    first_date_str = "YYYY. M. D."
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

    output_file = agency_path
    
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

    wb = openpyxl.load_workbook(output_file)
                
    write_sheet_data(wb, '전의', f"{first_date_str} 자 지방공무원 인사발령 전의 기초자료", jeoneuian_sections)
    
    if '전의' in wb.sheetnames:
        wb.active = wb.sheetnames.index('전의')
        for ws in wb.worksheets:
            if ws.views:
                ws.views.sheetView[0].tabSelected = False
        wb.active.views.sheetView[0].tabSelected = True
    
    wb.save(output_file)
    wb.close()

    print(f"\n작업 완료! 기존 파일에 '전의' 시트가 성공적으로 생성되었습니다!")
    print(f"'{os.path.basename(output_file)}' 파일이 화면에 자동으로 열립니다!")
    time.sleep(1)

    try:
        os.startfile(output_file)
    except Exception as e:
        print(f"엑셀 파일을 자동으로 여는 데 실패했습니다: {e}")

except PermissionError: 
    messagebox.showerror("파일 열림 오류", f"'{os.path.basename(output_file)}' 파일이 이미 열려 있어 덮어쓸 수 없습니다.\n엑셀을 완전히 종료하고 다시 실행해주세요.")
except Exception as e:
    error_msg = traceback.format_exc()
    print(f"\n[오류 발생] 데이터 분석 또는 엑셀 생성 중 문제가 발생했습니다.\n{error_msg}")
    messagebox.showerror("프로그램 오류", f"데이터 분석 또는 엑셀 생성 중 오류가 발생했습니다.\n\n터미널(콘솔) 창의 에러 메시지를 확인해주세요.\n원인: {e}")
