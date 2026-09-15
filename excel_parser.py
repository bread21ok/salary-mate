"""
excel_parser.py
================
운영현황.xlsx / 개인부담금.xlsx / 자치단체부담금.xlsx 3개 입력 파일을 읽는다.
- 개인부담금·자치단체부담금 파일은 상단에 병합셀 헤더가 있으므로
  openpyxl로 병합을 해제한 뒤, "당월보험료" 컬럼 위치를 동적으로 탐지한다.
"""

import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter

EXCLUDE_KEYWORDS = ["소계", "합계", "계"]


def _is_summary_row(value) -> bool:
    """소계/합계/계 등 집계 행인지 판단"""
    if value is None:
        return False
    text = str(value).strip()
    return text in EXCLUDE_KEYWORDS


# ---------------------------------------------------------------------
# 1. 운영현황.xlsx
# ---------------------------------------------------------------------
def load_operation_status(filepath: str) -> pd.DataFrame:
    """
    운영현황 시트: 3행이 헤더(0-index 2), 4행은 예시데이터(제외), 5행부터 실제 데이터
    """
    df = pd.read_excel(filepath, sheet_name="운영현황", header=2)
    # header=2로 읽으면 첫 데이터 행(=엑셀 4행)이 예시데이터이므로 제거
    df = df.iloc[1:].reset_index(drop=True)
    # 완전 빈 행 제거 (성명 없는 행)
    if "성명" in df.columns:
        df = df[df["성명"].notna()].reset_index(drop=True)
    return df


# ---------------------------------------------------------------------
# 공통: 병합셀 해제 + 헤더 동적 탐지 유틸
# ---------------------------------------------------------------------
def _unmerge_and_get_grid(filepath: str, sheet_name: str = None):
    """
    워크북을 열어 모든 병합셀을 해제하고, 해제된 값(좌상단 값을 전체 셀에 채움)을
    포함한 2차원 리스트(grid)를 반환한다.
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active

    merged_ranges = list(ws.merged_cells.ranges)
    for merged_range in merged_ranges:
        min_col, min_row, max_col, max_row = (
            merged_range.min_col,
            merged_range.min_row,
            merged_range.max_col,
            merged_range.max_row,
        )
        top_left_value = ws.cell(row=min_row, column=min_col).value
        ws.unmerge_cells(str(merged_range))
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                ws.cell(row=row, column=col).value = top_left_value

    grid = []
    for row in ws.iter_rows(values_only=True):
        grid.append(list(row))
    return grid


def _find_column_index(header_row_1, header_row_2, major_keyword: str, minor_keyword: str):
    """
    header_row_1(대제목 행)과 header_row_2(세부 컬럼명 행)를 대조하여
    major_keyword(예: '건강보험')이면서 minor_keyword(예: '당월보험료')인
    컬럼의 0-based 인덱스를 반환한다. 못 찾으면 None.
    """
    for idx, (major, minor) in enumerate(zip(header_row_1, header_row_2)):
        major_text = str(major).replace(" ", "") if major else ""
        minor_text = str(minor).replace(" ", "") if minor else ""
        if major_keyword in major_text and minor_keyword in minor_text:
            return idx
    return None


# ---------------------------------------------------------------------
# 2. 개인부담금.xlsx
# ---------------------------------------------------------------------
def load_personal_burden(filepath: str) -> pd.DataFrame:
    """
    개인부담금 파일: 3행=보험 대제목(병합), 4행=세부컬럼명, 5행부터 데이터
    반환 DataFrame 컬럼:
      구분, 예산구분, 채용구분, 소속, 성명, 주민등록번호,
      건강보험_당월보험료, 노인장기요양_당월보험료, 국민연금_당월보험료
    """
    grid = _unmerge_and_get_grid(filepath)

    header_major = grid[2]  # 엑셀 3행 (0-index 2)
    header_minor = grid[3]  # 엑셀 4행 (0-index 3)

    idx_health = _find_column_index(header_major, header_minor, "건강보험", "당월보험료")
    idx_ltc = _find_column_index(header_major, header_minor, "노인장기요양", "당월보험료")
    idx_pension = _find_column_index(header_major, header_minor, "국민연금", "당월보험료")

    if idx_health is None or idx_ltc is None or idx_pension is None:
        raise ValueError("개인부담금.xlsx 에서 '당월보험료' 컬럼 위치를 찾지 못했습니다.")

    # 고정 위치 컬럼 (스펙 기준: E~J열 = 0-based 4~9)
    IDX_GUBUN = 4       # E열: 구분
    IDX_BUDGET = 5      # F열: 예산구분
    IDX_HIRE_TYPE = 6   # G열: 채용구분
    IDX_DEPT = 7        # H열: 소속
    IDX_NAME = 8        # I열: 성명
    IDX_RRN = 9          # J열: 주민등록번호

    rows = []
    for row in grid[4:]:  # 5행부터
        if row is None:
            continue
        name = row[IDX_NAME] if len(row) > IDX_NAME else None
        if name is None or str(name).strip() == "":
            continue
        if _is_summary_row(name):
            continue

        rrn = row[IDX_RRN] if len(row) > IDX_RRN else None
        birth6 = str(rrn)[:6] if rrn else None

        rows.append({
            "구분": row[IDX_GUBUN] if len(row) > IDX_GUBUN else None,
            "예산구분": row[IDX_BUDGET] if len(row) > IDX_BUDGET else None,
            "채용구분": row[IDX_HIRE_TYPE] if len(row) > IDX_HIRE_TYPE else None,
            "소속": row[IDX_DEPT] if len(row) > IDX_DEPT else None,
            "성명": name,
            "생년월일6": birth6,
            "건강보험_당월보험료": row[idx_health] or 0,
            "노인장기요양_당월보험료": row[idx_ltc] or 0,
            "국민연금_당월보험료": row[idx_pension] or 0,
        })

    return pd.DataFrame(rows)


# ---------------------------------------------------------------------
# 3. 자치단체부담금.xlsx
# ---------------------------------------------------------------------
def load_local_burden(filepath: str) -> pd.DataFrame:
    """
    자치단체부담금 파일: 4행=보험 대제목(병합), 5행=세부컬럼명, 6행=합계행(제외), 7행부터 데이터
    반환 DataFrame 컬럼:
      예산구분, 채용구분, 소속, 성명, 생년월일6,
      건강보험_당월보험료, 노인장기요양_당월보험료, 국민연금_당월보험료
    """
    grid = _unmerge_and_get_grid(filepath)

    header_major = grid[3]  # 엑셀 4행 (0-index 3)
    header_minor = grid[4]  # 엑셀 5행 (0-index 4)

    idx_health = _find_column_index(header_major, header_minor, "건강보험", "당월보험료")
    idx_ltc = _find_column_index(header_major, header_minor, "노인장기요양", "당월보험료")
    idx_pension = _find_column_index(header_major, header_minor, "국민연금", "당월보험료")

    if idx_health is None or idx_ltc is None or idx_pension is None:
        raise ValueError("자치단체부담금.xlsx 에서 '당월보험료' 컬럼 위치를 찾지 못했습니다.")

    # 고정 위치 컬럼 (스펙 기준: B~F열 = 0-based 1~5)
    IDX_BUDGET = 1       # B열: 예산구분
    IDX_HIRE_TYPE = 2    # C열: 채용구분
    IDX_DEPT = 3         # D열: 소속
    IDX_NAME = 4         # E열: 성명
    IDX_RRN = 5          # F열: 주민등록번호

    rows = []
    # 6행(0-index 5)은 합계행이므로 7행(0-index 6)부터 시작
    for row in grid[6:]:
        if row is None:
            continue
        name = row[IDX_NAME] if len(row) > IDX_NAME else None
        if name is None or str(name).strip() == "":
            continue
        if _is_summary_row(name):
            continue

        rrn = row[IDX_RRN] if len(row) > IDX_RRN else None
        birth6 = str(rrn)[:6] if rrn else None

        rows.append({
            "예산구분": row[IDX_BUDGET] if len(row) > IDX_BUDGET else None,
            "채용구분": row[IDX_HIRE_TYPE] if len(row) > IDX_HIRE_TYPE else None,
            "소속": row[IDX_DEPT] if len(row) > IDX_DEPT else None,
            "성명": name,
            "생년월일6": birth6,
            "건강보험_당월보험료": row[idx_health] or 0,
            "노인장기요양_당월보험료": row[idx_ltc] or 0,
            "국민연금_당월보험료": row[idx_pension] or 0,
        })

    return pd.DataFrame(rows)
