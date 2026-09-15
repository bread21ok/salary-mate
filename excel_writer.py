"""
excel_writer.py
================
계산된 결과를 급여내역서_양식.xlsx 의 '급여내역서' 시트에 채워 넣는다.
- SUM 함수 셀(M·T·U·AA열)은 절대 덮어쓰지 않는다.
- 급여내역서 시트 외 다른 시트는 절대 건드리지 않는다.
- 데이터는 9행부터 시작하며, 양식은 100명분(9~108행)까지 만들어져 있다고 가정한다.
- 실제 인원보다 남는 빈 행은 뒤에서부터 앞으로 삭제한다.
"""

import io
import openpyxl

SHEET_NAME = "급여내역서"
DATA_START_ROW = 9
MAX_ROWS = 100  # 양식에 미리 만들어진 최대 인원

# 컬럼 문자 매핑 (SUM 셀인 M, T, U, AA 는 제외하고 값만 입력)
COLUMN_MAP = {
    "번호": "A",
    "성명": "B",
    "소속": "C",
    "생년월일": "D",
    "기본급": "E",
    "정액급식비": "F",
    "정기상여금": "G",
    "생활임금보전수당": "H",
    "연차수당": "I",
    "가족수당": "J",
    "시간외근무수당": "K",
    "기타수당": "L",
    # M열(보수내역합계)은 SUM 셀 - 건드리지 않음
    "건강보험_개인": "N",
    "노인장기요양_개인": "O",
    "국민연금_개인": "P",
    "고용보험_개인": "Q",
    "소득세": "R",
    "지방소득세": "S",
    # T열(공제내역합계), U열(실수령액)은 SUM 셀 - 건드리지 않음
    "건강보험_자치단체": "V",
    "노인장기요양_자치단체": "W",
    "국민연금_자치단체": "X",
    "고용보험_자치단체": "Y",
    "산재보험_자치단체": "Z",
    # AA열(합계)은 SUM 셀 - 건드리지 않음
}

SUM_COLUMNS = {"M", "T", "U", "AA"}


def write_payroll(template_filepath: str, records: list) -> io.BytesIO:
    """
    template_filepath : 원본 급여내역서_양식.xlsx 경로
    records            : 사람별 계산 결과 dict 리스트 (COLUMN_MAP의 키를 사용)
    반환                : 완성된 엑셀 파일을 담은 BytesIO
    """
    wb = openpyxl.load_workbook(template_filepath, data_only=False)  # 함수 유지
    ws = wb[SHEET_NAME]

    # 1) 값 입력 (급여내역서 시트만)
    for i, record in enumerate(records):
        row_num = DATA_START_ROW + i
        if row_num > DATA_START_ROW + MAX_ROWS - 1:
            break  # 양식에 준비된 최대 행수를 넘으면 더 이상 쓰지 않음(별도 처리 필요)

        for field, col_letter in COLUMN_MAP.items():
            if col_letter in SUM_COLUMNS:
                continue  # SUM 함수 셀은 절대 덮어쓰지 않음
            value = record.get(field)
            ws[f"{col_letter}{row_num}"] = value

    # 2) 빈 행 삭제: 실제 데이터가 없는 나머지 행을 뒤에서부터 앞으로 삭제
    used_rows = DATA_START_ROW + len(records)
    last_template_row = DATA_START_ROW + MAX_ROWS  # 배타적 경계 (100명분 다음 행)

    for row_num in range(last_template_row - 1, used_rows - 1, -1):
        ws.delete_rows(row_num, 1)

    # 3) BytesIO로 저장 (다른 시트는 위 과정에서 전혀 건드리지 않았으므로 그대로 유지됨)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
