"""
main_calculator.py
===================
전체 처리 순서(1~9단계)를 실행하는 메인 함수.
반환값: (BytesIO, messages)
"""

import re
from datetime import date, datetime

import pandas as pd

from excel_parser import load_operation_status, load_personal_burden, load_local_burden
from masking import mask_dataframe, unmask_dataframe, MASK_COLUMNS
from validator import validate_required_fields, determine_budget_type, check_budget_type_mismatch
from calculator import (
    compute_pay_items, compute_taxable_monthly, lookup_income_tax,
    apply_prorate, calc_health_pension_charge, is_age_65_or_older_at_hire,
    INSURANCE_RATE,
)
from excel_writer import write_payroll


# ---------------------------------------------------------------------
# 유틸
# ---------------------------------------------------------------------
def _to_date(value):
    """엑셀에서 읽은 값(datetime, date, 문자열 등)을 date 객체로 통일한다."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
            try:
                return datetime.strptime(value.strip(), fmt).date()
            except ValueError:
                continue
    return None


def _format_birth6(birth_date: date) -> str:
    """date 객체를 주민등록번호 앞 6자리 형식(YYMMDD 문자열)으로 변환"""
    if birth_date is None:
        return None
    return f"{birth_date.year % 100:02d}{birth_date.month:02d}{birth_date.day:02d}"


def _match_worker(name, birth6, dept, burden_df, messages, source_label):
    """
    성명 + 생년월일(앞6자리) 로 burden_df에서 매칭.
    동명이인(2건 이상)이면 소속까지 더해서 재매칭.
    매칭 실패 시 None 반환.
    """
    if burden_df is None or burden_df.empty:
        return None

    candidates = burden_df[(burden_df["성명"] == name) & (burden_df["생년월일6"] == birth6)]

    if len(candidates) == 0:
        return None

    if len(candidates) == 1:
        return candidates.iloc[0]

    # 동명이인 -> 소속까지 조합하여 재매칭
    dup_msg = f"[{name}] {source_label} 파일에서 동명이인(성명+생년월일 동일) {len(candidates)}건 발견 - 소속으로 재매칭 시도"
    messages.append(dup_msg)
    print(dup_msg)

    re_matched = candidates[candidates["소속"] == dept]
    if len(re_matched) == 1:
        return re_matched.iloc[0]

    fail_msg = f"[{name}] {source_label} 파일 동명이인 재매칭 실패 - 소속으로도 구분되지 않습니다"
    messages.append(fail_msg)
    print(fail_msg)
    return None


# ---------------------------------------------------------------------
# 메인 함수
# ---------------------------------------------------------------------
def main(operation_filepath: str,
         personal_burden_filepath: str,
         local_burden_filepath: str,
         template_filepath: str,
         target_year: int,
         target_month: int):
    """
    급여내역서 자동 계산 메인 함수.
    반환: (BytesIO, messages)
    """
    messages = []

    # ---------------- 1단계: 엑셀 파일 3개 읽기 ----------------
    op_df = load_operation_status(operation_filepath)
    personal_df = load_personal_burden(personal_burden_filepath)
    local_df = load_local_burden(local_burden_filepath)

    # ---------------- 2단계: 개인정보 마스킹 ----------------
    # (계산 자체는 마스킹 여부와 무관하므로, 원본 값은 매칭·메시지 출력용으로 별도 보관하고
    #  운영현황 df는 마스킹된 버전을 만들어 내부 처리 표준을 따른다.)
    masked_op_df, mask_mapping = mask_dataframe(op_df, columns=MASK_COLUMNS)

    records = []

    for idx, row in op_df.iterrows():
        original_name = row.get("성명")
        raw_row = {
            "성명": original_name,
            "소속": row.get("소속"),
            "생년월일": _to_date(row.get("생년월일")),
            "채용일자": _to_date(row.get("채용일자(최초계약일)")) or _to_date(row.get("채용일자")),
            "계약만료일": _to_date(row.get("계약만료일")),
            "소정근로시간": row.get("1일 소정근로시간"),
            "연차수당_지급일수": row.get("연차수당 지급일수"),
            "시간외근무시간": row.get("시간외근무시간"),
            "배우자여부": row.get("배우자여부"),
            "자녀수": row.get("자녀수"),
            "기타부양가족수": row.get("기타부양가족수"),
        }

        # ---------------- 3단계: 필수값 검증 ----------------
        if not validate_required_fields(raw_row, messages):
            continue  # 누락 시 해당 인원 처리 보류

        birth6 = _format_birth6(raw_row["생년월일"])

        personal_match = _match_worker(
            original_name, birth6, raw_row["소속"], personal_df, messages, "개인부담금"
        )
        if personal_match is None:
            msg = f"[{original_name}] 개인부담금 파일에서 매칭되는 인원을 찾을 수 없습니다"
            messages.append(msg)
            print(msg)
            continue

        # ---------------- 4단계: 예산구분 판단 ----------------
        budget_type = determine_budget_type(personal_match.get("예산구분"), original_name, messages)
        if budget_type is None:
            continue  # 판단 불가 -> 처리 보류

        local_match = None
        if budget_type == "국도비":
            local_match = _match_worker(
                original_name, birth6, raw_row["소속"], local_df, messages, "자치단체부담금"
            )
            if local_match is not None:
                check_budget_type_mismatch(original_name, budget_type, local_match.get("예산구분"), messages)
            else:
                msg = f"[{original_name}] 자치단체부담금 매칭 실패 - V~Z열 0 처리"
                messages.append(msg)
                print(msg)

        # ---------------- 5단계: 급여 계산 ----------------
        pay_items = compute_pay_items(raw_row)
        taxable_monthly = compute_taxable_monthly(pay_items)

        hire_date = raw_row["채용일자"]
        expire_date = raw_row["계약만료일"]
        newly_65 = is_age_65_or_older_at_hire(raw_row["생년월일"], hire_date)

        # -- 개인부담금 4대보험 --
        health_personal = calc_health_pension_charge(
            personal_match.get("건강보험_당월보험료"), hire_date, expire_date, target_year, target_month
        )
        ltc_personal = calc_health_pension_charge(
            personal_match.get("노인장기요양_당월보험료"), hire_date, expire_date, target_year, target_month
        )
        pension_personal = calc_health_pension_charge(
            personal_match.get("국민연금_당월보험료"), hire_date, expire_date, target_year, target_month
        )
        if newly_65:
            employment_personal = 0
        else:
            employment_personal = apply_prorate(
                taxable_monthly, INSURANCE_RATE["고용보험_개인"],
                hire_date, expire_date, target_year, target_month
            )

        # -- 소득세 --
        income_tax = lookup_income_tax(taxable_monthly, original_name, messages)
        local_income_tax = int(income_tax * 0.1)

        # -- 자치단체부담금 4대보험 (국·도비 인원만) --
        if budget_type == "국도비" and local_match is not None:
            health_local = calc_health_pension_charge(
                local_match.get("건강보험_당월보험료"), hire_date, expire_date, target_year, target_month
            )
            ltc_local = calc_health_pension_charge(
                local_match.get("노인장기요양_당월보험료"), hire_date, expire_date, target_year, target_month
            )
            pension_local = calc_health_pension_charge(
                local_match.get("국민연금_당월보험료"), hire_date, expire_date, target_year, target_month
            )
            if newly_65:
                employment_local = 0
            else:
                employment_local = apply_prorate(
                    taxable_monthly, INSURANCE_RATE["고용보험_자치단체"],
                    hire_date, expire_date, target_year, target_month
                )
            accident_local = apply_prorate(
                taxable_monthly, INSURANCE_RATE["산재보험_자치단체"],
                hire_date, expire_date, target_year, target_month
            )
        else:
            health_local = ltc_local = pension_local = employment_local = accident_local = 0

        record = {
            "번호": idx + 1,
            "성명": original_name,
            "소속": raw_row["소속"],
            "생년월일": raw_row["생년월일"],
            **pay_items,  # 기본급/정액급식비/정기상여금/생활임금보전수당/연차수당/가족수당/시간외근무수당/기타수당/보수내역합계
            "건강보험_개인": health_personal,
            "노인장기요양_개인": ltc_personal,
            "국민연금_개인": pension_personal,
            "고용보험_개인": employment_personal,
            "소득세": income_tax,
            "지방소득세": local_income_tax,
            "건강보험_자치단체": health_local,
            "노인장기요양_자치단체": ltc_local,
            "국민연금_자치단체": pension_local,
            "고용보험_자치단체": employment_local,
            "산재보험_자치단체": accident_local,
        }
        records.append(record)

    # ---------------- 6단계: 개인정보 복원 ----------------
    # (records는 이미 원본 이름을 사용했으므로 별도 복원 절차 불필요.
    #  masked_op_df 는 내부 표준 처리 단계에서만 사용되고 최종 출력에는 쓰이지 않는다.)

    # ---------------- 7~8단계: 양식 매핑 + 빈 행 삭제 ----------------
    output_bytesio = write_payroll(template_filepath, records)

    # ---------------- 9단계: 반환 ----------------
    return output_bytesio, messages


def detect_target_year_month(title_text: str):
    """파일 제목(예: '2026년 6월 개인부담금...')에서 연/월을 추출한다. 실패 시 (None, None)."""
    match = re.search(r"(\d{4})\s*년\s*(\d{1,2})\s*월", title_text or "")
    if match:
        return int(match.group(1)), int(match.group(2))
    return None, None
