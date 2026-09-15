"""
validator.py
============
필수값 누락 검증 + 예산구분(시비/국도비) 판단
"""

REQUIRED_FIELDS = ["성명", "생년월일", "채용일자", "계약만료일", "소정근로시간"]


def validate_required_fields(row: dict, messages: list) -> bool:
    """
    운영현황 한 사람 분의 row(dict)에서 필수 항목 누락 여부를 검사한다.
    누락이 있으면 messages에 메시지를 추가하고 False 반환 (해당 인원 처리 보류).
    """
    name = row.get("성명", "이름미상")
    ok = True
    for field in REQUIRED_FIELDS:
        value = row.get(field)
        if value is None or str(value).strip() == "":
            msg = f"[{name}] 근로자의 {field} 확인이 필요합니다"
            messages.append(msg)
            print(msg)
            ok = False
    return ok


def determine_budget_type(budget_text, name: str, messages: list):
    """
    개인부담금 파일 F열(예산구분) 텍스트를 보고 '시비' 또는 '국도비'를 판단한다.
    판단 불가하면 None을 반환하고 messages에 경고를 남긴다.
    """
    if budget_text is None:
        text = ""
    else:
        text = str(budget_text)

    if ("일반" in text) or ("시비" in text):
        return "시비"
    if ("특별" in text) or ("국비" in text) or ("도비" in text):
        return "국도비"

    msg = f"[{name}] 근로자의 예산구분 확인이 필요합니다"
    messages.append(msg)
    print(msg)
    return None


def check_budget_type_mismatch(name: str, personal_budget_type, local_budget_text, messages: list):
    """
    자치단체부담금 파일의 예산구분과 개인부담금 파일의 예산구분이 불일치하면
    경고 메시지를 남긴다. (실제 처리는 항상 개인부담금 F열 기준)
    """
    if local_budget_text is None:
        return

    local_text = str(local_budget_text)
    if personal_budget_type == "시비" and (("특별" in local_text) or ("국비" in local_text) or ("도비" in local_text)):
        msg = f"[{name}] 개인부담금·자치단체부담금 파일 간 예산구분 불일치 - 개인부담금 기준으로 처리합니다"
        messages.append(msg)
        print(msg)
    if personal_budget_type == "국도비" and (("일반" in local_text) or ("시비" in local_text)):
        msg = f"[{name}] 개인부담금·자치단체부담금 파일 간 예산구분 불일치 - 개인부담금 기준으로 처리합니다"
        messages.append(msg)
        print(msg)
