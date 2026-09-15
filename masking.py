"""
masking.py
==========
성명, 생년월일, 주소, 연락처, 계좌번호를 가역적 토큰으로 치환하고,
필요할 때 다시 원래 값으로 복원한다.
매핑 테이블은 메모리(딕셔너리)에만 보관하며 파일로 저장하지 않는다.
"""

import pandas as pd

MASK_COLUMNS = ["성명", "생년월일", "주소", "연락처", "계좌번호"]


def mask_dataframe(df: pd.DataFrame, columns=None):
    """
    df의 지정 컬럼들을 토큰(예: 성명__TOKEN_0001)으로 치환한다.
    반환값: (마스킹된 df 복사본, 복원용 매핑 딕셔너리)
      매핑 딕셔너리 구조: {"성명__TOKEN_0001": "홍길동", ...}
    """
    if columns is None:
        columns = [c for c in MASK_COLUMNS if c in df.columns]

    masked_df = df.copy().reset_index(drop=True)
    mapping = {}

    for col in columns:
        if col not in masked_df.columns:
            continue
        # 날짜/숫자 등 다른 dtype 컬럼에 문자열 토큰을 넣을 수 있도록 object 타입으로 변환
        masked_df[col] = masked_df[col].astype(object)
        for i, original_value in masked_df[col].items():
            token = f"{col}__TOKEN_{i:04d}"
            mapping[token] = original_value
            masked_df.at[i, col] = token

    return masked_df, mapping


def unmask_value(token, mapping: dict):
    """토큰 하나를 원래 값으로 복원. 매핑에 없으면 원본 그대로 반환."""
    return mapping.get(token, token)


def unmask_dataframe(df: pd.DataFrame, mapping: dict, columns=None):
    """df 전체(지정 컬럼)를 원래 값으로 복원한다."""
    if columns is None:
        columns = [c for c in MASK_COLUMNS if c in df.columns]

    restored_df = df.copy()
    for col in columns:
        if col not in restored_df.columns:
            continue
        restored_df[col] = restored_df[col].apply(lambda v: unmask_value(v, mapping))
    return restored_df
