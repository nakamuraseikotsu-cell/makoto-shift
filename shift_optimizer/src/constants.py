# -*- coding: utf-8 -*-
"""シフト最適化で参照する固定マスタ定数"""
from __future__ import annotations

import pandas as pd

# 動かせないスタッフ（他院ヘルプ不可・自院に固定）
# 院名 → そのスタッフの姓（または短縮名）のリスト
# xlsx 上のスタッフ名表記が「熊谷さん」「熊谷 健太」などでも、
# 部分一致（surname in cell_name）で識別する。
IMMOVABLE_STAFF: dict[str, list[str]] = {
    "国分寺":    ["熊谷", "黒澤"],
    "武蔵小金井": ["田村", "岡田"],
    "東小金井":  ["伊藤"],
    "小金井坂下": ["山本"],
    "人形町":    ["長島"],
}


# 退職・除外スタッフ（システム全体から完全に除外）
# 部分一致：氏名に下記いずれかの文字列を含むレコードを全てドロップする
EXCLUDED_STAFF: list[str] = [
    "池田",
]


def is_immovable(staff_name: str, home_clinic: str) -> bool:
    """スタッフ名 + 自院 が動かせないスタッフかどうかを判定（部分一致）"""
    name = (staff_name or "").strip()
    if not name:
        return False
    for surname in IMMOVABLE_STAFF.get(home_clinic, []):
        if surname in name:
            return True
    return False


def all_immovable_pairs() -> set[tuple[str, str]]:
    """{ (院名, 姓) } の集合で返す（表示用などに）"""
    return {(c, n) for c, names in IMMOVABLE_STAFF.items() for n in names}


def is_excluded_staff_name(name) -> bool:
    """名前が EXCLUDED_STAFF のいずれかを含むかどうか（部分一致）"""
    if name is None:
        return False
    try:
        s = str(name)
    except Exception:
        return False
    for excl in EXCLUDED_STAFF:
        if excl and excl in s:
            return True
    return False


def drop_rows_with_excluded_staff(
    df: pd.DataFrame,
    staff_columns: list[str] | None = None,
) -> pd.DataFrame:
    """DataFrame から、スタッフ名列に EXCLUDED_STAFF を含む行を削除して返す。

    staff_columns: 明示指定があればその列のみをチェック。
                   未指定なら、object dtype の全列をスキャンして判定。
    """
    if df is None or len(df) == 0 or not EXCLUDED_STAFF:
        return df

    if staff_columns is None:
        # object dtype の列を自動検出
        check_cols = [c for c in df.columns if df[c].dtype == 'object']
    else:
        check_cols = [c for c in staff_columns if c in df.columns]

    if not check_cols:
        return df

    mask = pd.Series(False, index=df.index)
    for col in check_cols:
        s = df[col].astype(str)
        for excl in EXCLUDED_STAFF:
            mask = mask | s.str.contains(excl, na=False, regex=False)

    return df.loc[~mask].reset_index(drop=True)
