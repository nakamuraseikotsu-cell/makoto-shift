# -*- coding: utf-8 -*-
"""前処理：列名標準化・型変換・欠損確認"""
from __future__ import annotations
import pandas as pd

# 想定する標準列 → 候補となる実列名
CANDIDATE_COLS = {
    'date':          ['date', '日付', 'Date'],
    'area':          ['area', '院', '院名', 'エリア', '店舗'],
    'day_of_week':   ['day_of_week', '曜日'],
    'weather':       ['天気', 'weather'],
    'temperature':   ['最高気温', '気温', 'temperature'],
    'patients':      ['total_patients', '来院数', '総患', '総 患', '患者数'],
    'new_patients':  ['new_patients', '新患', '新 患', '新規患者'],
    'staff_first':   ['shift_first_half', '前半人数', '前半'],
    'staff_second':  ['shift_second_half', '後半人数', '後半'],
    'staff_optimal': ['shift_optimal', '適正人数', 'スタッフ数', '稼働スタッフ数'],
}

# 曜日名（英→日）
WD_EN_JA = {'Mon': '月', 'Tue': '火', 'Wed': '水',
            'Thu': '木', 'Fri': '金', 'Sat': '土', 'Sun': '日'}


class Preprocessor:
    def __init__(self, assumptions: list[str]):
        self.assumptions = assumptions
        self.col_map: dict[str, str] = {}

    def find_columns(self, df: pd.DataFrame) -> dict[str, str]:
        """標準列→実列名のマッピングを推定"""
        result = {}
        for std, cands in CANDIDATE_COLS.items():
            for c in cands:
                if c in df.columns:
                    result[std] = c
                    break
        return result

    def run(self, df: pd.DataFrame) -> pd.DataFrame:
        self.col_map = self.find_columns(df)
        print('[Preprocessor] 列マッピング:')
        for std in CANDIDATE_COLS:
            mapped = self.col_map.get(std)
            print(f'   {std:14s} ← {mapped}')

        missing = [k for k in CANDIDATE_COLS if k not in self.col_map]
        if missing:
            self.assumptions.append(
                f'前処理: 該当列が無い項目あり {missing}（後段でフォールバック）'
            )

        # rename: 実列名 → 標準名
        rename_map = {v: k for k, v in self.col_map.items() if v != k}
        df = df.rename(columns=rename_map)

        if 'date' in df.columns:
            df['date'] = pd.to_datetime(df['date'], errors='coerce')

        if 'day_of_week' not in df.columns and 'date' in df.columns:
            df['day_of_week'] = df['date'].dt.strftime('%a').map(WD_EN_JA)

        # 数値化
        for c in ['patients', 'new_patients', 'staff_first',
                  'staff_second', 'staff_optimal', 'temperature']:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors='coerce')

        # 欠損サマリ
        n_total = len(df)
        miss_summary = {}
        for c in ['patients', 'staff_optimal', 'weather']:
            if c in df.columns:
                miss = df[c].isna().sum()
                miss_summary[c] = f'{miss}/{n_total} ({miss/n_total*100:.1f}%)'
        print(f'[Preprocessor] 欠損率: {miss_summary}')

        return df
