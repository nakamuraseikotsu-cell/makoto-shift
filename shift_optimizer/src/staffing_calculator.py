# -*- coding: utf-8 -*-
"""必要人員 & 過不足計算"""
from __future__ import annotations
import math
import pandas as pd


class StaffingCalculator:
    """予測来院数から必要人員数を算出し、休み希望を反映して過不足を計算"""

    def __init__(self, min_staff: int = 2, assumptions: list[str] | None = None):
        self.min_staff = min_staff
        self.assumptions = assumptions if assumptions is not None else []

    # --------------------------- 必要人員 ---------------------------
    def calculate_required(self, forecast_df: pd.DataFrame,
                            productivity: dict) -> pd.DataFrame:
        prod_adw = productivity['area_dow_weather']
        prod_ad = productivity['area_dow'].set_index(
            ['area', 'day_of_week'])['productivity'].to_dict()
        prod_a = productivity['area'].set_index('area')['productivity'].to_dict()
        global_prod = float(productivity['global_value'])

        # area×dow×weather のルックアップ辞書
        prod_adw_idx = {}
        if len(prod_adw):
            for _, r in prod_adw.iterrows():
                if r['count'] >= 3 and not pd.isna(r['productivity']):
                    prod_adw_idx[(r['area'], r['day_of_week'], r['weather'])] = r['productivity']

        rows = []
        for _, r in forecast_df.iterrows():
            key_adw = (r['area'], r['day_of_week'], r['weather_forecast'])
            key_ad = (r['area'], r['day_of_week'])
            if key_adw in prod_adw_idx:
                p, p_level = prod_adw_idx[key_adw], 'area_dow_weather'
            elif key_ad in prod_ad and not pd.isna(prod_ad[key_ad]):
                p, p_level = prod_ad[key_ad], 'area_dow'
            elif r['area'] in prod_a and not pd.isna(prod_a[r['area']]):
                p, p_level = prod_a[r['area']], 'area'
            else:
                p, p_level = global_prod, 'global'
            if p is None or pd.isna(p) or p <= 0:
                p, p_level = global_prod, 'global'
            required = max(self.min_staff, math.ceil(r['predicted_visits'] / p))
            rows.append({
                **r.to_dict(),
                'productivity_used': round(float(p), 2),
                'productivity_level': p_level,
                'required_staff': int(required),
            })
        df = pd.DataFrame(rows)
        self.assumptions.append(
            f'必要人員 = ceil(予測来院数 ÷ 1人あたり患者数), 最低 {self.min_staff}人'
        )
        return df

    # ----------------------------- 過不足 -----------------------------
    def calculate_gap(self, required_df: pd.DataFrame,
                       planned_shifts: pd.DataFrame,
                       leave_requests: pd.DataFrame) -> pd.DataFrame:
        df = required_df.copy()
        ps = planned_shifts.copy()
        if 'planned_staff' not in ps.columns:
            raise KeyError('planned_shifts に planned_staff 列が必要です')
        df = df.merge(ps[['date', 'area', 'planned_staff']],
                      on=['date', 'area'], how='left')
        df['planned_staff'] = df['planned_staff'].fillna(
            df['required_staff']).astype(int)

        # 休み希望集計（重要度ごとも記録）
        lr = leave_requests.copy()
        lr['希望日'] = pd.to_datetime(lr['希望日']).dt.strftime('%Y-%m-%d')
        lr_grp = (
            lr.groupby(['希望日', '院'])
              .agg(leave_requested=('スタッフID', 'count'),
                   high_priority=('重要度', lambda s: (s == '高').sum()))
              .reset_index()
              .rename(columns={'希望日': 'date', '院': 'area'})
        )
        df = df.merge(lr_grp, on=['date', 'area'], how='left')
        df['leave_requested'] = df['leave_requested'].fillna(0).astype(int)
        df['high_priority'] = df['high_priority'].fillna(0).astype(int)

        df['available_staff'] = (df['planned_staff'] - df['leave_requested']).clip(lower=0)
        df['gap'] = df['available_staff'] - df['required_staff']
        df['status'] = df['gap'].apply(
            lambda x: '不足' if x < 0 else ('余剰' if x > 0 else '適正')
        )

        # 推定取りこぼし患者数 / 売上影響（不足時のみ正値）
        # 1患者あたり売上を 4,000円 と仮定
        UNIT_REVENUE = 4000
        df['estimated_missed_patients'] = df.apply(
            lambda r: max(0, -r['gap']) * r['productivity_used'], axis=1
        ).round(1)
        df['estimated_sales_impact'] = (
            df['estimated_missed_patients'] * UNIT_REVENUE
        ).round(0).astype(int)
        self.assumptions.append(
            '推定取りこぼし = 不足人数 × 1人あたり患者数, 売上影響 = 患者数 × ¥4,000 (仮)'
        )
        return df
