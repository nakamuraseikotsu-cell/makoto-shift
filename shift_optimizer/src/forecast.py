# -*- coding: utf-8 -*-
"""生産性算出 & 来院予測（階層フォールバック + 祝日補正）"""
from __future__ import annotations
from datetime import datetime, timedelta
import pandas as pd
import numpy as np


# 2026 年の日本の祝日（ハードコード）
HOLIDAYS_2026 = {
    '2026-01-01', '2026-01-12', '2026-02-11', '2026-02-23',
    '2026-03-20', '2026-04-29', '2026-05-03', '2026-05-04',
    '2026-05-05', '2026-05-06', '2026-07-20', '2026-08-11',
    '2026-09-21', '2026-09-22', '2026-09-23', '2026-10-12',
    '2026-11-03', '2026-11-23',
}

WD_JP = {0: '月', 1: '火', 2: '水', 3: '木', 4: '金', 5: '土', 6: '日'}


class Forecaster:
    """過去実績から (院×曜日×天気) の生産性／平均来院数を集計し、対象月を予測"""

    def __init__(self, df: pd.DataFrame, target_month: str, assumptions: list[str]):
        self.df = df
        self.target_month = target_month
        self.assumptions = assumptions
        self.productivity: dict | None = None
        self.visits_avg: dict | None = None

    # ----------------------------- 生産性 -----------------------------
    def compute_productivity(self) -> dict:
        df = self.df.copy()
        # patients / staff_optimal 両方ある行のみ
        m = (df['patients'].notna() & df['staff_optimal'].notna()
             & (df['staff_optimal'] > 0))
        valid = df[m].copy()
        valid['productivity'] = valid['patients'] / valid['staff_optimal']

        prod = {}
        # 1) area × dow × weather
        if 'weather' in valid.columns:
            prod['area_dow_weather'] = (
                valid.dropna(subset=['weather'])
                     .groupby(['area', 'day_of_week', 'weather'])['productivity']
                     .agg(['mean', 'count']).reset_index()
                     .rename(columns={'mean': 'productivity'})
            )
        else:
            prod['area_dow_weather'] = pd.DataFrame(
                columns=['area', 'day_of_week', 'weather', 'productivity', 'count']
            )
        # 2) area × dow
        prod['area_dow'] = (
            valid.groupby(['area', 'day_of_week'])['productivity']
                 .agg(['mean', 'count']).reset_index()
                 .rename(columns={'mean': 'productivity'})
        )
        # 3) area
        prod['area'] = (
            valid.groupby(['area'])['productivity']
                 .agg(['mean', 'count']).reset_index()
                 .rename(columns={'mean': 'productivity'})
        )
        # 4) global
        prod['global_value'] = (
            float(valid['productivity'].mean()) if len(valid) else 6.0
        )

        # 同様に来院数の平均
        va = {}
        if 'weather' in valid.columns:
            va['area_dow_weather'] = (
                valid.dropna(subset=['weather'])
                     .groupby(['area', 'day_of_week', 'weather'])['patients']
                     .agg(['mean', 'count']).reset_index()
            )
        else:
            va['area_dow_weather'] = pd.DataFrame(
                columns=['area', 'day_of_week', 'weather', 'mean', 'count']
            )
        va['area_dow'] = (
            valid.groupby(['area', 'day_of_week'])['patients']
                 .agg(['mean', 'count']).reset_index()
        )
        va['area'] = (
            valid.groupby(['area'])['patients']
                 .agg(['mean', 'count']).reset_index()
        )
        va['global_value'] = float(valid['patients'].mean()) if len(valid) else 20.0

        self.productivity = prod
        self.visits_avg = va
        print(f'[Forecaster] 生産性集計: area×dow×weather={len(prod["area_dow_weather"])}行, '
              f'area×dow={len(prod["area_dow"])}行, area={len(prod["area"])}行, '
              f'global={prod["global_value"]:.2f}')
        return prod

    # --------------------------- フォールバック ---------------------------
    @staticmethod
    def _reliability(n: int) -> str:
        if n >= 20:
            return 'high'
        if n >= 10:
            return 'medium'
        return 'low'

    def _lookup(self, store: dict, area: str, dow: str, weather: str,
                value_col: str) -> tuple[float, str, str]:
        """階層フォールバックで値を取得"""
        # 1) area × dow × weather
        df_lv = store['area_dow_weather']
        if len(df_lv):
            row = df_lv[(df_lv['area'] == area) & (df_lv['day_of_week'] == dow)
                        & (df_lv['weather'] == weather)]
            if len(row) and row.iloc[0]['count'] >= 3:
                v = row.iloc[0][value_col]
                if not pd.isna(v):
                    return float(v), self._reliability(int(row.iloc[0]['count'])), 'area_dow_weather'
        # 2) area × dow
        df_lv = store['area_dow']
        row = df_lv[(df_lv['area'] == area) & (df_lv['day_of_week'] == dow)]
        if len(row) and row.iloc[0]['count'] >= 3:
            v = row.iloc[0][value_col]
            if not pd.isna(v):
                return float(v), self._reliability(int(row.iloc[0]['count'])), 'area_dow'
        # 3) area
        df_lv = store['area']
        row = df_lv[df_lv['area'] == area]
        if len(row):
            v = row.iloc[0][value_col]
            if not pd.isna(v):
                return float(v), self._reliability(int(row.iloc[0]['count'])), 'area'
        # 4) global
        return float(store['global_value']), 'low', 'global'

    # ----------------------------- 予測 -----------------------------
    def predict_visits(self, weather_forecast: pd.DataFrame,
                       clinic_master: pd.DataFrame) -> pd.DataFrame:
        year, month = map(int, self.target_month.split('-'))
        next_m = datetime(year + 1, 1, 1) if month == 12 else datetime(year, month + 1, 1)
        start = datetime(year, month, 1)
        days = []
        d = start
        while d < next_m:
            days.append(d); d += timedelta(days=1)

        wf = weather_forecast.copy()
        wf['日付'] = pd.to_datetime(wf['日付']).dt.strftime('%Y-%m-%d')
        weather_lookup = wf.set_index(['日付', 'エリア']).to_dict('index')

        c2warea = clinic_master.set_index('院名')['天気エリア'].to_dict()

        records = []
        for d in days:
            ds = d.strftime('%Y-%m-%d')
            dow = WD_JP[d.weekday()]
            is_holiday = ds in HOLIDAYS_2026
            for _, c in clinic_master.iterrows():
                clinic = c['院名']
                warea = c2warea.get(clinic, '小金井')
                wkey = (ds, warea)
                if wkey in weather_lookup:
                    weather = weather_lookup[wkey]['天気']
                    temp = weather_lookup[wkey]['最高気温']
                else:
                    weather = '晴'
                    temp = np.nan
                visits, reliability, level = self._lookup(
                    self.visits_avg, clinic, dow, weather, 'mean'
                )
                if is_holiday:
                    visits = visits * 0.7
                records.append({
                    'date': ds, 'area': clinic, 'day_of_week': dow,
                    'is_holiday': is_holiday,
                    'weather_forecast': weather,
                    'temperature_forecast': temp,
                    'predicted_visits': round(visits, 1),
                    'forecast_reliability': reliability,
                    'forecast_level': level,
                })
        df = pd.DataFrame(records)
        self.assumptions.append(
            '来院予測: 階層フォールバック (area×曜日×天気 → area×曜日 → area → 全院平均)'
        )
        self.assumptions.append(
            f'祝日要因: HOLIDAYS_2026 をハードコード、祝日は通常日 × 0.7'
        )
        print(f'[Forecaster] 来院予測完了: {len(df)} 行 '
              f'(対象月 {self.target_month}, {len(days)}日 × {len(clinic_master)}院)')
        return df
