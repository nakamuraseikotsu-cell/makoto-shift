# -*- coding: utf-8 -*-
"""データ読込（実データ専用版）

* ダミーデータ自動生成は **完全に廃止** しました。
* 売上 xlsx と シフト xlsx から構築された ultimate_shift_master.csv を患者数の母集団とし、
  対象月の planned_shifts / 希望休 は 実シフト xlsx から直接抽出します。
* 対象月の実データが見つからない場合は RealDataMissingError を送出して処理を停止します。
* clinic_master は静的メタデータ（運用上の固定情報）として内蔵。
"""
from __future__ import annotations
from pathlib import Path
import pandas as pd

from .real_data import load_real_shift_data
from .constants import drop_rows_with_excluded_staff


class RealDataMissingError(Exception):
    """対象月の実データが存在しない場合に送出"""


# 静的な院マスタ（運用上の固定情報。実データの範疇）
_CLINIC_MASTER_ROWS = [
    {'院名': '国分寺',     'エリア': '中央線', '天気エリア': '小金井',
     '開店': '09:00', '閉店': '20:00', '標準スタッフ数': 4},
    {'院名': '小金井坂下', 'エリア': '中央線', '天気エリア': '小金井',
     '開店': '09:00', '閉店': '20:00', '標準スタッフ数': 3},
    {'院名': '東小金井',   'エリア': '中央線', '天気エリア': '小金井',
     '開店': '09:00', '閉店': '20:00', '標準スタッフ数': 3},
    {'院名': '武蔵小金井', 'エリア': '中央線', '天気エリア': '小金井',
     '開店': '09:00', '閉店': '20:00', '標準スタッフ数': 4},
    {'院名': '人形町',     'エリア': '都心',   '天気エリア': '人形町',
     '開店': '10:00', '閉店': '21:00', '標準スタッフ数': 3},
]


class DataLoader:
    """売上 xlsx → 過去患者集計（ultimate_shift_master.csv）
       シフト xlsx → 対象月の planned_shifts / 希望休"""

    def __init__(self, base_dir, master_csv, target_month, assumptions,
                 real_data_folder=None):
        self.base = Path(base_dir)
        # 実 xlsx の置き場。明示指定があればそれを使い（Googleドライブ等）、
        # 無ければ base の親（crm_scraper/）にフォールバック。
        self.real_data_folder = (
            Path(real_data_folder) if real_data_folder else self.base.parent
        )
        self.master_csv = Path(master_csv)
        self.target_month = target_month
        self.assumptions = assumptions
        self._shift_data_cache = None  # 同一インスタンス内の一時記憶（毎回 new するので跨がない）

    # =====================================================================
    # master（過去患者×シフト×天気の統合）
    # =====================================================================
    def load_master(self) -> pd.DataFrame:
        if not self.master_csv.exists():
            raise RealDataMissingError(
                f'ultimate_shift_master.csv が見つかりません: {self.master_csv}\n'
                f'先に build_ultimate_master.py を実行して 売上 xlsx から構築してください。'
            )
        df = pd.read_csv(self.master_csv, encoding='utf-8-sig')
        before = len(df)
        # 退職・除外スタッフ（池田 等）が万一含まれていれば全列スキャンで除外
        df = drop_rows_with_excluded_staff(df)
        if len(df) != before:
            print(f'[DataLoader] master 退職スタッフ行を除外: {before} → {len(df)}')
        print(f'[DataLoader] master 読込: {len(df)} 行 (ultimate_shift_master.csv)')
        return df

    # =====================================================================
    # 院マスタ（静的）
    # =====================================================================
    def load_clinic_master(self) -> pd.DataFrame:
        df = pd.DataFrame(_CLINIC_MASTER_ROWS)
        print(f'[DataLoader] clinic_master: 静的定義 ({len(df)} 件)')
        return df

    # 互換 alias
    def load_or_create_clinic_master(self) -> pd.DataFrame:
        return self.load_clinic_master()

    # =====================================================================
    # シフト xlsx からの実データ抽出
    # =====================================================================
    def _load_shift_data(self) -> dict:
        if self._shift_data_cache is not None:
            return self._shift_data_cache
        print(f'[DataLoader] シフトxlsx読込開始 ({self.real_data_folder})')
        data = load_real_shift_data(self.real_data_folder, self.target_month)
        print(f'[DataLoader] シフトxlsx対象月={self.target_month} '
              f'マッチ院={sorted(data["found_areas"])} '
              f'欠落={data["missing_areas"]}')
        for area, sn in data['matched_sheets'].items():
            print(f'   - {area}: シート「{sn}」を採用')
        self._shift_data_cache = data
        return data

    # ---------------- planned_shifts (実データ) ----------------
    def load_or_create_planned_shifts(self, required_df=None) -> pd.DataFrame:
        data = self._load_shift_data()
        df = data['planned']
        if df.empty:
            raise RealDataMissingError(
                f'シフトxlsxに {self.target_month} のデータがありません。\n'
                f' - 対象月のシートが各シフト表に作成されているか確認してください。\n'
                f' - 読取対象ファイル: {data["files_checked"]}\n'
                f' - {self.target_month} を含まない院: {data["missing_areas"]}'
            )
        if data['missing_areas']:
            self.assumptions.append(
                f'planned_shifts: {self.target_month} のシートが無い院 '
                f'{data["missing_areas"]} は集計対象外'
            )
        print(f'[DataLoader] planned_shifts (実データ): {len(df)} 行, '
              f'院数={df["area"].nunique()}')
        return df

    # ---------------- 人事系の補助データ (実データ) ----------------
    def load_paid_leave(self) -> pd.DataFrame:
        """有給休暇取得記録 [date, area, staff_name, type]"""
        data = self._load_shift_data()
        df = data.get('paid_leave', pd.DataFrame(
            columns=['date', 'area', 'staff_name', 'type']
        ))
        df = drop_rows_with_excluded_staff(df, staff_columns=['staff_name'])
        print(f'[DataLoader] paid_leave (実データ): {len(df)} 件')
        return df

    def load_help_actions_actual(self) -> pd.DataFrame:
        """実シフト上で発生済みのヘルプ移動 [date, staff_name, src_clinic,
        dst_clinic, cell_value]"""
        data = self._load_shift_data()
        df = data.get('help_actions', pd.DataFrame(
            columns=['date', 'staff_name', 'src_clinic',
                     'dst_clinic', 'cell_value']
        ))
        df = drop_rows_with_excluded_staff(df, staff_columns=['staff_name'])
        print(f'[DataLoader] help_actions (実データ): {len(df)} 件')
        return df

    def load_fixed_leave(self) -> pd.DataFrame:
        """固定休（結婚式・運動会等の年間予定）[date, area, staff_name, type]"""
        data = self._load_shift_data()
        df = data.get('fixed_leave', pd.DataFrame(
            columns=['date', 'area', 'staff_name', 'type']
        ))
        df = drop_rows_with_excluded_staff(df, staff_columns=['staff_name'])
        print(f'[DataLoader] fixed_leave (実データ): {len(df)} 件')
        return df

    # ---------------- worked_staff (実データ) ----------------
    def load_worked_staff(self) -> pd.DataFrame:
        """日別×スタッフ名の出勤明細を返す（後段の応援案で利用）"""
        data = self._load_shift_data()
        df = data.get('worked', pd.DataFrame(
            columns=['date', 'area', 'staff_name']
        ))
        # 安全網: 退職・除外スタッフを最終フィルタ
        before = len(df)
        df = drop_rows_with_excluded_staff(df, staff_columns=['staff_name'])
        if before != len(df):
            print(f'[DataLoader] worked_staff 退職スタッフ除外: {before} → {len(df)}')
        print(f'[DataLoader] worked_staff (実データ): {len(df)} 行, '
              f'院数={df["area"].nunique() if len(df) else 0}')
        return df

    # ---------------- leave_requests (実データ) ----------------
    def load_or_create_leave_requests(self, clinic_master=None) -> pd.DataFrame:
        data = self._load_shift_data()
        df = data['leave']
        if df.empty:
            # 希望休が0件の月はあり得るので警告のみで継続
            print(f'[DataLoader] leave_requests (実データ): 0件 '
                  f'（{self.target_month} に「希」セルなし）')
            self.assumptions.append(
                f'leave_requests: {self.target_month} のシフトxlsx内に「希望休」記載なし'
            )
            return pd.DataFrame(columns=['申請日', 'スタッフID', 'スタッフ名',
                                          '院', '希望日', '希望種別', '重要度', '備考'])
        # 安全網: 退職・除外スタッフ（池田 等）を最終フィルタ
        before = len(df)
        df = drop_rows_with_excluded_staff(
            df, staff_columns=['スタッフ名', 'スタッフID']
        )
        if before != len(df):
            print(f'[DataLoader] leave_requests 退職スタッフ除外: '
                  f'{before} → {len(df)}')
        print(f'[DataLoader] leave_requests (実データ): {len(df)} 件 '
              f'（院別: {df["院"].value_counts().to_dict()}）')
        return df

    # =====================================================================
    # 天気予報 — ultimate_shift_master.csv の実天気から target_month を抽出。
    # 実天気が無い対象月は「天気未使用」で続行（生成はしない）。
    # =====================================================================
    def load_or_create_weather_forecast(self, df_master) -> pd.DataFrame:
        empty = pd.DataFrame(columns=['日付', 'エリア', '天気', '最高気温'])
        if df_master is None or df_master.empty:
            return empty
        df = df_master.copy()
        if '天気' not in df.columns or 'date' not in df.columns:
            print('[DataLoader] weather_forecast: 天気列なし → 未使用')
            return empty
        df['date'] = pd.to_datetime(df['date'], errors='coerce')
        year, month = map(int, self.target_month.split('-'))
        mask = (df['date'].dt.year == year) & (df['date'].dt.month == month) & df['天気'].notna()
        sub = df[mask].copy()
        if sub.empty:
            print(f'[DataLoader] weather_forecast: {self.target_month} の実天気データなし → 未使用')
            self.assumptions.append(
                f'weather: {self.target_month} の実天気データがないため、天気特徴量は使用しない'
            )
            return empty
        sub['天気エリア'] = sub['area'].map(
            lambda a: '人形町' if a == '人形町' else '小金井'
        )
        out = (sub.groupby(['date', '天気エリア'])
                  .agg({'天気': 'first', '最高気温': 'mean'})
                  .reset_index()
                  .rename(columns={'date': '日付', '天気エリア': 'エリア'}))
        out['日付'] = out['日付'].dt.strftime('%Y-%m-%d')
        print(f'[DataLoader] weather_forecast (実データ): {len(out)} 行')
        return out
