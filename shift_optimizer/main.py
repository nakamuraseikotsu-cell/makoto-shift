# -*- coding: utf-8 -*-
"""
シフト最適化プロトタイプ（CLI エントリ）

使い方:
    python main.py --target-month 2026-06
    python main.py --target-month 2026-06 --min-staff 2
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

# src ディレクトリを import path に追加
sys.path.insert(0, str(Path(__file__).parent))

from src.data_loader import DataLoader, RealDataMissingError  # noqa: E402
from src.preprocess import Preprocessor            # noqa: E402
from src.forecast import Forecaster                # noqa: E402
from src.staffing_calculator import StaffingCalculator  # noqa: E402
from src.reallocator import Reallocator            # noqa: E402
from src.reporting import Reporter                 # noqa: E402

# 文字化け対策
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass


def parse_args():
    p = argparse.ArgumentParser(description='シフト最適化プロトタイプ')
    p.add_argument('--target-month', default='2026-06',
                   help='対象月 YYYY-MM（既定 2026-06）')
    p.add_argument('--min-staff', type=int, default=2,
                   help='最低運営人数（既定 2 人）')
    p.add_argument('--base-dir', default=str(Path(__file__).parent),
                   help='shift_optimizer フォルダのパス')
    p.add_argument('--master-csv', default=None,
                   help='ultimate_shift_master.csv のパス（既定 ../ 配下）')
    return p.parse_args()


def run_full_analysis(target_month: str, min_staff: int,
                       base_dir, master_csv,
                       real_data_folder=None,
                       movable_whitelist=None) -> dict:
    """Streamlit / CLI 共通の分析エントリ。
    関連 DataFrame を辞書で返す。CSV/MDの書き出しも併せて行う。

    real_data_folder: シフト xlsx の探索フォルダ（Drive など）。
                       未指定なら base_dir の親フォルダを使う。
    movable_whitelist: 院長が明示指定した「ヘルプ要員スタッフ」の dict
                       {院名: [スタッフ名, ...]}。指定時は当該院の応援候補が
                       このリストに限定される（Reallocator に伝搬）。
    """
    base = Path(base_dir)
    master_csv = Path(master_csv)
    yyyymm = target_month.replace('-', '')
    output_dir = base / 'outputs' / yyyymm
    output_dir.mkdir(parents=True, exist_ok=True)
    assumptions: list[str] = []

    # 1) データ読込
    loader = DataLoader(base, master_csv, target_month, assumptions,
                         real_data_folder=real_data_folder)
    df_master = loader.load_master()

    # 2) 前処理
    pre = Preprocessor(assumptions)
    df = pre.run(df_master)

    # 3) 院マスタ
    clinic_master = loader.load_or_create_clinic_master()

    # 4) 生産性算出
    forecaster = Forecaster(df, target_month, assumptions)
    productivity = forecaster.compute_productivity()

    # 5) 天気予報
    weather_forecast = loader.load_or_create_weather_forecast(df_master)

    # 6) 来院予測
    forecast_df = forecaster.predict_visits(weather_forecast, clinic_master)

    # 7) 必要人員
    calc = StaffingCalculator(min_staff=min_staff, assumptions=assumptions)
    required_df = calc.calculate_required(forecast_df, productivity)

    # 8) 予定シフト
    planned_shifts = loader.load_or_create_planned_shifts(required_df)
    # シフト xlsx の解析で対象月シートが無かった院（明示的に取り出す）
    _shift_data = loader._load_shift_data()
    missing_areas = list(_shift_data.get('missing_areas', []))

    # 9) 休み希望
    leave_requests = loader.load_or_create_leave_requests(clinic_master)

    # 10) 過不足判定
    gap_df = calc.calculate_gap(required_df, planned_shifts, leave_requests)

    # 11) 応援移動（院単位の集計案）
    reallocator = Reallocator(clinic_master, assumptions)
    suggestions = reallocator.suggest_moves(gap_df)

    # 11b) 応援アクション指示（スタッフ名単位・動かせないスタッフを除外）
    worked_df = loader.load_worked_staff()
    # 固定休（結婚式・運動会等）は応援対象から除外
    fixed_leave_df = loader.load_fixed_leave()
    staff_help_actions = reallocator.suggest_staff_help_actions(
        gap_df, worked_df, fixed_leave_df=fixed_leave_df,
        movable_whitelist=movable_whitelist,
    )

    # 11c) 人事関連（有給取得・ヘルプ実績）
    paid_leave_df = loader.load_paid_leave()
    help_actions_actual_df = loader.load_help_actions_actual()

    # 12) ファイル書き出し（Reporter）
    reporter = Reporter(output_dir, target_month, assumptions)
    reporter.write_all(required_df, gap_df, suggestions, productivity)
    # 応援アクション指示も CSV として保存
    try:
        staff_help_actions.to_csv(
            output_dir / 'staff_help_actions.csv',
            index=False, encoding='utf-8-sig',
        )
    except Exception:
        pass

    return {
        'target_month': target_month,
        'output_dir': output_dir,
        'required_df': required_df,
        'gap_df': gap_df,
        'suggestions': suggestions,
        'staff_help_actions': staff_help_actions,
        'worked_df': worked_df,
        'productivity': productivity,
        'assumptions': assumptions,
        'clinic_master': clinic_master,
        'leave_requests': leave_requests,
        'planned_shifts': planned_shifts,
        'paid_leave_df': paid_leave_df,
        'help_actions_actual_df': help_actions_actual_df,
        'fixed_leave_df': fixed_leave_df,
        'missing_areas': missing_areas,
    }


def main():
    args = parse_args()
    base = Path(args.base_dir)
    master_csv = args.master_csv or str(base.parent / 'ultimate_shift_master.csv')

    print('=' * 80)
    print(f' シフト最適化プロトタイプ（実データ専用版）')
    print(f'  対象月: {args.target_month}  最低人数: {args.min_staff}')
    print('=' * 80)

    try:
        result = run_full_analysis(args.target_month, args.min_staff, base, master_csv)
        reporter = Reporter(result['output_dir'], args.target_month, result['assumptions'])
        reporter.print_console(result['gap_df'], result['suggestions'])
        print()
        print(f'出力フォルダ: {result["output_dir"]}')
    except RealDataMissingError as e:
        print()
        print('!' * 80)
        print(' ❌ 実データ不足のため処理を停止しました')
        print('!' * 80)
        print(str(e))
        print()
        print('対処方法:')
        print(f'  1) 各院のシフトxlsxに「{args.target_month}」のシートを作成してください')
        print('  2) または、過去月（例: --target-month 2026-05）を指定して動作確認してください')
        sys.exit(2)


if __name__ == '__main__':
    main()
