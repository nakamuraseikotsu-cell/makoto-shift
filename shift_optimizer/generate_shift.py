# -*- coding: utf-8 -*-
"""シフト自動最適化 CLI

使い方:
    py shift_optimizer\\generate_shift.py --target-month 2026-04

出力:
    outputs/YYYYMM/
        auto_shift_schedule.csv      … 配置全件（date, staff, home, assigned, status, cell）
        auto_shift_matrix_<院>.csv   … 各院の縦＝スタッフ／横＝日付 のマトリックス
        auto_shift_shortages.csv     … 充足できなかった日リスト
        auto_shift_summary.md        … 概要レポート
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent))

from src.shift_generator import optimize_shift, REQUIRED_BY_DOW  # noqa: E402

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass


def parse_args():
    p = argparse.ArgumentParser(description='シフト自動最適化（PuLP）')
    p.add_argument('--target-month', default='2026-04', help='YYYY-MM')
    p.add_argument('--base-dir', default=str(Path(__file__).parent))
    p.add_argument('--time-limit', type=int, default=90,
                   help='ソルバ制限時間（秒、既定90）')
    return p.parse_args()


def write_matrices(df_assign, output_dir: Path):
    """院別マトリックスCSVを出力"""
    for clinic, sub in df_assign.groupby('home_clinic'):
        # date を列、staff_name を行、 cell を値に
        pivot = sub.pivot_table(index='staff_name', columns='date',
                                 values='cell', aggfunc='first').fillna('')
        # 日付順に列を整列
        pivot = pivot.reindex(sorted(pivot.columns), axis=1)
        out_path = output_dir / f'auto_shift_matrix_{clinic}.csv'
        pivot.to_csv(out_path, encoding='utf-8-sig')


def df_to_md(df, max_rows=None):
    if df is None or df.empty:
        return '_(該当なし)_'
    if max_rows:
        df = df.head(max_rows)
    cols = list(df.columns)
    header = '| ' + ' | '.join(str(c) for c in cols) + ' |'
    sep = '| ' + ' | '.join(['---'] * len(cols)) + ' |'
    body = []
    for _, r in df.iterrows():
        body.append('| ' + ' | '.join(
            ('' if v is None or (hasattr(v, '__len__') and not str(v).strip()) else str(v))
            for v in r.values
        ) + ' |')
    return '\n'.join([header, sep] + body)


def write_summary(result, target_month: str, output_dir: Path):
    df_a = result['assignments']
    df_s = result['shortages']
    df_w = result['work_days']

    md = []
    md.append(f'# シフト自動最適化 結果サマリー：{target_month}\n')
    md.append(f'生成日時：{datetime.now().strftime("%Y-%m-%d %H:%M")}\n')

    md.append('## 1. 結果概要\n')
    md.append(f'- 最適化ステータス：**{result["status"]}**')
    md.append(f'- マッチした院：{result["matched_areas"]}')
    if result['missing_areas']:
        md.append(f'- 名簿が取れなかった院：{result["missing_areas"]} '
                  f'（対象月シート未作成）')
    md.append(f'- 中央線スタッフ数：{len(result["chuo_staff"])} 名')
    md.append(f'- 人形町スタッフ数：{len(result["ningyo_staff"])} 名')
    md.append(f'- 反映した希望休：{result["leave_count"]} 件')
    md.append(f'- 出勤割り当て合計：{int((df_a["status"].isin(["自院出勤","応援出勤"])).sum())} 行')
    md.append(f'- 応援出勤の延べ件数：{int((df_a["status"]=="応援出勤").sum())} 回')
    md.append(f'- 不足発生：{len(df_s)} 行\n')

    md.append('## 2. 必要人員（曜日別）\n')
    md.append('| 院 | 月 | 火 | 水 | 木 | 金 | 土 | 日 |')
    md.append('| --- | --- | --- | --- | --- | --- | --- | --- |')
    for clinic, req in REQUIRED_BY_DOW.items():
        md.append(f'| {clinic} | {req["月"]} | {req["火"]} | {req["水"]} | '
                  f'{req["木"]} | {req["金"]} | {req["土"]} | {req["日"]} |')
    md.append('')

    md.append('## 3. スタッフ別 月間出勤日数\n')
    md.append(df_to_md(df_w.sort_values(['home_clinic', 'staff_name'])))
    md.append('')

    md.append('## 4. 不足が出た日（全件）\n')
    if df_s.empty:
        md.append('✅ すべての日 / 院で必要人数を充足しています。\n')
    else:
        md.append(df_to_md(df_s))
        md.append('')
        md.append(f'**不足合計：{int(df_s["gap"].sum())} 人日**\n')

    md.append('## 5. 応援サンプル（上位30件）\n')
    sup = df_a[df_a['status'] == '応援出勤'].head(30)
    if sup.empty:
        md.append('_(応援出勤なし)_')
    else:
        md.append(df_to_md(sup[['date', 'day_of_week', 'staff_name',
                                 'home_clinic', 'assigned_clinic']]))
    md.append('')

    md.append('## 6. 制約条件（実装時）\n')
    md.append('- 各スタッフ 月間出勤 ≤ 22 日（22 日へ寄せる目標）')
    md.append('- 最大 6 連勤（7日窓 sum ≤ 6）')
    md.append('- 「希」セルは強制公休')
    md.append('- 1日1院（中央線スタッフは4院いずれかへ配置）')
    md.append('- 人形町は独立（他院との行き来なし）')
    md.append('- 必要人数を満たせない分は不足としてログ出力')

    (output_dir / 'auto_shift_summary.md').write_text(
        '\n'.join(md), encoding='utf-8'
    )


def run_generation_core(target_month: str, base_dir, crm_folder,
                         time_limit: int = 90) -> dict:
    """Streamlit / CLI 共通の自動生成エントリ"""
    base = Path(base_dir)
    crm_folder = Path(crm_folder)
    yyyymm = target_month.replace('-', '')
    output_dir = base / 'outputs' / yyyymm
    output_dir.mkdir(parents=True, exist_ok=True)

    result = optimize_shift(target_month, crm_folder, time_limit=time_limit)

    # CSV 出力
    result['assignments'].to_csv(
        output_dir / 'auto_shift_schedule.csv',
        index=False, encoding='utf-8-sig')
    result['shortages'].to_csv(
        output_dir / 'auto_shift_shortages.csv',
        index=False, encoding='utf-8-sig')
    result['work_days'].to_csv(
        output_dir / 'auto_shift_work_days.csv',
        index=False, encoding='utf-8-sig')
    write_matrices(result['assignments'], output_dir)
    write_summary(result, target_month, output_dir)

    result['output_dir'] = output_dir
    result['target_month'] = target_month
    return result


def main():
    args = parse_args()
    base = Path(args.base_dir)
    crm_folder = base.parent

    print('=' * 80)
    print(f' シフト自動最適化 (PuLP)  対象月: {args.target_month}')
    print('=' * 80)

    try:
        result = run_generation_core(args.target_month, base, crm_folder,
                                       time_limit=args.time_limit)
    except RuntimeError as e:
        print()
        print('!' * 80)
        print(' ❌ シフト最適化を実行できません')
        print('!' * 80)
        print(str(e))
        sys.exit(2)

    # ===== コンソール表示 =====
    print()
    print('=' * 80)
    print(f' 最適化サマリー: {args.target_month}')
    print('=' * 80)
    print(f'  状態: {result["status"]}')
    print(f'  中央線スタッフ: {len(result["chuo_staff"])} 名 / 人形町: {len(result["ningyo_staff"])} 名')
    print(f'  反映した希望休: {result["leave_count"]} 件')
    print(f'  応援出勤の延べ件数: {int((result["assignments"]["status"]=="応援出勤").sum())} 回')
    print()
    df_s = result['shortages']
    if df_s.empty:
        print('  ✅ 全日・全院で必要人数を充足しています')
    else:
        print(f'  ▼ 不足リスト（{len(df_s)}件 / 合計 {int(df_s["gap"].sum())}人日）:')
        for _, r in df_s.iterrows():
            print(f'    {r["date"]} ({r["day_of_week"]}) '
                  f'{r["clinic"]}: {r["gap"]}人不足 '
                  f'(必要 {r["required"]}, 配置 {r["assigned"]})')
    print()
    print(f'  出力フォルダ: {result["output_dir"]}')


if __name__ == '__main__':
    main()
