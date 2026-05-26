# -*- coding: utf-8 -*-
"""応援移動候補の提案（同エリア優先）"""
from __future__ import annotations
from dataclasses import dataclass
import pandas as pd

from .constants import (
    IMMOVABLE_STAFF,
    drop_rows_with_excluded_staff,
)


# =============================================================================
# ベクトル化ヘルパ（apply axis=1 の代替）
# =============================================================================
def _build_immovable_mask(
    wdf: pd.DataFrame,
    movable_whitelist: dict | None,
) -> pd.Series:
    """worked_df に対する「動かせない」ブーリアンマスクをベクトル化生成する。

    - movable_whitelist[area] が明示指定された院 →
      リスト外のスタッフは全員 immovable
    - whitelist に院が無い場合 → IMMOVABLE_STAFF (姓部分一致) で判定

    apply(axis=1) を完全に避け、列演算と str.contains だけで合成する。
    """
    mask = pd.Series(False, index=wdf.index)
    if wdf.empty:
        return mask
    area_col = wdf['area']
    name_col = wdf['staff_name'].astype(str)
    wl_areas = set(movable_whitelist.keys()) if movable_whitelist else set()

    # 1) 院長指定の whitelist 院: リスト外を immovable
    if movable_whitelist:
        for area, allowed in movable_whitelist.items():
            area_mask = (area_col == area)
            mask |= area_mask & ~name_col.isin(set(allowed))

    # 2) 既定の IMMOVABLE_STAFF 判定（whitelist に含まれない院のみ）
    for area, surnames in IMMOVABLE_STAFF.items():
        if area in wl_areas:
            continue
        area_mask = (area_col == area)
        for surname in surnames:
            mask |= area_mask & name_col.str.contains(
                surname, regex=False, na=False
            )
    return mask


@dataclass
class _DropResult:
    df: pd.DataFrame
    dropped: int


def _drop_fixed_leave(
    movable_df: pd.DataFrame,
    fixed_leave_df: pd.DataFrame | None,
) -> _DropResult:
    """固定休 (date, staff_name) に該当する行を merge で anti-join 除外する。"""
    if fixed_leave_df is None or len(fixed_leave_df) == 0:
        return _DropResult(movable_df, 0)
    keys = (
        fixed_leave_df[['date', 'staff_name']]
        .drop_duplicates()
        .assign(date=lambda d: d['date'].astype(str), _fixed=True)
    )
    joined = (
        movable_df.assign(date=movable_df['date'].astype(str))
        .merge(keys, on=['date', 'staff_name'], how='left')
    )
    dropped = int(joined['_fixed'].fillna(False).sum())
    if dropped == 0:
        return _DropResult(movable_df, 0)
    return _DropResult(
        joined[joined['_fixed'].isna()].drop(columns='_fixed'),
        dropped,
    )


class Reallocator:
    """過不足DFから、同日内で 余剰院→不足院 の応援移動案を生成"""

    def __init__(self, clinic_master: pd.DataFrame,
                 assumptions: list[str] | None = None):
        self.clinic_master = clinic_master
        self.assumptions = assumptions if assumptions is not None else []
        self.area_map = clinic_master.set_index('院名')['エリア'].to_dict()

    def suggest_moves(self, gap_df: pd.DataFrame) -> pd.DataFrame:
        suggestions = []
        for date, day_group in gap_df.groupby('date'):
            short = day_group[day_group['gap'] < 0].copy()
            surplus = day_group[day_group['gap'] > 0].copy()
            if short.empty or surplus.empty:
                continue
            # 不足が大きい順、余剰が大きい順
            short = short.sort_values('gap')
            surplus = surplus.sort_values('gap', ascending=False).copy()

            for _, s in short.iterrows():
                need = abs(int(s['gap']))
                if need <= 0:
                    continue
                s_area = self.area_map.get(s['area'], '')
                # 候補：余剰の中から同エリア優先
                cand = surplus.copy()
                cand['_same_area'] = cand['area'].map(
                    lambda a: 1 if self.area_map.get(a, '') == s_area else 0
                )
                cand = cand.sort_values(['_same_area', 'gap'],
                                        ascending=[False, False])
                for idx, c in cand.iterrows():
                    if need <= 0:
                        break
                    available = int(c['gap'])
                    if available <= 0:
                        continue
                    move_n = min(need, available)
                    # 推奨度スコア：同エリアは +10, 余剰量に応じて +最大5
                    score = (10 if c['_same_area'] == 1 else 5) + min(available, 5)
                    suggestions.append({
                        'date': date,
                        'shortage_clinic': s['area'],
                        'shortage_n': need,
                        'support_clinic': c['area'],
                        'support_available': available,
                        'recommend_move': move_n,
                        'same_area': '同エリア' if c['_same_area'] == 1 else '他エリア',
                        'priority_score': score,
                    })
                    need -= move_n
                    surplus.loc[surplus['area'] == c['area'], 'gap'] -= move_n

        if not suggestions:
            cols = ['date', 'shortage_clinic', 'shortage_n', 'support_clinic',
                    'support_available', 'recommend_move',
                    'same_area', 'priority_score']
            return pd.DataFrame(columns=cols)

        df = pd.DataFrame(suggestions).sort_values(
            ['date', 'priority_score'], ascending=[True, False]
        ).reset_index(drop=True)
        self.assumptions.append(
            '応援移動: 同日内のみ、同エリア優先 → 他エリアにフォールバック。推奨度=同エリア+10, 余剰量で加点'
        )
        return df

    # =====================================================================
    # スタッフ名レベルの応援アクション指示
    # -----------------------------------------------------------------------
    # gap_df + worked_df から、日付・スタッフ名・移動元・移動先 の組合せを生成。
    # 動かせないスタッフ (IMMOVABLE_STAFF) は src として絶対に選ばない。
    # =====================================================================
    def suggest_staff_help_actions(
        self,
        gap_df: pd.DataFrame,
        worked_df: pd.DataFrame,
        fixed_leave_df: pd.DataFrame | None = None,
        movable_whitelist: dict[str, set[str]] | dict[str, list[str]] | None = None,
    ) -> pd.DataFrame:
        """日次の応援指示を、スタッフ名・移動元・移動先 単位で返す。

        fixed_leave_df: 固定休（結婚式・運動会等）— 該当 (date, staff_name)
                       は応援対象から完全に除外する。
        movable_whitelist: 院長が明示的に許可した「ヘルプ要員」スタッフ。
                          {院名: [スタッフ名, ...]} の dict。
                          指定時は **このリストに含まれないスタッフは src 候補
                          から完全に除外** され、`IMMOVABLE_STAFF` ルールよりも
                          優先される（院長判断 > 既定ルール）。
                          None または当該院キー欠落の場合は、従来通り
                          IMMOVABLE_STAFF に基づいて自動判定する。
        Returns: DataFrame [date, staff_name, src_clinic, dst_clinic,
                            same_area, shortage_n_remaining]
        """
        cols = ['date', 'staff_name', 'src_clinic', 'dst_clinic',
                'same_area', 'shortage_n_remaining']
        if gap_df is None or gap_df.empty:
            return pd.DataFrame(columns=cols)
        if worked_df is None or worked_df.empty:
            self.assumptions.append(
                '応援アクション: worked データが空のためスタッフ名指定は不可'
            )
            return pd.DataFrame(columns=cols)

        # 日付ごとに、 (院, [動かせるスタッフ名一覧]) を構築
        wdf = worked_df.copy()
        # 退職・除外スタッフ（池田 等）を最終安全網として再除外
        wdf = drop_rows_with_excluded_staff(
            wdf, staff_columns=['staff_name']
        )

        # immovable 判定マスクをベクトル化して構築
        wdf['_immovable'] = _build_immovable_mask(wdf, movable_whitelist)
        if movable_whitelist:
            n_excluded = int(wdf['_immovable'].sum())
            self.assumptions.append(
                f'応援アクション: 院長指定の movable_whitelist を適用 '
                f'(対象院 {sorted(movable_whitelist.keys())}) — '
                f'候補外 {n_excluded} レコードを除外'
            )
        movable_df = wdf[~wdf['_immovable']].copy()

        # 固定休（年間予定の確保された休み）に該当する (date, staff) を
        # ベクトル化 merge で除外
        fixed_excluded = _drop_fixed_leave(movable_df, fixed_leave_df)
        if fixed_excluded.dropped > 0:
            movable_df = fixed_excluded.df
            self.assumptions.append(
                f'応援アクション: 固定休 (結婚式・運動会等) '
                f'{fixed_excluded.dropped} 件を応援候補から除外'
            )

        # 各日付・各院での「動かせるスタッフ」をリスト化
        # (date, area) -> list[staff_name]
        movable_df_str_date = movable_df.assign(
            _date_str=movable_df['date'].astype(str)
        )
        movable_idx: dict[tuple, list[str]] = (
            movable_df_str_date
            .groupby(['_date_str', 'area'])['staff_name']
            .apply(list)
            .to_dict()
        )

        actions: list[dict] = []
        # 日付単位で、 余剰院→不足院 へスタッフを充当
        # 1スタッフは1日1回しか移動できない（後段で重複排除）
        used_staff_per_day: set[tuple] = set()  # (date, src_area, staff_name)

        for date, day_group in gap_df.groupby('date'):
            short = day_group[day_group['gap'] < 0].copy()
            surplus = day_group[day_group['gap'] > 0].copy()
            if short.empty or surplus.empty:
                continue
            short = short.sort_values('gap')  # 不足が大きい順 (gapが小さい順)
            surplus = surplus.sort_values('gap', ascending=False).copy()
            # 各院の残り余剰枠（dict 化はベクトル代入で）
            surplus_remaining: dict[str, int] = dict(zip(
                surplus['area'].tolist(),
                surplus['gap'].astype(int).tolist(),
            ))

            for _, s in short.iterrows():
                need = abs(int(s['gap']))
                dst_clinic = s['area']
                dst_area = self.area_map.get(dst_clinic, '')
                if need <= 0:
                    continue
                # 同エリア優先で src 候補を順序付け
                src_candidates = sorted(
                    surplus_remaining.keys(),
                    key=lambda c: (
                        0 if self.area_map.get(c, '') == dst_area else 1,
                        -surplus_remaining[c],
                    ),
                )
                for src_clinic in src_candidates:
                    if need <= 0:
                        break
                    available = surplus_remaining.get(src_clinic, 0)
                    if available <= 0:
                        continue
                    # 動かせるスタッフから N人を取り出す
                    same_area = (
                        self.area_map.get(src_clinic, '') == dst_area
                    )
                    pool = movable_idx.get((str(date), src_clinic), [])
                    pool = [s_name for s_name in pool
                            if (str(date), src_clinic, s_name)
                                not in used_staff_per_day]
                    if not pool:
                        continue
                    take_n = min(need, available, len(pool))
                    for s_name in pool[:take_n]:
                        actions.append({
                            'date': str(date),
                            'staff_name': s_name,
                            'src_clinic': src_clinic,
                            'dst_clinic': dst_clinic,
                            'same_area': '同エリア' if same_area else '他エリア',
                            'shortage_n_remaining': need - 1,
                        })
                        used_staff_per_day.add(
                            (str(date), src_clinic, s_name)
                        )
                        need -= 1
                        surplus_remaining[src_clinic] -= 1

        if not actions:
            return pd.DataFrame(columns=cols)
        df = pd.DataFrame(actions).sort_values(
            ['date', 'dst_clinic', 'src_clinic']
        ).reset_index(drop=True)
        self.assumptions.append(
            f'応援アクション: スタッフ名単位で {len(df)} 件の移動案を生成 / '
            f'動かせないスタッフ ({sum(len(v) for v in IMMOVABLE_STAFF.values())} 名) '
            f'は src 候補から除外'
        )
        return df
