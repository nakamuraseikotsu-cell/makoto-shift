# -*- coding: utf-8 -*-
"""レポート出力 & コンソール表示"""
from __future__ import annotations
from pathlib import Path
from datetime import datetime
import pandas as pd


DISCIPLINE_TEMPLATES = '''# 労務関連 文書テンプレート（{target_month} 分）

> ⚠️ **重要な注意事項**
> このテンプレートは事務手続を円滑化するための **参考雛形** です。
> **懲戒処分を判定・決定するものではありません。**
> 個別の判断は、必ず **就業規則・雇用契約・労務手続に基づき個別判断してください。**
> 本書面の使用前に、社内ルールおよび労働基準法・関連法令との適合性をご確認のうえご利用ください。

---

## 1. 「注意」テンプレート（書面通知の草案）

```
宛先：{{スタッフ名}} 様
発信日：{{発信日}}
所属：{{院名}}
件名：勤務状況に関するお知らせ

平素より業務にご尽力いただき、誠にありがとうございます。
近日の勤務状況において、以下の点を確認させていただきました。

【確認内容】
- 該当日：{{該当日}}
- 内容：{{事実関係（例: 出勤予定時刻からの連絡なしの遅延、シフト調整への協力依頼に対する未回答 等）}}

つきましては、就業規則および雇用契約書の取り決めに沿って、
今後の対応について所属長と話し合いの場を持たせていただきたくお願い申し上げます。

なお本書面は事実確認および注意喚起を目的としたものであり、懲戒処分の決定ではありません。
ご不明点があれば本書面受領後 3 営業日以内にご連絡ください。

所属長：{{所属長氏名}}
連絡先：{{連絡先}}
```

---

## 2. 「面談勧告」テンプレート

```
宛先：{{スタッフ名}} 様
発信日：{{発信日}}
件名：勤務シフト調整に関するご面談のお願い

このたび、対象月（{target_month}）の業務量と人員配置の予測において、
{{院名}}での貴殿の出勤予定とご希望休との調整について、
対面でのご相談が必要と判断いたしました。

【ご相談したい内容】
- 必要人員数と現在のシフト予定との差
- ご希望休の優先順位の確認
- 業務調整（応援出勤、休日変更、勤務時間調整など）のご相談

ご都合のよろしい日時を以下の候補からお選びいただき、ご返信をお願い申し上げます。
候補日：{{日付候補1}} / {{日付候補2}} / {{日付候補3}}

本面談は人員配置の最適化を目的とした業務調整のご相談であり、
懲戒手続を前提とするものではありません。

所属長：{{所属長氏名}}
連絡先：{{連絡先}}
```

---

## 3. 「就業規則確認依頼」テンプレート

```
宛先：{{スタッフ名}} 様
発信日：{{発信日}}
件名：就業規則の関連条項のご確認のお願い

業務運営上の必要性から、下記の条項について、改めてご確認いただきますよう
お願い申し上げます。

【ご確認いただきたい条項】
- 第〇〇条（勤務時間およびシフト）
- 第〇〇条（休暇申請の手続）
- 第〇〇条（業務命令・応援出勤への協力）

ご質問、ご相談、解釈の確認等がございましたら、所属長または労務担当まで
お気軽にお申し付けください。

本書面は事実確認・情報共有を目的としたものであり、
いかなる処分の決定でもありません。
個別の状況については、雇用契約・就業規則・労務手続に基づき
適切な手順で個別判断いたします。

所属長：{{所属長氏名}}
連絡先：{{連絡先}}
```

---

## 4. 対応履歴 記録用テーブル

| 項目 | 内容 |
|------|------|
| 依頼日 | {{依頼日}} |
| 本人回答 | {{回答内容 / 未回答}} |
| 代替案提示 | {{有 / 無}} |
| 面談実施日 | {{実施日}} |
| 所属長コメント | {{コメント}} |

> このテーブルは記録・引継ぎを目的とした個別ファイル化を想定しています。
> 同じ内容を `outputs/{{yyyymm}}/discipline_log/個別.md` のように複製してご利用ください。

---

**最終確認者：__________________ ／ 日付：__________________**

---

> 再掲：本書面群は **懲戒処分の決定文ではありません**。
> 個別事案は必ず **就業規則・雇用契約・労務手続** に基づき、
> 所属長と労務担当の合議で個別判断してください。
'''


def df_to_md(df: pd.DataFrame, max_rows: int | None = None) -> str:
    """tabulate に依存しない簡易 Markdown テーブル"""
    if df is None or df.empty:
        return '_（データなし）_'
    if max_rows is not None:
        df = df.head(max_rows)
    cols = list(df.columns)
    header = '| ' + ' | '.join(str(c) for c in cols) + ' |'
    sep = '| ' + ' | '.join(['---'] * len(cols)) + ' |'
    body_rows = []
    for _, r in df.iterrows():
        body_rows.append(
            '| ' + ' | '.join(
                '' if pd.isna(v) else str(v) for v in r.values
            ) + ' |'
        )
    return '\n'.join([header, sep] + body_rows)


class Reporter:
    def __init__(self, output_dir: Path, target_month: str,
                 assumptions: list[str]):
        self.out = Path(output_dir)
        self.target_month = target_month
        self.assumptions = assumptions

    # ----------------------------- 一括 -----------------------------
    def write_all(self, required_df: pd.DataFrame, gap_df: pd.DataFrame,
                  suggestions: pd.DataFrame, productivity: dict):
        self.out.mkdir(parents=True, exist_ok=True)

        # CSV 3本
        gap_df.to_csv(self.out / 'staffing_forecast_daily.csv',
                      index=False, encoding='utf-8-sig')
        gap_df.to_csv(self.out / 'staffing_gap_report.csv',
                      index=False, encoding='utf-8-sig')
        suggestions.to_csv(self.out / 'staffing_reallocation_suggestions.csv',
                           index=False, encoding='utf-8-sig')

        # assumptions
        with open(self.out / 'assumptions_report.md', 'w',
                  encoding='utf-8') as f:
            f.write(f'# 仮定一覧（{self.target_month}）\n\n')
            f.write('本プロトタイプで採用した仮定・データ補完の一覧です。\n\n')
            for i, a in enumerate(self.assumptions, 1):
                f.write(f'{i}. {a}\n')

        # 労務テンプレート
        with open(self.out / 'discipline_templates.md', 'w',
                  encoding='utf-8') as f:
            f.write(DISCIPLINE_TEMPLATES.format(target_month=self.target_month))

        # summary
        self._write_summary(gap_df, suggestions)
        print(f'[Reporter] 出力完了: {self.out}')

    # --------------------------- summary ---------------------------
    def _write_summary(self, gap_df: pd.DataFrame, suggestions: pd.DataFrame):
        shortage = gap_df[gap_df['gap'] < 0].copy()
        shortage['gap_abs'] = shortage['gap'].abs()
        top_short = shortage.sort_values(
            'gap_abs', ascending=False).head(10)

        surplus = gap_df[gap_df['gap'] > 0].sort_values(
            'gap', ascending=False).head(10)

        cols_main = ['date', 'area', 'day_of_week', 'weather_forecast',
                     'predicted_visits', 'required_staff',
                     'planned_staff', 'leave_requested',
                     'available_staff', 'gap']
        cols_main = [c for c in cols_main if c in gap_df.columns]

        cols_impact = ['date', 'area', 'gap',
                       'estimated_missed_patients', 'estimated_sales_impact']
        cols_impact = [c for c in cols_impact if c in shortage.columns]

        # 推定影響合計
        total_missed = float(shortage.get(
            'estimated_missed_patients', pd.Series([0])).sum()) if len(shortage) else 0
        total_sales = float(shortage.get(
            'estimated_sales_impact', pd.Series([0])).sum()) if len(shortage) else 0

        # 代替配置の可否
        if len(suggestions) > 0:
            alt_dates = set(suggestions[['date', 'shortage_clinic']]
                            .apply(tuple, axis=1))
        else:
            alt_dates = set()
        shortage['代替配置'] = shortage[['date', 'area']].apply(
            lambda r: '可' if (r['date'], r['area']) in alt_dates else '不可',
            axis=1,
        ) if len(shortage) else pd.Series([], dtype=str)

        with open(self.out / 'summary_report.md', 'w', encoding='utf-8') as f:
            f.write(f'# シフト最適化 月次サマリー：{self.target_month}\n\n')
            f.write(f'生成日時：{datetime.now().strftime("%Y-%m-%d %H:%M")}\n\n')

            # 概要
            f.write('## 概要\n\n')
            f.write(f'- 対象月：**{self.target_month}**\n')
            f.write(f'- 集計対象日数：{gap_df["date"].nunique()} 日\n')
            f.write(f'- 対象院数：{gap_df["area"].nunique()} 院\n')
            f.write(f'- 不足発生：{(gap_df["gap"] < 0).sum()} 行\n')
            f.write(f'- 余剰発生：{(gap_df["gap"] > 0).sum()} 行\n')
            f.write(f'- 適正：{(gap_df["gap"] == 0).sum()} 行\n\n')

            # 予測の前提
            f.write('## 予測の前提\n\n')
            f.write('- 来院予測は (院 × 曜日 × 天気) の過去平均を起点に、サンプル不足時は (院 × 曜日) → (院) → 全院平均 にフォールバック。\n')
            f.write('- 信頼度: 過去サンプル数 ≥ 20 → high, ≥ 10 → medium, それ未満 → low。\n')
            f.write('- 祝日は通常日の **70%** に補正（HOLIDAYS_2026 ハードコード）。\n')
            f.write('- 天気予報は対象月の実データ未取得のため、過去同月のパターンからサンプリング。\n\n')

            # 必要人数算出ロジック
            f.write('## 必要人数算出ロジック\n\n')
            f.write('- `必要人数 = ceil(予測来院数 ÷ 1人あたり対応患者数)`\n')
            f.write('- 最低運営人数を設定（既定 2 人）\n')
            f.write('- 過不足 = `(予定 − 休み希望) − 必要人数`\n\n')

            # 仮定一覧
            f.write('## データ不足による仮定\n\n')
            for a in self.assumptions:
                f.write(f'- {a}\n')
            f.write('\n')

            # 不足院 top10
            f.write('## 不足院トップ10\n\n')
            if len(top_short):
                f.write(df_to_md(top_short[cols_main]))
            else:
                f.write('- なし')
            f.write('\n\n')

            # 余剰院 top10
            f.write('## 余剰院トップ10\n\n')
            if len(surplus):
                f.write(df_to_md(surplus[cols_main]))
            else:
                f.write('- なし')
            f.write('\n\n')

            # 応援移動案
            f.write('## 推奨応援移動案（上位15件）\n\n')
            if len(suggestions):
                f.write(df_to_md(suggestions.head(15)))
            else:
                f.write('- 応援案：該当なし（不足が無い／余剰院が無い）')
            f.write('\n\n')

            # 業務影響
            f.write('## 不足による業務影響（推定）\n\n')
            if len(shortage):
                show = shortage.sort_values('gap').head(15)
                cols_show = cols_impact + (
                    ['代替配置'] if '代替配置' in show.columns else []
                )
                f.write(df_to_md(show[cols_show]))
                f.write('\n\n')
                f.write(f'- 合計推定取りこぼし患者数：**約 {total_missed:.0f} 名**\n')
                f.write(f'- 合計売上影響：**約 ¥{total_sales:,.0f}**（1患者あたり ¥4,000 を仮定）\n')
                f.write('\n（実データに置換することで、実態に近づきます）\n\n')
            else:
                f.write('- 不足発生なし。\n\n')

            # 今後の改善
            f.write('## 今後の改善点\n\n')
            f.write('- 売上データを取り込み、1患者あたり実単価を反映\n')
            f.write('- 職種（柔整師 / 受付 / アルバイト）別の必要人数算出に拡張\n')
            f.write('- スタッフ個別スキル・経験年数を加味した最適化\n')
            f.write('- 天気予報APIとの自動連携（OpenWeather, JMA等）\n')
            f.write('- 機械学習モデル（線形回帰 / 勾配ブースティング）の導入と精度比較\n')
            f.write('- 隣接エリア定義をマスター化（移動時間・運賃込みのコスト最適化）\n')
            f.write('- 勤怠DBと連携した「実績ベース予定」の自動更新\n')
            f.write('- 労務テンプレートの個別生成（CSV取込 → 各人別 md 出力）\n')

    # ----------------------------- console -----------------------------
    def print_console(self, gap_df: pd.DataFrame, suggestions: pd.DataFrame):
        print()
        print('=' * 80)
        print(f'  シフト過不足リスト：{self.target_month}')
        print('=' * 80)

        shortage = gap_df[gap_df['gap'] < 0].copy().sort_values('gap')
        if shortage.empty:
            print('  ✅ 不足の出る日はありません')
        else:
            print(f'  ▼ 不足日リスト（{len(shortage)} 件 / 先頭20件）')
            cols = ['date', 'area', 'day_of_week', 'weather_forecast',
                    'predicted_visits', 'required_staff',
                    'planned_staff', 'leave_requested',
                    'available_staff', 'gap']
            cols = [c for c in cols if c in shortage.columns]
            print(shortage[cols].head(20).to_string(index=False))

        print()
        if len(suggestions):
            print(f'  ▼ 応援移動候補（上位10件）')
            top = suggestions.head(10)
            for _, r in top.iterrows():
                print(f'    {r["date"]} {r["shortage_clinic"]}: '
                      f'{r["shortage_n"]}人不足 → '
                      f'{r["support_clinic"]} ({r["same_area"]}) から '
                      f'{r["recommend_move"]}人応援（推奨度 {r["priority_score"]}）')
        else:
            print('  応援移動候補：なし')
        print('=' * 80)
