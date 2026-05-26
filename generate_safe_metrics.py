# -*- coding: utf-8 -*-
"""ローカルPC専用: 生CSVから機密情報を除いた safe_metrics.json を生成する。

このスクリプトは **誠グループの管理 PC でのみ実行** することを想定している。
出力された `safe_metrics.json` には:
  - 各院の客単価（CRM 売上履歴の平均）
  - スタッフ名簿（所属院別、退職者 池田 は除外）
  - 院ごとの曜日別必要人員（誠グループの運用ルール）
  - 動かせないスタッフ（自院固定）
  - 実習生・新人（ヘルプ除外対象）
  - リスク管理ルール（小金井坂下: 山本休 + 稲田単独 禁止 等）
のみが含まれ、**患者個人情報・売上明細・客様データ等は一切含まない**。

Streamlit Cloud にデプロイされた Web アプリは、この JSON を読み込んで
院長がアップロードしたシフト Excel と照合する形で動作する。

使い方:
    cd C:\\Users\\中村文亮\\Desktop\\crm_scraper
    python generate_safe_metrics.py

出力:
    crm_scraper/safe_metrics.json
"""
from __future__ import annotations
import json
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

# このスクリプトと同じフォルダに置かれている生 CSV を入力とする
SCRIPT_DIR = Path(__file__).resolve().parent
MASTER_CSV = SCRIPT_DIR / "ultimate_shift_master.csv"
CRM_CSV = SCRIPT_DIR / "final_analysis_data.csv"
OUTPUT_JSON = SCRIPT_DIR / "safe_metrics.json"

# =============================================================================
# 誠グループ固定の運用ルール（コード内に明示）
# =============================================================================
# 退職者・除外スタッフ（部分一致でレコードから完全除外）
EXCLUDED_STAFF: list[str] = ["池田"]

# 動かせないスタッフ（自院固定 — 他院ヘルプ対象外）
IMMOVABLE_STAFF: dict[str, list[str]] = {
    "国分寺":    ["熊谷", "黒澤"],
    "武蔵小金井": ["田村", "岡田"],
    "東小金井":   ["伊藤"],
    "小金井坂下": ["山本"],
    "人形町":    ["長島"],
}

# 実習生・新人（ヘルプ対象から除外。自院では勤務可能）
TRAINEE_EXCLUDED: list[str] = ["島田", "清水"]

# 5 院のメタデータ
CLINICS: list[str] = ["国分寺", "武蔵小金井", "東小金井", "小金井坂下", "人形町"]

# 1 日あたりの施術対応枠（機会損失計算に使用）
CAPACITY_PER_DAY: dict[str, int] = {
    "国分寺": 12, "武蔵小金井": 12, "東小金井": 12,
    "小金井坂下": 12, "人形町": 24,  # 人形町は施術時間半分のため 2 倍
}

# 院×曜日 の必要人員（誠グループの運用ルール）
REQUIRED_BY_DOW: dict[str, dict[str, int]] = {
    "国分寺":     {"月": 3, "火": 2, "水": 2, "木": 3, "金": 3, "土": 4, "日": 4},
    "小金井坂下": {"月": 2, "火": 2, "水": 1, "木": 1, "金": 2, "土": 3, "日": 2},
    "東小金井":   {"月": 2, "火": 2, "水": 2, "木": 2, "金": 2, "土": 3, "日": 3},
    "武蔵小金井": {"月": 3, "火": 3, "水": 2, "木": 3, "金": 3, "土": 4, "日": 4},
    "人形町":     {"月": 2, "火": 2, "水": 2, "木": 2, "金": 2, "土": 1, "日": 2},
}

# リスク管理ルール
RISK_RULES: dict[str, dict] = {
    "sakashita_solo_open": {
        "description": "小金井坂下 オープン作業リスク回避",
        "primary_staff": "山本",
        "secondary_staff": "稲田",
        "rule": "primary が休みの日に secondary が単独配置になることを禁止・警告する",
    },
    "saturday_help_shortage": {
        "description": "土曜日 (DOW=5) のヘルプ要請漏れ検知",
        "threshold_available_staff": 3,
        "rule": "土曜に available_staff <= 3 の院があり、かつ他院に余剰がある場合に強い警告",
    },
}

# CSV→院名 標準化マップ（CRM の clinic_name を 5 院の正式名に統一）
_CLINIC_KEYWORD_MAP = [
    ("国分寺",     "国分寺"),
    ("武蔵小金井", "武蔵小金井"),
    ("東小金井",   "東小金井"),
    ("坂下",       "小金井坂下"),
    ("人形町",     "人形町"),
]


def _standardize_clinic_name(name: object) -> str:
    """CRM の clinic_name → 5 院の正式名称に統一する。"""
    s = str(name)
    for kw, std in _CLINIC_KEYWORD_MAP:
        if kw in s:
            return std
    return s


def _parse_money(val) -> float:
    """『10.4万』『104,000』『104000』を 104000.0 に変換。"""
    if pd.isna(val):
        return 0.0
    s = str(val).replace(",", "").strip()
    if "万" in s:
        try:
            return float(s.replace("万", "")) * 10000.0
        except ValueError:
            return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def _is_excluded(name: object) -> bool:
    """退職者・除外スタッフかどうかの部分一致判定。"""
    if name is None or pd.isna(name):
        return False
    s = str(name)
    return any(x and x in s for x in EXCLUDED_STAFF)


# =============================================================================
# 入力 CSV からの集計
# =============================================================================
def calculate_unit_prices(crm_csv: Path) -> dict[str, int]:
    """final_analysis_data.csv から院別の客単価（売上 ÷ 来店数）を算出する。

    戻り値: {院名: 客単価(整数, 円)}
    """
    if not crm_csv.exists():
        print(f"[WARN] {crm_csv.name} が見つかりません → 客単価を ¥6,000 で埋めます")
        return {c: 6000 for c in CLINICS}

    # 文字コード自動判別
    for enc in ("utf-8", "cp932"):
        try:
            df = pd.read_csv(crm_csv, encoding=enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise RuntimeError(f"{crm_csv} の読込に失敗 (utf-8 / cp932 ともに不可)")

    needed = ["clinic_name", "総合売上", "来店数"]
    if any(c not in df.columns for c in needed):
        missing = [c for c in needed if c not in df.columns]
        raise ValueError(f"CRM CSV に必要列がありません: {missing}")

    df["clinic_std"] = df["clinic_name"].map(_standardize_clinic_name)
    df["sales"] = df["総合売上"].map(_parse_money)
    df["visits"] = pd.to_numeric(df["来店数"], errors="coerce").fillna(0)

    grouped = df.groupby("clinic_std")[["sales", "visits"]].sum()
    grouped["unit_price"] = grouped["sales"] / grouped["visits"].replace(0, pd.NA)
    out: dict[str, int] = {}
    for clinic in CLINICS:
        if clinic in grouped.index and not pd.isna(grouped.loc[clinic, "unit_price"]):
            out[clinic] = int(round(grouped.loc[clinic, "unit_price"]))
        else:
            out[clinic] = 6000  # フォールバック
    return out


def extract_staff_roster(master_csv: Path) -> dict[str, list[str]]:
    """ultimate_shift_master.csv からスタッフ名と所属院を抽出する。

    退職者 池田 は完全除外。
    戻り値: {院名: [スタッフ名, ...]}
    """
    if not master_csv.exists():
        print(f"[WARN] {master_csv.name} が見つかりません → 空の名簿を生成します")
        return {c: [] for c in CLINICS}

    df = pd.read_csv(master_csv, encoding="utf-8-sig")

    # スタッフ名列を推定（master CSV は集計形式なのでスタッフ列が無い場合あり）
    name_cols = [c for c in df.columns if c in ("staff_name", "スタッフ名", "氏名")]
    area_cols = [c for c in df.columns if c in ("area", "院", "院名", "エリア")]

    rosters: dict[str, list[str]] = {c: [] for c in CLINICS}

    if not name_cols or not area_cols:
        print(
            f"[INFO] master_csv にスタッフ名列が見つからないため、"
            f"名簿は xlsx 解析時に動的構築されます"
        )
        return rosters

    name_col = name_cols[0]
    area_col = area_cols[0]

    sub = df[[name_col, area_col]].dropna()
    sub = sub[~sub[name_col].map(_is_excluded)]
    sub[area_col] = sub[area_col].map(_standardize_clinic_name)
    for clinic, grp in sub.groupby(area_col):
        if clinic in rosters:
            rosters[clinic] = sorted(set(grp[name_col].astype(str)))
    return rosters


# =============================================================================
# メイン
# =============================================================================
def build_safe_metrics() -> dict:
    """各種データを集約して safe_metrics 辞書を構築する。"""
    print(f"[generate_safe_metrics] 入力: {MASTER_CSV.name}, {CRM_CSV.name}")

    unit_prices = calculate_unit_prices(CRM_CSV)
    print(f"  ✓ 客単価算出: {unit_prices}")

    staff_roster = extract_staff_roster(MASTER_CSV)
    n_staff = sum(len(v) for v in staff_roster.values())
    print(f"  ✓ 名簿抽出: 合計 {n_staff} 名 (池田は完全除外)")

    return {
        "version": "1.0",
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "source_files": {
            "master_csv": MASTER_CSV.name,
            "crm_csv": CRM_CSV.name,
        },
        "clinics": CLINICS,
        "capacity_per_day": CAPACITY_PER_DAY,
        "unit_prices": unit_prices,
        "required_by_dow": REQUIRED_BY_DOW,
        "staff_roster": staff_roster,
        "immovable_staff": IMMOVABLE_STAFF,
        "trainee_excluded": TRAINEE_EXCLUDED,
        "excluded_staff": EXCLUDED_STAFF,
        "risk_rules": RISK_RULES,
        "notes": (
            "このファイルはローカル PC で generate_safe_metrics.py を実行して"
            "生成された安全な集計データです。患者個人情報・売上明細は含まれません。"
            "Streamlit Cloud にデプロイされた Web アプリが、院長がアップロード"
            "したシフト Excel と照合する形で使用します。"
        ),
    }


def main() -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    print("=" * 70)
    print(" generate_safe_metrics.py — 安全な集計 JSON を生成します")
    print("=" * 70)

    if not MASTER_CSV.exists() and not CRM_CSV.exists():
        print(
            f"⚠ 入力 CSV が両方とも見つかりません:\n"
            f"  - {MASTER_CSV}\n  - {CRM_CSV}\n"
            f"スクリプトと同じフォルダに上記 2 ファイルを配置してから再実行してください。"
        )
        return 1

    metrics = build_safe_metrics()
    OUTPUT_JSON.write_text(
        json.dumps(metrics, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print()
    print(f"✓ 出力: {OUTPUT_JSON}")
    print(f"  ファイルサイズ: {OUTPUT_JSON.stat().st_size:,} bytes")
    print()
    print("次のステップ:")
    print("  1) safe_metrics.json の中身を確認（個人情報が含まれていないこと）")
    print("  2) git add safe_metrics.json && git commit && git push")
    print("  3) Streamlit Cloud が自動再デプロイ → 院長アクセス可能に")
    return 0


if __name__ == "__main__":
    sys.exit(main())
