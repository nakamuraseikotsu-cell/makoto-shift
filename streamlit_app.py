# -*- coding: utf-8 -*-
"""シフト最適化 Web パネル（Streamlit Cloud デプロイ版）

主な特徴:
  - サイドバーで対象月を選択
  - 画面上から **シフト表 Excel / マスター CSV / CRM CSV** を直接アップロード
    （ローカルパス・Google Drive 等への依存はゼロ）
  - アップロードされたファイルはサーバ側の **セッション専用 temp dir** にのみ
    保存され、元のローカルファイルには一切影響しない
  - 結果は **Excel (.xlsx, 複数シート) または CSV** でダウンロード可能
  - 設定変更後の「再計算」も、アップロード済データを保持したまま実行可能
"""
from __future__ import annotations
import calendar
import io
import json
import os
import re
import shutil
import sys
import tempfile
from datetime import datetime, date as _date
from pathlib import Path

import openpyxl
import pandas as pd
import streamlit as st

# shift_optimizer を import path に追加（crm_scraper/ を sys.path に）
_SCRIPT_DIR = Path(__file__).resolve().parent
SHIFT_OPTIMIZER_DIR = _SCRIPT_DIR / "shift_optimizer"
if str(_SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPT_DIR))

from shift_optimizer.main import (  # noqa: E402
    run_full_analysis, RealDataMissingError,
)
from shift_optimizer.src.constants import IMMOVABLE_STAFF  # noqa: E402


# ==========================================
# 設定
# ==========================================
# デフォルトデータディレクトリ
# -----------------------------------------------------------------------------
# スクリプトと同じフォルダに以下のファイル名で配置されていれば、
# アプリ起動時に自動でセッション temp dir へコピーされて読み込まれる。
# アップロードがあればそちらが優先（同名ファイルは上書きされる）。
#   - ultimate_shift_master.csv  （統合マスタ）
#   - final_analysis_data.csv    （CRM データ）
#   - 【○○】*.xlsx               （5院ぶんのシフト表）
# -----------------------------------------------------------------------------
DEFAULT_DATA_DIR = _SCRIPT_DIR
DEFAULT_MASTER_CSV = 'ultimate_shift_master.csv'
DEFAULT_CRM_CSV = 'final_analysis_data.csv'

# 1日あたりの施術対応枠（人）
CAPACITY_PER_DAY = {
    "国分寺": 12,
    "武蔵小金井": 12,
    "東小金井": 12,
    "小金井坂下": 12,
    "人形町": 24,   # 施術時間半分のため2倍
}

# 院名 → シフト表ファイル名のプレフィックス
CLINIC_FILE_PREFIX = {
    "国分寺": "【国分寺】",
    "武蔵小金井": "【武蔵小金井】",
    "東小金井": "【東小金井】",
    "小金井坂下": "【坂下】",
    "人形町": "【人形町】",
}

# ==========================================
# session_state キー登録（クリア漏れ防止）
# 分析結果やそれに関連する全てのキーをここに列挙すること。
# 新しい state key を導入したら必ず本リストに追加すること。
# ==========================================
_ANALYSIS_STATE_KEYS: tuple[str, ...] = (
    'analysis',              # 現行のメイン分析結果
    '_pending_reanalysis',   # 再分析予約フラグ
    # 互換性 / 将来用に予約しているキー（存在しなくても pop は無害）
    'analysis_results',
    'df_summary',
    'shortages',
    'surpluses',
    'unit_prices',
    'staff_help_actions',
    'gap_df',
    'paid_leave_df',
    'help_actions_actual_df',
    'fixed_leave_df',
    'worked_df',
)

# 上記以外でも、これらのプレフィックスで始まるキーは分析関連とみなして消す
_ANALYSIS_STATE_PREFIXES: tuple[str, ...] = (
    'analysis_',
    'df_',
    'shortage_',
)


def _reset_analysis_state() -> list[str]:
    """分析結果に関連する session_state を「完全に」初期化。

    Returns: 消した key のリスト（toast/ログ表示用）
    """
    cleared: list[str] = []
    # 既知キーの明示削除
    for key in _ANALYSIS_STATE_KEYS:
        if key in st.session_state:
            del st.session_state[key]
            cleared.append(key)
    # プレフィックス一致の動的削除
    for key in list(st.session_state.keys()):
        if any(key.startswith(p) for p in _ANALYSIS_STATE_PREFIXES):
            if key not in cleared:
                del st.session_state[key]
                cleared.append(key)
    return cleared


def _clear_streamlit_caches() -> None:
    """Streamlit の cache_data / cache_resource を防御的にクリア。
    本アプリでは @st.cache_data は使っていないが、サードパーティ
    ライブラリが内部で使うケースに備えて毎回クリアする。
    """
    try:
        st.cache_data.clear()
    except Exception:
        pass
    try:
        st.cache_resource.clear()
    except Exception:
        pass

# ==========================================
# ページ設定
# ==========================================
st.set_page_config(
    page_title="シフト最適化 Web パネル",
    page_icon="🩺",
    layout="wide",
    # スマホ/タブレットでは初期サイドバーは閉じておく
    # （narrow 幅で全体が見やすくなる。デスクトップはユーザが開けば良い）
    initial_sidebar_state="auto",
)

# -------------------------------------------------------------------
# モバイル / タブレット向け CSS
# -------------------------------------------------------------------
# - viewport meta は Streamlit が自動付与するため、ここでは
#   狭幅レイアウトでのフォントサイズ・余白を最適化するだけに留める。
# - 主要なリセット:
#     * h1〜h3 のフォントを clamp で可変化
#     * メインブロックの上下パディングを縮小
#     * st.columns 内の要素が画面幅を溢れた場合に折り返す
#     * DataFrame・テーブルが画面幅で横スクロール可能
# -------------------------------------------------------------------
st.markdown(
    """
    <style>
    /* === グローバル: モバイル/タブレット最適化 === */
    .main .block-container {
        padding-top: 1.2rem;
        padding-left: clamp(0.6rem, 2vw, 2rem);
        padding-right: clamp(0.6rem, 2vw, 2rem);
        max-width: 100%;
    }
    h1, h2, h3 {
        word-break: keep-all;
        overflow-wrap: anywhere;
    }
    h1 { font-size: clamp(1.4rem, 4vw, 2.2rem) !important; }
    h2 { font-size: clamp(1.2rem, 3vw, 1.8rem) !important; }
    h3 { font-size: clamp(1.0rem, 2.5vw, 1.4rem) !important; }
    /* DataFrame は横スクロール可、画面幅にフィット */
    div[data-testid="stDataFrame"] { width: 100% !important; }
    div[data-testid="stDataFrame"] > div { overflow-x: auto; }
    /* メトリック・カード類が縦に並ぶときの間隔を詰める */
    div[data-testid="stMetric"] { padding: 0.4rem 0.6rem; }
    /* マルチセレクトのタグが折り返す */
    div[data-baseweb="select"] span { white-space: normal !important; }

    @media (max-width: 768px) {
        /* スマホ: カラムを縦積みにし、ボタンを大きく */
        button[kind="primary"], button[kind="secondary"] {
            font-size: 0.95rem !important;
            padding: 0.6rem 0.8rem !important;
        }
        /* expander タイトルを少し大きく（タップしやすく） */
        details summary { font-size: 1.0rem !important; padding: 0.6rem !important; }
        /* サイドバーを開いたときの幅を狭く */
        section[data-testid="stSidebar"] { min-width: 80vw !important; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ==========================================
# 🔒 パスワードロック（テスト公開時の社外閲覧防止）
# ------------------------------------------------------------
# パスワードは次の優先順で取得する:
#   1) st.secrets["APP_PASSWORD"]   (推奨: Streamlit Cloud のシークレット)
#   2) 環境変数 APP_PASSWORD
#   3) 開発用デフォルト  "makoto2026"
# 注意: GitHub に公開するコードのため、上記 3) は本番運用では
#       必ず Streamlit Cloud 側で 1) を設定して上書きすること。
# ==========================================
_DEFAULT_PASSWORD = "makoto2026"  # 開発用 (本番は st.secrets で上書き)


def _resolve_app_password() -> str:
    """st.secrets → 環境変数 → デフォルト の順でパスワードを解決する。"""
    try:
        pw = st.secrets["APP_PASSWORD"]
        if pw:
            return str(pw)
    except (KeyError, FileNotFoundError, Exception):
        pass
    env_pw = os.environ.get("APP_PASSWORD")
    if env_pw:
        return env_pw
    return _DEFAULT_PASSWORD


def _require_login() -> bool:
    """簡易ログインゲート。認証済みなら True を返し、本体スクリプトの実行に進む。
    未認証ならロック画面を描画して False を返す（呼び出し側で st.stop()）。
    """
    if st.session_state.get("_authenticated", False):
        return True

    expected = _resolve_app_password()

    # ロック画面 UI（スマホでもセンタリングして見やすく）
    st.markdown(
        """
        <div style="max-width:460px; margin:8vh auto 0 auto; padding:24px;
                    background:linear-gradient(135deg,#1e3a8a 0%,#7c3aed 100%);
                    border-radius:16px; color:white;
                    box-shadow:0 10px 30px rgba(15,23,42,0.25);">
            <div style="font-size:clamp(1.4rem, 4vw, 1.9rem);
                        font-weight:bold; text-align:center;">
                🔒 シフト最適化 Web パネル
            </div>
            <div style="text-align:center; opacity:0.85; margin-top:6px;
                        font-size:clamp(0.85rem, 2vw, 1.0rem);">
                テスト公開中 — 関係者のみ閲覧可能
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.write("")  # 余白
    with st.form("_login_form", clear_on_submit=False):
        pw = st.text_input(
            "パスワード",
            type="password",
            placeholder="共有されたパスワードを入力してください",
            help="運用担当者から共有されたパスワードを入力してください",
        )
        submitted = st.form_submit_button(
            "🔓 ログイン", type="primary", use_container_width=True,
        )

    if submitted:
        if pw == expected:
            st.session_state["_authenticated"] = True
            st.toast("✅ ログインしました", icon="🔓")
            st.rerun()
        else:
            st.error("❌ パスワードが違います。再度入力してください。")
    st.caption(
        "※ パスワードをお忘れの場合は、運用担当者までお問い合わせください。"
    )
    return False


if not _require_login():
    st.stop()


# ==========================================
# ヘルパー
# ==========================================
def parse_money(val):
    """『10.4万』『104,000』『104000』を 104000 に変換"""
    if pd.isna(val):
        return 0
    val_str = str(val).replace(',', '')
    if '万' in val_str:
        try:
            return float(val_str.replace('万', '')) * 10000
        except Exception:
            return 0
    try:
        return float(val_str)
    except Exception:
        return 0


# 院名標準化（CRM の clinic_name → 5院の正式名称）
_CLINIC_KEYWORD_MAP = (
    ('国分寺', '国分寺'),
    ('武蔵小金井', '武蔵小金井'),
    ('東小金井', '東小金井'),
    ('坂下', '小金井坂下'),
    ('人形町', '人形町'),
)


def _standardize_clinic_name(name: object) -> str:
    s = str(name)
    for kw, std in _CLINIC_KEYWORD_MAP:
        if kw in s:
            return std
    return s


# ------- mtime ベースのキャッシュ層 -------
# `@st.cache_data` は引数のハッシュでキャッシュキーを生成するため、ファイル
# のパスだけでなく **mtime** も引数に含めることで、xlsx/CSV が編集された
# 瞬間に自動でキャッシュを無効化する。
@st.cache_data(show_spinner=False)
def _cached_unit_prices(crm_path: str, mtime: float) -> tuple[dict | None, str | None]:
    """CRM CSV を読み込み、院別客単価を算出してキャッシュ。
    mtime は cache key 用に渡すだけ（関数本体では未使用）。"""
    del mtime  # unused: cache key only
    try:
        df = pd.read_csv(crm_path, encoding="utf-8")
    except Exception:
        try:
            df = pd.read_csv(crm_path, encoding="cp932")
        except Exception as e:
            return None, f"CRMデータの読み込みに失敗しました: {e}"

    needed = ["clinic_name", "総合売上", "来店数"]
    missing_cols = [c for c in needed if c not in df.columns]
    if missing_cols:
        return None, f"CRMデータに以下の列が見つかりません: {missing_cols}"

    df["標準店舗名"] = df["clinic_name"].map(_standardize_clinic_name)
    df["売上数値"] = df["総合売上"].map(parse_money)
    df["来店数値"] = pd.to_numeric(df["来店数"], errors='coerce').fillna(0)
    summary = df.groupby("標準店舗名")[["売上数値", "来店数値"]].sum()
    summary["客単価"] = summary["売上数値"] / summary["来店数値"]
    return summary["客単価"].to_dict(), None


def calculate_unit_prices(crm_path):
    """CRMデータから院別の客単価を算出。
    `crm_path` の mtime をキーにキャッシュ済の結果を返すため、ファイル変更時
    だけ再読み込みが走る（重い CSV groupby を毎回回避）。"""
    if not os.path.exists(crm_path):
        return None, f"CRMデータが見つかりません: {crm_path}"
    return _cached_unit_prices(crm_path, os.path.getmtime(crm_path))


# =============================================================================
# セッション専用 temp dir 管理（アップロードファイルの一時保存先）
# -----------------------------------------------------------------------------
# Streamlit Cloud では各ユーザのブラウザセッション毎に独立した temp dir を
# 生成する。アップロードされたシフト xlsx / マスター CSV / CRM CSV はこの
# ディレクトリにのみ書き出され、元のローカルファイルは一切変更されない。
# セッション終了時（ブラウザを閉じる等）に OS が自動的にクリーンアップする。
# =============================================================================
def get_data_folder() -> str:
    """セッション専用 temp dir のパスを返す（無ければ作成）。"""
    if '_upload_tempdir' not in st.session_state:
        st.session_state['_upload_tempdir'] = tempfile.mkdtemp(
            prefix='shift_session_'
        )
    return st.session_state['_upload_tempdir']


def reset_session_tempdir() -> None:
    """アップロード済みファイルを全削除し、temp dir を破棄→再作成。"""
    old = st.session_state.pop('_upload_tempdir', None)
    if old and os.path.isdir(old):
        try:
            shutil.rmtree(old, ignore_errors=True)
        except Exception:
            pass
    st.session_state.pop('_written_file_ids', None)


def find_shift_xlsx(clinic_name: str) -> Path | None:
    """院名からアップロード済シフト xlsx を検索する。"""
    prefix = CLINIC_FILE_PREFIX.get(clinic_name)
    if not prefix:
        return None
    folder = Path(get_data_folder())
    if not folder.exists():
        return None
    candidates = sorted(folder.glob(f'{prefix}*.xlsx'))
    for p in candidates:
        if 'シフト' in p.name:
            return p
    return candidates[0] if candidates else None


def collect_shift_xlsx_mtimes(folder: str | None = None) -> dict:
    """各院シフト xlsx の最終更新時刻（表示用文字列）。"""
    del folder  # find_shift_xlsx が session temp dir を内部解決する
    out = {}
    for clinic in CLINIC_FILE_PREFIX:
        xlsx = find_shift_xlsx(clinic)
        if xlsx and xlsx.exists():
            try:
                out[clinic] = datetime.fromtimestamp(
                    os.path.getmtime(xlsx)
                ).strftime("%Y-%m-%d %H:%M")
            except Exception:
                out[clinic] = "-"
        else:
            out[clinic] = "-"
    return out


def _shift_xlsx_mtimes_tuple() -> tuple:
    """キャッシュキー用: (clinic, mtime_float) のソート済みタプル。"""
    items = []
    for clinic in CLINIC_FILE_PREFIX:
        xlsx = find_shift_xlsx(clinic)
        if xlsx and xlsx.exists():
            try:
                items.append((clinic, os.path.getmtime(xlsx)))
            except Exception:
                items.append((clinic, -1.0))
        else:
            items.append((clinic, -1.0))
    return tuple(items)


# =============================================================================
# アップロードファイルを temp dir に同期するユーティリティ
# -----------------------------------------------------------------------------
# Streamlit の UploadedFile オブジェクトは session_state にまたがって永続化
# されるが、`run_full_analysis` は **ディスク上の xlsx パス** を必要とする。
# そこで temp dir に書き出しておく。`file_id` を追跡して、同じファイルなら
# 二度目以降は書き直さない（無駄な I/O を避ける）。
# =============================================================================
def _file_id_of(uploaded) -> tuple:
    """UploadedFile を一意に識別するキー（書き直し回避用）。"""
    return (
        getattr(uploaded, 'file_id', None) or uploaded.name,
        uploaded.name,
        uploaded.size,
    )


def sync_uploads_to_tempdir(
    shift_uploads: list | None,
    master_upload,
    crm_upload,
) -> dict:
    """アップロードを temp dir に同期する。既に書き出し済みならスキップ。

    アップロードされたファイルが同じ院のデフォルトファイル（事前に seed 済み）
    を上書きできるよう、対象院の既存 xlsx は削除してから新版を書き出す。
    """
    tmpdir = Path(get_data_folder())
    tmpdir.mkdir(parents=True, exist_ok=True)
    written_ids = st.session_state.setdefault('_written_file_ids', set())
    overridden = st.session_state.setdefault('_default_overridden', {})

    status: dict = {'shifts': [], 'master': None, 'crm': None}

    from shift_optimizer.src.real_data import detect_area_from_filename

    for uf in (shift_uploads or []):
        fid = _file_id_of(uf)
        area = detect_area_from_filename(uf.name)
        # 上書き: 同院の既存 xlsx（デフォルト seed や前回アップロード）を削除
        if area:
            prefix = CLINIC_FILE_PREFIX.get(area)
            if prefix:
                for existing in tmpdir.glob(f'{prefix}*.xlsx'):
                    if existing.name != uf.name:
                        try:
                            existing.unlink()
                        except Exception:
                            pass
            overridden[area] = True
        target = tmpdir / uf.name
        if fid not in written_ids or not target.exists():
            target.write_bytes(uf.getbuffer())
            written_ids.add(fid)
        status['shifts'].append({
            'name': uf.name, 'area': area,
            'size': uf.size, 'path': str(target),
        })

    if master_upload is not None:
        fid = _file_id_of(master_upload)
        target = tmpdir / DEFAULT_MASTER_CSV
        if fid not in written_ids or not target.exists():
            target.write_bytes(master_upload.getbuffer())
            written_ids.add(fid)
        overridden[DEFAULT_MASTER_CSV] = True
        status['master'] = {'name': master_upload.name,
                            'size': master_upload.size,
                            'path': str(target)}

    if crm_upload is not None:
        fid = _file_id_of(crm_upload)
        target = tmpdir / DEFAULT_CRM_CSV
        if fid not in written_ids or not target.exists():
            target.write_bytes(crm_upload.getbuffer())
            written_ids.add(fid)
        overridden[DEFAULT_CRM_CSV] = True
        status['crm'] = {'name': crm_upload.name,
                         'size': crm_upload.size,
                         'path': str(target)}

    return status


# =============================================================================
# デフォルトデータの自動 seed
# -----------------------------------------------------------------------------
# スクリプトと同じフォルダ (DEFAULT_DATA_DIR) に配置された固定ファイルを
# セッション temp dir へコピーする。既にアップロード版がある場合は触らない。
# 一度 seed したら session_state にフラグを立て、同セッション内で再 seed しない。
# =============================================================================
def seed_defaults_if_present() -> dict:
    """DEFAULT_DATA_DIR からデフォルトファイルを temp dir へコピーする。

    戻り値: {
        'master': 'default' | 'uploaded' | 'absent',
        'crm':    'default' | 'uploaded' | 'absent',
        'shifts': {clinic: 'default' | 'uploaded' | 'absent', ...},
        'seeded': bool,  # 今回 seed を実行したかどうか
    }
    """
    tmpdir = Path(get_data_folder())
    tmpdir.mkdir(parents=True, exist_ok=True)

    # 同セッションで2回目以降は seed をスキップ（ステータスのみ返す）
    already = st.session_state.get('_defaults_seeded', False)

    status: dict = {'master': 'absent', 'crm': 'absent',
                    'shifts': {c: 'absent' for c in CLINIC_FILE_PREFIX},
                    'seeded': False}

    def _seed_one(default_name: str, target_name: str | None = None) -> str:
        """1ファイルだけseed。戻り値 = 'default'/'uploaded'/'absent'"""
        target_name = target_name or default_name
        src = DEFAULT_DATA_DIR / default_name
        dst = tmpdir / target_name
        if dst.exists() and st.session_state.get('_default_overridden', {}).get(target_name):
            return 'uploaded'
        if dst.exists() and not src.exists():
            # アップロードされた版（seed ソースが無いので uploaded 判定）
            return 'uploaded'
        if dst.exists() and src.exists() and not already:
            # 既に seed 済みの可能性。mtime で確認
            return 'default'
        if not src.exists():
            return 'absent'
        if not dst.exists():
            shutil.copy2(src, dst)
            status['seeded'] = True
            return 'default'
        return 'default'

    # マスタCSV / CRMの seed
    status['master'] = _seed_one(DEFAULT_MASTER_CSV)
    status['crm'] = _seed_one(DEFAULT_CRM_CSV)

    # シフト xlsx の seed (各院ごとにプレフィックス検索)
    for clinic, prefix in CLINIC_FILE_PREFIX.items():
        existing_in_tmp = sorted(tmpdir.glob(f'{prefix}*.xlsx'))
        if existing_in_tmp:
            # 既にあるならアップロード版か seed 済みか
            # _default_overridden フラグでアップロードか判定（簡易）
            overridden = st.session_state.get('_default_overridden', {}).get(clinic)
            status['shifts'][clinic] = 'uploaded' if overridden else 'default'
            continue
        # デフォルトを探す
        defaults = sorted(DEFAULT_DATA_DIR.glob(f'{prefix}*.xlsx'))
        if not defaults:
            status['shifts'][clinic] = 'absent'
            continue
        # 「シフト」を含むものを優先
        pick = next((p for p in defaults if 'シフト' in p.name), defaults[0])
        try:
            shutil.copy2(pick, tmpdir / pick.name)
            status['shifts'][clinic] = 'default'
            status['seeded'] = True
        except Exception:
            status['shifts'][clinic] = 'absent'

    st.session_state['_defaults_seeded'] = True
    return status


def _src_badge(src: str) -> str:
    """読み込み元 (default/uploaded/absent) を絵文字バッジで表示する。"""
    return {
        'default': '🟢 デフォルト',
        'uploaded': '🔵 アップロード上書き',
        'absent': '⬜ 未読み込み',
    }.get(src, '⬜ 未読み込み')


# =============================================================================
# safe_metrics.json ローダー
# -----------------------------------------------------------------------------
# generate_safe_metrics.py で生成された安全な集計データを読み込む。
# CRM 売上CSV や統合マスタCSV を Web 上で扱う必要が無くなるため、Streamlit
# Cloud にも安心してデプロイできる。
# =============================================================================
SAFE_METRICS_PATH = _SCRIPT_DIR / 'safe_metrics.json'


@st.cache_data(show_spinner=False)
def _load_safe_metrics_cached(path_str: str, mtime: float) -> dict:
    del mtime
    with open(path_str, 'r', encoding='utf-8') as f:
        return json.load(f)


def load_safe_metrics() -> dict | None:
    """safe_metrics.json を mtime キーでキャッシュ読込する。
    無い場合は None を返す（古い CSV パイプラインへフォールバック）。"""
    if not SAFE_METRICS_PATH.exists():
        return None
    try:
        return _load_safe_metrics_cached(
            str(SAFE_METRICS_PATH), os.path.getmtime(SAFE_METRICS_PATH)
        )
    except Exception as e:
        st.warning(f"safe_metrics.json の読込に失敗: {e}")
        return None


# =============================================================================
# 院長がアップロードした xlsx のプリンタフリーズ対策
# -----------------------------------------------------------------------------
# Excel ファイルが「ページレイアウト」「改ページプレビュー」モードで保存されて
# いると、利用者の PC でファイルを開いた瞬間にデフォルトプリンタへ問い合わせ
# が走り、UI がフリーズすることがある。アップロードされた xlsx について、
# load 直前に **全シートの sheet_view を 'normal' に強制** することで完全防止。
# =============================================================================
def force_normal_view_on_tempfile(xlsx_path: Path) -> int:
    """xlsx ファイルの全シートを「標準ビュー」に強制する。
    Returns: 書き換えたシート数。失敗時は 0 (例外を握りつぶす)。"""
    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=False)
    except Exception:
        return 0
    changed = 0
    try:
        for ws in wb.worksheets:
            try:
                sv = ws.sheet_view
                if getattr(sv, 'view', None) != 'normal':
                    sv.view = 'normal'
                    changed += 1
                # zoom 系も初期化 (古いビュー設定の残骸を除去)
                if hasattr(sv, 'zoomScalePageLayoutView'):
                    sv.zoomScalePageLayoutView = None
                if hasattr(sv, 'zoomScaleSheetLayoutView'):
                    sv.zoomScaleSheetLayoutView = None
            except Exception:
                pass
        if changed:
            wb.save(xlsx_path)
    except Exception:
        pass
    finally:
        try:
            wb.close()
        except Exception:
            pass
    return changed


# =============================================================================
# light_analyze — xlsx + safe_metrics.json のみで動作する軽量分析
# -----------------------------------------------------------------------------
# 旧 run_full_analysis は ultimate_shift_master.csv を必要としたが、
# 本関数は **safe_metrics.json から必要人員 / 客単価を取得** することで
# CSV 依存をゼロにする。アップロードされた xlsx だけで完結する。
# =============================================================================
def _build_required_df_from_metrics(
    target_month: str, safe_metrics: dict,
) -> pd.DataFrame:
    """safe_metrics の required_by_dow から、対象月の日別×院別 required_df を生成。"""
    year, month = map(int, target_month.split('-'))
    n_days = calendar.monthrange(year, month)[1]
    wd_jp = '月火水木金土日'
    rows = []
    req_map = safe_metrics.get('required_by_dow', {})
    for day in range(1, n_days + 1):
        d = datetime(year, month, day)
        dow_jp = wd_jp[d.weekday()]
        date_str = d.date().isoformat()
        for clinic, dow_map in req_map.items():
            rows.append({
                'date': date_str,
                'area': clinic,
                'day_of_week': dow_jp,
                'required_staff': int(dow_map.get(dow_jp, 2)),
                'predicted_visits': 0,
                'productivity_used': 1.0,
                'productivity_level': 'safe_metrics_json',
            })
    return pd.DataFrame(rows)


def light_analyze(
    target_month: str,
    folder: str,
    safe_metrics: dict,
    movable_whitelist: dict | None = None,
) -> dict:
    """safe_metrics.json + 院長アップロードの xlsx だけで動作する分析。

    旧 run_full_analysis 同等の戻り値構造を返すが、master CSV / CRM CSV を
    一切要求しない。
    """
    from shift_optimizer.src.real_data import (
        load_real_shift_data, find_shift_files,
    )
    from shift_optimizer.src.staffing_calculator import StaffingCalculator
    from shift_optimizer.src.reallocator import Reallocator

    folder_p = Path(folder)

    # ★ プリンタフリーズ対策: アップロードされた xlsx に sheet_view=normal を強制
    sanitized = 0
    for xlsx_path in find_shift_files(folder_p):
        sanitized += force_normal_view_on_tempfile(xlsx_path)

    # 1) xlsx から実データ抽出
    shift_data = load_real_shift_data(folder_p, target_month)
    planned = shift_data['planned']
    leave = shift_data['leave']
    worked = shift_data['worked']
    paid_leave = shift_data.get('paid_leave', pd.DataFrame())
    help_actions = shift_data.get('help_actions', pd.DataFrame())
    fixed_leave = shift_data.get('fixed_leave', pd.DataFrame())

    if planned.empty:
        return {
            'error': (
                f"{target_month} のシートが各シフト xlsx に見つかりませんでした。"
                f"対象月のシートが作成されているか確認してください。"
            ),
            'missing_areas': shift_data.get('missing_areas', []),
        }

    # 2) safe_metrics から required_df を生成（master CSV 不要）
    required_df = _build_required_df_from_metrics(target_month, safe_metrics)

    # 3) 過不足計算
    calc = StaffingCalculator(min_staff=2, assumptions=[])
    gap_df = calc.calculate_gap(required_df, planned, leave)

    # 4) 応援アクション（trainee 除外を whitelist で実現）
    clinic_master = pd.DataFrame([
        {'院名': c,
         'エリア': '中央線' if c != '人形町' else '都心'}
        for c in safe_metrics['clinics']
    ])

    # trainee_excluded を movable_whitelist から自動的に外す
    trainee_names = set(safe_metrics.get('trainee_excluded', []))
    if movable_whitelist and trainee_names:
        movable_whitelist = {
            clinic: [n for n in names
                     if not any(t in str(n) for t in trainee_names)]
            for clinic, names in movable_whitelist.items()
        }

    reallocator = Reallocator(clinic_master, [])
    staff_help_actions = reallocator.suggest_staff_help_actions(
        gap_df, worked,
        fixed_leave_df=fixed_leave,
        movable_whitelist=movable_whitelist,
    )

    return {
        'target_month': target_month,
        'gap_df': gap_df,
        'staff_help_actions': staff_help_actions,
        'worked_df': worked,
        'planned_shifts': planned,
        'leave_requests': leave,
        'paid_leave_df': paid_leave,
        'help_actions_actual_df': help_actions,
        'fixed_leave_df': fixed_leave,
        'required_df': required_df,
        'unit_prices': safe_metrics.get('unit_prices', {}),
        'missing_areas': shift_data.get('missing_areas', []),
        'assumptions': [
            f'必要人員: safe_metrics.json の required_by_dow を使用',
            f'客単価: safe_metrics.json の unit_prices を使用 (CRM CSV不要)',
            f'プリンタフリーズ対策: {sanitized} シートを normal ビューへ変換',
        ],
    }


def _whitelist_to_key(wl: dict | None) -> tuple | None:
    """movable_whitelist を hashable な形に正規化（キャッシュキー用）。"""
    if not wl:
        return None
    return tuple(sorted(
        (k, tuple(sorted(v))) for k, v in wl.items() if isinstance(v, (list, tuple, set))
    ))


@st.cache_data(show_spinner=False)
def _cached_compute_real_shortages(
    target_month: str,
    folder: str,
    master_mtime: float,
    xlsx_mtimes_key: tuple,
    whitelist_key: tuple | None,
    safe_metrics_mtime: float,
):
    """compute_real_shortages の重い部分をキャッシュ。
    mtime / whitelist / safe_metrics の変化でキャッシュ無効化。

    実行モード:
      - safe_metrics.json があれば → light_analyze (master CSV 不要)
      - safe_metrics.json が無く master CSV があれば → 旧 run_full_analysis
    """
    del master_mtime, xlsx_mtimes_key, safe_metrics_mtime  # cache key only

    # キャッシュキーから movable_whitelist を再構築
    movable_whitelist: dict | None = None
    if whitelist_key:
        movable_whitelist = {k: list(v) for k, v in whitelist_key}

    safe_metrics = load_safe_metrics()
    if safe_metrics is not None:
        # ★ ハイブリッド方式: safe_metrics.json + xlsx で完結
        result = light_analyze(
            target_month=target_month,
            folder=folder,
            safe_metrics=safe_metrics,
            movable_whitelist=movable_whitelist,
        )
        return result

    # 旧パイプライン (master CSV を要求する)
    master_csv = Path(folder) / "ultimate_shift_master.csv"
    result = run_full_analysis(
        target_month=target_month,
        min_staff=2,
        base_dir=SHIFT_OPTIMIZER_DIR,
        master_csv=master_csv,
        real_data_folder=folder,
        movable_whitelist=movable_whitelist,
    )
    return result


def compute_real_shortages(target_month: str, folder: str,
                            movable_whitelist: dict | None = None):
    """xlsx (+ safe_metrics.json) から院別「不足人日」と応援アクションを計算。

    Returns: (shortages_dict, gap_df, staff_help_actions_df,
              missing_areas, error_msg, extras)
    """
    safe_metrics = load_safe_metrics()
    safe_metrics_mtime = (
        os.path.getmtime(SAFE_METRICS_PATH) if SAFE_METRICS_PATH.exists() else -1.0
    )
    master_csv = Path(folder) / "ultimate_shift_master.csv"

    # safe_metrics.json も master CSV も無いと分析できない
    if safe_metrics is None and not master_csv.exists():
        return None, None, None, [], (
            "safe_metrics.json も ultimate_shift_master.csv も見つかりません。"
            "管理者に safe_metrics.json をリポジトリへ配置するよう依頼してください。"
        ), {}

    try:
        result = _cached_compute_real_shortages(
            target_month=target_month,
            folder=folder,
            master_mtime=(os.path.getmtime(master_csv) if master_csv.exists() else -1.0),
            xlsx_mtimes_key=_shift_xlsx_mtimes_tuple(),
            whitelist_key=_whitelist_to_key(movable_whitelist),
            safe_metrics_mtime=safe_metrics_mtime,
        )
    except RealDataMissingError as e:
        return None, None, None, [], f"実シフトデータが不足しています:\n{e}", {}
    except Exception as e:
        return None, None, None, [], f"分析中にエラーが発生しました: {e}", {}

    # light_analyze がエラーを返した場合
    if result.get('error'):
        return None, None, None, result.get('missing_areas', []), result['error'], {}

    # 集計（ベクトル化）— apply(lambda) から clip ベースに変更
    gap_df = result['gap_df'].copy()
    gap_df['shortage_pd'] = (-gap_df['gap']).clip(lower=0).astype(int)
    gap_df['surplus_pd'] = gap_df['gap'].clip(lower=0).astype(int)
    grp_short = gap_df.groupby('area')['shortage_pd'].sum().astype(int)
    grp_surp = gap_df.groupby('area')['surplus_pd'].sum().astype(int)
    shortages = {c: int(grp_short.get(c, 0)) for c in CLINIC_FILE_PREFIX}
    surpluses = {c: int(grp_surp.get(c, 0)) for c in CLINIC_FILE_PREFIX}

    staff_help_actions = result.get('staff_help_actions')
    if staff_help_actions is None:
        staff_help_actions = pd.DataFrame(
            columns=['date', 'staff_name', 'src_clinic',
                     'dst_clinic', 'same_area', 'shortage_n_remaining']
        )

    # missing_areas は main.run_full_analysis() が直接返す
    missing_areas = list(result.get('missing_areas') or [])

    extras = {
        'surpluses': surpluses,
        'paid_leave_df': result.get('paid_leave_df'),
        'help_actions_actual_df': result.get('help_actions_actual_df'),
        'fixed_leave_df': result.get('fixed_leave_df'),
        'worked_df': result.get('worked_df'),
    }
    return (shortages, gap_df, staff_help_actions, missing_areas,
            None, extras)


# ==========================================
# 過不足ステータスのカラーリング判定
# ==========================================
# 各院のひと月合計 (不足人日, 余剰人日) から下記いずれかを返す
#   'short'  : 不足 → 赤
#   'optimal': 適正 → 緑
#   'over'   : 過剰 → 黄
def classify_clinic_status(shortage_pd: int, surplus_pd: int) -> str:
    if shortage_pd > 0 and shortage_pd >= surplus_pd:
        return 'short'
    if surplus_pd > 0 and surplus_pd > shortage_pd:
        return 'over'
    return 'optimal'


# (背景色, 枠線色, ラベル文字色, ラベル, アイコン)
STATUS_STYLE = {
    'short':   ('#fee2e2', '#dc2626', '#991b1b', '不足', '🔴'),
    'optimal': ('#dcfce7', '#16a34a', '#166534', '最適', '🟢'),
    'over':    ('#fef9c3', '#ca8a04', '#854d0e', '過剰', '🟡'),
}

# 月4回以上のヘルプ移動を「定着リスク」として警告する閾値
HELP_WARN_THRESHOLD = 4


# ==========================================
# 集計ヘルパ（ベクトル化された分析ユーティリティ）
# ==========================================
def detect_saturday_help_misses(gap_df: pd.DataFrame) -> list[dict]:
    """土曜日の「3人体制以下の不足院」と「余剰院」の並存を検出する。
    pandas のベクトル演算のみで構築し、apply / iterrows を一切使わない。"""
    if not isinstance(gap_df, pd.DataFrame) or gap_df.empty:
        return []
    gdf = gap_df.copy()
    gdf['date'] = pd.to_datetime(gdf['date'], errors='coerce')
    gdf = gdf.dropna(subset=['date'])
    if gdf.empty:
        return []
    sat = gdf[gdf['date'].dt.weekday == 5]
    if sat.empty:
        return []
    avail = (sat['available_staff']
             if 'available_staff' in sat.columns else sat['planned_staff'])
    short_mask = (avail <= 3) & (sat['gap'] < 0)
    surplus_mask = sat['gap'] > 0
    short_by_date = sat.loc[short_mask].groupby('date')['area'].apply(list)
    surplus_by_date = sat.loc[surplus_mask].groupby('date')['area'].apply(list)
    common_dates = short_by_date.index.intersection(surplus_by_date.index)
    return [
        {
            'date': d.date().isoformat(),
            'short': short_by_date.loc[d],
            'surplus': surplus_by_date.loc[d],
        }
        for d in common_dates
    ]


def detect_sakashita_solo_risk(worked_df: pd.DataFrame) -> list[dict]:
    """小金井坂下で「山本休 + 稲田単独」となる日を検出する（ベクトル化）。"""
    if not isinstance(worked_df, pd.DataFrame) or worked_df.empty:
        return []
    sub = worked_df[worked_df['area'] == '小金井坂下'].copy()
    if sub.empty:
        return []
    names = sub['staff_name'].astype(str)
    sub['_is_yamamoto'] = names.str.contains('山本', na=False, regex=False)
    sub['_is_inada'] = names.str.contains('稲田', na=False, regex=False)
    sub['_is_other'] = ~(sub['_is_yamamoto'] | sub['_is_inada'])
    agg = sub.groupby('date').agg(
        has_yamamoto=('_is_yamamoto', 'any'),
        has_inada=('_is_inada', 'any'),
        n_other=('_is_other', 'sum'),
        names=('staff_name', lambda s: list(s)),
    )
    risk = agg[(~agg['has_yamamoto']) & agg['has_inada'] & (agg['n_other'] == 0)]
    return [
        {'date': str(d), 'staff': r['names']}
        for d, r in risk.iterrows()
    ]


def build_hr_summary(
    worked_df: pd.DataFrame | None,
    paid_leave_df: pd.DataFrame | None,
    help_actions_df: pd.DataFrame | None,
    fixed_leave_df: pd.DataFrame | None,
) -> tuple[pd.DataFrame, dict]:
    """人事サマリ DataFrame と統計値を groupby / map ベースで一括生成する。
    返り値: (hr_df, stats) — stats は {zero_paid_n, many_help_n, fixed_days_n}
    """
    # スタッフ→自院 のマップを 3 ソースから union（ベクトル化）
    parts = []
    src_table = [
        (worked_df, 'staff_name', 'area'),
        (paid_leave_df, 'staff_name', 'area'),
        (help_actions_df, 'staff_name', 'src_clinic'),
    ]
    for df, name_col, area_col in src_table:
        if isinstance(df, pd.DataFrame) and len(df) > 0 \
                and name_col in df.columns and area_col in df.columns:
            parts.append(
                df[[name_col, area_col]]
                .rename(columns={name_col: 'staff_name', area_col: '院'})
                .dropna(subset=['staff_name'])
            )
    if not parts:
        cols = ['院', 'スタッフ名', '有給取得日数', '有給アラート',
                'ヘルプ回数', 'ヘルプアラート', '固定休日数']
        empty = pd.DataFrame(columns=cols)
        return empty, {'zero_paid_n': 0, 'many_help_n': 0, 'fixed_days_n': 0}

    base = (pd.concat(parts, ignore_index=True)
            .drop_duplicates(subset=['staff_name'], keep='first'))

    # 有給日数
    paid_n = (paid_leave_df.groupby('staff_name').size()
              if isinstance(paid_leave_df, pd.DataFrame) and len(paid_leave_df) > 0
              else pd.Series(dtype=int))
    # ヘルプ回数（他院移動のみ）
    if (isinstance(help_actions_df, pd.DataFrame) and len(help_actions_df) > 0
            and {'src_clinic', 'dst_clinic'}.issubset(help_actions_df.columns)):
        hd = help_actions_df[
            help_actions_df['src_clinic'] != help_actions_df['dst_clinic']
        ]
        help_n = hd.groupby('staff_name').size()
    else:
        help_n = pd.Series(dtype=int)
    # 固定休日数
    fixed_n = (fixed_leave_df.groupby('staff_name').size()
               if isinstance(fixed_leave_df, pd.DataFrame) and len(fixed_leave_df) > 0
               else pd.Series(dtype=int))

    base['有給取得日数'] = base['staff_name'].map(paid_n).fillna(0).astype(int)
    base['ヘルプ回数'] = base['staff_name'].map(help_n).fillna(0).astype(int)
    base['固定休日数'] = base['staff_name'].map(fixed_n).fillna(0).astype(int)
    base['有給アラート'] = base['有給取得日数'].where(
        base['有給取得日数'] != 0, '🚨 介入要'
    ).where(base['有給取得日数'] == 0, '')
    base['ヘルプアラート'] = (
        base['ヘルプ回数'] >= HELP_WARN_THRESHOLD
    ).map({True: '⚠️ 注意', False: ''})

    hr_df = (base.rename(columns={'staff_name': 'スタッフ名'})
                 [['院', 'スタッフ名', '有給取得日数', '有給アラート',
                   'ヘルプ回数', 'ヘルプアラート', '固定休日数']]
                 .sort_values(['院', 'スタッフ名'])
                 .reset_index(drop=True))

    stats = {
        'zero_paid_n': int((hr_df['有給取得日数'] == 0).sum()),
        'many_help_n': int((hr_df['ヘルプ回数'] >= HELP_WARN_THRESHOLD).sum()),
        'fixed_days_n': int(hr_df['固定休日数'].sum()),
    }
    return hr_df, stats


def _hr_style_highlight(df: pd.DataFrame):
    """有給0=赤背景、ヘルプ過多=橙背景 を Styler に適用する。"""
    paid0 = df['有給取得日数'] == 0
    help_warn = df['ヘルプ回数'] >= HELP_WARN_THRESHOLD
    out = pd.DataFrame('', index=df.index, columns=df.columns)
    out.loc[paid0, '有給取得日数'] = (
        'color:#dc2626; font-weight:bold; background:#fee2e2;'
    )
    out.loc[help_warn, 'ヘルプ回数'] = (
        'color:#92400e; font-weight:bold; background:#fef3c7;'
    )
    return out


# ==========================================
# 結果ダウンロード（Excel/CSV）生成
# ==========================================
def build_results_xlsx_bytes(
    analysis: dict,
    hr_df: pd.DataFrame | None,
    capacity_per_day: dict,
) -> bytes:
    """分析結果から複数シートの Excel ファイルを生成して bytes で返す。
    利用者が Excel で開いて手直しできる形式（行・列の編集自由）。"""
    buf = io.BytesIO()
    target_month = analysis.get('target_month', '')
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        # シート1: 概要
        shortages = analysis.get('shortages', {}) or {}
        surpluses = analysis.get('surpluses', {}) or {}
        unit_prices = analysis.get('unit_prices', {}) or {}
        rows = []
        for c in shortages:
            up = unit_prices.get(c) or 6000
            if pd.isna(up):
                up = 6000
            cap = capacity_per_day.get(c, 12)
            short = int(shortages.get(c, 0))
            rows.append({
                '院': c, '不足人日': short, '余剰人日': int(surpluses.get(c, 0)),
                '1日対応枠': cap, '客単価': int(up),
                '機会損失額': short * cap * int(up),
            })
        pd.DataFrame(rows).to_excel(writer, sheet_name='概要', index=False)

        # シート2-6: 詳細データ
        sheets = [
            ('応援アクション指示', 'staff_help_actions'),
            ('過不足詳細',         'gap_df'),
            ('有給取得',           'paid_leave_df'),
            ('他院ヘルプ実績',     'help_actions_actual_df'),
            ('固定休',             'fixed_leave_df'),
        ]
        for sheet_name, key in sheets:
            df = analysis.get(key)
            if isinstance(df, pd.DataFrame) and len(df) > 0:
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        # シート7: 人事サマリ
        if isinstance(hr_df, pd.DataFrame) and len(hr_df) > 0:
            hr_df.to_excel(writer, sheet_name='人事サマリ', index=False)

        # シート8: メタ情報
        meta = pd.DataFrame([
            {'項目': '対象月', '値': target_month},
            {'項目': '分析実行時刻', '値':
                analysis.get('analyzed_at', '')},
            {'項目': '機会損失合計', '値':
                sum(r['機会損失額'] for r in rows)},
        ])
        meta.to_excel(writer, sheet_name='メタ情報', index=False)

    return buf.getvalue()


def df_to_csv_bytes(df: pd.DataFrame | None) -> bytes:
    """DataFrame を UTF-8 BOM 付き CSV bytes に変換（Excel で文字化けしない）。"""
    if df is None or len(df) == 0:
        return b''
    return df.to_csv(index=False).encode('utf-8-sig')


# ==========================================
# 日付連動の運用スケジュール警告
# ==========================================
def render_schedule_alert(today: _date | None = None) -> None:
    """システム日付に応じて、画面上部にシフト作成フローの進捗アラートを表示"""
    d = today or _date.today()
    day = d.day
    if 10 <= day <= 15:
        st.info(
            f"📝 **【希望休提出期間です（15日〆）】** "
            f"本日 {d.isoformat()} — 各院スタッフは {d.year}年{d.month}月15日 までに "
            f"希望休をシフト表に入力してください。",
            icon="📝",
        )
    elif 15 < day <= 20:
        st.warning(
            f"🛠️ **【シフト作成期間です】** 本日 {d.isoformat()} — "
            f"{d.year}年{d.month}月20日 までに各院でシフトを完成・公開してください。",
            icon="🛠️",
        )
    elif day >= 21:
        st.error(
            f"⚠️ **21日を過ぎています（本日 {d.isoformat()}）。** "
            f"未公開の院は至急是正し、システムを更新してください。",
            icon="⚠️",
        )
    else:
        # 1-9日: 集計確認・準備期間
        st.caption(
            f"📅 本日 {d.isoformat()} — 集計・準備期間（10日から希望休提出開始）"
        )


# ==========================================
# サイドバー
# ==========================================
# --- セッション起動時にデフォルトファイルを temp dir へ seed する ---
_seed_status = seed_defaults_if_present()
st.session_state['_seed_status_cache'] = _seed_status
if _seed_status.get('seeded'):
    # 初回 seed 時のみトースト
    n_seeded_shifts = sum(1 for v in _seed_status['shifts'].values() if v == 'default')
    msg_parts = []
    if _seed_status['master'] == 'default':
        msg_parts.append("マスタCSV")
    if _seed_status['crm'] == 'default':
        msg_parts.append("CRM CSV")
    if n_seeded_shifts:
        msg_parts.append(f"シフト{n_seeded_shifts}院")
    if msg_parts:
        st.toast(
            f"📥 デフォルトデータを自動読込: {' / '.join(msg_parts)}",
            icon="✅",
        )

# ---- 院長専用アップロード（シフト Excel のみ） ----
st.sidebar.header("📤 シフト表をアップロード")
_safe_metrics_present = SAFE_METRICS_PATH.exists()
if _safe_metrics_present:
    st.sidebar.success(
        "✅ **safe_metrics.json 読込済み**\n\n"
        "客単価・必要人員・固定スタッフルールは管理者が事前に設定済みです。"
        "院長は **自院のシフト表 Excel** をドロップするだけで分析できます。"
    )
else:
    st.sidebar.warning(
        "⚠ safe_metrics.json が見つかりません。\n"
        "管理者に generate_safe_metrics.py を実行して、リポジトリに"
        "safe_metrics.json を配置してもらってください。"
    )

st.sidebar.caption(
    "🛡️ アップロードされたファイルはサーバの一時メモリにのみ置かれ、"
    "元のローカルファイルは一切変更されません。"
)

shift_uploads = st.sidebar.file_uploader(
    "🗓 シフト表 Excel（自院分でOK）",
    type=['xlsx'],
    accept_multiple_files=True,
    key='upload_shifts',
    help=(
        "ファイル名の先頭が「【国分寺】」「【武蔵小金井】」「【東小金井】」"
        "「【坂下】」「【人形町】」のいずれかである .xlsx をドロップ。"
        "1院だけのアップロードでも分析できます（他院はデフォルト or 未読込として扱い）。"
    ),
)

# アップロードを temp dir に書き出し（同院の seed 済 xlsx は削除して上書き）
_upload_status = sync_uploads_to_tempdir(shift_uploads, None, None)

# アップロード後に seed status を再評価して画面側へ反映
_seed_status = seed_defaults_if_present()
st.session_state['_seed_status_cache'] = _seed_status

# サイドバー: 現在の読込状況（コンパクト表示）
st.sidebar.divider()
st.sidebar.markdown("**📋 シフト表の認識状況**")
for c in CLINIC_FILE_PREFIX:
    st.sidebar.caption(
        f"{_src_badge(_seed_status['shifts'].get(c, 'absent'))} {c}"
    )

st.sidebar.divider()
st.sidebar.markdown("### 📅 対象月を選択")
target_month = st.sidebar.text_input(
    "対象月入力",
    value="2026-04",
    help="YYYY-MM 形式で入力（例: 2026-02 / 2026-03 / 2026-04）",
    label_visibility="collapsed",
)
month_valid = bool(re.fullmatch(r'\d{4}-\d{2}', target_month or ''))
if not month_valid:
    st.sidebar.warning("⚠ YYYY-MM 形式で入力してください")

st.sidebar.divider()

# ★ 目立つプライマリーボタン
analyze_clicked = st.sidebar.button(
    f"▶  {target_month} を分析する",
    type="primary",
    use_container_width=True,
    disabled=not month_valid,
    help="このボタンを押すと、上記の対象月で分析を実行します",
)

# ★ 強制リロードボタン（Excel編集後はこちらを推奨）
reload_clicked = st.sidebar.button(
    "🔄 最新のデータを再読み込み",
    use_container_width=True,
    disabled=not month_valid,
    help=(
        "Excelファイルを編集した直後に押してください。\n"
        "キャッシュをすべてクリアし、ディスクから最新を読み直して再分析します。"
    ),
)
st.sidebar.caption(
    "💡 Excelを書き換えた後はこのボタンを押すと、"
    "前回の表示を破棄して最新ファイルから再計算します。"
)

# リロード処理 — Streamlitキャッシュ + 分析関連 session_state を「完全に」初期化し、
#                月入力が有効なら直後に再分析を自動実行する
if reload_clicked:
    # 1) Streamlit の cache_data / cache_resource を防御的にクリア
    _clear_streamlit_caches()
    # 2) 分析結果に関連する session_state を全消し
    cleared_keys = _reset_analysis_state()
    # 3) 月入力が有効なら 必ず 再分析を予約（前回結果の有無に依存しない）
    if month_valid:
        st.session_state['_pending_reanalysis'] = True
    # 4) ユーザに何をクリアしたか可視化
    if cleared_keys:
        st.toast(
            f"🧹 キャッシュ + state {len(cleared_keys)} 件をクリア: "
            f"{', '.join(cleared_keys[:5])}"
            + (f" 他{len(cleared_keys)-5}件" if len(cleared_keys) > 5 else ''),
            icon="🔄",
        )
    else:
        st.toast("🧹 キャッシュをクリア（state は元から空でした）", icon="🔄")
    # 5) 画面を即時更新（次回 run で再分析が走る）
    st.rerun()

if 'analysis' in st.session_state:
    st.sidebar.divider()
    if st.sidebar.button(
        "🧹 分析結果をクリア", use_container_width=True,
        help="この画面の分析結果を破棄します（再分析はされません）",
    ):
        # クリアボタンも同じ helper で漏れなく
        cleared_keys = _reset_analysis_state()
        st.toast(
            f"🧹 state {len(cleared_keys)} 件をクリア", icon="🧹",
        )
        st.rerun()

# 🗑 アップロードクリア（次の利用者のためにセッションをリセット）
st.sidebar.divider()
st.sidebar.markdown("### 🗑 アップロード/セッションのリセット")
st.sidebar.caption(
    "アップロード済みファイルをサーバ側のメモリから削除し、最初の状態に戻します。"
    "**元のローカルファイルには影響しません**。"
)
if st.sidebar.button(
    "🗑 アップロードを全部クリア",
    use_container_width=True,
    help="このボタンを押すと、アップロードしたファイルがサーバ側から消去されます",
):
    reset_session_tempdir()
    _reset_analysis_state()
    _clear_streamlit_caches()
    # アップロードウィジェットのキーもリセット
    for k in ['upload_shifts', 'movable_whitelist']:
        st.session_state.pop(k, None)
    for k in list(st.session_state.keys()):
        if k.startswith('ms_help_'):
            st.session_state.pop(k, None)
    st.toast("🗑 アップロードをすべてクリアしました", icon="🧹")
    st.rerun()

# ★ デバッグ用：session_state 診断パネル
st.sidebar.divider()
with st.sidebar.expander("🔍 session_state を確認", expanded=False):
    keys = list(st.session_state.keys())
    if not keys:
        st.caption("（空）— state にデータは残っていません ✅")
    else:
        st.caption(
            f"現在 **{len(keys)} 件** のキーが残っています："
        )
        for k in sorted(keys):
            v = st.session_state[k]
            if isinstance(v, dict):
                size = f"dict({len(v)} keys)"
            elif hasattr(v, '__len__'):
                try:
                    size = f"{type(v).__name__}({len(v)})"
                except Exception:
                    size = type(v).__name__
            else:
                size = repr(v)[:30]
            st.caption(f"・`{k}` = {size}")
    if st.button("🧨 全 state を完全削除（debug）",
                 use_container_width=True,
                 help="あらゆるキーを問答無用で消します。"
                      "再ロードボタンで取れない state を消す最終手段。"):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        _clear_streamlit_caches()
        st.toast("🧨 すべての session_state を消去しました", icon="🧨")
        st.rerun()


# ==========================================
# メインタイトル＋対象月バナー
# ==========================================
header_col1, header_col2 = st.columns([3, 1])
with header_col1:
    st.title("🩺 シフト最適化 Web パネル")
    st.markdown(
        "毎月17日までに必要人員を算出し、20日時点で過不足と"
        "**機会損失額**が見える状態を作るためのWebツール"
    )
with header_col2:
    st.markdown(
        f"""
        <div style="background:linear-gradient(135deg,#2563eb 0%,#7c3aed 100%);
                    border-radius:14px;
                    padding:clamp(10px, 2.5vw, 18px) clamp(8px, 2vw, 12px);
                    text-align:center; color:white;
                    box-shadow:0 4px 12px rgba(15,23,42,0.15);">
            <div style="font-size:clamp(9px, 1.6vw, 11px);
                        letter-spacing:2px; opacity:0.85;">
                ANALYSIS MONTH
            </div>
            <div style="font-size:clamp(18px, 4.5vw, 28px); font-weight:bold;
                        margin-top:4px; letter-spacing:1px;">
                {target_month}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# ==========================================
# 日付連動の運用スケジュール警告（最上段に表示）
# ==========================================
render_schedule_alert()

st.divider()


# =====================================================================
# 📤 ファイルアップロード
# ---------------------------------------------------------------------
# ローカルパスや Google Drive への依存を廃止し、すべてブラウザから
# アップロードされたファイルで分析する。
#  - シフト Excel: 5院ぶん（必須）
#  - ultimate_shift_master.csv: 過去患者・スタッフ実績の集計（必須）
#  - final_analysis_data.csv: CRM 売上データ（任意 — 客単価算出に使用）
#
# ★ 重要: アップロードされたファイルは **サーバの一時ディレクトリにのみ**
#   保存され、元のローカルファイルには一切影響しない。
# =====================================================================
# =====================================================================
# データ読み込み状況パネル（メイン画面・簡潔表示）
# ---------------------------------------------------------------------
#  サイドバーで一時ファイルをアップロードした場合はそちらを優先表示し、
#  なければ DEFAULT_DATA_DIR のデフォルトファイルが自動使用される。
# =====================================================================
seed_status = st.session_state.get('_seed_status_cache', {})
_safe_metrics_for_status = load_safe_metrics()


with st.container(border=True):
    st.markdown("### 📊 データ読み込み状況")
    sub_cols = st.columns(6)
    # safe_metrics.json (管理者が事前生成・リポジトリ同梱)
    sub_cols[0].caption("**設定 JSON**")
    if _safe_metrics_for_status:
        sub_cols[0].markdown("🟢 読込済み")
        sub_cols[0].caption(
            f"生成: {_safe_metrics_for_status.get('generated_at', '?')[:10]}"
        )
    else:
        sub_cols[0].markdown("⬜ 未配置")
    # 5院シフト xlsx
    shift_status = seed_status.get('shifts', {}) or {}
    for i, clinic in enumerate(CLINIC_FILE_PREFIX):
        sub_cols[i + 1].caption(f"**{clinic}**")
        sub_cols[i + 1].markdown(_src_badge(shift_status.get(clinic, 'absent')))

    # 凡例
    st.caption(
        "凡例: 🟢 リポジトリ内のデフォルトファイル使用 ／ "
        "🔵 サイドバーでアップロードされたファイルで上書き中 ／ "
        "⬜ 未読み込み（サイドバーからシフト Excel をアップロードしてください）"
    )
    # safe_metrics.json から客単価を表示
    if _safe_metrics_for_status:
        ups = _safe_metrics_for_status.get('unit_prices', {})
        st.caption(
            "💰 **客単価 (safe_metrics.json より)**: "
            + " / ".join(f"{c} ¥{int(ups.get(c, 6000)):,}" for c in CLINIC_FILE_PREFIX)
        )

st.divider()


# ==========================================
# 院長向け：ヘルプ要員選択（院ごとマルチセレクト）
# -----------------------------------------------------------------
#  各院の院長が「他院へヘルプに出してよいスタッフ」を明示的に選択する。
#  ここで選択されたスタッフ「だけ」が応援候補となり、選択されていない
#  スタッフは応援対象から完全に除外される（Reallocator に伝搬）。
# ==========================================

@st.cache_data(show_spinner=False)
def _cached_rosters_from_xlsx(
    month: str, folder: str, xlsx_mtimes_key: tuple,
) -> dict[str, list[str]]:
    """シフト xlsx から {院名: [スタッフ名]} を抽出してキャッシュ。
    xlsx の mtime が変わるとキャッシュは自動失効する。"""
    del xlsx_mtimes_key  # cache key only
    try:
        from shift_optimizer.src.shift_generator import (
            extract_rosters_and_leave,
        )
        info = extract_rosters_and_leave(Path(folder), month)
        return info.get('rosters', {}) or {}
    except Exception:
        return {}


def _load_rosters_from_xlsx(month: str, folder: str) -> dict[str, list[str]]:
    """シフト xlsx から名簿を取得（mtime keyed cached）。"""
    return _cached_rosters_from_xlsx(
        month, folder, _shift_xlsx_mtimes_tuple()
    )


def _get_or_load_rosters(month: str, folder: str) -> dict[str, list[str]]:
    """名簿を取得する。worked_df があればそこから、無ければ xlsx から。
    後者は `@st.cache_data` 経由で重い openpyxl 処理を回避する。"""
    rosters: dict[str, list[str]] = {}
    analysis = st.session_state.get('analysis') or {}
    if analysis.get('target_month') == month:
        worked = analysis.get('worked_df')
        if isinstance(worked, pd.DataFrame) and len(worked) > 0:
            rosters = (
                worked.dropna(subset=['staff_name'])
                .groupby('area')['staff_name']
                .agg(lambda s: sorted(set(s)))
                .to_dict()
            )
    if set(rosters.keys()) != set(CLINIC_FILE_PREFIX.keys()):
        for c, names in _load_rosters_from_xlsx(month, folder).items():
            rosters.setdefault(c, list(names))
    return rosters


# =====================================================================
# 院長向け「ヘルプ要員」選択
# ---------------------------------------------------------------------
#  スマホ・タブレットで使いやすいよう、以下の二段構えで構成する:
#   1) 外側 st.expander (expanded=True) — 常に画面上に存在する設定エリア
#   2) 内側 st.tabs (5院ぶん) — タップで切り替え、一度に一院だけ表示
#  これにより、スマホ画面でも縦長スクロールにならず、見たい院だけ操作できる。
#
#  選択された staff だけが Reallocator の応援候補になる仕組みは下記:
#      session_state['movable_whitelist'] = {clinic: [staff,...]}
#         ↓
#      compute_real_shortages(..., movable_whitelist=...)
#         ↓
#      run_full_analysis(..., movable_whitelist=...)
#         ↓
#      Reallocator.suggest_staff_help_actions(..., movable_whitelist=...)
# =====================================================================
_wl_state = st.session_state.get('movable_whitelist') or {}
_wl_total = sum(
    len(v) for v in _wl_state.values()
    if isinstance(v, (list, set, tuple))
)
_wl_summary = (
    f"（現在 {_wl_total} 名選択中）" if _wl_total
    else "（未選択／既定で全員可）"
)

with st.expander(
    f"👥 院長向け：ヘルプに出せるスタッフを各院で選択 {_wl_summary}",
    expanded=True,
):
    st.caption(
        "各院タブのマルチセレクトで **チェックされたスタッフだけ** が "
        "他院へのヘルプ要員候補になります。**チェックされていないスタッフは "
        "応援対象から完全に除外されます**。"
    )

    folder_for_roster = get_data_folder()
    rosters_for_ui = _get_or_load_rosters(target_month, folder_for_roster)

    # 上段: 名簿取得ボタン（モバイルでも押しやすいよう独立行）
    if st.button(
        "📋 名簿を取得 / 更新",
        help="シフト xlsx からスタッフ名簿を取得し、選択肢を更新します",
        use_container_width=True,
    ):
        # mtime-key の `@st.cache_data` を明示的にクリアして再取得
        _cached_rosters_from_xlsx.clear()
        with st.spinner("シフト xlsx から名簿を取得中…"):
            rosters_for_ui = _load_rosters_from_xlsx(
                target_month, folder_for_roster
            )
        if rosters_for_ui:
            st.toast(
                f"📋 {len(rosters_for_ui)} 院の名簿を取得しました",
                icon="✅",
            )
        else:
            st.warning(
                f"⚠ {target_month} の名簿を xlsx から取得できませんでした。"
                "対象月のシートが各シフト xlsx に存在するか確認してください。"
            )

    # 既存 selection の保持先
    whitelist_state_key = 'movable_whitelist'
    if whitelist_state_key not in st.session_state:
        st.session_state[whitelist_state_key] = {}

    clinic_order = list(CLINIC_FILE_PREFIX.keys())

    # -----------------------------------------------------------------
    # ★ st.tabs ではなく st.radio (horizontal) で「擬似タブ」を実装する。
    #   理由: st.tabs はラベル変更や rerun でアクティブタブが先頭にリセット
    #   される問題があるため、session_state で確実に保持できる radio を使う。
    #   選択中院の選択肢数(N)はラジオラベル内ではなく、選択後にカード内で
    #   表示することで、ラベル変化によるさらなる再描画リセットも防止する。
    # -----------------------------------------------------------------
    if 'active_clinic_tab' not in st.session_state:
        st.session_state['active_clinic_tab'] = clinic_order[0]
    elif st.session_state['active_clinic_tab'] not in clinic_order:
        st.session_state['active_clinic_tab'] = clinic_order[0]

    # ★ 重要: ラジオの format_func は **動的な値を含めない** こと。
    #   選択数 (N) を含めると rerun のたびにラベル文字列が変わり、Streamlit
    #   の options.index(format_func(value)) 探索が失敗してタブが先頭にリセット
    #   される（武蔵小金井・小金井坂下・人形町などで顕著だった）。
    #   選択数は別のサマリ行で表示することで、ラベルを完全に安定化する。
    def _clinic_label(c: str) -> str:
        return f"🏢 {c}"

    active_clinic = st.radio(
        "院を選択（タブ操作）",
        options=clinic_order,
        format_func=_clinic_label,
        horizontal=True,
        key='active_clinic_tab',
        label_visibility='collapsed',
    )

    # 各院の選択中人数を可視化する 1 行サマリ（ラジオラベルとは別管理）
    _count_summary_parts = []
    for c in clinic_order:
        val = st.session_state.get(f'ms_help_{c}')
        if val is None:
            wl = st.session_state.get(whitelist_state_key) or {}
            val = wl.get(c) or []
        n = len(val or [])
        if n > 0:
            _count_summary_parts.append(f"**{c}** {n}名")
        else:
            _count_summary_parts.append(f"{c} —")
    st.caption("👥 選択中: " + " ／ ".join(_count_summary_parts))

    # ----- 表示は active_clinic の1院ぶんのみ -----
    # 他の院の選択値は session_state にウィジェットキー (ms_help_<院>) として
    # 保持されているため、ここで再レンダーしなくても消えない。
    new_whitelist: dict[str, list[str]] = {}
    for c in clinic_order:
        if c == active_clinic:
            continue  # 下のブロックで描画
        # アクティブ外の院: 既存セッション値を引き継ぐ
        new_whitelist[c] = (
            st.session_state.get(f'ms_help_{c}')
            or st.session_state[whitelist_state_key].get(c)
            or []
        )

    clinic = active_clinic
    staff_options = rosters_for_ui.get(clinic, [])
    immov_for_clinic = IMMOVABLE_STAFF.get(clinic, [])

    # 既定: 動かせないスタッフ (IMMOVABLE_STAFF) 以外は全員チェック
    default_selected = [
        n for n in staff_options
        if not any(s in n for s in immov_for_clinic)
    ]
    existing = st.session_state[whitelist_state_key].get(clinic)
    if existing is not None:
        preset = [n for n in existing if n in staff_options]
    else:
        preset = default_selected

    with st.container(border=True):
        st.caption(
            f"**{clinic}**  ／  候補スタッフ: **{len(staff_options)}** 名"
            + (
                f"  ／  自院固定 (既定除外): {', '.join(immov_for_clinic)}"
                if immov_for_clinic else ""
            )
        )

        sel = st.multiselect(
            "他院へヘルプに出せるスタッフ",
            options=staff_options,
            default=preset,
            key=f'ms_help_{clinic}',
            help=(
                f"{clinic} の中で、他院へヘルプに出せるスタッフを"
                f"選択してください。"
            ),
            placeholder=(
                "名簿未取得 — 上の「📋 名簿を取得」を押してください"
                if not staff_options else "ヘルプ要員を選択"
            ),
            label_visibility="visible",
        )
        new_whitelist[clinic] = sel

        # フィードバック
        if staff_options:
            excluded = [n for n in staff_options if n not in sel]
            st.markdown(
                f"✅ ヘルプ可: **{len(sel)} 名** ／ "
                f"🚫 ヘルプ不可: **{len(excluded)} 名**"
            )
            if excluded:
                with st.expander(
                    f"🚫 ヘルプから除外されるスタッフ ({len(excluded)} 名)",
                    expanded=False,
                ):
                    st.markdown(
                        "・" + "  ／  ".join(f"`{n}`" for n in excluded)
                    )

    # 状態を保存（次回 analyze 時に compute_real_shortages へ渡される）
    st.session_state[whitelist_state_key] = new_whitelist

    # 適用ボタン: 現在の選択でその場で再分析を予約する
    total_sel = sum(len(v) for v in new_whitelist.values())
    st.divider()
    st.markdown(
        f"📊 **現在の合計ヘルプ要員: {total_sel} 名** "
        "— 変更を反映するには下のボタンを押してください。"
    )
    if st.button(
        "✅ 選択を反映して再分析",
        type="primary",
        use_container_width=True,
        disabled=not month_valid,
        help="現在のチェック内容で応援アクションを再計算します",
    ):
        st.session_state['_pending_reanalysis'] = True
        st.rerun()

st.divider()


# ==========================================
# 分析の実行 → session_state に保存
# 「▶ 分析する」または リロードによる自動再分析 で起動
# ==========================================
# 注意: pop は常に呼び出すこと（or の short-circuit で消費漏れすると、
#       次の run でも True のまま残り、無意味な再実行を引き起こす）
_pending_flag = st.session_state.pop('_pending_reanalysis', False)
_trigger_analysis = analyze_clicked or _pending_flag

if _trigger_analysis and month_valid:
    folder = get_data_folder()
    crm_path = os.path.join(folder, "final_analysis_data.csv")
    master_path = os.path.join(folder, "ultimate_shift_master.csv")

    # safe_metrics.json が利用可能か事前判定
    _safe_metrics_dict = load_safe_metrics()
    _have_safe_metrics = _safe_metrics_dict is not None

    # 必須ファイルの存在チェック（アップロード未完了の防止）
    missing_required: list[str] = []

    # safe_metrics があれば master CSV は不要、無ければ必要
    if not _have_safe_metrics and not os.path.exists(master_path):
        missing_required.append(
            "safe_metrics.json または ultimate_shift_master.csv"
        )

    n_shift_xlsx = sum(
        1 for c in CLINIC_FILE_PREFIX
        if (lambda x: x is not None and x.exists())(find_shift_xlsx(c))
    )
    if n_shift_xlsx == 0:
        missing_required.append("シフト Excel (1院以上)")

    if missing_required:
        st.error(
            "⚠ **必要なデータが不足しています**\n\n"
            + "\n".join(f"- ❌ {f}" for f in missing_required)
            + "\n\nサイドバーの **「📤 シフト表をアップロード」** から"
              "ファイルを選択してください。"
        )
    else:
        # 客単価: safe_metrics.json 優先 → 旧 CRM CSV → 既定¥6,000
        if _have_safe_metrics:
            unit_prices = dict(_safe_metrics_dict.get('unit_prices', {}))
            for c in CLINIC_FILE_PREFIX:
                unit_prices.setdefault(c, 6000)
            crm_mtime = (
                f"safe_metrics.json ({_safe_metrics_dict.get('generated_at', '?')})"
            )
        elif os.path.exists(crm_path):
            unit_prices, err_up = calculate_unit_prices(crm_path)
            if err_up:
                st.warning(
                    f"CRM 読み込み失敗 → 既定 ¥6,000: {err_up}"
                )
                unit_prices = {c: 6000 for c in CLINIC_FILE_PREFIX}
            crm_mtime = datetime.fromtimestamp(
                os.path.getmtime(crm_path)
            ).strftime("%Y-%m-%d %H:%M")
        else:
            unit_prices = {c: 6000 for c in CLINIC_FILE_PREFIX}
            crm_mtime = "未指定（既定 ¥6,000 で計算）"

        with st.spinner(f"{target_month} を分析中…"):
            err = None
            if True:
                # ② シフト xlsx + masterCSV から実「不足人日」と応援指示を算出
                # 院長の選択した「ヘルプ要員」白リストを反映
                _wl = st.session_state.get('movable_whitelist') or {}
                # 空リスト = 「誰も指定なし」と区別するため、空 dict のときは
                # None を渡して既定ロジック (IMMOVABLE_STAFF) で動作させる
                _wl_arg = _wl if any(v for v in _wl.values()) or any(
                    k in _wl for k in CLINIC_FILE_PREFIX
                ) else None
                (shortages, gap_df, staff_help_actions, missing_areas,
                 real_err, extras) = compute_real_shortages(
                    target_month, folder,
                    movable_whitelist=_wl_arg,
                )
                if real_err:
                    st.error(
                        f"⚠ **不足人日の算出に失敗**\n\n{real_err}\n\n"
                        f"対処: 各院のシフト xlsx に `{target_month}` "
                        f"シートが作成されているか確認してください。"
                    )
                else:
                    # シフト xlsx の最新 mtime を院ごとに記録
                    xlsx_mtimes = collect_shift_xlsx_mtimes(folder)
                    st.session_state['analysis'] = {
                        'target_month': target_month,
                        'unit_prices': unit_prices,
                        'shortages': shortages,
                        'surpluses': extras.get('surpluses', {}),
                        'gap_df': gap_df,
                        'staff_help_actions': staff_help_actions,
                        'paid_leave_df': extras.get('paid_leave_df'),
                        'help_actions_actual_df':
                            extras.get('help_actions_actual_df'),
                        'fixed_leave_df': extras.get('fixed_leave_df'),
                        'worked_df': extras.get('worked_df'),
                        'movable_whitelist_used': _wl_arg,
                        'missing_areas': missing_areas,
                        'data_folder': folder,
                        'crm_mtime': crm_mtime,
                        'analyzed_at': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        'crm_path': crm_path,
                        'xlsx_mtimes': xlsx_mtimes,
                    }
                    if missing_areas:
                        st.warning(
                            f"⚠ {target_month} のシートが無い院: "
                            f"{missing_areas}（集計対象外として処理）"
                        )
                    st.success(
                        f"✓ {target_month} の実シフト分析が完了しました "
                        f"（CRM最終更新: {crm_mtime}, "
                        f"不足合計: {sum(shortages.values())} 人日, "
                        f"応援指示: {len(staff_help_actions)} 件）"
                    )


# ==========================================
# 結果表示（session_state から読む）
# ==========================================
if 'analysis' in st.session_state:
    data = st.session_state['analysis']
    displayed_month = data['target_month']

    # サイドバーの対象月と表示中月がズレている場合の注意
    if displayed_month != target_month:
        st.warning(
            f"📌 現在表示中: **{displayed_month}** の結果 ／ "
            f"サイドバーは **{target_month}** に切り替わっています。\n\n"
            f"反映するには左サイドバーの **「▶ {target_month} を分析する」** "
            f"ボタンを押してください。"
        )

    shortages = data['shortages']
    unit_prices = data['unit_prices']
    total_shortage = sum(shortages.values())

    # ──── データ最新性インジケータ ────
    crm_mtime = data.get('crm_mtime', '不明')
    analyzed_at = data.get('analyzed_at', '不明')
    crm_path_str = data.get('crm_path', '')
    stored_xlsx_mtimes = data.get('xlsx_mtimes', {}) or {}
    data_folder = data.get('data_folder', get_data_folder())

    # CRM CSV と シフト xlsx の現在 mtime をディスクから再取得して比較
    stale_files: list[str] = []
    try:
        if crm_path_str and os.path.exists(crm_path_str):
            current_crm = datetime.fromtimestamp(
                os.path.getmtime(crm_path_str)
            ).strftime("%Y-%m-%d %H:%M")
            if current_crm != crm_mtime:
                stale_files.append(f"CRM ({current_crm})")
    except Exception:
        pass
    try:
        current_xlsx_mtimes = collect_shift_xlsx_mtimes(data_folder)
        for clinic, prev in stored_xlsx_mtimes.items():
            curr = current_xlsx_mtimes.get(clinic, '-')
            if curr != '-' and prev != '-' and curr != prev:
                stale_files.append(f"{clinic} ({curr})")
    except Exception:
        pass

    # 古いファイルへの警告だけは常時可視化、補足情報は expander に集約
    if stale_files:
        st.warning(
            "⚠ 分析後に変更されたファイル: " + " / ".join(stale_files) + "\n\n"
            "サイドバーの「🔄 最新のデータを再読み込み」ボタンで再分析してください。",
            icon="📌",
        )
    with st.expander(
        f"⏱ データ最新性インジケータ（CRM: {crm_mtime} ／ 分析実行: {analyzed_at}）",
        expanded=False,
    ):
        info_cols = st.columns([1, 1])
        info_cols[0].caption(f"📂 CRM最終更新: **{crm_mtime}**")
        info_cols[1].caption(f"⏱ 分析実行: **{analyzed_at}**")
        if stored_xlsx_mtimes:
            st.markdown("**📋 各院シフト Excel の最終更新時刻**")
            mtime_cols = st.columns(len(stored_xlsx_mtimes))
            for i, (clinic, mt) in enumerate(stored_xlsx_mtimes.items()):
                mtime_cols[i].caption(f"**{clinic}**\n\n{mt}")

    # ──── 応援要請ヘッダー（対象月を反映） ────
    st.markdown(
        f"### 🚨 【応援要請】─ **{displayed_month}** の30日間で "
        f"**合計 {total_shortage} 人日** の不足が発生しています"
    )
    st.caption(
        "👇 各院名のボタンをクリックすると、その院の "
        "**シフト表（Excel）** が直接開きます"
    )

    # ──── 💡応援アクション指示（スタッフ名・日付・src→dst 単位） ────
    staff_help_df = data.get('staff_help_actions')
    if staff_help_df is None or len(staff_help_df) == 0:
        st.info(
            "✅ 応援アクション指示: なし — 自院の余剰スタッフだけでは"
            "不足を埋められない、もしくは全院が同じ日に不足/余剰のため、"
            "現時点での具体的な人員移動案はありません。"
        )
    else:
        # 応援アクションは件数が多いと縦に長くなりがちなのでスマホでは
        # expander で畳めるようにする（デスクトップ表示でも長さ調整に有用）。
        with st.expander(
            f"💡 応援アクション指示 — **{len(staff_help_df)} 件**",
            expanded=True,
        ):
            st.caption(
                "下記の通り、動かせるスタッフのうち最適な人員を他院に応援に"
                "出してください。動かせないスタッフ（自院固定）は含まれません。"
            )
            df_sorted = staff_help_df.sort_values(
                ['date', 'dst_clinic', 'src_clinic', 'staff_name']
            ).reset_index(drop=True)

            # 日付表記をベクトル化（iterrows 廃止）
            dt_series = pd.to_datetime(
                df_sorted['date'].astype(str).str[:10], errors='coerce'
            )
            wd_jp = '月火水木金土日'
            fmt_dates = [
                f'{d.month}月{d.day}日（{wd_jp[d.weekday()]}）'
                if pd.notna(d) else str(orig)
                for d, orig in zip(dt_series, df_sorted['date'])
            ]
            lines = [
                f"- 💡 **【応援アクション指示】{name}さんを {fd} "
                f"{src}院 ➔ {dst}院 にヘルプを出して再度シフト表を"
                f"アップロードしてください**  〔{sa}〕"
                for name, fd, src, dst, sa in zip(
                    df_sorted['staff_name'], fmt_dates,
                    df_sorted['src_clinic'], df_sorted['dst_clinic'],
                    df_sorted['same_area'],
                )
            ]
            st.markdown("\n".join(lines))

            # データフレームでも一覧（モバイルでも横スクロール可能）
            st.dataframe(
                df_sorted, use_container_width=True, hide_index=True,
            )
            st.download_button(
                "📥 応援アクション指示を CSV ダウンロード",
                data=df_sorted.to_csv(index=False).encode('utf-8-sig'),
                file_name=f"staff_help_actions_{displayed_month}.csv",
                mime="text/csv",
                use_container_width=True,
            )

    # 動かせないスタッフ一覧（折りたたみ）
    with st.expander("🔒 動かせないスタッフ（他院応援の対象外）", expanded=False):
        immov_cols = st.columns(len(IMMOVABLE_STAFF))
        for i, (clinic, names) in enumerate(IMMOVABLE_STAFF.items()):
            immov_cols[i].markdown(
                f"**{clinic}**\n\n" + ("\n".join(f"・{n}" for n in names) or "—")
            )

    # ──── 院別カラム表示（クリッカブルな院名ボタン + ステータスカラー） ────
    surpluses = data.get('surpluses', {}) or {}
    # 凡例
    st.markdown(
        "<div style='display:flex; gap:14px; margin:6px 0 12px 0; "
        "font-size:0.9em;'>"
        "<span>🔴 <b>不足</b>（要応援）</span>"
        "<span>🟢 <b>最適</b>（適正人数）</span>"
        "<span>🟡 <b>過剰</b>（余剰人員）</span>"
        "</div>",
        unsafe_allow_html=True,
    )

    cols = st.columns(len(shortages))
    total_loss = 0

    for i, (clinic, shortage) in enumerate(shortages.items()):
        unit_price = unit_prices.get(clinic, 6000)
        if pd.isna(unit_price):
            unit_price = 6000
        capacity = CAPACITY_PER_DAY.get(clinic, 12)
        loss_amount = shortage * capacity * unit_price
        total_loss += loss_amount

        surplus = int(surpluses.get(clinic, 0))
        status = classify_clinic_status(int(shortage), surplus)
        bg, border, label_color, label, icon = STATUS_STYLE[status]

        with cols[i]:
            # ステータスカラーカード
            st.markdown(
                f"""
                <div style='background:{bg}; border:2px solid {border};
                            border-radius:12px; padding:10px 12px;
                            text-align:center; margin-bottom:6px;'>
                    <div style='font-size:0.85em; color:{label_color};
                                font-weight:bold; letter-spacing:1px;'>
                        {icon} {label}
                    </div>
                    <div style='font-size:0.95em; color:#334155; margin-top:4px;'>
                        不足 <b>{int(shortage)}</b> 人日 ／
                        余剰 <b>{surplus}</b> 人日
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            # 院名表示（Streamlit Cloud では os.startfile() は使えないため
            # クリッカブル → Excel 起動 は廃止。代わりに院別のヘルプ要請数を表示）
            st.markdown(
                f"<div style='text-align:center; font-weight:bold; "
                f"font-size:1.05em; margin:6px 0;'>🏢 {clinic}</div>",
                unsafe_allow_html=True,
            )

            # 院ごとの指標
            st.caption(f"{capacity} 枠/日 × 単価 {int(unit_price):,} 円")
            st.markdown(
                f"<div style='color:#dc2626; font-size:1.05em; "
                f"font-weight:bold; margin-top:6px;'>"
                f"📉 損失額: −{int(loss_amount):,} 円"
                f"</div>",
                unsafe_allow_html=True,
            )

    # ============================================================
    # ⚠ 運用リスクアラート（土曜の不足 ＋ 小金井坂下の単独オープン）
    # ============================================================
    st.markdown("---")
    st.markdown("### 🚨 運用リスクアラート")
    risk_messages: list[str] = []
    gap_df_full = data.get('gap_df')
    worked_df = data.get('worked_df')

    # ---- (A) 土曜日のヘルプ要請漏れ検知（ベクトル化ヘルパ） ----
    saturday_alerts = detect_saturday_help_misses(gap_df_full)
    if saturday_alerts:
        with st.container(border=True):
            st.error(
                f"⚠️ **【土曜日のヘルプ要請漏れ】 — {len(saturday_alerts)} 件**\n\n"
                "下記の土曜日は『3人体制以下の不足院』と『余剰院』が同日に並存しています。"
                "余剰院 → 不足院 への応援要請を必ず手配してください。"
            )
            for a in saturday_alerts:
                st.markdown(
                    f"- **{a['date']}（土）**:  "
                    f"不足院 = {', '.join(a['short'])}  "
                    f"／  余剰院 = {', '.join(a['surplus'])}"
                )
        risk_messages.append('土曜ヘルプ要請漏れ')

    # ---- (B) 小金井坂下: 山本休 ＋ 稲田単独 オープンリスク（ベクトル化） ----
    sakashita_alerts = detect_sakashita_solo_risk(worked_df)
    if sakashita_alerts:
        with st.container(border=True):
            st.error(
                f"⚠️ **【小金井坂下 オープン作業リスク】 — "
                f"{len(sakashita_alerts)} 日**\n\n"
                "下記の日は **山本さんが休み** で **稲田さん 1人出勤** に "
                "なっています。朝のオープン時間帯のリスクを避けるため、"
                "他スタッフ（ヘルプ含む）を必ず配置してください。"
            )
            for a in sakashita_alerts:
                st.markdown(
                    f"- **{a['date']}**: 出勤者 = "
                    f"{', '.join(a['staff']) or '（誰もいない）'}"
                )
        risk_messages.append('小金井坂下 単独オープンリスク')

    if not risk_messages:
        st.success("✅ 重大な運用リスクは検知されませんでした。")

    # ──── 総合機会損失額 ────
    st.markdown("---")
    st.markdown(
        f"### ⚠️ **{displayed_month}** の総合 機会損失額: "
        f"<span style='color:#dc2626; font-size:clamp(1.1em, 4vw, 1.5em); "
        f"font-weight:bold; display:inline-block; word-break:keep-all;'>"
        f"−{int(total_loss):,} 円"
        f"</span>",
        unsafe_allow_html=True,
    )
    st.caption(
        "※上記金額は「不足しているシフトがすべて埋まり、"
        "かつ予約がすべて埋まった場合」に取りこぼしている想定売上です。"
    )

    # ============================================================
    # 👥 人事・労務管理セクション
    # ============================================================
    st.markdown("---")
    st.markdown("## 👥 人事・労務管理")
    st.caption(
        f"対象月 **{displayed_month}** のシフト表から、有給取得・他院ヘルプ・"
        "固定休の実績を自動集計しています。"
    )

    paid_leave_df = data.get('paid_leave_df')
    help_actions_actual_df = data.get('help_actions_actual_df')
    fixed_leave_df = data.get('fixed_leave_df')

    hr_df, hr_stats = build_hr_summary(
        worked_df, paid_leave_df, help_actions_actual_df, fixed_leave_df,
    )

    if hr_df.empty:
        st.info("ℹ 当月のスタッフデータがまだ集計できません。")
    else:
        try:
            styled = hr_df.style.apply(_hr_style_highlight, axis=None)
            st.dataframe(styled, use_container_width=True, hide_index=True)
        except Exception:
            # スタイル失敗時は素のDataFrameで表示（最低限機能を担保）
            st.dataframe(hr_df, use_container_width=True, hide_index=True)

        sumcol1, sumcol2, sumcol3 = st.columns(3)
        sumcol1.metric(
            "🩺 有給取得ゼロ", f"{hr_stats['zero_paid_n']} 名",
            help="人事介入の優先対象（赤字ハイライト）",
        )
        sumcol2.metric(
            "🔁 ヘルプ過多", f"{hr_stats['many_help_n']} 名",
            help=f"月 {HELP_WARN_THRESHOLD} 回以上 — 患者エンゲージメント低下リスク",
        )
        sumcol3.metric(
            "🎉 固定休 (応援除外)", f"{hr_stats['fixed_days_n']} 日",
            help="結婚式・運動会等の年間予定で確保された休み。応援対象から自動除外。",
        )

        if isinstance(fixed_leave_df, pd.DataFrame) and len(fixed_leave_df) > 0:
            with st.expander("🎉 固定休（応援対象外）の一覧", expanded=False):
                st.dataframe(
                    fixed_leave_df[['date', 'area', 'staff_name', 'type']]
                    .sort_values(['date', 'area', 'staff_name'])
                    .reset_index(drop=True),
                    use_container_width=True, hide_index=True,
                )

        st.download_button(
            "📥 人事・労務サマリを CSV ダウンロード",
            data=hr_df.to_csv(index=False).encode('utf-8-sig'),
            file_name=f"hr_summary_{displayed_month}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    # ============================================================
    # 📦 結果の一括ダウンロード（手直し可能な Excel / CSV 形式）
    # ------------------------------------------------------------
    # 利用者が表計算ソフトで開いて自由に編集できるよう、PDF ではなく
    # Excel (.xlsx, 複数シート) と CSV (各データ別) で提供する。
    # ============================================================
    st.markdown("---")
    st.markdown("## 📦 結果をダウンロード（手直し可能な形式）")
    st.caption(
        "下記からファイルを取得して、Excel / 表計算ソフトで開いて自由に編集できます。"
        "用途に合わせて Excel（複数シート一括）または CSV（個別）を選んでください。"
    )

    dl_data_full = {
        'target_month': displayed_month,
        'shortages': data.get('shortages', {}),
        'surpluses': data.get('surpluses', {}),
        'unit_prices': data.get('unit_prices', {}),
        'staff_help_actions': data.get('staff_help_actions'),
        'gap_df': data.get('gap_df'),
        'paid_leave_df': data.get('paid_leave_df'),
        'help_actions_actual_df': data.get('help_actions_actual_df'),
        'fixed_leave_df': data.get('fixed_leave_df'),
        'analyzed_at': data.get('analyzed_at', ''),
    }

    dl_col1, dl_col2 = st.columns([1, 1])
    with dl_col1:
        st.markdown("**🟢 まとめてダウンロード（推奨）**")
        try:
            xlsx_bytes = build_results_xlsx_bytes(
                dl_data_full, hr_df if not hr_df.empty else None,
                CAPACITY_PER_DAY,
            )
            st.download_button(
                "📥 全結果を Excel (.xlsx) でダウンロード",
                data=xlsx_bytes,
                file_name=f"shift_analysis_{displayed_month}.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True,
                type="primary",
                help=(
                    "概要・応援アクション・過不足詳細・人事サマリ等が "
                    "1 つの xlsx に複数シートで入ります（編集自由）。"
                ),
            )
        except Exception as e:
            st.warning(f"Excel ファイル生成に失敗: {e}")

    with dl_col2:
        st.markdown("**🔵 個別ダウンロード（CSV）**")
        st.download_button(
            "📥 応援アクション指示 (CSV)",
            data=df_to_csv_bytes(data.get('staff_help_actions')),
            file_name=f"staff_help_actions_{displayed_month}.csv",
            mime="text/csv",
            use_container_width=True,
            disabled=(
                data.get('staff_help_actions') is None
                or len(data.get('staff_help_actions', [])) == 0
            ),
        )
        st.download_button(
            "📥 過不足詳細 (CSV)",
            data=df_to_csv_bytes(data.get('gap_df')),
            file_name=f"gap_detail_{displayed_month}.csv",
            mime="text/csv",
            use_container_width=True,
            disabled=(
                data.get('gap_df') is None
                or len(data.get('gap_df', [])) == 0
            ),
        )
        st.download_button(
            "📥 人事・労務サマリ (CSV)",
            data=df_to_csv_bytes(hr_df if not hr_df.empty else None),
            file_name=f"hr_summary_{displayed_month}.csv",
            mime="text/csv",
            use_container_width=True,
            disabled=hr_df.empty,
        )

    st.caption(
        "💡 **再計算したいときは**: 設定（対象月・ヘルプ要員選択など）を変更してから "
        "もう一度「▶ 分析する」を押してください。アップロード済みのファイルは"
        "保持されているので、再アップロードは不要です。"
    )

else:
    # 未実行時の案内
    st.info(
        "👋 ご利用にあたって\n\n"
        "1. 上の **「📤 ファイルをアップロード」** で 5 院ぶんのシフト Excel と"
        " マスター CSV を選択してください\n"
        "2. サイドバーで **対象月** を選んで、"
        f" **「▶ {target_month} を分析する」** ボタンを押してください\n"
        "3. 結果は **Excel (.xlsx) または CSV** でダウンロードできます\n\n"
        "🛡️ アップロードしたファイルは、サーバー側の一時メモリに"
        "置かれるだけで、PC 内の元ファイルには一切影響しません。"
    )
    st.markdown("#### 機能ガイド")
    st.markdown(
        "- 📅 **対象月**：サイドバーで YYYY-MM 形式で入力\n"
        "- 📤 **ファイルアップロード**：シフト Excel と CSV をブラウザから直接\n"
        "- 👥 **ヘルプ要員選択**：各院タブで「他院へ出せるスタッフ」を選択\n"
        "- ▶ **分析実行**：プライマリーボタン（青色）でワンクリック\n"
        "- 📥 **結果ダウンロード**：Excel（複数シート）または CSV（個別）"
    )
