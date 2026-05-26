# 🚀 Streamlit Community Cloud デプロイ手順書

各院の院長に **テスト利用してもらうための Web 公開** までの手順を、初心者の方でも進められるようステップバイステップで記載しています。

> **所要時間: 約 20〜30 分**（GitHub アカウントをすでにお持ちの場合は約 10 分）

---

## ✅ 事前準備チェックリスト

- [ ] GitHub アカウント（無料）— https://github.com/signup
- [ ] Streamlit Community Cloud アカウント（無料／GitHub アカウントで連携可）
- [ ] Git for Windows がインストールされている — https://git-scm.com/download/win
- [ ] このプロジェクト（`crm_scraper` フォルダ）が手元にある

---

## 📦 ステップ 1: ローカル準備（5 分）

### 1-1. PowerShell でプロジェクトフォルダを開く

```powershell
cd "C:\Users\中村文亮\Desktop\crm_scraper"
```

### 1-2. アップロードされるファイルを確認

`.gitignore` で個人情報（xlsx・csv・.env など）は自動除外されます。念のため、これからアップロードされる「予定」のファイル一覧を確認しておきましょう。

```powershell
git init
git status --short
```

> **コミット対象は `streamlit_app.py` / `shift_optimizer/*.py` / `requirements.txt` / `.gitignore` / `DEPLOY.md` 等だけ** になっているはずです。`*.xlsx` や `*.csv`、`.env` が含まれていないことを必ず確認してください。

### 1-3. ローカルコミット

```powershell
git add .
git commit -m "Initial commit: シフト最適化Webパネル"
```

---

## 🐙 ステップ 2: GitHub にプライベートリポジトリを作成（5 分）

### 2-1. GitHub にアクセス
ブラウザで https://github.com/new を開く。

### 2-2. リポジトリ設定

| 項目 | 設定値 |
|---|---|
| Repository name | `shift-optimizer-web`（任意） |
| Description | （任意）「シフト最適化 Web パネル」 |
| **Visibility** | **🔒 Private（必ず Private を選択）** |
| README / .gitignore / license | **すべて未チェック**（既にローカルにあるため） |

**「Create repository」** をクリック。

### 2-3. ローカル → GitHub にプッシュ

リポジトリ作成後に表示される「…or push an existing repository from the command line」のコマンドをコピペします。例:

```powershell
git remote add origin https://github.com/<あなたのGitHubユーザー名>/shift-optimizer-web.git
git branch -M main
git push -u origin main
```

> 初回プッシュ時にユーザー名・パスワード（または Personal Access Token）の入力を求められたら入力してください。

### 2-4. プッシュ後の確認

GitHub の Web 画面で、`streamlit_app.py` などが表示されていること、**xlsx・csv ファイルが表示されていないこと** を必ず確認してください。

---

## ☁️ ステップ 3: Streamlit Community Cloud にデプロイ（10 分）

### 3-1. Streamlit Cloud にサインイン

https://share.streamlit.io/ を開き、「**Continue with GitHub**」で先ほどのアカウントを連携。

### 3-2. 新しいアプリを作成

1. ダッシュボード右上の **「Create app」** → **「Deploy a public app from GitHub」** をクリック。
   （プライベートリポジトリでも、Streamlit Cloud に GitHub 連携済みなら選択可能。）

2. 以下を入力:

   | 項目 | 設定値 |
   |---|---|
   | Repository | `<あなたのユーザー名>/shift-optimizer-web` |
   | Branch | `main` |
   | Main file path | `streamlit_app.py` |
   | App URL (subdomain) | 任意（例: `shift-optimizer-<your-name>`） |

3. 右下の **「Deploy!」** ボタンをクリック。

> 初回ビルドは `requirements.txt` を `pip install` するため、**5〜10 分程度** かかります。ビルドログがリアルタイム表示されます。

### 3-3. 🔑 パスワードをシークレットに登録（重要）

デフォルトの `makoto2026` のままだと **このコードを見た人なら誰でも閲覧可能** になってしまいます。**必ず Streamlit Cloud 側でパスワードを上書き** してください。

1. デプロイ画面右上の「**⋮**」メニュー → **「Settings」** をクリック
2. 左サイドメニューで **「Secrets」** を選択
3. 以下を入力（**TOML 形式**）:

   ```toml
   APP_PASSWORD = "ここに本番用の長いパスワード"
   ```

   > 例: `APP_PASSWORD = "Koganei-Shift-2026-Spring"`（推測されにくい文字列を推奨）

4. **「Save」** をクリック → アプリが自動再起動

### 3-4. 動作確認

- 公開 URL（例: `https://shift-optimizer-xxx.streamlit.app/`）にアクセス
- 🔒 ロック画面が表示される → パスワード入力 → ログイン成功でメイン画面表示

---

## 📤 ステップ 4: 各院の院長に共有

院長にお知らせするテンプレート例:

```
件名: シフト最適化 Web パネル テスト公開のご連絡

各院長様

このたび、シフト最適化 Web パネルのテスト公開を開始しました。
スマホ・タブレット・PC のいずれからもアクセス可能です。

🌐 URL:       https://shift-optimizer-xxx.streamlit.app/
🔑 パスワード: <Streamlit Cloud で設定したパスワード>

機能:
 - 各院の不足/最適/過剰の色分け表示
 - ヘルプ要員の選択 (各院タブで複数選択可)
 - 有給取得状況・ヘルプ移動回数の可視化
 - 土曜のヘルプ要請漏れ・小金井坂下の単独オープン リスク警告

不具合・要望がございましたら、運用担当までご連絡ください。
```

---

## ⚠️ 注意事項

### データファイルについて
- 個人情報を含む xlsx・csv は `.gitignore` で除外しているため、**Streamlit Cloud 上のアプリにはデータが存在しません**。
- テスト公開時の動作確認は、UI の確認のみ可能です。
- 実データを使った分析を Cloud 上で行いたい場合は、別途以下のいずれかを検討してください:
  1. Google Drive 連携 (`google-api-python-client` 等の追加と認証設定)
  2. データファイルを Streamlit のアップロード機能 (`st.file_uploader`) でユーザがアップロードする
  3. 引き続きローカル PC で運用し、Streamlit Cloud はデモ用とする

### パスワード管理
- **`makoto2026` はソースコードに含まれている公開デフォルト** です。本番では必ず `st.secrets` で上書きしてください。
- パスワードを変更したい場合: Streamlit Cloud の Settings → Secrets で `APP_PASSWORD` を編集 → Save で即時反映。

### ログアウト
- ブラウザのタブを閉じれば自動的にログアウトされます（session_state がリセットされるため）。

---

## 🔄 デプロイ後にコードを更新するには

ローカルでコードを編集 → 以下のコマンドで GitHub にプッシュすれば、Streamlit Cloud が自動でアプリを再ビルド・再起動します。

```powershell
cd "C:\Users\中村文亮\Desktop\crm_scraper"
git add .
git commit -m "fix: 何を直したかメモ"
git push
```

---

## 🆘 トラブルシューティング

| 症状 | 対処 |
|---|---|
| Streamlit Cloud で ModuleNotFoundError | `requirements.txt` に該当パッケージを追記して再 push |
| パスワード画面で何度入力しても弾かれる | Streamlit Cloud の Secrets に保存した値とコピペした文字列が一致しているか確認（全半角スペース・改行に注意） |
| アプリが「No app data available」と出る | xlsx ファイルが Cloud に存在しないのが原因。UI 確認のみなら問題なし。実データ運用は上記「データファイルについて」を参照 |
| ビルドが終わらない（10 分以上） | リポジトリの「Reboot app」をクリック。それでもダメなら `requirements.txt` のバージョン指定を確認 |
