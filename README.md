# RogoAI Mail Hub v1

**Gmail洗浄機戦略：複数プロバイダメールを安全に一元管理**

[![YouTube Channel](https://img.shields.io/badge/YouTube-老後AI-red?logo=youtube)](https://www.youtube.com/@seniorAI-japan)
[![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

---

## 📖 概要

RogoAI Mail Hub v1.0は、複数のプロバイダメール（OCN、Nifty、So-net、Biglobe等）をGmailに転送し、ローカルで安全に管理するためのメールクライアントです。

### Gmail洗浄機戦略とは？

```
プロバイダーメール
    ↓ (転送設定)
Gmail (スパムフィルター = 洗浄機)
    ↓ (POP3受信)
Mail Hub (ローカル保管 = 安全)
    ↓ (SMTP送信)
元のアドレスで返信
```

1. **プロバイダーメールをGmailに転送**
   - Gmailの強力なスパムフィルターで「洗浄」
   
2. **Mail HubでPOP3受信**
   - 清らかなメールだけをローカル取得
   - SQLiteデータベースに安全保管
   
3. **ランサムウェア対策**
   - クラウドから切り離してローカル管理
   - オフラインでもメール閲覧可能
   
4. **元のアドレスで返信**
   - 受信者には元のプロバイダーアドレスから送信されたように見える

---

## ✨ 主な機能

- 📬 **複数アドレスの一元管理**: 転送されたメールを元のアドレス別に分類表示
- 🎨 **プロバイダ別色分け**: 一目でどのアドレス宛てか判別可能
- 📎 **添付ファイル対応**: 添付ファイルの保存・開封が可能
- 🔍 **検索機能**: 送信者、件名、本文から検索
- 💾 **ローカル保存**: メール本文をSQLiteに保存（オフライン閲覧可能）
- ✉️ **返信機能**: 元のアドレスから返信（SMTP経由）
- 🌐 **HTML表示**: HTMLメールを正しく表示（tkinterweb使用）

---

## 🚀 クイックスタート

### Windows向け（推奨）

#### 実行ファイル版（EXE）

1. [Releases](https://github.com/RogoAI-Takejij/251225_MailHub_GmailSQLite/releases)から最新版をダウンロード
   - `MailHub_v1.zip` (日本語版)

**注意：** 現在のバージョン（v13）は日本語版のみです。英語版は将来のバージョンで対応予定です。

2. ZIPを解凍して`MailHub_v1.exe`を実行

3. 初回起動時に設定を入力
   - Gmailアドレス
   - Gmailアプリパスワード
   - プロバイダー設定（転送元アドレス）

#### Python実行版

```bash
# リポジトリをクローン
git clone https://github.com/RogoAI-Takejij/251225_MailHub_GmailSQLite.git
cd 251225_MailHub_GmailSQLite

# 起動
python main.py
```

---

## ⚙️ 初回設定

### 1. Gmailアプリパスワードの取得

1. [Googleアカウント](https://myaccount.google.com/) → セキュリティ
2. **2段階認証を有効化**（必須）
3. 「アプリパスワード」を生成
   - アプリ: 「メール」
   - デバイス: 「Windowsパソコン」
4. 生成された**16桁のパスワード**をメモ

### 2. プロバイダーメールをGmailに転送

各プロバイダーの管理画面で、Gmailアドレスへの転送設定を行います。

**例: OCNメール**
1. OCNマイページ → メール設定
2. 転送設定 → 転送先にGmailアドレスを入力
3. 「元のメールも残す」を選択（推奨）

### 3. Mail Hub設定

1. Mail Hubを起動
2. 「設定」タブを開く
3. **Gmail設定**を入力
   - メールアドレス: `your-email@gmail.com`
   - アプリパスワード: `xxxx xxxx xxxx xxxx`（16桁）
   - POPサーバー: `pop.gmail.com`
   - POPポート: `995`
   - SMTPサーバー: `smtp.gmail.com`
   - SMTPポート: `587`

4. **プロバイダー設定**を追加
   - 転送元メールアドレス（例: `yourname@ocn.ne.jp`）
   - 表示色を選択

5. 「接続テスト」で確認

6. 「保存」をクリック

---

## 📁 ディレクトリ構造

```
251225_MailHub_GmailSQLite/
├── main.py                      # メインプログラム
├── icon.ico                     # アプリケーションアイコン
├── README.md                    # このファイル（日本語）
├── README_EN.md                 # 英語版README
├── LICENSE                      # MITライセンス
├── LICENSE_THIRD_PARTY.md       # サードパーティライセンス
├── CHANGELOG.md                 # 変更履歴
├── .gitignore                   # Git除外設定
│
├── lib/                         # 内包ライブラリ
│   ├── tkinterweb/             # HTML表示ライブラリ（MIT）
│   └── tkinterweb_tkhtml/      # HTML表示ライブラリ依存
│
├── config/                      # 設定ファイル（自動生成・Git除外）
│   └── config.json             # ユーザー設定
│
└── storage/                     # データ保存（自動生成・Git除外）
    └── emails.db               # メールデータベース
```

**注意**: 
- `config/`と`storage/`は個人情報を含むため、Gitには含まれません
- 初回起動時に自動生成されます

---

## 🛠️ 使い方

### メールの取得

1. 「メール」タブで「受信」ボタンをクリック
2. GmailからPOP3で最新メールをダウンロード
3. プロバイダ別に色分けされて表示

### メールの閲覧

- メール一覧から選択すると、下部に本文が表示されます
- HTMLメールは正しく整形されて表示されます
- 添付ファイルがある場合、「添付ファイルを開く」ボタンが表示されます

### メールの返信

1. 返信したいメールを選択
2. 「返信」ボタンをクリック
3. 本文を入力して送信
4. **元のプロバイダーアドレスから送信されます**（SMTP経由）

### 検索機能

- 送信者、件名、本文から検索可能
- 複数キーワードでの絞り込み対応

---

## ⚠️ 重要な注意事項

### 2026年以降のPOP3について

現在のバージョン（v13）はPOP3を使用していますが、**Gmail側のPOP3サポートは継続予定**です。

2026年1月に廃止されるのは：
- ❌ GmailのWeb UIで「他社メールをPOP3で取り込む機能」

継続するもの：
- ✅ 外部クライアント（Mail Hubなど）からGmailへのPOP3アクセス

詳細は[Googleの公式発表](https://support.google.com/mail/)をご確認ください。

### セキュリティ

- **アプリパスワードは厳重に管理してください**
- `config.json`には暗号化されたパスワードが保存されます
- このファイルは絶対に他人と共有しないでください

---

## 🐛 トラブルシューティング

### メールが取得できない

✅ **確認事項:**
- Gmailのアプリパスワードが正しいか
- 2段階認証が有効になっているか
- Gmail設定で「POP」が有効になっているか
  - Gmail設定 → 転送とPOP/IMAP → 「POPを有効にする」

### プロバイダが表示されない

✅ **確認事項:**
- 設定タブの「プロバイダ設定」で登録されているか
- Gmailへの転送設定が正しいか
- Gmailで実際にメールを受信しているか

### 添付ファイルが開けない

✅ **確認事項:**
- 一時ファイルの保存先に書き込み権限があるか
- ウイルス対策ソフトがブロックしていないか

### HTMLメールが表示されない

✅ **確認事項:**
- `lib/tkinterweb/`フォルダが存在するか
- Pythonバージョンが3.8以上か

---

## 🔧 上級者向けカスタマイズ

### 設定・データ保存場所の変更

デフォルトでは、Mail Hub本体（main.py）と同じフォルダ内の`config/`と`storage/`に保存されます。

**ほとんどのユーザーは変更不要です。** 特別な理由がある場合のみ、環境変数で保存場所を変更できます。

### 環境変数での設定

- `MAILHUB_CONFIG_DIR`: 設定ファイル（config.json）の保存先
- `MAILHUB_STORAGE_DIR`: メールデータベース（emails.db）の保存先

### 使用方法

#### Start_MailHub.bat（標準版）

通常のユーザーはこちらを使用：
```batch
Start_MailHub.bat
```
- デフォルトの場所（main.pyと同じフォルダ）に保存
- 設定不要で即座に使用可能

#### Start_MailHub_Custom.bat（上級者向け）

カスタム保存場所を指定したい場合：

1. `Start_MailHub_Custom.bat`をテキストエディタで開く
2. 以下の行のコメント（REM）を外して、パスを設定：

```batch
set MAILHUB_CONFIG_DIR=D:\MySettings\MailHub
set MAILHUB_STORAGE_DIR=D:\MyData\MailHub
```

3. `Start_MailHub_Custom.bat`を実行

### 注意事項

⚠️ **以下の場所は避けてください**:
- ネットワークドライブ（アクセス権限の問題）
- 日本語を含むパス（文字化けの可能性）
- システムフォルダ（権限エラー）
- 外付けドライブ（取り外すと起動不能）

✅ **推奨される場所**:
- ローカルディスク上の英数字のみのパス
- 書き込み権限のあるフォルダ

---

## 🔧 開発者向け情報

### 必要な環境

- Python 3.8以上
- 標準ライブラリのみ使用（追加インストール不要）

### 内包ライブラリ

このプロジェクトは以下のライブラリを`lib/`フォルダに内包しています：

- **tkinterweb**: HTMLメール表示用（MITライセンス）
- **tkinterweb_tkhtml**: tkinterwebの依存ライブラリ

詳細は`LICENSE_THIRD_PARTY.md`をご覧ください。

### 開発環境のセットアップ

```bash
# リポジトリをクローン
git clone https://github.com/RogoAI-Takejij/251225_MailHub_GmailSQLite.git
cd 251225_MailHub_GmailSQLite

# 実行
python main.py
```

### EXEビルド

```bash
# PyInstallerをインストール
pip install pyinstaller

# ビルドスクリプトを実行
build_exe.bat
```

---

## 📝 ライセンス

MIT License - 詳細は[LICENSE](LICENSE)をご覧ください。

サードパーティライブラリのライセンスは[LICENSE_THIRD_PARTY.md](LICENSE_THIRD_PARTY.md)をご覧ください。

---

## 👤 作者

**Created by Takejii (RogoAI)**

- YouTube: [@seniorAI-japan](https://www.youtube.com/@seniorAI-japan)
- GitHub: [RogoAI-Takeji](https://github.com/RogoAI-Takeji)

---

## 🙏 サポート

質問や問題がある場合：
1. [GitHub Issues](https://github.com/RogoAI-Takejij/251225_MailHub_GmailSQLite/issues)
2. [YouTube動画のコメント欄](https://www.youtube.com/@seniorAI-japan)

---

## 📚 関連動画

- [Mail Hub紹介動画](https://youtube.com/...)
- [Gmail洗浄機戦略の解説](https://youtube.com/...)
- [セキュリティ対策の重要性](https://youtube.com/...)

---

**Enjoy safe email management! 📧🔒**