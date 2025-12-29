import tkinter as tk
from tkinter import ttk, messagebox, filedialog, colorchooser
import imaplib
import smtplib
import email
from email.header import decode_header
from email.utils import parsedate_to_datetime, parseaddr
from datetime import timezone, timedelta, datetime
import sqlite3
import json
import os
import base64
import threading
import webbrowser
import tempfile
import re
import hashlib
import atexit
import urllib.parse
import urllib.request
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import sys

# ==========================================
# 内包ライブラリのパス追加
# ==========================================
# プロジェクト内のlibフォルダを優先的に読み込む
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LIB_DIR = os.path.join(BASE_DIR, "lib")
if os.path.exists(LIB_DIR):
    sys.path.insert(0, LIB_DIR)

# ==========================================
# 定数・設定 (ディレクトリ構成対応版)
# ==========================================
# 実行ファイルの場所を基準にする
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 環境変数でカスタマイズ可能（上級者向け）
# 通常ユーザーは何も設定せず、デフォルト（main.pyと同じフォルダ）を使用
CONFIG_DIR = os.getenv("MAILHUB_CONFIG_DIR") or os.path.join(BASE_DIR, "config")
STORAGE_DIR = os.getenv("MAILHUB_STORAGE_DIR") or os.path.join(BASE_DIR, "storage")

# ファイルパス
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.json")
DB_FILE = os.path.join(STORAGE_DIR, "emails.db")

# ディレクトリがない場合は自動作成（念のため）
os.makedirs(CONFIG_DIR, exist_ok=True)
os.makedirs(STORAGE_DIR, exist_ok=True)

DEFAULT_CONFIG = {
    "email": "",
    "password": "",
    "imap_server": "imap.gmail.com",
    "imap_folder": '"[Gmail]/&MFkweTBmMG4w4TD8MOs-"', 
    "providers": [],
    "provider_colors": {},  # 新規: プロバイダ別背景色
    "skip_html_warning": False,  # 新規: HTML完全版警告をスキップ
    # 取得範囲設定（新規）
    "first_launch": True,  # 初回起動フラグ
    "fetch_mode": "latest_only",  # デフォルト取得モード
    "custom_days": 30,  # カスタム日数
    # 自動取得設定（新規）
    "auto_fetch_on_startup": True,  # 起動時自動取得（デフォルトON）
    "auto_fetch_interval": False,  # 定期自動取得（デフォルトOFF）
    "auto_fetch_interval_minutes": 30,  # 定期取得間隔（分）
    # 上級者向け保存先カスタマイズ（新規）
    "storage_dir": None,  # None = デフォルト（main.pyと同じフォルダ）
    "config_dir": None,  # None = デフォルト（main.pyと同じフォルダ）
}

# ==========================================
# 1. ConfigManager
# ==========================================
class ConfigManager:
    def __init__(self, filepath):
        self.filepath = filepath
        self.config = self.load()

    def load(self):
        if not os.path.exists(self.filepath):
            return DEFAULT_CONFIG.copy()
        try:
            with open(self.filepath, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                if cfg.get("password"):
                    try:
                        cfg["password"] = base64.b64decode(cfg["password"]).decode("utf-8")
                    except:
                        pass
                return cfg
        except:
            return DEFAULT_CONFIG.copy()

    def save(self):
        save_data = self.config.copy()
        if save_data.get("password"):
            save_data["password"] = base64.b64encode(save_data["password"].encode("utf-8")).decode("utf-8")
        
        with open(self.filepath, "w", encoding="utf-8") as f:
            json.dump(save_data, f, indent=4, ensure_ascii=False)

    def get(self, key):
        return self.config.get(key)

    def set(self, key, value):
        self.config[key] = value

# ==========================================
# 2. DatabaseManager
# ==========================================
class DatabaseManager:
    def __init__(self, db_path):
        self.db_path = db_path
        self.init_db()

    def init_db(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        
        # テーブル作成（存在しない場合のみ）
        c.execute("""
            CREATE TABLE IF NOT EXISTS emails (
                message_id TEXT PRIMARY KEY,
                original_to TEXT,
                subject TEXT,
                sender TEXT,
                date_disp TEXT,
                timestamp TEXT,
                raw_data TEXT
            )
        """)
        
        # プロモルールテーブル作成
        c.execute("""
            CREATE TABLE IF NOT EXISTS promo_rules (
                rule_id INTEGER PRIMARY KEY AUTOINCREMENT,
                sender_pattern TEXT UNIQUE,
                added_date TEXT,
                match_count INTEGER DEFAULT 0
            )
        """)
        
        # マイグレーション: providerカラム追加
        needs_provider_migration = False
        try:
            c.execute("SELECT provider FROM emails LIMIT 1")
        except sqlite3.OperationalError:
            print("[MIGRATION] Adding 'provider' column...")
            c.execute("ALTER TABLE emails ADD COLUMN provider TEXT")
            needs_provider_migration = True
        
        # マイグレーション: read_flagカラム追加
        try:
            c.execute("SELECT read_flag FROM emails LIMIT 1")
        except sqlite3.OperationalError:
            print("[MIGRATION] Adding 'read_flag' column...")
            c.execute("ALTER TABLE emails ADD COLUMN read_flag INTEGER DEFAULT 0")
        
        # マイグレーション: is_promoカラム追加
        try:
            c.execute("SELECT is_promo FROM emails LIMIT 1")
        except sqlite3.OperationalError:
            print("[MIGRATION] Adding 'is_promo' column...")
            c.execute("ALTER TABLE emails ADD COLUMN is_promo INTEGER DEFAULT 0")
        
        # マイグレーション: folderカラム追加
        try:
            c.execute("SELECT folder FROM emails LIMIT 1")
        except sqlite3.OperationalError:
            print("[MIGRATION] Adding 'folder' column...")
            c.execute("ALTER TABLE emails ADD COLUMN folder TEXT DEFAULT NULL")
        
        # マイグレーション: is_repliedカラム追加
        try:
            c.execute("SELECT is_replied FROM emails LIMIT 1")
        except sqlite3.OperationalError:
            print("[MIGRATION] Adding 'is_replied' column...")
            c.execute("ALTER TABLE emails ADD COLUMN is_replied INTEGER DEFAULT 0")
        
        # マイグレーション: is_deletedカラム追加
        try:
            c.execute("SELECT is_deleted FROM emails LIMIT 1")
        except sqlite3.OperationalError:
            print("[MIGRATION] Adding 'is_deleted' column...")
            c.execute("ALTER TABLE emails ADD COLUMN is_deleted INTEGER DEFAULT 0")
        
        # マイグレーション: promo_rules.target_folderカラム追加
        try:
            c.execute("SELECT target_folder FROM promo_rules LIMIT 1")
        except sqlite3.OperationalError:
            print("[MIGRATION] Adding 'target_folder' column to promo_rules...")
            c.execute("ALTER TABLE promo_rules ADD COLUMN target_folder TEXT DEFAULT NULL")
        
        # foldersテーブル作成
        c.execute("""
            CREATE TABLE IF NOT EXISTS folders (
                folder_id INTEGER PRIMARY KEY AUTOINCREMENT,
                provider TEXT NOT NULL,
                folder_name TEXT NOT NULL,
                folder_type TEXT DEFAULT 'custom',
                created_date TEXT DEFAULT (datetime('now')),
                UNIQUE(provider, folder_name)
            )
        """)
        
        # deleted_messagesテーブル作成
        c.execute("""
            CREATE TABLE IF NOT EXISTS deleted_messages (
                message_id TEXT PRIMARY KEY,
                deleted_date TEXT DEFAULT (datetime('now')),
                delete_mode TEXT
            )
        """)
        
        # マイグレーション: attachmentsカラム追加
        try:
            c.execute("SELECT attachments FROM emails LIMIT 1")
        except sqlite3.OperationalError:
            print("[MIGRATION] Adding 'attachments' column...")
            c.execute("ALTER TABLE emails ADD COLUMN attachments TEXT")  # JSON形式で保存
            conn.commit()
        
        conn.commit()
        
        # 既存データのproviderカラムを埋める
        if needs_provider_migration:
            print("[MIGRATION] Populating provider column for existing emails...")
            c.execute("SELECT message_id, original_to FROM emails WHERE provider IS NULL")
            rows = c.fetchall()
            
            # MailFetcherインスタンス作成
            from __main__ import MailFetcher
            fetcher = MailFetcher()
            
            for msg_id, orig_to in rows:
                provider = fetcher.extract_provider(orig_to)
                c.execute("UPDATE emails SET provider=? WHERE message_id=?", (provider, msg_id))
            conn.commit()
            print(f"[MIGRATION] Updated {len(rows)} emails with provider info")
        
        # 既存データのprovider修正（修飾子除去）
        print("[MIGRATION] Checking for provider data with decorations...")
        c.execute("SELECT COUNT(*) FROM emails WHERE provider LIKE '%>%' OR provider LIKE '%<%'")
        dirty_count = c.fetchone()[0]
        
        if dirty_count > 0:
            print(f"[MIGRATION] Found {dirty_count} emails with malformed provider data. Fixing...")
            
            c.execute("SELECT message_id, original_to FROM emails")
            all_rows = c.fetchall()
            
            from __main__ import MailFetcher
            fetcher = MailFetcher()
            
            fixed_count = 0
            for msg_id, orig_to in all_rows:
                if orig_to:
                    correct_provider = fetcher.extract_provider(orig_to)
                    c.execute("UPDATE emails SET provider=? WHERE message_id=?", (correct_provider, msg_id))
                    if c.rowcount > 0:
                        fixed_count += 1
            
            conn.commit()
            print(f"[MIGRATION] Fixed {fixed_count} provider entries")
        
        conn.close()

    def reset_db(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("DROP TABLE IF EXISTS emails")
        conn.commit()
        conn.close()
        self.init_db()

    def save_emails(self, email_list):
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        new_count = 0
        
        # MailFetcherインスタンス作成（extract_provider使用のため）
        from __main__ import MailFetcher
        fetcher = MailFetcher()
        
        print(f"[DEBUG] save_emails: 受信したメール数={len(email_list)}件")
        
        # プロモルール取得（target_folder情報も含む）
        cur.execute("SELECT sender_pattern, target_folder FROM promo_rules")
        promo_rules = cur.fetchall()
        
        try:
            for item in email_list:
                # プロバイダ抽出（修飾子除去＆正規化）
                provider = fetcher.extract_provider(item["to"])
                
                # プロモ判定
                is_promo = 0
                target_folder = None
                sender_clean = fetcher.clean_address(item["from"])
                
                for pattern, folder in promo_rules:
                    if self.match_pattern(sender_clean, pattern):
                        is_promo = 1
                        target_folder = folder  # フォルダ情報も取得
                        # マッチカウント更新
                        cur.execute("UPDATE promo_rules SET match_count = match_count + 1 WHERE sender_pattern=?", (pattern,))
                        break
                
                # 添付ファイル情報を抽出
                import email
                msg = email.message_from_string(item["raw_data"])
                attachments = fetcher.extract_attachments(msg)
                
                # 添付ファイル情報をJSON形式で保存
                import json
                attachments_json = json.dumps(attachments, ensure_ascii=False) if attachments else None
                
                try:
                    cur.execute("""
                        INSERT INTO emails (message_id, original_to, subject, sender, date_disp, timestamp, raw_data, provider, read_flag, is_promo, folder, attachments)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, 0, ?, ?, ?)
                    """, (
                        item["message_id"], item["to"], item["subject"], item["from"], 
                        item["date_disp"], item["timestamp"], item["raw_data"], provider, is_promo, target_folder, attachments_json
                    ))
                    if cur.rowcount > 0:
                        new_count += 1
                except sqlite3.IntegrityError:
                    pass 
            conn.commit()
            print(f"[DEBUG] save_emails完了: 新規保存={new_count}件、重複スキップ={len(email_list) - new_count}件")
        except Exception as e:
            print(f"[DEBUG] DB Save Error: {e}")
        finally:
            conn.close()
        return new_count
    
    def match_pattern(self, text, pattern):
        """SQL LIKEパターンマッチング（%をワイルドカードとして扱う）"""
        import re
        # %を.*に、_を.に変換
        regex = pattern.replace("%", ".*").replace("_", ".")
        return bool(re.search(regex, text, re.IGNORECASE))

    def mark_as_read(self, message_id):
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        cur.execute("UPDATE emails SET read_flag=1 WHERE message_id=?", (message_id,))
        conn.commit()
        conn.close()

    def get_providers(self):
        """登録されている全プロバイダのリスト取得"""
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        try:
            cur.execute("SELECT DISTINCT provider FROM emails WHERE provider IS NOT NULL ORDER BY provider")
            providers = [row[0] for row in cur.fetchall()]
        except sqlite3.OperationalError:
            # カラムがまだ存在しない場合（マイグレーション前）
            providers = []
        conn.close()
        return providers

    def get_last_fetch_time(self):
        """最後に取得したメールの日時（IMAP検索用）"""
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        cur.execute("SELECT MAX(timestamp) FROM emails")
        result = cur.fetchone()[0]
        conn.close()
        
        if not result:
            return None
        
        # ISO形式 → IMAP検索形式に変換
        # "2024-12-24T10:30:00+09:00" → "24-Dec-2024"
        try:
            dt = datetime.fromisoformat(result)
            return dt.strftime("%d-%b-%Y")
        except:
            return None

    def get_email_count(self):
        """現在保存されているメールの総数"""
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM emails")
        count = cur.fetchone()[0]
        conn.close()
        return count

    def get_oldest_email_time(self):
        """最も古いメールの日時"""
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        cur.execute("SELECT MIN(timestamp) FROM emails")
        result = cur.fetchone()[0]
        conn.close()
        return result if result else "なし"
    
    def create_folder(self, provider, folder_name, folder_type='custom'):
        """フォルダ作成"""
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        try:
            cur.execute("""
                INSERT INTO folders (provider, folder_name, folder_type)
                VALUES (?, ?, ?)
            """, (provider, folder_name, folder_type))
            conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False
        finally:
            conn.close()
    
    def get_folders(self, provider):
        """プロバイダのフォルダ一覧取得"""
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        cur.execute("""
            SELECT folder_name, folder_type FROM folders 
            WHERE provider=? ORDER BY folder_type, folder_name
        """, (provider,))
        folders = cur.fetchall()
        conn.close()
        return folders
    
    def delete_folder(self, provider, folder_name):
        """フォルダ削除"""
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        cur.execute("DELETE FROM folders WHERE provider=? AND folder_name=?", (provider, folder_name))
        conn.commit()
        conn.close()
    
    def move_to_folder(self, message_id, folder_path):
        """メールをフォルダへ移動"""
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        cur.execute("UPDATE emails SET folder=? WHERE message_id=?", (folder_path, message_id))
        conn.commit()
        conn.close()
    
    def mark_as_replied(self, message_id):
        """返信済みフラグを立てる（フォルダ移動なし）"""
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        cur.execute("UPDATE emails SET is_replied=1 WHERE message_id=?", (message_id,))
        conn.commit()
        conn.close()
    
    def save_sent_email(self, email_data):
        """送信済みメールをDBに保存"""
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        cur.execute("""
            INSERT OR REPLACE INTO emails (
                message_id, original_to, subject, sender, 
                date_disp, timestamp, raw_data, provider, folder, is_replied
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            email_data["message_id"],
            email_data["original_to"],
            email_data["subject"],
            email_data["sender"],
            email_data["date_disp"],
            email_data["timestamp"],
            email_data["raw_data"],
            email_data["provider"],
            "__sent__",  # 送信済みフォルダ
            0  # is_replied（送信メール自体は返信済みではない）
        ))
        conn.commit()
        conn.close()
    
    def get_deleted_message_ids(self):
        """削除済みメッセージIDリスト"""
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        cur.execute("SELECT message_id FROM deleted_messages")
        ids = [row[0] for row in cur.fetchall()]
        conn.close()
        return ids
    
    def permanently_delete_email(self, message_id, delete_mode, config):
        """メールを完全削除"""
        from datetime import datetime
        
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        
        # emailsテーブルから削除
        cur.execute("DELETE FROM emails WHERE message_id=?", (message_id,))
        
        # deleted_messagesに登録
        cur.execute("""
            INSERT OR REPLACE INTO deleted_messages (message_id, deleted_date, delete_mode)
            VALUES (?, ?, ?)
        """, (message_id, datetime.now().isoformat(), delete_mode))
        
        conn.commit()
        conn.close()
        
        # Gmail Cloudからも削除（モードBの場合）
        if delete_mode == "gmail_cloud":
            try:
                self._delete_from_gmail_cloud(message_id, config)
            except Exception as e:
                print(f"[WARNING] Gmail Cloud削除失敗: {e}")
    
    def _delete_from_gmail_cloud(self, message_id, config):
        """Gmail CloudからIMAP削除"""
        import imaplib
        
        email_addr = config.get("email")
        password = config.get("password")
        server = config.get("imap_server") or "imap.gmail.com"
        
        if not email_addr or not password:
            return
        
        try:
            mail = imaplib.IMAP4_SSL(server)
            mail.login(email_addr, password)
            mail.select('"[Gmail]/All Mail"')
            
            # Message-IDで検索
            status, data = mail.search(None, f'HEADER Message-ID "{message_id}"')
            
            if data and data[0]:
                mail_ids = data[0].split()
                for mail_id in mail_ids:
                    mail.store(mail_id, '+FLAGS', '\\Deleted')
                mail.expunge()
            
            mail.logout()
        except Exception as e:
            raise Exception(f"IMAP削除失敗: {e}")

# ==========================================
# 3. MailFetcher
# ==========================================
class MailFetcher:
    def __init__(self):
        pass

    def decode_h(self, header_val):
        if not header_val: return ""
        dec = decode_header(header_val)
        ret = ""
        for b, enc in dec:
            if isinstance(b, bytes):
                try: ret += b.decode(enc if enc else 'utf-8', 'replace')
                except: ret += b.decode('utf-8', 'replace')
            else: ret += str(b)
        return ret

    def clean_address(self, addr_str):
        """
        メールアドレスから修飾子を完全除去＆正規化
        
        入力例:
        - "Taro Tanaka <taro@example.com>"
        - "<taro@example.com>"
        - "taro@example.com>"
        - "TARO@EXAMPLE.COM"
        
        出力: "taro@example.com" (全て統一)
        """
        if not addr_str:
            return ""
        
        # parseaddrで名前とアドレスを分離
        name, addr = parseaddr(addr_str)
        
        # アドレス部分を取得（失敗時は元の文字列）
        result = addr if addr else addr_str
        
        # 念のため、残存する<>を除去
        result = result.strip().replace("<", "").replace(">", "")
        
        # 空白除去
        result = result.strip()
        
        # 小文字統一（大文字小文字の揺れ対策）
        return result.lower()
    
    def extract_provider(self, email_address):
        """
        メールアドレスからプロバイダ（ドメイン）を抽出
        
        入力: "fager@roy.hi-ho.ne.jp" or "<fager@roy.hi-ho.ne.jp>"
        出力: "roy.hi-ho.ne.jp"
        """
        clean_addr = self.clean_address(email_address)
        
        if not clean_addr or "@" not in clean_addr:
            return "unknown"
        
        # @以降を取得
        domain = clean_addr.split("@")[-1]
        
        # 余分な文字を除去（念のため）
        domain = domain.strip().lower()
        
        return domain

    def fetch_central(self, config, limit=None, progress_callback=None):
        """Gmail IMAP経由でメール取得（取得範囲設定対応）"""
        print("[DEBUG] fetch_central started.")
        email_addr = config.get("email")
        password = config.get("password")
        server = config.get("imap_server")
        folder = config.get("imap_folder")
        fetch_range = config.get("fetch_range") or "week"  # デフォルト: 過去1週間

        if not email_addr or not password:
            raise ValueError("メールアドレスまたはパスワードが設定されていません")

        fetched_data = []
        mail = imaplib.IMAP4_SSL(server)
        mail.login(email_addr, password)
        mail.select(folder)

        # 取得範囲に応じた検索条件を構築
        from datetime import datetime, timedelta
        
        search_criteria = "ALL"
        
        if fetch_range == "latest":
            # 最新のみ → 前回取得以降の未取得メールをすべて取得
            
            import sqlite3
            try:
                conn_db = sqlite3.connect(DB_FILE)
                cur_db = conn_db.cursor()
                
                # DBの最新メールのタイムスタンプとmessage_idを取得
                cur_db.execute("SELECT MAX(timestamp), message_id FROM emails WHERE timestamp = (SELECT MAX(timestamp) FROM emails)")
                row = cur_db.fetchone()
                max_timestamp = row[0] if row and row[0] else None
                
                # すでに取得済みのmessage_idリストを取得
                cur_db.execute("SELECT message_id FROM emails")
                existing_ids = set(r[0] for r in cur_db.fetchall())
                conn_db.close()
                
                if max_timestamp:
                    # 既存メールがある場合：最新タイムスタンプ以降のメールを取得
                    from datetime import datetime
                    
                    # ISO形式のタイムスタンプをdatetimeに変換
                    max_dt = datetime.fromisoformat(max_timestamp)
                    
                    # IMAP検索用の日付形式に変換（その日以降）
                    search_date = max_dt.strftime("%d-%b-%Y")
                    
                    status, messages = mail.search(None, f'SINCE {search_date}')
                    
                    if not messages[0]:
                        print(f"[DEBUG] 最新のみモード: 新着メールなし（最終取得: {max_timestamp}）")
                        mail.logout()
                        return []
                    
                    all_ids = messages[0].split()
                    
                    # 各メールのMessage-IDをチェックして未取得のみ抽出
                    latest_ids = []
                    checked_count = 0
                    
                    for eid in all_ids:
                        # メールヘッダーのみ取得（高速化）
                        res, msg_data = mail.fetch(eid, "(BODY.PEEK[HEADER])")
                        if msg_data and msg_data[0]:
                            header = msg_data[0][1]
                            temp_msg = email.message_from_bytes(header)
                            check_msg_id = temp_msg.get("Message-ID")
                            
                            checked_count += 1
                            
                            # DBに存在しないMessage-IDのみ追加
                            if check_msg_id and check_msg_id not in existing_ids:
                                latest_ids.append(eid)
                    
                    if not latest_ids:
                        print(f"[DEBUG] 最新のみモード: 新着メールなし（{checked_count}件チェック済み）")
                        mail.logout()
                        return []
                    
                    print(f"[DEBUG] 最新のみモード: {checked_count}件中、未取得{len(latest_ids)}件を取得")
                else:
                    # 初回起動：最新100件のみ取得
                    status, messages = mail.search(None, "ALL")
                    if not messages[0]:
                        mail.logout()
                        return []
                    all_ids = messages[0].split()
                    latest_ids = all_ids[-100:]
                    print(f"[DEBUG] 最新のみモード（初回）: 最新100件を取得")
                    
            except Exception as e:
                print(f"[WARNING] DB最新タイムスタンプ取得エラー: {e}")
                import traceback
                traceback.print_exc()
                # エラー時は最新100件取得
                status, messages = mail.search(None, "ALL")
                if not messages[0]:
                    mail.logout()
                    return []
                id_list = messages[0].split()
                latest_ids = id_list[-100:]
                print(f"[DEBUG] 最新のみモード（フォールバック）: 最新100件を取得")
                id_list = messages[0].split()
                latest_ids = id_list[-100:]
                print(f"[DEBUG] 最新のみモード（フォールバック）: 最新100件を取得")
        elif fetch_range == "week":
            # 過去1週間（件数制限なし）
            since_date = (datetime.now() - timedelta(days=7)).strftime("%d-%b-%Y")
            search_criteria = f'SINCE {since_date}'
            status, messages = mail.search(None, search_criteria)
            if not messages[0]:
                mail.logout()
                return []
            latest_ids = messages[0].split()
            print(f"[DEBUG] 過去1週間モード: {len(latest_ids)}件")
        elif fetch_range == "month":
            # 過去1ヶ月（件数制限なし）
            since_date = (datetime.now() - timedelta(days=30)).strftime("%d-%b-%Y")
            search_criteria = f'SINCE {since_date}'
            print(f"[DEBUG] 過去1ヶ月モード: 検索条件={search_criteria}")
            status, messages = mail.search(None, search_criteria)
            if not messages[0]:
                mail.logout()
                return []
            latest_ids = messages[0].split()
            print(f"[DEBUG] 過去1ヶ月モード: Gmail IMAP検索結果={len(latest_ids)}件")
        elif fetch_range == "3months":
            # 過去3ヶ月
            since_date = (datetime.now() - timedelta(days=90)).strftime("%d-%b-%Y")
            search_criteria = f'SINCE {since_date}'
            status, messages = mail.search(None, search_criteria)
            if not messages[0]:
                mail.logout()
                return []
            latest_ids = messages[0].split()
        elif fetch_range == "year":
            # 過去1年
            since_date = (datetime.now() - timedelta(days=365)).strftime("%d-%b-%Y")
            search_criteria = f'SINCE {since_date}'
            status, messages = mail.search(None, search_criteria)
            if not messages[0]:
                mail.logout()
                return []
            latest_ids = messages[0].split()
        elif fetch_range == "all":
            # すべて
            status, messages = mail.search(None, "ALL")
            if not messages[0]:
                mail.logout()
                return []
            latest_ids = messages[0].split()
        elif fetch_range == "custom":
            # カスタム期間
            custom_days = config.get("custom_days") or 30
            since_date = (datetime.now() - timedelta(days=int(custom_days))).strftime("%d-%b-%Y")
            search_criteria = f'SINCE {since_date}'
            status, messages = mail.search(None, search_criteria)
            if not messages[0]:
                mail.logout()
                return []
            latest_ids = messages[0].split()
        else:
            # デフォルト: 最新200件
            status, messages = mail.search(None, "ALL")
            if not messages[0]:
                mail.logout()
                return []
            id_list = messages[0].split()
            latest_ids = id_list[-200:]
        
        # limit指定がある場合は制限（後方互換性）
        if limit:
            latest_ids = latest_ids[-limit:]
        
        total_count = len(latest_ids)
        print(f"[DEBUG] Fetching {total_count} emails (range: {fetch_range})")

        for idx, eid in enumerate(latest_ids, 1):
            # 進捗コールバック
            if progress_callback:
                progress_callback(idx, total_count)
            
            res, msg_data = mail.fetch(eid, "(RFC822)")
            raw = msg_data[0][1]
            msg = email.message_from_bytes(raw)

            subj = self.decode_h(msg.get("Subject"))
            frm = self.decode_h(msg.get("From"))
            msg_id = msg.get("Message-ID")
            
            date_str_raw = msg.get("Date")
            dt_object = datetime.now(timezone.utc)
            date_display = "(No Date)"
            try:
                if date_str_raw:
                    dt_obj = parsedate_to_datetime(date_str_raw)
                    JST = timezone(timedelta(hours=9))
                    dt_jst = dt_obj.astimezone(JST)
                    date_display = dt_jst.strftime("%Y/%m/%d %H:%M:%S")
                    dt_object = dt_jst
            except:
                pass
            timestamp_str = dt_object.isoformat()

            if not msg_id: msg_id = f"NOID_{timestamp_str}_{frm}"

            to_raw = msg.get("To")
            to_dec = self.decode_h(to_raw)
            if email_addr in to_dec:
                disp_to = f"{email_addr} (Direct)"
            else:
                disp_to = to_dec 

            fetched_data.append({
                "message_id": msg_id,
                "to": disp_to,
                "subject": subj,
                "from": frm,
                "date_disp": date_display,
                "timestamp": timestamp_str,
                "raw_data": raw.decode("utf-8", errors="replace")
            })

        mail.logout()
        return fetched_data

    def test_connection_imap(self, server, email_addr, password):
        mail = imaplib.IMAP4_SSL(server)
        mail.login(email_addr, password)
        mail.logout()

    def test_connection_smtp(self, host, port, email_addr, password):
        """SMTP接続テスト（タイムアウト10秒、ポート465対応）"""
        import socket
        
        port = int(port)
        
        # タイムアウト設定
        original_timeout = socket.getdefaulttimeout()
        socket.setdefaulttimeout(10)
        
        try:
            if port == 465:
                # SSL接続
                server = smtplib.SMTP_SSL(host, port, timeout=10)
            else:
                # TLS接続
                server = smtplib.SMTP(host, port, timeout=10)
                server.starttls()
            
            server.login(email_addr, password)
            server.quit()
        finally:
            # タイムアウトを元に戻す
            socket.setdefaulttimeout(original_timeout)

    def send_email(self, provider_config, global_config, to, subject, body, attachments=None):
        use_fallback = provider_config.get("fallback_gmail", False)
        
        if use_fallback:
            smtp_host = "smtp.gmail.com"
            smtp_port = 587
            smtp_user = global_config.get("email")
            smtp_pass = global_config.get("password")
            from_addr = provider_config["email"]
        else:
            smtp_host = provider_config.get("smtp_host")
            smtp_port = int(provider_config.get("smtp_port", 587))
            smtp_user = provider_config["email"]
            smtp_pass = provider_config["password"]
            from_addr = provider_config["email"]

        msg = MIMEMultipart()
        msg["From"] = from_addr
        msg["To"] = to
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))

        if attachments:
            for fpath in attachments:
                with open(fpath, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(fpath)}")
                msg.attach(part)

        server = smtplib.SMTP(smtp_host, smtp_port)
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)
        server.quit()
        return (True, "送信成功")
    
    def extract_attachments(self, msg):
        """添付ファイル情報を抽出"""
        attachments = []
        
        if msg.is_multipart():
            for part in msg.walk():
                content_disposition = part.get("Content-Disposition", "")
                
                # 添付ファイルの判定
                if "attachment" in content_disposition:
                    filename = part.get_filename()
                    
                    if filename:
                        # ファイル名をデコード
                        decoded_filename = self.decode_h(filename)
                        
                        # サイズ取得
                        payload = part.get_payload(decode=True)
                        size = len(payload) if payload else 0
                        
                        # Content-Type取得
                        content_type = part.get_content_type()
                        
                        attachments.append({
                            "filename": decoded_filename,
                            "size": size,
                            "content_type": content_type
                        })
        
        return attachments

    def extract_text_body(self, msg):
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                ctype = part.get_content_type()
                if ctype == "text/plain":
                    payload = part.get_payload(decode=True)
                    charset = part.get_content_charset() or 'utf-8'
                    try: body += payload.decode(charset, errors='replace')
                    except: body += payload.decode('utf-8', errors='replace')
        else:
            payload = msg.get_payload(decode=True)
            charset = msg.get_content_charset() or 'utf-8'
            try: body = payload.decode(charset, errors='replace')
            except: body = payload.decode('utf-8', errors='replace')
        return body

    def extract_html_body(self, msg):
        html = ""
        if msg.is_multipart():
            for part in msg.walk():
                ctype = part.get_content_type()
                if ctype == "text/html":
                    payload = part.get_payload(decode=True)
                    charset = part.get_content_charset() or 'utf-8'
                    try: html += payload.decode(charset, errors='replace')
                    except: html += payload.decode('utf-8', errors='replace')
        else:
            if msg.get_content_type() == "text/html":
                payload = msg.get_payload(decode=True)
                charset = msg.get_content_charset() or 'utf-8'
                try: html = payload.decode(charset, errors='replace')
                except: html = payload.decode('utf-8', errors='replace')
        return html

# ==========================================
# 4. MailViewer (HTMLレンダリング対応)
# ==========================================
class MailViewer:
    temp_files = []  # クラス変数で一時ファイル追跡
    
    def __init__(self, parent, raw_data, subject, config_mgr=None, message_id=None, attachments=None):
        self.raw_data = raw_data
        self.subject = subject
        self.fetcher = MailFetcher()
        self.config_mgr = config_mgr  # 設定管理を保持
        self.message_id = message_id  # メッセージID
        self.attachments = attachments or []  # 添付ファイル情報
        
        self.viewer = tk.Toplevel(parent)
        self.viewer.title(f"メール詳細: {subject}")
        self.viewer.geometry("800x650")
        
        # 上部: 表示モード選択
        mode_frame = tk.Frame(self.viewer, bg="#f0f0f0")
        mode_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(mode_frame, text="表示モード:", bg="#f0f0f0").pack(side=tk.LEFT, padx=10)
        tk.Button(mode_frame, text="📄 テキスト版", command=lambda: self.switch_mode("text")).pack(side=tk.LEFT, padx=2)
        tk.Button(mode_frame, text="🌐 HTML版（安全）", command=lambda: self.switch_mode("html_safe")).pack(side=tk.LEFT, padx=2)
        tk.Button(mode_frame, text="🌐💀 HTML版（完全）", command=lambda: self.switch_mode("html_full")).pack(side=tk.LEFT, padx=2)
        
        # 添付ファイル表示
        if self.attachments:
            attachment_frame = tk.Frame(self.viewer, bg="#e3f2fd", relief=tk.RIDGE, bd=2)
            attachment_frame.pack(fill=tk.X, padx=10, pady=5)
            
            tk.Label(attachment_frame, text=f"📎 添付ファイル ({len(self.attachments)}件)", 
                    bg="#e3f2fd", font=("Arial", 10, "bold")).pack(anchor=tk.W, padx=10, pady=5)
            
            for att in self.attachments:
                filename = att.get("filename", "unknown")
                size = att.get("size", 0)
                content_type = att.get("content_type", "")
                size_str = f"{size / 1024:.1f} KB" if size < 1024*1024 else f"{size / (1024*1024):.1f} MB"
                
                att_row = tk.Frame(attachment_frame, bg="#e3f2fd")
                att_row.pack(fill=tk.X, padx=10, pady=2)
                
                tk.Label(att_row, text=f"📄 {filename} ({size_str})", bg="#e3f2fd").pack(side=tk.LEFT)
                
                # ボタンを右から左の順に配置（pack side=RIGHTは右から順に並ぶ）
                # 画像ファイルの場合はプレビューボタンも追加
                if content_type.startswith("image/"):
                    tk.Button(att_row, text="👁️ 表示", command=lambda f=filename: self.preview_attachment(f), 
                             bg="#2196F3", fg="white", width=8).pack(side=tk.RIGHT, padx=2)
                
                tk.Button(att_row, text="💾 保存", command=lambda f=filename: self.save_attachment(f), 
                         bg="#4CAF50", fg="white", width=8).pack(side=tk.RIGHT, padx=2)
                tk.Button(att_row, text="📂 開く", command=lambda f=filename: self.open_attachment(f), 
                         bg="#FF9800", fg="white", width=8).pack(side=tk.RIGHT, padx=2)
        
        # コンテンツエリア
        self.content_frame = tk.Frame(self.viewer)
        self.content_frame.pack(fill=tk.BOTH, expand=True)
        
        # デフォルトはテキスト版（安全）
        try:
            self.switch_mode("html_safe")
        except Exception as e:
            print(f"[WARNING] HTML表示に失敗、テキストモードで開きます: {e}")
            self.switch_mode("text")
    
    def switch_mode(self, mode):
        # 既存コンテンツクリア
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        
        msg = email.message_from_string(self.raw_data)
        
        if mode == "text":
            self.show_text_mode(msg)
        elif mode == "html_safe":
            self.show_html_safe_mode(msg)
        elif mode == "html_full":
            self.show_html_full_mode(msg)
    
    def show_text_mode(self, msg):
        """プレーンテキスト表示"""
        body = self.fetcher.extract_text_body(msg)
        
        text_widget = tk.Text(self.content_frame, wrap=tk.WORD, font=("Arial", 10))
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # ヘッダー情報
        header = f"件名: {self.subject}\n"
        header += f"送信者: {self.fetcher.decode_h(msg.get('From'))}\n"
        header += f"日付: {self.fetcher.decode_h(msg.get('Date'))}\n"
        header += "=" * 50 + "\n\n"
        
        text_widget.insert("1.0", header + body)
        text_widget.config(state=tk.DISABLED)
    
    def show_html_safe_mode(self, msg):
        """HTML安全版（スクリプト除去）"""
        html_body = self.fetcher.extract_html_body(msg)
        
        if not html_body:
            # HTMLがない場合はテキスト版にフォールバック
            self.show_text_mode(msg)
            return
        
        # 最小限のサニタイズ
        html_safe = re.sub(r'<script[^>]*>.*?</script>', '', html_body, flags=re.DOTALL|re.IGNORECASE)
        html_safe = re.sub(r'\s+on\w+\s*=\s*["\'][^"\']*["\']', '', html_safe, flags=re.IGNORECASE)
        
        # 警告バナー追加
        warning = """
        <div style="background:#fff3cd;border:2px solid #856404;padding:10px;margin:10px;border-radius:5px;font-family:Arial;">
            ⚠️ このメールにはHTMLコンテンツが含まれています。外部画像やリンクが表示される場合があります。<br>
            💡 Gmailフィルタ通過済みのため、基本的に安全です。
        </div>
        """
        html_safe = warning + html_safe
        
        # tkinterwebを試行（なければブラウザ起動）
        try:
            import tkinterweb
            # デバッグメッセージを無効化
            browser = tkinterweb.HtmlFrame(self.content_frame, messages_enabled=False)
            browser.load_html(html_safe)
            browser.pack(fill=tk.BOTH, expand=True)
        except ImportError:
            # tkinterwebがない場合はブラウザで開く
            self.open_in_browser(html_safe)
            tk.Label(self.content_frame, text="ブラウザでHTMLを開きました", fg="blue", font=("Arial", 12)).pack(pady=50)
        except Exception as e:
            print(f"[ERROR] HTML表示エラー: {e}")
            import traceback
            traceback.print_exc()
            # エラー時はテキストモードにフォールバック
            self.show_text_mode(msg)
    
    def show_html_full_mode(self, msg):
        """HTML完全版（ユーザー自己責任）"""
        html_body = self.fetcher.extract_html_body(msg)
        
        if not html_body:
            self.show_text_mode(msg)
            return
        
        # 警告スキップ設定確認
        skip_warning = False
        if self.config_mgr:
            skip_warning = self.config_mgr.get("skip_html_warning") or False
        
        if not skip_warning:
            # カスタム警告ダイアログ（チェックボックス付き）
            proceed, dont_show_again = self.show_html_warning_dialog()
            
            if dont_show_again and self.config_mgr:
                self.config_mgr.set("skip_html_warning", True)
                self.config_mgr.save()
            
            if not proceed:
                # キャンセルされた場合は何もしない
                return
        
        try:
            import tkinterweb
            # デバッグメッセージを無効化
            browser = tkinterweb.HtmlFrame(self.content_frame, messages_enabled=False)
            browser.load_html(html_body)
            browser.pack(fill=tk.BOTH, expand=True)
        except ImportError:
            # tkinterwebがない場合はブラウザで開く
            self.open_in_browser(html_body)
            tk.Label(self.content_frame, text="ブラウザでHTMLを開きました\n（完全版・フィルタなし）", 
                    fg="blue", font=("Arial", 12), justify=tk.CENTER).pack(pady=50)
        except Exception as e:
            print(f"[ERROR] HTML完全版表示エラー: {e}")
            import traceback
            traceback.print_exc()
            # エラー時はテキストモードにフォールバック
            self.show_text_mode(msg)
    
    def show_html_warning_dialog(self):
        """HTML完全版警告ダイアログ（チェックボックス付き）"""
        dialog = tk.Toplevel(self.viewer)
        dialog.title("確認")
        dialog.geometry("450x250")
        dialog.transient(self.viewer)
        dialog.grab_set()
        
        # 警告メッセージ
        msg_frame = tk.Frame(dialog, bg="white")
        msg_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        tk.Label(msg_frame, text="⚠️ HTML完全版を表示します", 
                font=("Arial", 12, "bold"), bg="white").pack(pady=(0, 10))
        
        warning_text = (
            "外部スクリプトが実行される可能性があります。\n\n"
            "※ Gmailフィルタ通過済みですが、\n"
            "リスクを理解した上で続行してください。"
        )
        tk.Label(msg_frame, text=warning_text, justify=tk.LEFT, bg="white").pack()
        
        # チェックボックス
        var_dont_show = tk.BooleanVar()
        tk.Checkbutton(msg_frame, text="次回からこの警告を表示しない", 
                      variable=var_dont_show, bg="white").pack(pady=(15, 0))
        
        # ボタン
        result = {"proceed": False, "dont_show": False}
        
        def on_ok():
            result["proceed"] = True
            result["dont_show"] = var_dont_show.get()
            dialog.destroy()
        
        def on_cancel():
            result["proceed"] = False
            result["dont_show"] = False
            dialog.destroy()
        
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=20, pady=(0, 20))
        
        tk.Button(btn_frame, text="OK", command=on_ok, bg="#4CAF50", fg="white", 
                 width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="キャンセル", command=on_cancel, bg="#f44336", fg="white", 
                 width=15).pack(side=tk.LEFT, padx=5)
        
        # ダイアログを中央に配置
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        dialog.wait_window()
        return result["proceed"], result["dont_show"]
    
    def open_attachment(self, filename):
        """添付ファイルを一時保存して開く"""
        import email
        import tempfile
        import os
        import subprocess
        import sys
        
        # メールから添付ファイルを抽出
        msg = email.message_from_string(self.raw_data)
        
        for part in msg.walk():
            if part.get_filename():
                decoded_filename = self.fetcher.decode_h(part.get_filename())
                
                if decoded_filename == filename:
                    try:
                        payload = part.get_payload(decode=True)
                        
                        # 一時ファイルに保存
                        # 拡張子を保持
                        _, ext = os.path.splitext(filename)
                        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
                        temp_file.write(payload)
                        temp_file.close()
                        
                        # ファイルを開く
                        if sys.platform == "win32":
                            os.startfile(temp_file.name)
                        elif sys.platform == "darwin":
                            subprocess.run(["open", temp_file.name])
                        else:
                            subprocess.run(["xdg-open", temp_file.name])
                        
                        messagebox.showinfo("ファイルを開きました", f"{filename} を開きました")
                    except Exception as e:
                        messagebox.showerror("エラー", f"ファイルを開けませんでした:\n{e}")
                    break
    
    def preview_attachment(self, filename):
        """画像ファイルをプレビュー表示"""
        import email
        from PIL import Image, ImageTk
        import io
        
        # メールから添付ファイルを抽出
        msg = email.message_from_string(self.raw_data)
        
        for part in msg.walk():
            if part.get_filename():
                decoded_filename = self.fetcher.decode_h(part.get_filename())
                
                if decoded_filename == filename:
                    try:
                        payload = part.get_payload(decode=True)
                        
                        # 画像を開く
                        image = Image.open(io.BytesIO(payload))
                        
                        # プレビューウィンドウ作成
                        preview_win = tk.Toplevel(self.viewer)
                        preview_win.title(f"プレビュー: {filename}")
                        
                        # 画像サイズ調整（最大800x600）
                        max_width, max_height = 800, 600
                        img_width, img_height = image.size
                        
                        if img_width > max_width or img_height > max_height:
                            ratio = min(max_width/img_width, max_height/img_height)
                            new_size = (int(img_width*ratio), int(img_height*ratio))
                            image = image.resize(new_size, Image.Resampling.LANCZOS)
                        
                        # Tkinter用に変換
                        photo = ImageTk.PhotoImage(image)
                        
                        # ラベルに表示
                        label = tk.Label(preview_win, image=photo)
                        label.image = photo  # 参照を保持
                        label.pack()
                        
                    except Exception as e:
                        messagebox.showerror("エラー", f"画像を表示できませんでした:\n{e}")
                    break
    
    def save_attachment(self, filename):
        """添付ファイルを保存"""
        from tkinter import filedialog
        import email
        
        # メールから添付ファイルを抽出
        msg = email.message_from_string(self.raw_data)
        
        for part in msg.walk():
            if part.get_filename():
                decoded_filename = self.fetcher.decode_h(part.get_filename())
                
                if decoded_filename == filename:
                    # 保存先選択
                    save_path = filedialog.asksaveasfilename(
                        title="添付ファイルを保存",
                        initialfile=filename,
                        defaultextension=""
                    )
                    
                    if save_path:
                        try:
                            payload = part.get_payload(decode=True)
                            with open(save_path, "wb") as f:
                                f.write(payload)
                            messagebox.showinfo("保存完了", f"ファイルを保存しました:\n{save_path}")
                        except Exception as e:
                            messagebox.showerror("保存失敗", f"ファイル保存に失敗しました:\n{e}")
                    break
    
    def open_in_browser(self, html_content):
        """外部ブラウザでHTML表示"""
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", suffix=".html", delete=False) as f:
            f.write(html_content)
            html_path = f.name
            self.temp_files.append(html_path)
        
        # Windows対応: パス区切りを正規化
        html_path = html_path.replace("\\", "/")
        
        # URLエンコード不要な形式でブラウザ起動
        import urllib.parse
        file_url = urllib.parse.urljoin('file:', urllib.request.pathname2url(html_path))
        
        print(f"[DEBUG] Opening HTML in browser: {file_url}")
        
        try:
            webbrowser.open(file_url)
        except Exception as e:
            print(f"[ERROR] Failed to open browser: {e}")
            # フォールバック: 直接ファイルパスで開く試行
            try:
                webbrowser.open(html_path)
            except:
                pass

# 一時ファイルクリーンアップ
def cleanup_temp_files():
    for path in MailViewer.temp_files:
        try:
            os.remove(path)
        except:
            pass

atexit.register(cleanup_temp_files)

# ==========================================
# 5. MailHubApp (メインアプリ)
# ==========================================
class MailHubApp:
    def __init__(self, root):
        self.root = root
        self.root.title("RogoAI Mail Hub v1.0")
        self.root.geometry("1200x700")
        
        self.config_mgr = ConfigManager(CONFIG_FILE)
        self.db_mgr = DatabaseManager(DB_FILE)
        self.fetcher = MailFetcher()
        
        self.current_filter = None  # 現在のフィルタ（None=全件、provider名=特定プロバイダ）
        self.current_folder = None  # 現在のサブフォルダ
        self.current_search = ""    # 現在の検索キーワード
        self.current_promo_filter = False  # プロモフィルタ（False=通常、True=プロモのみ）
        
        # Message-ID <-> 安全なIIDのマッピング
        self.iid_to_msgid = {}  # iid -> message_id
        self.msgid_to_iid = {}  # message_id -> iid
        
        # ページング
        self.current_page = 1       # 現在のページ
        self.items_per_page = 200   # 1ページあたりの件数
        self.total_items = 0        # 総件数
        
        self.setup_ui()
        self.refresh_provider_list()
        self.refresh_folder_tree()
        self.refresh_tree_from_db()
    
    def prev_page(self):
        """前のページへ"""
        if self.current_page > 1:
            self.current_page -= 1
            self.refresh_tree_from_db()
    
    def next_page(self):
        """次のページへ"""
        total_pages = (self.total_items + self.items_per_page - 1) // self.items_per_page
        if self.current_page < total_pages:
            self.current_page += 1
            self.refresh_tree_from_db()
    
    def first_page(self):
        """最初のページへ"""
        if self.current_page != 1:
            self.current_page = 1
            self.refresh_tree_from_db()
    
    def last_page(self):
        """最後のページへ"""
        total_pages = max(1, (self.total_items + self.items_per_page - 1) // self.items_per_page)
        if self.current_page != total_pages:
            self.current_page = total_pages
            self.refresh_tree_from_db()
    
    def setup_ui(self):
        # トップバー
        top_bar = tk.Frame(self.root, bg="#2196F3", height=50)
        top_bar.pack(fill=tk.X)
        
        self.btn_inbox = tk.Button(top_bar, text="📥 受信箱", command=self.show_inbox, 
                                   relief=tk.SUNKEN, bg="white", width=15)
        self.btn_inbox.pack(side=tk.LEFT, padx=5, pady=5)
        
        self.btn_config = tk.Button(top_bar, text="⚙️ 設定", command=self.show_config, 
                                    relief=tk.RAISED, bg="#eee", width=15)
        self.btn_config.pack(side=tk.LEFT, padx=5, pady=5)
        
        # 検索バー
        tk.Label(top_bar, text="🔍", bg="#2196F3", fg="white", font=("Arial", 14)).pack(side=tk.LEFT, padx=(20,5))
        
        # 検索窓の変更を監視するためのStringVar
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.on_search_entry_change)
        
        self.search_entry = tk.Entry(top_bar, width=30, textvariable=self.search_var)
        self.search_entry.pack(side=tk.LEFT, padx=5)
        self.search_entry.bind("<Return>", lambda e: self.do_search())
        tk.Button(top_bar, text="検索", command=self.do_search, bg="white").pack(side=tk.LEFT, padx=5)
        tk.Button(top_bar, text="❓", command=self.show_search_help, bg="#2196F3", fg="white", 
                 font=("Arial", 10, "bold"), width=3).pack(side=tk.LEFT, padx=2)
        
        # ビューコンテナ
        self.view_inbox = tk.Frame(self.root)
        self.view_config = tk.Frame(self.root)
        
        self.setup_inbox_view()
        self.setup_config_view()
        
        self.view_inbox.pack(fill=tk.BOTH, expand=True)
    
    def setup_inbox_view(self):
        # 3ペインレイアウト
        main_pane = tk.PanedWindow(self.view_inbox, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True)
        
        # 左ペイン: フォルダツリー
        left_frame = tk.Frame(main_pane, width=200)
        main_pane.add(left_frame)
        
        # フォルダラベルとプロモ更新ボタン
        folder_header = tk.Frame(left_frame)
        folder_header.pack(fill=tk.X, padx=5, pady=5)
        
        tk.Label(folder_header, text="📂 フォルダ", font=("Arial", 12, "bold")).pack(side=tk.LEFT)
        
        self.btn_promo_update = tk.Button(
            folder_header, 
            text="🔄 プロモ更新", 
            command=self.apply_promo_rules_to_existing,
            font=("Arial", 9),
            bg="#9b59b6",
            fg="white",
            relief=tk.RAISED,
            cursor="hand2"
        )
        self.btn_promo_update.pack(side=tk.RIGHT, padx=5)
        
        # 初期状態チェック
        self.update_promo_button_state()
        
        self.folder_tree = ttk.Treeview(left_frame, show="tree")
        self.folder_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.folder_tree.bind("<<TreeviewSelect>>", self.on_folder_select)
        self.folder_tree.bind("<Button-3>", self.show_folder_context_menu)  # 右クリック
        
        # 中央ペイン: メール一覧
        center_frame = tk.Frame(main_pane)
        main_pane.add(center_frame, stretch="always")
        
        # ツールバー
        toolbar = tk.Frame(center_frame)
        toolbar.pack(fill=tk.X, padx=5, pady=5)
        
        self.btn_fetch = tk.Button(toolbar, text="📨 受信", command=self.start_fetch_task, 
                                   bg="#4CAF50", fg="white", width=12)
        self.btn_fetch.pack(side=tk.LEFT, padx=2)
        
        tk.Button(toolbar, text="✉️ 新規作成", command=self.open_compose_window, 
                 bg="#9C27B0", fg="white", width=12).pack(side=tk.LEFT, padx=2)
        
        tk.Button(toolbar, text="↩️ 返信", command=self.open_reply_window, 
                 bg="#2196F3", fg="white", width=12).pack(side=tk.LEFT, padx=2)
        
        tk.Button(toolbar, text="🔄 更新", command=self.refresh_tree_from_db, 
                 bg="#607D8B", fg="white", width=12).pack(side=tk.LEFT, padx=2)
        
        self.lbl_status = tk.Label(toolbar, text="準備完了", fg="green")
        self.lbl_status.pack(side=tk.LEFT, padx=10)
        
        # 取得進捗表示
        self.progress_frame = tk.Frame(toolbar)
        self.progress_frame.pack(side=tk.RIGHT, padx=10)
        
        self.lbl_progress = tk.Label(self.progress_frame, text="", fg="blue", font=("Arial", 9))
        self.lbl_progress.pack(side=tk.TOP)
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, length=200, mode='determinate')
        self.progress_bar.pack(side=tk.TOP, pady=2)
        
        # 初期状態では非表示
        self.progress_frame.pack_forget()
        
        # 検索中バナー（メールリスト上部）
        self.search_banner = tk.Frame(center_frame, bg="#fff3cd", relief=tk.RAISED, bd=2)
        self.search_banner_label = tk.Label(
            self.search_banner, 
            text="", 
            bg="#fff3cd", 
            fg="#856404", 
            font=("Arial", 10, "bold")
        )
        self.search_banner_label.pack(side=tk.LEFT, padx=10, pady=5)
        
        tk.Button(
            self.search_banner, 
            text="✗ クリア", 
            command=self.clear_search, 
            bg="#f39c12", 
            fg="white",
            font=("Arial", 9, "bold")
        ).pack(side=tk.RIGHT, padx=10, pady=5)
        
        # 初期状態では非表示
        # self.search_banner.pack_forget()  # 後で制御
        
        # メールリスト（スクロールバー付き）
        self.tree_frame = tk.Frame(center_frame)
        self.tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 縦スクロールバー
        tree_scroll = tk.Scrollbar(self.tree_frame, orient=tk.VERTICAL)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        cols = ("宛先", "件名", "送信者", "日付")
        self.tree = ttk.Treeview(self.tree_frame, columns=cols, show="headings", height=15, yscrollcommand=tree_scroll.set)
        
        # スクロールバーとTreeviewを連携
        tree_scroll.config(command=self.tree.yview)
        
        # カラムヘッダー設定とソート機能追加
        self.sort_column = None  # 現在のソート列
        self.sort_reverse = False  # ソート順（False=昇順, True=降順）
        
        for col in cols:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_tree_column(c))
            
        self.tree.column("宛先", width=200)
        self.tree.column("件名", width=300)
        self.tree.column("送信者", width=200)
        self.tree.column("日付", width=150)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree.bind("<Double-1>", self.open_viewer)
        self.tree.bind("<<TreeviewSelect>>", self.on_mail_select)
        self.tree.bind("<Button-3>", self.show_mail_context_menu)  # 右クリック
        
        # ページングコントロール
        paging_frame = tk.Frame(center_frame, bg="#f0f0f0", height=40)
        paging_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # ページ情報ラベル
        self.page_info_label = tk.Label(paging_frame, text="ページ 1 / 1 (0 件)", bg="#f0f0f0", font=("Arial", 10))
        self.page_info_label.pack(side=tk.LEFT, padx=10)
        
        # ページ移動ボタン
        btn_frame = tk.Frame(paging_frame, bg="#f0f0f0")
        btn_frame.pack(side=tk.RIGHT, padx=10)
        
        tk.Button(btn_frame, text="◀️ 前", command=self.prev_page, width=8).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_frame, text="次 ▶️", command=self.next_page, width=8).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_frame, text="⏮️ 最初", command=self.first_page, width=8).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_frame, text="最後 ⏭️", command=self.last_page, width=8).pack(side=tk.LEFT, padx=2)
        
        # 右ペイン: プレビュー
        right_frame = tk.Frame(main_pane, width=300)
        main_pane.add(right_frame)
        
        tk.Label(right_frame, text="📧 プレビュー", font=("Arial", 12, "bold")).pack(anchor=tk.W, padx=5, pady=5)
        
        self.preview_text = tk.Text(right_frame, wrap=tk.WORD, state=tk.DISABLED, font=("Arial", 9))
        self.preview_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 検索バナーを初期状態で非表示
        self.search_banner.pack_forget()
    
    def setup_config_view(self):
        # 設定タブ
        notebook = ttk.Notebook(self.view_config)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # ========================================
        # タブ1: Gmail集約アドレス設定（左右分割版）
        # ========================================
        tab1 = tk.Frame(notebook)
        notebook.add(tab1, text="Gmail集約アドレス設定")
        
        # 左右分割
        left_frame1 = tk.Frame(tab1)
        left_frame1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        right_frame1 = tk.Frame(tab1)
        right_frame1.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # --- 左側：Gmail接続設定 ---
        gmail_frame = tk.LabelFrame(left_frame1, text="Gmail接続設定", font=("Arial", 11, "bold"))
        gmail_frame.pack(fill=tk.X, padx=5, pady=5)
        
        tk.Label(gmail_frame, text="Gmailアドレス:", font=("Arial", 10)).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.ent_cen_email = tk.Entry(gmail_frame, width=35, font=("Arial", 10))
        self.ent_cen_email.grid(row=0, column=1, padx=10, pady=5)
        self.ent_cen_email.insert(0, self.config_mgr.get("email") or "")
        
        tk.Label(gmail_frame, text="アプリパスワード:", font=("Arial", 10)).grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.ent_cen_pass = tk.Entry(gmail_frame, width=35, show="*", font=("Arial", 10))
        self.ent_cen_pass.grid(row=1, column=1, padx=10, pady=5)
        self.ent_cen_pass.insert(0, self.config_mgr.get("password") or "")  # 保存済みパスワードを表示
        
        tk.Button(gmail_frame, text="接続テスト", font=("Arial", 9), bg="#3498db", fg="white", command=self.test_gmail_connection).grid(row=2, column=1, sticky="e", padx=10, pady=5)
        
        # --- 左側：メール取得範囲設定 ---
        fetch_frame = tk.LabelFrame(left_frame1, text="メール取得範囲設定", font=("Arial", 11, "bold"))
        fetch_frame.pack(fill=tk.X, padx=5, pady=5)
        
        tk.Label(fetch_frame, text="「受信箱」でメール取得する際の範囲", font=("Arial", 9), fg="#7f8c8d").pack(anchor="w", padx=10, pady=5)
        
        self.fetch_mode_var = tk.StringVar(value=self.config_mgr.get("fetch_mode") or "latest_only")
        fetch_modes = [
            ("latest_only", "最新のみ（前回以降） ✓ 推奨"),
            ("last_week", "過去1週間"),
            ("last_month", "過去1ヶ月"),
            ("last_3months", "過去3ヶ月"),
            ("last_year", "過去1年"),
            ("all", "すべて"),
            ("custom", "カスタム日数指定...")
        ]
        
        for value, text in fetch_modes:
            rb = tk.Radiobutton(fetch_frame, text=text, variable=self.fetch_mode_var, value=value, font=("Arial", 9))
            rb.pack(anchor="w", padx=20, pady=2)
        
        custom_frame = tk.Frame(fetch_frame)
        custom_frame.pack(anchor="w", padx=40, pady=2)
        tk.Label(custom_frame, text="日数:", font=("Arial", 9)).pack(side=tk.LEFT)
        self.custom_days_var = tk.IntVar(value=self.config_mgr.get("custom_days") or 30)
        tk.Spinbox(custom_frame, from_=1, to=3650, textvariable=self.custom_days_var, width=8, font=("Arial", 9)).pack(side=tk.LEFT, padx=5)
        tk.Label(custom_frame, text="日分", font=("Arial", 9)).pack(side=tk.LEFT)
        
        tk.Label(fetch_frame, text="💡 通常は「最新のみ」で十分です", font=("Arial", 9), fg="#27ae60").pack(anchor="w", padx=15, pady=5)
        
        # --- 右側：自動取得設定 ---
        auto_frame = tk.LabelFrame(right_frame1, text="自動取得設定", font=("Arial", 11, "bold"))
        auto_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.startup_var = tk.BooleanVar(value=self.config_mgr.get("auto_fetch_on_startup"))
        startup_check = tk.Checkbutton(auto_frame, text="アプリ起動時に自動取得 ✓ 推奨", variable=self.startup_var, font=("Arial", 10), fg="#27ae60")
        startup_check.pack(anchor="w", padx=15, pady=5)
        tk.Label(auto_frame, text="  （アプリ起動直後に1回だけ取得）", font=("Arial", 8), fg="#7f8c8d").pack(anchor="w", padx=30, pady=0)
        
        self.interval_var = tk.BooleanVar(value=self.config_mgr.get("auto_fetch_interval"))
        interval_check = tk.Checkbutton(auto_frame, text="定期的に自動取得", variable=self.interval_var, font=("Arial", 10))
        interval_check.pack(anchor="w", padx=15, pady=5)
        tk.Label(auto_frame, text="  （アプリ起動中のみ動作）", font=("Arial", 8), fg="#7f8c8d").pack(anchor="w", padx=30, pady=0)
        
        interval_frame = tk.Frame(auto_frame)
        interval_frame.pack(anchor="w", padx=30, pady=5)
        tk.Label(interval_frame, text="間隔:", font=("Arial", 9)).pack(side=tk.LEFT)
        self.interval_combo = ttk.Combobox(interval_frame, values=["15分ごと", "30分ごと ✓推奨", "1時間ごと", "2時間ごと"], state="readonly", width=15, font=("Arial", 9))
        self.interval_combo.pack(side=tk.LEFT, padx=5)
        self.interval_combo.current(1)
        
        advice_frame = tk.Frame(auto_frame, bg="#ecf0f1", relief=tk.GROOVE, bd=1)
        advice_frame.pack(fill=tk.X, padx=15, pady=10)
        advice_text = "💡 推奨設定:\n ✅ 起動時: ON\n ⬜ 定期: OFF\n\n📝 組み合わせ例:\n ・起動時のみ（推奨）\n ・定期のみ（長時間起動）\n ・両方（常に最新）\n\n重複は自動除外"
        tk.Label(advice_frame, text=advice_text, font=("Arial", 9), bg="#ecf0f1", fg="#16a085", justify=tk.LEFT).pack(pady=5, padx=10)
        
        # 保存ボタン（タブ1の下部）
        button_frame1 = tk.Frame(tab1)
        button_frame1.pack(side=tk.BOTTOM, pady=10)
        tk.Button(button_frame1, text="保存", command=self.save_central_config, font=("Arial", 11, "bold"), bg="#27ae60", fg="white", width=15, height=2).pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame1, text="既定に戻す", command=self.reset_gmail_settings, font=("Arial", 10), bg="#95a5a6", fg="white", width=12).pack(side=tk.LEFT, padx=10)
        
        # ========================================
        # タブ2: 送信アカウント管理（スクロール＋注意書き下部）
        # ========================================
        tab2 = tk.Frame(notebook)
        notebook.add(tab2, text="送信アカウント管理")
        
        # スクロール対応
        canvas2 = tk.Canvas(tab2)
        scrollbar2 = ttk.Scrollbar(tab2, orient="vertical", command=canvas2.yview)
        scrollable_frame2 = tk.Frame(canvas2)
        
        scrollable_frame2.bind("<Configure>", lambda e: canvas2.configure(scrollregion=canvas2.bbox("all")))
        canvas2.create_window((0, 0), window=scrollable_frame2, anchor="nw")
        canvas2.configure(yscrollcommand=scrollbar2.set)
        
        canvas2.pack(side="left", fill="both", expand=True)
        scrollbar2.pack(side="right", fill="y")
        
        tk.Label(scrollable_frame2, text="登録済みアカウント", font=("Arial", 12, "bold")).pack(pady=5)
        
        prov_cols = ("Email", "SMTPサーバー", "ポート", "Gmail代理")
        self.tree_prov = ttk.Treeview(scrollable_frame2, columns=prov_cols, show="headings", height=8)
        for col in prov_cols:
            self.tree_prov.heading(col, text=col)
        self.tree_prov.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        tk.Button(scrollable_frame2, text="❌ 削除", command=self.delete_provider, bg="#f44336", fg="white").pack(pady=5)
        
        tk.Label(scrollable_frame2, text="新規アカウント追加", font=("Arial", 12, "bold")).pack(pady=10)
        
        add_frame = tk.Frame(scrollable_frame2)
        add_frame.pack(fill=tk.X, padx=20, pady=5)
        
        tk.Label(add_frame, text="メールアドレス:").grid(row=0, column=0, sticky=tk.W)
        self.ent_prov_email = tk.Entry(add_frame, width=30)
        self.ent_prov_email.grid(row=0, column=1, padx=5, pady=2)
        
        tk.Label(add_frame, text="SMTPサーバー:").grid(row=1, column=0, sticky=tk.W)
        self.ent_prov_host = tk.Entry(add_frame, width=30)
        self.ent_prov_host.grid(row=1, column=1, padx=5, pady=2)
        
        tk.Label(add_frame, text="ポート:").grid(row=2, column=0, sticky=tk.W)
        self.ent_prov_port = tk.Entry(add_frame, width=30)
        self.ent_prov_port.grid(row=2, column=1, padx=5, pady=2)
        self.ent_prov_port.insert(0, "587")
        
        tk.Label(add_frame, text="パスワード:").grid(row=3, column=0, sticky=tk.W)
        self.ent_prov_pass = tk.Entry(add_frame, width=30, show="*")
        self.ent_prov_pass.grid(row=3, column=1, padx=5, pady=2)
        
        self.var_fallback = tk.BooleanVar()
        tk.Checkbutton(add_frame, text="Gmail経由で送信（SMTP接続不可の場合）", variable=self.var_fallback).grid(row=4, column=1, sticky=tk.W, pady=5)
        
        tk.Button(add_frame, text="➕ 追加", command=self.add_provider, bg="#2196F3", fg="white", width=20).grid(row=5, column=1, pady=10)
        
        # 注意書き（下部）
        warning_frame = tk.Frame(scrollable_frame2, bg="#fff3cd", relief=tk.GROOVE, bd=2)
        warning_frame.pack(fill=tk.X, padx=10, pady=10)
        tk.Label(warning_frame, text="⚠️ プロバイダ別 送信設定の重要な注意", font=("Arial", 11, "bold"), bg="#fff3cd", fg="#e74c3c").pack(anchor="w", padx=15, pady=5)
        smtp_info = "【Microsoft系メール】 live.jp / outlook.jp / hotmail.co.jp\n   ❌ 送信（SMTP）設定：不可（Microsoft側が外部SMTPを拒否）\n   ✅ 受信（IMAP）：可能\n   📧 返信時：Gmail経由で自動送信\n\n【その他のプロバイダ】 OCN / So-net / Nifty 等\n   ✅ 送信（SMTP）設定：可能\n   ⚠️ 未登録の場合：Gmail経由で送信"
        tk.Label(warning_frame, text=smtp_info, font=("Arial", 9), bg="#fff3cd", fg="#2c3e50", justify=tk.LEFT).pack(anchor="w", padx=15, pady=10)
        
        # ========================================
        # タブ3: プロバイダ色設定（重複削除済み）
        # ========================================
        tab3 = tk.Frame(notebook)
        notebook.add(tab3, text="プロバイダ色設定")
        
        tk.Label(tab3, text="プロバイダ別背景色カスタマイズ", font=("Arial", 12, "bold")).pack(pady=10)
        
        color_cols = ("プロバイダ", "現在の色")
        self.tree_colors = ttk.Treeview(tab3, columns=color_cols, show="headings", height=10)
        for col in color_cols:
            self.tree_colors.heading(col, text=col)
        self.tree_colors.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        tk.Button(tab3, text="🎨 色を変更", command=self.change_provider_color, bg="#9C27B0", fg="white", width=20).pack(pady=5)
        tk.Button(tab3, text="🔄 自動色に戻す", command=self.reset_provider_color, bg="#607D8B", fg="white", width=20).pack(pady=5)
        
        # ========================================
        # タブ4: データベース管理（左右分割版）
        # ========================================
        tab4 = tk.Frame(notebook)
        notebook.add(tab4, text="ゴミ箱設定")
        
        # 削除動作の設定
        delete_mode_frame = tk.LabelFrame(tab4, text="削除動作の設定", font=("Arial", 11, "bold"))
        delete_mode_frame.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Label(delete_mode_frame, text="ゴミ箱から完全削除する際の動作を選択:", font=("Arial", 9)).pack(anchor="w", padx=15, pady=5)
        
        self.delete_mode_var = tk.StringVar(value=self.config_mgr.get("delete_mode") or "local_only")
        
        # モードA: Mail Hubのみ
        mode_a_frame = tk.Frame(delete_mode_frame, bg="#e8f5e9", relief=tk.GROOVE, bd=2)
        mode_a_frame.pack(fill=tk.X, padx=15, pady=5)
        
        tk.Radiobutton(
            mode_a_frame, 
            text="Mail Hubのみから削除（Gmail Cloudは残る）【推奨】",
            variable=self.delete_mode_var,
            value="local_only",
            font=("Arial", 10, "bold"),
            bg="#e8f5e9"
        ).pack(anchor="w", padx=10, pady=5)
        
        # モードB: Gmail Cloudからも削除
        mode_b_frame = tk.Frame(delete_mode_frame, bg="#fff3e0", relief=tk.GROOVE, bd=2)
        mode_b_frame.pack(fill=tk.X, padx=15, pady=5)
        
        tk.Radiobutton(
            mode_b_frame,
            text="Gmail Cloudからも削除（完全削除）",
            variable=self.delete_mode_var,
            value="gmail_cloud",
            font=("Arial", 10, "bold"),
            bg="#fff3e0"
        ).pack(anchor="w", padx=10, pady=5)
        
        # 保存ボタン
        tk.Button(delete_mode_frame, text="保存", command=self.save_delete_mode, font=("Arial", 10), bg="#2196F3", fg="white", width=20).pack(pady=10)
        
        # 削除済みリスト管理
        deleted_list_frame = tk.LabelFrame(tab4, text="削除済みリスト管理", font=("Arial", 11, "bold"))
        deleted_list_frame.pack(fill=tk.X, padx=10, pady=10)
        
        deleted_count = len(self.db_mgr.get_deleted_message_ids())
        tk.Label(deleted_list_frame, text=f"登録件数: {deleted_count:,}件", font=("Arial", 10)).pack(anchor="w", padx=15, pady=5)
        
        tk.Button(deleted_list_frame, text="リストを表示", font=("Arial", 9), bg="#9b59b6", fg="white", width=30, command=self.show_deleted_list).pack(pady=5)
        
        tk.Label(deleted_list_frame, text="⚠️ 上級者向け機能", font=("Arial", 9), fg="#e67e22").pack(pady=2)
        tk.Button(deleted_list_frame, text="リストをクリア（すべて復活）", font=("Arial", 9), bg="#e74c3c", fg="white", width=30, command=self.clear_deleted_list).pack(pady=5)
        
        # ========================================
        # タブ5: データベース
        # ========================================
        tab5 = tk.Frame(notebook)
        notebook.add(tab5, text="データベース")
        
        # 左右分割
        left_frame5 = tk.Frame(tab5)
        left_frame5.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        right_frame5 = tk.Frame(tab5)
        right_frame5.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # --- 左側：データベース情報 ---
        db_info_frame = tk.LabelFrame(left_frame5, text="データベース情報", font=("Arial", 11, "bold"))
        db_info_frame.pack(fill=tk.X, padx=5, pady=5)
        
        db_grid = tk.Frame(db_info_frame)
        db_grid.pack(anchor="w", padx=15, pady=10)
        
        tk.Label(db_grid, text="保存場所:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", pady=3)
        tk.Label(db_grid, text=DB_FILE, font=("Arial", 8), fg="#2c3e50", wraplength=400, justify=tk.LEFT).grid(row=0, column=1, sticky="w", padx=10, pady=3)
        
        tk.Label(db_grid, text="メール件数:", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky="w", pady=3)
        email_count = self.db_mgr.get_email_count()
        tk.Label(db_grid, text=f"{email_count:,}件", font=("Arial", 9), fg="#2c3e50").grid(row=1, column=1, sticky="w", padx=10, pady=3)
        
        tk.Label(db_grid, text="サイズ:", font=("Arial", 10, "bold")).grid(row=2, column=0, sticky="w", pady=3)
        if os.path.exists(DB_FILE):
            db_size = os.path.getsize(DB_FILE) / (1024 * 1024)
            tk.Label(db_grid, text=f"{db_size:.2f} MB", font=("Arial", 9), fg="#2c3e50").grid(row=2, column=1, sticky="w", padx=10, pady=3)
        else:
            tk.Label(db_grid, text="0.00 MB", font=("Arial", 9), fg="#2c3e50").grid(row=2, column=1, sticky="w", padx=10, pady=3)
        
        # --- 左側：データベース操作 ---
        db_operation_frame = tk.LabelFrame(left_frame5, text="データベース操作", font=("Arial", 11, "bold"), fg="red")
        db_operation_frame.pack(fill=tk.X, padx=5, pady=10)
        
        tk.Button(db_operation_frame, text="データベースをリセット", font=("Arial", 10), bg="#e74c3c", fg="white", width=25, command=self.reset_database).pack(pady=5)
        tk.Label(db_operation_frame, text="⚠️ すべてのメールが削除", font=("Arial", 9), fg="#e74c3c").pack(pady=2)
        
        # --- 右側：設定ファイル情報 ---
        config_info_frame = tk.LabelFrame(right_frame5, text="設定ファイル情報", font=("Arial", 11, "bold"))
        config_info_frame.pack(fill=tk.X, padx=5, pady=5)
        
        config_grid = tk.Frame(config_info_frame)
        config_grid.pack(anchor="w", padx=15, pady=10)
        
        tk.Label(config_grid, text="保存場所:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", pady=3)
        tk.Label(config_grid, text=CONFIG_FILE, font=("Arial", 8), fg="#2c3e50", wraplength=400, justify=tk.LEFT).grid(row=0, column=1, sticky="w", padx=10, pady=3)
        
        tk.Label(config_grid, text="状態:", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky="w", pady=3)
        if os.path.exists(CONFIG_FILE):
            config_size = os.path.getsize(CONFIG_FILE) / 1024
            tk.Label(config_grid, text=f"✅ 作成済み ({config_size:.2f} KB)", font=("Arial", 9), fg="#27ae60").grid(row=1, column=1, sticky="w", padx=10, pady=3)
        else:
            tk.Label(config_grid, text="⚠️ 未作成（設定保存後に作成）", font=("Arial", 9), fg="#f39c12").grid(row=1, column=1, sticky="w", padx=10, pady=3)
        
        # --- 右側：設定ファイル操作 ---
        config_operation_frame = tk.LabelFrame(right_frame5, text="設定ファイル操作", font=("Arial", 11, "bold"))
        config_operation_frame.pack(fill=tk.X, padx=5, pady=5)
        
        tk.Button(config_operation_frame, text="設定ファイルを開く", font=("Arial", 10), bg="#3498db", fg="white", width=25, command=self.open_config_file).pack(pady=5)
        tk.Label(config_operation_frame, text="📝 メモ帳で編集（作成後のみ）", font=("Arial", 9), fg="#7f8c8d").pack(pady=2)
        
        tk.Button(config_operation_frame, text="フォルダを開く", font=("Arial", 10), bg="#9b59b6", fg="white", width=25, command=self.open_config_folder).pack(pady=5)
        tk.Label(config_operation_frame, text="📂 エクスプローラーで表示", font=("Arial", 9), fg="#7f8c8d").pack(pady=2)
        
        tk.Button(config_operation_frame, text="設定ファイルを初期化", font=("Arial", 10), bg="#f39c12", fg="white", width=25, command=self.reset_config).pack(pady=5)
        tk.Label(config_operation_frame, text="⚠️ すべての設定が削除", font=("Arial", 9), fg="#f39c12").pack(pady=2)
        
        # 注意書き
        if not os.path.exists(CONFIG_FILE):
            note_frame = tk.Frame(config_operation_frame, bg="#fff3cd", relief=tk.GROOVE, bd=1)
            note_frame.pack(fill=tk.X, padx=5, pady=10)
            note_text = "💡 初回起動時はファイル未作成\n\n「Gmail集約アドレス設定」で\n設定を保存すると作成されます"
            tk.Label(note_frame, text=note_text, font=("Arial", 9), bg="#fff3cd", fg="#e67e22", justify=tk.LEFT).pack(pady=8, padx=10)
        
        # --- 右側：上級者向けカスタマイズ ---
        advanced_frame = tk.LabelFrame(right_frame5, text="上級者向けカスタマイズ", font=("Arial", 10, "bold"), fg="#e67e22")
        advanced_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=10)
        
        advice_text = (
            "💡 config.jsonで保存先を変更できます\n\n"
            "config.jsonに追加:\n"
            '  "storage_dir": "D:\\\\MyMailData",\n'
            '  "config_dir": null\n\n'
            "注意:\n"
            "・Windowsは \\\\ で区切る\n"
            "・null = デフォルト\n"
            "・変更後は再起動\n\n"
            "例:\n"
            '  "storage_dir": "C:\\\\MailHub\\\\data"\n'
            '  → emails.dbの保存先変更'
        )
        
        tk.Label(advanced_frame, text=advice_text, font=("Arial", 9), fg="#2c3e50", justify=tk.LEFT, bg="#ecf0f1").pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # ========================================
        # タブ6: プロモ・ルール管理
        # ========================================
        tab6 = tk.Frame(notebook)
        notebook.add(tab6, text="プロモ・ルール管理")
        
        # 説明ヘッダー
        header_frame = tk.Frame(tab6, bg="#3498db")
        header_frame.pack(fill=tk.X)
        tk.Label(header_frame, text="プロモ・ボックス 自動振り分けルール管理", 
                font=("Arial", 12, "bold"), bg="#3498db", fg="white").pack(pady=10)
        
        # ルール一覧フレーム
        list_frame = tk.LabelFrame(tab6, text="登録済みルール一覧", font=("Arial", 11, "bold"))
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # ルールリスト（スクロール付き）
        list_container = tk.Frame(list_frame)
        list_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Treeview for rules
        # Treeview for rules
        columns = ("pattern", "target_folder", "match_count", "added_date")
        self.promo_rules_tree = ttk.Treeview(list_container, columns=columns, show="headings", height=15)
        
        self.promo_rules_tree.heading("pattern", text="送信者パターン")
        self.promo_rules_tree.heading("target_folder", text="振り分け先")
        self.promo_rules_tree.heading("match_count", text="マッチ件数")
        self.promo_rules_tree.heading("added_date", text="登録日時")
        
        self.promo_rules_tree.column("pattern", width=300)
        self.promo_rules_tree.column("target_folder", width=150, anchor=tk.CENTER)
        self.promo_rules_tree.column("match_count", width=100, anchor=tk.CENTER)
        self.promo_rules_tree.column("added_date", width=180, anchor=tk.CENTER)
        
        scrollbar = tk.Scrollbar(list_container, command=self.promo_rules_tree.yview)
        self.promo_rules_tree.config(yscrollcommand=scrollbar.set)
        
        self.promo_rules_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ボタンフレーム
        button_frame = tk.Frame(tab6)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Button(button_frame, text="🔄 更新", command=self.refresh_promo_rules, 
                 bg="#3498db", fg="white", font=("Arial", 10, "bold"), width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="🗑️ 選択削除", command=self.delete_selected_promo_rule, 
                 bg="#e74c3c", fg="white", font=("Arial", 10, "bold"), width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="🗑️ 全削除", command=self.delete_all_promo_rules, 
                 bg="#c0392b", fg="white", font=("Arial", 10, "bold"), width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="➕ 手動追加", command=self.add_promo_rule_manual, 
                 bg="#27ae60", fg="white", font=("Arial", 10, "bold"), width=12).pack(side=tk.LEFT, padx=5)
        
        # 説明文
        info_frame = tk.Frame(tab6, bg="#ecf0f1", relief=tk.GROOVE, bd=2)
        info_frame.pack(fill=tk.X, padx=10, pady=5)
        
        info_text = (
            "📌 プロモ・ルールについて\n\n"
            "• メールをプロモ・ボックスに移動すると、自動的にルールが作成されます\n"
            "• 次回メール取得時から、同じドメインからのメールが自動振り分けされます\n"
            "• ルールを削除すると、以降は自動振り分けされなくなります\n"
            "• 既にプロモ・ボックスにあるメールは削除されません"
        )
        tk.Label(info_frame, text=info_text, font=("Arial", 9), bg="#ecf0f1", 
                fg="#2c3e50", justify=tk.LEFT).pack(padx=15, pady=10)
        
        # 初回ロード
        self.refresh_promo_rules()
    
    def refresh_promo_rules(self):
        """プロモ・ルール一覧を更新"""
        # 既存項目をクリア
        for item in self.promo_rules_tree.get_children():
            self.promo_rules_tree.delete(item)
        
        # DBから取得
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("SELECT sender_pattern, target_folder, match_count, added_date FROM promo_rules ORDER BY added_date DESC")
        rules = cur.fetchall()
        conn.close()
        
        # Treeviewに追加
        for pattern, target_folder, match_count, added_date in rules:
            folder_display = target_folder if target_folder else "プロモ・ボックス"
            self.promo_rules_tree.insert("", tk.END, values=(pattern, folder_display, match_count, added_date))
        
        # 件数表示
        self.promo_rules_tree.update_idletasks()
    
    def delete_selected_promo_rule(self):
        """選択したプロモ・ルールを削除"""
        selected = self.promo_rules_tree.selection()
        if not selected:
            messagebox.showwarning("未選択", "削除するルールを選択してください")
            return
        
        patterns = []
        for item in selected:
            values = self.promo_rules_tree.item(item, "values")
            patterns.append(values[0])
        
        if not messagebox.askyesno("確認", f"{len(patterns)}件のルールを削除しますか？\n\n削除後は自動振り分けされなくなります"):
            return
        
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        for pattern in patterns:
            cur.execute("DELETE FROM promo_rules WHERE sender_pattern=?", (pattern,))
        conn.commit()
        conn.close()
        
        messagebox.showinfo("完了", f"{len(patterns)}件のルールを削除しました")
        self.refresh_promo_rules()
        self.update_promo_button_state()
    
    def delete_all_promo_rules(self):
        """全プロモ・ルールを削除"""
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM promo_rules")
        count = cur.fetchone()[0]
        conn.close()
        
        if count == 0:
            messagebox.showinfo("情報", "削除するルールがありません")
            return
        
        if not messagebox.askyesno("確認", f"全{count}件のルールを削除しますか？\n\n⚠️ この操作は取り消せません"):
            return
        
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("DELETE FROM promo_rules")
        conn.commit()
        conn.close()
        
        messagebox.showinfo("完了", f"全{count}件のルールを削除しました")
        self.refresh_promo_rules()
        self.update_promo_button_state()
    
    def add_promo_rule_manual(self):
        """手動でプロモ・ルールを追加"""
        dialog = tk.Toplevel(self.root)
        dialog.title("プロモ・ルール手動追加")
        dialog.geometry("500x250")
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text="送信者パターンを入力", font=("Arial", 11, "bold")).pack(pady=10)
        
        info_frame = tk.Frame(dialog, bg="#fff3cd", relief=tk.GROOVE, bd=2)
        info_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(info_frame, text="例: %@example.com% （このドメインからのメール全て）\n例: %newsletter@% （このアドレスから始まるメール全て）", 
                font=("Arial", 9), bg="#fff3cd", justify=tk.LEFT).pack(padx=10, pady=5)
        
        tk.Label(dialog, text="パターン:", font=("Arial", 10)).pack(pady=5)
        entry = tk.Entry(dialog, font=("Arial", 10), width=40)
        entry.pack(pady=5)
        entry.insert(0, "%@")
        
        result = {"confirmed": False}
        
        def on_ok():
            pattern = entry.get().strip()
            if not pattern or pattern == "%@":
                messagebox.showwarning("入力エラー", "有効なパターンを入力してください")
                return
            
            if "%" not in pattern:
                messagebox.showwarning("入力エラー", "パターンには % を含めてください")
                return
            
            result["pattern"] = pattern
            result["confirmed"] = True
            dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=15)
        
        tk.Button(btn_frame, text="追加", command=on_ok, bg="#27ae60", fg="white", 
                 font=("Arial", 10, "bold"), width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="キャンセル", command=on_cancel, bg="#e74c3c", fg="white", 
                 font=("Arial", 10, "bold"), width=12).pack(side=tk.LEFT, padx=5)
        
        dialog.wait_window()
        
        if result.get("confirmed"):
            pattern = result["pattern"]
            
            conn = sqlite3.connect(DB_FILE)
            cur = conn.cursor()
            try:
                cur.execute("INSERT INTO promo_rules (sender_pattern, added_date, match_count, target_folder) VALUES (?, datetime('now'), 0, NULL)", (pattern,))
                conn.commit()
                messagebox.showinfo("完了", f"ルールを追加しました\n\nパターン: {pattern}")
                self.refresh_promo_rules()
                self.update_promo_button_state()
            except sqlite3.IntegrityError:
                messagebox.showwarning("重複", "このパターンは既に登録されています")
            finally:
                conn.close()
    
    def save_delete_mode(self):
        """削除モード保存"""
        mode = self.delete_mode_var.get()
        self.config_mgr.set("delete_mode", mode)
        self.config_mgr.save()
        
        mode_name = "Mail Hubのみ" if mode == "local_only" else "Gmail Cloudからも削除"
        messagebox.showinfo("保存完了", f"削除モードを変更しました\n\n設定: {mode_name}")
    
    def show_deleted_list(self):
        """削除済みリスト表示"""
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("""
            SELECT message_id, deleted_date, delete_mode 
            FROM deleted_messages 
            ORDER BY deleted_date DESC
        """)
        rows = cur.fetchall()
        conn.close()
        
        # ダイアログ作成
        dialog = tk.Toplevel(self.root)
        dialog.title(f"削除済みメールリスト ({len(rows)}件)")
        dialog.geometry("700x500")
        
        # リスト
        cols = ("Message-ID", "削除日", "削除モード")
        tree = ttk.Treeview(dialog, columns=cols, show="headings", height=20)
        
        tree.heading("Message-ID", text="Message-ID")
        tree.heading("削除日", text="削除日")
        tree.heading("削除モード", text="削除モード")
        
        tree.column("Message-ID", width=400)
        tree.column("削除日", width=150)
        tree.column("削除モード", width=120)
        
        for msg_id, date, mode in rows:
            mode_text = "ローカルのみ" if mode == "local_only" else "Cloud含む"
            tree.insert("", "end", values=(msg_id, date, mode_text))
        
        tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # ボタン
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="閉じる", command=dialog.destroy, width=15).pack(side=tk.LEFT, padx=5)
    
    def clear_deleted_list(self):
        """削除済みリストクリア"""
        count = len(self.db_mgr.get_deleted_message_ids())
        
        result = messagebox.askyesno(
            "確認",
            f"削除済みリスト（{count:,}件）をクリアします。\n\n"
            "次回のメール取得時に、削除したメールが\n"
            "再度表示される可能性があります。\n\n"
            "よろしいですか？"
        )
        
        if result:
            conn = sqlite3.connect(DB_FILE)
            cur = conn.cursor()
            cur.execute("DELETE FROM deleted_messages")
            conn.commit()
            conn.close()
            messagebox.showinfo("完了", "削除済みリストをクリアしました")
    
    def open_config_file(self):
        """設定ファイルをメモ帳で開く"""
        try:
            if not os.path.exists(CONFIG_FILE):
                messagebox.showwarning("ファイルなし", "設定ファイルがまだ作成されていません。\n\n最初に設定を保存してください。")
                return
            
            import subprocess
            if os.name == 'nt':  # Windows
                subprocess.Popen(['notepad.exe', CONFIG_FILE])
            else:  # Mac/Linux
                subprocess.Popen(['open', CONFIG_FILE] if sys.platform == 'darwin' else ['xdg-open', CONFIG_FILE])
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルを開けませんでした:\n{e}")
    
    def open_config_folder(self):
        """設定ファイルのフォルダを開く"""
        try:
            import subprocess
            folder = os.path.dirname(CONFIG_FILE)
            if os.name == 'nt':  # Windows
                subprocess.Popen(['explorer', folder])
            else:  # Mac/Linux
                subprocess.Popen(['open', folder] if sys.platform == 'darwin' else ['xdg-open', folder])
        except Exception as e:
            messagebox.showerror("エラー", f"フォルダを開けませんでした:\n{e}")
    
    def test_gmail_connection(self):
        """Gmail接続テスト"""
        email_val = self.ent_cen_email.get().strip()
        pass_val = self.ent_cen_pass.get().strip()
        
        if not email_val or not pass_val:
            messagebox.showwarning("不足", "メールアドレスとパスワードを入力してください")
            return
        
        try:
            self.root.config(cursor="watch")
            self.root.update()
            self.fetcher.test_connection_imap(self.config_mgr.get("imap_server"), email_val, pass_val)
            self.root.config(cursor="")
            messagebox.showinfo("成功", "Gmail接続に成功しました！")
        except Exception as e:
            self.root.config(cursor="")
            messagebox.showerror("失敗", f"接続に失敗しました:\n{e}")
    
    def reset_config(self):
        """設定ファイルを初期化"""
        result = messagebox.askyesno(
            "確認",
            "本当に設定ファイルを初期化しますか？\n\nすべての設定（Gmail設定、プロバイダ設定など）が削除されます。"
        )
        if result:
            try:
                if os.path.exists(CONFIG_FILE):
                    os.remove(CONFIG_FILE)
                self.config_mgr.config = DEFAULT_CONFIG.copy()
                self.config_mgr.save()
                messagebox.showinfo("完了", "設定ファイルを初期化しました。\n\nアプリを再起動してください。")
            except Exception as e:
                messagebox.showerror("エラー", f"初期化に失敗しました:\n{e}")
    
    def reset_gmail_settings(self):
        """Gmail設定を既定値に戻す"""
        self.fetch_mode_var.set("latest_only")
        self.startup_var.set(True)
        self.interval_var.set(False)
        self.interval_combo.current(1)
        self.custom_days_var.set(30)
        messagebox.showinfo("リセット", "既定値に戻しました\n\n※ 起動時自動取得: ON（推奨）")
    
    def refresh_folder_tree(self):
        """フォルダツリー更新（プロモ・ボックス追加）"""
        self.folder_tree.delete(*self.folder_tree.get_children())
        
        # DBから全メール数とプロバイダ別メール数取得
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        
        # 全メール数（プロモ除外）
        cur.execute("SELECT COUNT(*) FROM emails WHERE is_promo=0 OR is_promo IS NULL")
        total_count = cur.fetchone()[0]
        
        # プロモメール数
        cur.execute("SELECT COUNT(*) FROM emails WHERE is_promo=1")
        promo_count = cur.fetchone()[0]
        
        # プロバイダ別メール数（プロモ除外）
        cur.execute("SELECT provider, COUNT(*) FROM emails WHERE provider IS NOT NULL AND (is_promo=0 OR is_promo IS NULL) GROUP BY provider")
        provider_counts = {row[0]: row[1] for row in cur.fetchall()}
        
        # 全件表示（プロモ除外）
        self.folder_tree.insert("", tk.END, "all", text=f"📧 受信メール ({total_count})")
        
        # プロモボックス（特別枠 + サブフォルダ）
        if promo_count > 0:
            promo_node = self.folder_tree.insert("", tk.END, "promo", text=f"📂 プロモ・ボックス ({promo_count})", tags=("promo",))
        else:
            promo_node = self.folder_tree.insert("", tk.END, "promo", text=f"📂 プロモ・ボックス (0)", tags=("promo", "empty"))
        
        # プロモ・ボックスのゴミ箱
        cur.execute("SELECT COUNT(*) FROM emails WHERE is_promo=1 AND folder='__trash__'")
        promo_trash_count = cur.fetchone()[0]
        self.folder_tree.insert(promo_node, tk.END, "__promo__:__trash__", text=f"🗑️ ゴミ箱 ({promo_trash_count})")
        
        # プロモ・ボックスのカスタムフォルダ
        promo_folders = self.db_mgr.get_folders("__promo__")
        for folder_name, folder_type in promo_folders:
            if folder_type == 'custom':
                cur.execute("SELECT COUNT(*) FROM emails WHERE is_promo=1 AND folder=?", (folder_name,))
                promo_custom_count = cur.fetchone()[0]
                self.folder_tree.insert(promo_node, tk.END, f"__promo__:{folder_name}",
                                       text=f"📂 {folder_name} ({promo_custom_count})", tags=("custom",))
        
        # 登録されている全プロバイダ（SMTP設定 + DB存在）を統合
        db_providers = set(provider_counts.keys())
        smtp_providers = set()
        
        # SMTP設定されたプロバイダを追加
        providers_config = self.config_mgr.get("providers") or []
        for p in providers_config:
            email = p.get("email", "")
            if "@" in email:
                domain = email.split("@")[-1].lower()
                smtp_providers.add(domain)
        
        # 統合（DBにあるプロバイダ + SMTP設定プロバイダ）
        all_providers = sorted(db_providers | smtp_providers)
        
        # フォルダ表示（システムフォルダ + カスタムフォルダ対応）
        for provider in all_providers:
            count = provider_counts.get(provider, 0)
            
            # プロバイダノード
            if count > 0:
                provider_node = self.folder_tree.insert("", tk.END, provider, text=f"📧 {provider} ({count})")
            else:
                provider_node = self.folder_tree.insert("", tk.END, provider, text=f"📧 {provider} (0)", tags=("empty",))
            
            # システムフォルダ
            system_folders = [
                ("__sent__", "📤 送信済み"),
                ("__drafts__", "📝 下書き"),
                ("__trash__", "🗑️ ゴミ箱"),
            ]
            
            for folder_key, folder_label in system_folders:
                # システムフォルダのメール数をカウント
                cur.execute("""
                    SELECT COUNT(*) FROM emails 
                    WHERE provider=? AND folder=?
                """, (provider, folder_key))
                folder_count = cur.fetchone()[0]
                
                self.folder_tree.insert(provider_node, tk.END, f"{provider}:{folder_key}", 
                                       text=f"{folder_label} ({folder_count})")
            
            # カスタムフォルダ
            folders = self.db_mgr.get_folders(provider)
            for folder_name, folder_type in folders:
                if folder_type == 'custom':
                    # カスタムフォルダのメール数
                    cur.execute("""
                        SELECT COUNT(*) FROM emails 
                        WHERE provider=? AND folder=?
                    """, (provider, folder_name))
                    custom_count = cur.fetchone()[0]
                    
                    self.folder_tree.insert(provider_node, tk.END, f"{provider}:{folder_name}",
                                           text=f"📂 {folder_name} ({custom_count})", tags=("custom",))
        
        conn.close()
        
        # スタイル設定
        
        conn.close()
        
        # スタイル設定
        self.folder_tree.tag_configure("promo", background="#FFF4E0")  # 薄オレンジ
        self.folder_tree.tag_configure("empty", foreground="#999999")
        
        # デフォルト選択
        if self.folder_tree.get_children():
            self.folder_tree.selection_set("all")
    
    
    
    def make_safe_iid(self, message_id):
        """Message-IDを安全なTreeview IIDに変換"""
        import hashlib
        if message_id in self.msgid_to_iid:
            return self.msgid_to_iid[message_id]
        
        # SHA256ハッシュ化
        safe_iid = hashlib.sha256(message_id.encode('utf-8')).hexdigest()[:32]
        
        # マッピング登録
        self.iid_to_msgid[safe_iid] = message_id
        self.msgid_to_iid[message_id] = safe_iid
        
        return safe_iid
    
    def get_msgid_from_selection(self, sel):
        """選択されたTreeview IIDから元のMessage-IDを取得"""
        if not sel:
            return None
        iid = sel[0] if isinstance(sel, tuple) else sel
        return self.iid_to_msgid.get(iid, iid)  # マッピングになければそのまま返す（後方互換性）
    
    def get_msgids_from_selection(self, selected):
        """複数選択されたTreeview IIDsから元のMessage-IDリストを取得"""
        return [self.iid_to_msgid.get(iid, iid) for iid in selected]
    def update_promo_button_state(self):
        """プロモ更新ボタンの状態を更新"""
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM promo_rules")
        count = cur.fetchone()[0]
        conn.close()
        
        if count > 0:
            self.btn_promo_update.config(state=tk.NORMAL, bg="#9b59b6", cursor="hand2")
        else:
            self.btn_promo_update.config(state=tk.DISABLED, bg="#bdc3c7", cursor="arrow")
    
    def apply_promo_rules_to_existing(self):
        """既存メールに対してプロモルールを適用"""
        # ルール取得
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("SELECT sender_pattern, target_folder FROM promo_rules")
        promo_rules = cur.fetchall()
        
        if not promo_rules:
            messagebox.showinfo("情報", "プロモルールが登録されていません")
            conn.close()
            return
        
        # 確認ダイアログ
        result = messagebox.askyesno(
            "確認", 
            f"{len(promo_rules)}件のプロモルールを既存メールに適用します。\n\n"
            "該当するメールをプロモ・ボックスに移動しますか？"
        )
        
        if not result:
            conn.close()
            return
        
        # プログレスバー表示
        progress_dialog = tk.Toplevel(self.root)
        progress_dialog.title("プロモルール適用中")
        progress_dialog.geometry("400x150")
        progress_dialog.transient(self.root)
        progress_dialog.grab_set()
        
        tk.Label(progress_dialog, text="プロモルールを適用しています...", 
                font=("Arial", 11, "bold")).pack(pady=20)
        
        progress_bar = ttk.Progressbar(progress_dialog, mode='indeterminate')
        progress_bar.pack(fill=tk.X, padx=30, pady=10)
        progress_bar.start(10)
        
        status_label = tk.Label(progress_dialog, text="処理中...", font=("Arial", 9))
        status_label.pack(pady=10)
        
        progress_dialog.update()
        
        # 処理実行
        moved_count = 0
        
        # 通常メール（is_promo=0）のみを対象
        cur.execute("SELECT message_id, sender FROM emails WHERE is_promo=0 OR is_promo IS NULL")
        emails = cur.fetchall()
        
        total = len(emails)
        processed = 0
        
        for msg_id, sender in emails:
            sender_clean = self.fetcher.clean_address(sender)
            
            # ルールにマッチするか確認
            for pattern, target_folder in promo_rules:
                if self.db_mgr.match_pattern(sender_clean, pattern):
                    # プロモに移動
                    cur.execute("UPDATE emails SET is_promo=1, folder=? WHERE message_id=?", 
                               (target_folder, msg_id))
                    # マッチカウント更新
                    cur.execute("UPDATE promo_rules SET match_count = match_count + 1 WHERE sender_pattern=?", (pattern,))
                    moved_count += 1
                    break
            
            processed += 1
            if processed % 50 == 0:
                status_label.config(text=f"処理中... {processed}/{total} 件")
                progress_dialog.update()
        
        conn.commit()
        conn.close()
        
        # プログレスバー非表示
        progress_bar.stop()
        progress_dialog.destroy()
        
        # 結果表示
        if moved_count > 0:
            messagebox.showinfo("完了", f"{moved_count}件のメールをプロモ・ボックスに移動しました")
            self.refresh_tree_from_db()
            # プロモルール管理画面が開いていれば更新
            if hasattr(self, 'promo_rules_tree'):
                self.refresh_promo_rules()
            self.refresh_folder_tree()
        else:
            messagebox.showinfo("完了", "該当するメールはありませんでした")
    def on_folder_select(self, event):
        """フォルダ選択時の処理"""
        sel = self.folder_tree.selection()
        if not sel:
            return
        
        folder_id = sel[0]
        
        # フォルダIDの解析
        if folder_id == "all":
            # 受信メール（プロモ除外）
            self.current_filter = None
            self.current_folder = None
            self.current_promo_filter = False
        elif folder_id == "promo":
            # プロモ・ボックス全体
            self.current_filter = None
            self.current_folder = None
            self.current_promo_filter = True
        elif ":" in folder_id:
            # サブフォルダ（プロバイダ:フォルダ名 or __promo__:フォルダ名）
            provider, folder_name = folder_id.split(":", 1)
            
            if provider == "__promo__":
                # プロモのサブフォルダ
                self.current_filter = None
                self.current_folder = folder_name
                self.current_promo_filter = True
            else:
                # 通常プロバイダのサブフォルダ
                self.current_filter = provider
                self.current_folder = folder_name
                self.current_promo_filter = False
        else:
            # プロバイダ直下
            self.current_filter = folder_id
            self.current_folder = None
            self.current_promo_filter = False
        
        # ページを1にリセット
        self.current_page = 1
        
        self.refresh_tree_from_db()
    
    def on_search_entry_change(self, *args):
        """検索窓の内容が変更されたときの処理"""
        keyword = self.search_var.get().strip()
        
        # 検索窓が空になり、かつ検索中の場合 → 自動クリア
        if not keyword and self.current_search:
            self.clear_search()
    
    def build_search_condition(self, search_text):
        """
        検索文字列からSQL条件を構築 (AND/OR対応)
        
        仕様:
        - スペース区切り → AND検索
        - "OR"キーワード → OR検索
        
        例:
        - "AI 開発" → "AI" AND "開発"
        - "AI OR 開発" → "AI" OR "開発"
        - "Python AI OR 機械学習" → "Python" AND "AI" OR "機械学習"
        
        戻り値: (sql_condition, params_list) のタプル、または None
        """
        if not search_text:
            return None
        
        # ORで分割 (大文字小文字区別なし)
        import re
        or_parts = re.split(r'\s+OR\s+', search_text, flags=re.IGNORECASE)
        
        or_conditions = []
        all_params = []
        
        for or_part in or_parts:
            # 各OR部分をスペースで分割してAND条件を作成
            and_keywords = or_part.strip().split()
            
            if not and_keywords:
                continue
            
            and_conditions = []
            
            for keyword in and_keywords:
                # 各キーワードは subject, sender, raw_data のいずれかにマッチ
                and_conditions.append("(subject LIKE ? OR sender LIKE ? OR raw_data LIKE ?)")
                search_param = f"%{keyword}%"
                all_params.extend([search_param, search_param, search_param])
            
            # AND条件を結合
            if and_conditions:
                or_conditions.append("(" + " AND ".join(and_conditions) + ")")
        
        if not or_conditions:
            return None
        
        # OR条件を結合
        final_condition = " OR ".join(or_conditions)
        
        return (final_condition, all_params)
    
    def show_search_help(self):
        """検索機能のヘルプを表示"""
        help_dialog = tk.Toplevel(self.root)
        help_dialog.title("検索機能ヘルプ")
        help_dialog.geometry("600x500")
        help_dialog.transient(self.root)
        help_dialog.grab_set()
        
        # タイトル
        tk.Label(help_dialog, text="🔍 検索機能の使い方", 
                font=("Arial", 14, "bold"), bg="#2196F3", fg="white").pack(fill=tk.X, pady=(0, 20))
        
        # スクロール可能なフレーム
        canvas = tk.Canvas(help_dialog)
        scrollbar = tk.Scrollbar(help_dialog, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # ヘルプコンテンツ
        help_content = [
            ("基本検索", "キーワードを入力すると、件名・送信者・本文から検索します"),
            ("", "例: AI → 「AI」を含むメールを検索"),
            
            ("AND検索（複数条件）", "スペースで区切ると、すべてのキーワードを含むメールを検索"),
            ("", "例: AI 開発 → 「AI」と「開発」の両方を含むメール"),
            
            ("OR検索（いずれか）", "「OR」で区切ると、いずれかのキーワードを含むメールを検索"),
            ("", "例: Python OR Java → 「Python」または「Java」を含むメール"),
            
            ("組み合わせ検索", "ANDとORを組み合わせて複雑な検索も可能"),
            ("", "例: AI 開発 OR プログラミング"),
            ("", "   → 「AI」と「開発」の両方を含む、または「プログラミング」を含む"),
            
            ("添付ファイル検索", "「添付」と入力すると添付ファイル付きメールを検索"),
            ("", "例: 添付 または 添付あり"),
            
            ("検索のクリア", "検索窓を空にするか、バナーの「✗ クリア」ボタンで解除"),
        ]
        
        for i, (title, desc) in enumerate(help_content):
            frame = tk.Frame(scrollable_frame, bg="white" if i % 2 == 0 else "#f5f5f5")
            frame.pack(fill=tk.X, padx=10, pady=2)
            
            if title:
                tk.Label(frame, text=title, font=("Arial", 10, "bold"), 
                        bg=frame["bg"], fg="#1976D2", anchor="w").pack(fill=tk.X, padx=10, pady=(5, 2))
            
            tk.Label(frame, text=desc, font=("Arial", 9), 
                    bg=frame["bg"], fg="#424242", anchor="w", wraplength=550, 
                    justify=tk.LEFT).pack(fill=tk.X, padx=20 if title else 30, pady=(0, 5))
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 閉じるボタン
        tk.Button(help_dialog, text="閉じる", command=help_dialog.destroy, 
                 bg="#2196F3", fg="white", width=20, font=("Arial", 10, "bold")).pack(pady=10)
        
        # ダイアログを中央に配置
        help_dialog.update_idletasks()
        x = (help_dialog.winfo_screenwidth() // 2) - (help_dialog.winfo_width() // 2)
        y = (help_dialog.winfo_screenheight() // 2) - (help_dialog.winfo_height() // 2)
        help_dialog.geometry(f"+{x}+{y}")
    
    def do_search(self):
        """検索実行"""
        keyword = self.search_entry.get().strip()
        self.current_search = keyword
        
        # ページを1にリセット
        self.current_page = 1
        
        if keyword:
            # 検索ボックス背景色変更
            self.search_entry.config(bg="#fff3cd")
            
            # バナー表示
            self.search_banner_label.config(text=f"🔍 検索中: \"{keyword}\"")
            self.search_banner.pack(fill=tk.X, before=self.tree_frame, padx=5, pady=(5, 0))
        
        self.refresh_tree_from_db()
    
    def clear_search(self):
        """検索クリア"""
        # 無限ループ防止：すでにクリア済みなら何もしない
        if not self.current_search:
            return
        
        # trace一時停止してクリア（無限ループ防止）
        try:
            self.search_var.trace_remove("write", self.search_var.trace_info()[0][1])
        except:
            pass
        
        self.search_entry.delete(0, tk.END)
        self.current_search = ""
        
        # trace再開
        self.search_var.trace_add("write", self.on_search_entry_change)
        
        # 検索ボックス背景色リセット
        self.search_entry.config(bg="white")
        
        # バナー非表示
        self.search_banner.pack_forget()
        self.refresh_tree_from_db()
    
    def refresh_tree_from_db(self):
        """メール一覧更新（フィルタ・検索・プロモ対応）"""
        self.tree.delete(*self.tree.get_children())
        
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        
        # SQL構築（is_replied, attachments追加）
        sql = "SELECT message_id, original_to, subject, sender, date_disp, provider, read_flag, is_replied, attachments FROM emails"
        conditions = []
        params = []
        
        # 削除済み除外（ゴミ箱のメールは表示）
        conditions.append("(is_deleted=0 OR is_deleted IS NULL)")
        
        # プロモフィルタ
        if self.current_promo_filter:
            conditions.append("is_promo=1")
        else:
            conditions.append("(is_promo=0 OR is_promo IS NULL)")
        
        # プロバイダフィルタ
        if self.current_filter:
            conditions.append("provider=?")
            params.append(self.current_filter)
        
        # サブフォルダフィルタ
        if self.current_folder:
            conditions.append("folder=?")
            params.append(self.current_folder)
        
        # 検索フィルタ
        if self.current_search:
            # 特殊検索: "添付" または "添付あり"
            if self.current_search.strip() in ["添付", "添付あり", "attachment", "attachments"]:
                conditions.append("(attachments IS NOT NULL AND attachments != '' AND attachments != '[]')")
            else:
                # AND/OR検索対応
                search_condition = self.build_search_condition(self.current_search)
                if search_condition:
                    conditions.append(search_condition[0])
                    params.extend(search_condition[1])
        
        if conditions:
            sql += " WHERE " + " AND ".join(conditions)
        
        # 総件数を取得
        count_sql = "SELECT COUNT(*) FROM emails"
        if conditions:
            count_sql += " WHERE " + " AND ".join(conditions)
        
        cur.execute(count_sql, params)
        self.total_items = cur.fetchone()[0]
        
        # ページング計算
        total_pages = max(1, (self.total_items + self.items_per_page - 1) // self.items_per_page)
        
        # 現在のページが範囲外の場合は調整
        if self.current_page > total_pages:
            self.current_page = total_pages
        if self.current_page < 1:
            self.current_page = 1
        
        # オフセット計算
        offset = (self.current_page - 1) * self.items_per_page
        
        # ページ情報更新
        self.page_info_label.config(text=f"ページ {self.current_page} / {total_pages} ({self.total_items} 件)")
        
        sql += f" ORDER BY timestamp DESC LIMIT {self.items_per_page} OFFSET {offset}"
        
        cur.execute(sql, params)
        rows = cur.fetchall()
        conn.close()
        
        # 色設定取得
        provider_colors = self.config_mgr.get("provider_colors") or {}
        
        for r in rows:
            msg_id, to, subj, frm, date, provider, read_flag, is_replied, attachments = r
            
            # 添付ファイルアイコン追加
            has_attachments = attachments and attachments.strip() not in ["", "null", "[]"]
            
            # 返信済みアイコン追加
            if is_replied:
                display_subject = f"↩️ {subj}"
            else:
                display_subject = subj
            
            # 添付ファイル表示を追加（件名の先頭）
            if has_attachments:
                display_subject = f"【添付あり】{display_subject}"
            
            # 色タグ生成
            if provider not in provider_colors:
                provider_colors[provider] = self.generate_pastel_color(provider)
                self.config_mgr.set("provider_colors", provider_colors)
                self.config_mgr.save()
            
            tags = [f"provider_{provider}"]
            if not read_flag:
                tags.append("unread")
            
            safe_iid = self.make_safe_iid(msg_id)
            self.tree.insert("", tk.END, iid=safe_iid, values=(to, display_subject, frm, date), tags=tuple(tags))
            
            # タグ設定
            self.tree.tag_configure(f"provider_{provider}", background=provider_colors[provider])
            self.tree.tag_configure("unread", font=("Arial", 10, "bold"))
        
        # 色設定リスト更新
        self.refresh_color_list()
    
    def generate_pastel_color(self, seed_str):
        """文字列からパステルカラー生成"""
        hash_val = int(hashlib.md5(seed_str.encode()).hexdigest()[:6], 16)
        r = 200 + (hash_val % 56)
        g = 200 + ((hash_val >> 8) % 56)
        b = 200 + ((hash_val >> 16) % 56)
        return f"#{r:02x}{g:02x}{b:02x}"
    
    def update_progress(self, current, total):
        """取得進捗更新"""
        if total > 0:
            percentage = int((current / total) * 100)
            self.progress_bar['value'] = percentage
            self.lbl_progress.config(text=f"{current}/{total}件取得中... ({percentage}%)")
            self.root.update()
    
    def show_progress(self):
        """プログレスバー表示"""
        self.progress_frame.pack(side=tk.RIGHT, padx=10)
        self.progress_bar['value'] = 0
        self.lbl_progress.config(text="取得準備中...")
        self.root.update()
    
    def hide_progress(self):
        """プログレスバー非表示"""
        self.progress_frame.pack_forget()
        self.root.update()
    
    def sort_tree_column(self, col):
        """カラムヘッダークリックでソート（トグル）"""
        # 同じ列をクリックした場合は昇順/降順をトグル
        if self.sort_column == col:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = col
            self.sort_reverse = False  # 新しい列は昇順から開始
        
        # カラムインデックス取得
        col_index = {"宛先": 0, "件名": 1, "送信者": 2, "日付": 3}[col]
        
        # 現在表示中のアイテムを取得してソート
        items = [(self.tree.set(item, col), item) for item in self.tree.get_children('')]
        
        # 日付の場合は特別処理（文字列比較だと正しくソートされない）
        if col == "日付":
            # "YYYY/MM/DD HH:MM:SS" 形式なので文字列比較でOK
            items.sort(reverse=self.sort_reverse)
        else:
            # その他は通常の文字列ソート
            items.sort(reverse=self.sort_reverse)
        
        # ソート結果でツリーを並び替え
        for index, (val, item) in enumerate(items):
            self.tree.move(item, '', index)
        
        # ヘッダーに矢印表示を更新
        self.update_column_headers(col)
    
    def update_column_headers(self, sorted_col):
        """カラムヘッダーにソート矢印を表示"""
        cols = ["宛先", "件名", "送信者", "日付"]
        
        for col in cols:
            if col == sorted_col:
                # ソート中の列には矢印を追加
                arrow = " ▼" if self.sort_reverse else " ▲"
                self.tree.heading(col, text=f"{col}{arrow}")
            else:
                # その他の列は矢印なし
                self.tree.heading(col, text=col)
    
    def refresh_color_list(self):
        """色設定リスト更新"""
        self.tree_colors.delete(*self.tree_colors.get_children())
        
        provider_colors = self.config_mgr.get("provider_colors") or {}
        for provider, color in provider_colors.items():
            self.tree_colors.insert("", tk.END, values=(provider, color))
    
    def change_provider_color(self):
        """プロバイダ色変更"""
        sel = self.tree_colors.selection()
        if not sel:
            messagebox.showwarning("未選択", "色を変更するプロバイダを選択してください")
            return
        
        provider = self.tree_colors.item(sel, "values")[0]
        current_color = self.tree_colors.item(sel, "values")[1]
        
        color = colorchooser.askcolor(title=f"{provider}の色を選択", initialcolor=current_color)
        if color[1]:
            colors = self.config_mgr.get("provider_colors") or {}
            colors[provider] = color[1]
            self.config_mgr.set("provider_colors", colors)
            self.config_mgr.save()
            self.refresh_tree_from_db()
            messagebox.showinfo("完了", f"{provider}の色を変更しました")
    
    def reset_provider_color(self):
        """プロバイダ色を自動生成色に戻す"""
        sel = self.tree_colors.selection()
        if not sel:
            messagebox.showwarning("未選択", "リセットするプロバイダを選択してください")
            return
        
        provider = self.tree_colors.item(sel, "values")[0]
        
        colors = self.config_mgr.get("provider_colors") or {}
        colors[provider] = self.generate_pastel_color(provider)
        self.config_mgr.set("provider_colors", colors)
        self.config_mgr.save()
        self.refresh_tree_from_db()
        messagebox.showinfo("完了", f"{provider}の色を自動生成色に戻しました")
    
    def on_mail_select(self, event):
        """メール選択時のプレビュー表示"""
        sel = self.tree.selection()
        if not sel:
            return
        
        # selはTreeviewの安全なIID、msg_idは元のMessage-ID
        safe_iid = sel[0]
        msg_id = self.get_msgid_from_selection(sel)
        
        # 既読マーク（DBには元のMessage-IDを使用）
        self.db_mgr.mark_as_read(msg_id)
        
        # プレビュー表示
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("SELECT raw_data, subject, sender FROM emails WHERE message_id=?", (msg_id,))
        row = cur.fetchone()
        conn.close()
        
        if row:
            raw, subj, sender = row
            msg = email.message_from_string(raw)
            body = self.fetcher.extract_text_body(msg)
            
            # テキスト除去（タグ除去）
            body_clean = re.sub(r'<[^>]+>', '', body)
            
            self.preview_text.config(state=tk.NORMAL)
            self.preview_text.delete("1.0", tk.END)
            
            preview = f"件名: {subj}\n"
            preview += f"送信者: {sender}\n"
            preview += f"日付: {self.fetcher.decode_h(msg.get('Date'))}\n"
            preview += "=" * 40 + "\n\n"
            preview += body_clean[:1000]
            if len(body_clean) > 1000:
                preview += "\n\n... (続きはダブルクリックで表示)"
            
            self.preview_text.insert("1.0", preview)
            self.preview_text.config(state=tk.DISABLED)
            
            # ツリーの太字解除（既読化）- Treeviewにはsafe_iidを使用
            tags = list(self.tree.item(safe_iid, "tags"))
            if "unread" in tags:
                tags.remove("unread")
                self.tree.item(safe_iid, tags=tuple(tags))
    
    def open_viewer(self, event=None):
        """メール詳細ビューア起動（下書きは編集モード）"""
        sel = self.tree.selection()
        if not sel:
            return
        
        msg_id = self.get_msgid_from_selection(sel)
        
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("SELECT raw_data, subject, folder, original_to, attachments FROM emails WHERE message_id=?", (msg_id,))
        row = cur.fetchone()
        conn.close()
        
        if row:
            raw_data, subj, folder, original_to, attachments_json = row
            
            # 下書きの場合は編集ウィンドウを開く
            if folder == "__drafts__":
                self.open_draft_editor(msg_id, original_to, subj, raw_data)
                return
            
            if not raw_data:
                messagebox.showwarning("データ空", "このメールの本文データがDBに保存されていません。")
                return
            
            # 添付ファイル情報をパース
            import json
            attachments = json.loads(attachments_json) if attachments_json else []
            
            MailViewer(self.root, raw_data, subj, self.config_mgr, msg_id, attachments)  # msg_idとattachmentsを追加
    
    def open_draft_editor(self, draft_id, to, subject, body):
        """下書き編集ウィンドウ"""
        # SMTPアカウント取得
        smtp_accounts = self.config_mgr.get("smtp_accounts") or []
        
        if not smtp_accounts:
            providers = self.config_mgr.get("providers") or []
            for p in providers:
                smtp_accounts.append({
                    "email": p.get("email", ""),
                    "password": p.get("password", ""),
                    "smtp_server": p.get("smtp_host", "smtp.gmail.com"),
                    "smtp_port": 465
                })
        
        if not smtp_accounts:
            messagebox.showerror("エラー", "SMTP設定が登録されていません。")
            return
        
        win = tk.Toplevel(self.root)
        win.title(f"📝 下書き編集: {subject}")
        win.geometry("800x600")
        
        # 送信元選択
        tk.Label(win, text="送信元:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        
        from_var = tk.StringVar()
        from_combo = ttk.Combobox(win, textvariable=from_var, width=50, state="readonly")
        from_combo['values'] = [acc['email'] for acc in smtp_accounts]
        from_combo.current(0)
        from_combo.grid(row=0, column=1, sticky="w", padx=10, pady=5)
        
        # 宛先
        tk.Label(win, text="宛先:", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky="w", padx=10, pady=5)
        to_entry = tk.Entry(win, width=50)
        to_entry.insert(0, to)
        to_entry.grid(row=1, column=1, sticky="w", padx=10, pady=5)
        
        # 件名
        tk.Label(win, text="件名:", font=("Arial", 10, "bold")).grid(row=2, column=0, sticky="w", padx=10, pady=5)
        subject_entry = tk.Entry(win, width=50)
        subject_entry.insert(0, subject)
        subject_entry.grid(row=2, column=1, sticky="w", padx=10, pady=5)
        
        # 本文
        tk.Label(win, text="本文:", font=("Arial", 10, "bold")).grid(row=3, column=0, sticky="nw", padx=10, pady=5)
        body_text = tk.Text(win, width=60, height=20, wrap=tk.WORD)
        body_text.insert("1.0", body)
        body_text.grid(row=3, column=1, padx=10, pady=5)
        
        def do_send():
            """送信処理"""
            to_val = to_entry.get().strip()
            subject_val = subject_entry.get().strip()
            body_val = body_text.get("1.0", tk.END).strip()
            from_email = from_var.get()
            
            if not to_val or not subject_val:
                messagebox.showerror("エラー", "宛先と件名は必須です")
                return
            
            smtp_account = next((acc for acc in smtp_accounts if acc['email'] == from_email), None)
            
            if not smtp_account:
                messagebox.showerror("エラー", "送信アカウント情報が見つかりません")
                return
            
            try:
                import smtplib
                from email.mime.text import MIMEText
                
                msg = MIMEText(body_val, "plain", "utf-8")
                msg["Subject"] = subject_val
                msg["From"] = from_email
                msg["To"] = to_val
                
                with smtplib.SMTP_SSL(smtp_account['smtp_server'], smtp_account.get('smtp_port', 465)) as server:
                    server.login(smtp_account['email'], smtp_account['password'])
                    server.send_message(msg)
                
                # 下書きを削除
                conn = sqlite3.connect(DB_FILE)
                cur = conn.cursor()
                cur.execute("DELETE FROM emails WHERE message_id=?", (draft_id,))
                conn.commit()
                conn.close()
                
                # 送信済みとして保存
                from datetime import datetime
                import uuid
                
                sent_email_data = {
                    "message_id": str(uuid.uuid4()),
                    "original_to": to_val,
                    "subject": subject_val,
                    "sender": from_email,
                    "date_disp": datetime.now().strftime("%Y/%m/%d %H:%M:%S"),
                    "timestamp": datetime.now().isoformat(),
                    "raw_data": body_val,
                    "provider": from_email.split("@")[-1] if "@" in from_email else "unknown"
                }
                
                self.db_mgr.save_sent_email(sent_email_data)
                
                self.refresh_tree_from_db()
                self.refresh_folder_tree()
                
                messagebox.showinfo("成功", "メールを送信しました")
                win.destroy()
                
            except Exception as e:
                messagebox.showerror("送信失敗", f"メール送信に失敗しました:\n{e}")
        
        def update_draft():
            """下書き更新"""
            to_val = to_entry.get().strip()
            subject_val = subject_entry.get().strip()
            body_val = body_text.get("1.0", tk.END).strip()
            
            if not subject_val:
                messagebox.showerror("エラー", "件名は必須です")
                return
            
            conn = sqlite3.connect(DB_FILE)
            cur = conn.cursor()
            cur.execute("""
                UPDATE emails 
                SET original_to=?, subject=?, raw_data=?
                WHERE message_id=?
            """, (to_val, subject_val, body_val, draft_id))
            conn.commit()
            conn.close()
            
            self.refresh_tree_from_db()
            
            messagebox.showinfo("保存完了", "下書きを更新しました")
            win.destroy()
        
        def delete_draft():
            """下書き削除"""
            if messagebox.askyesno("確認", "この下書きを削除しますか？"):
                conn = sqlite3.connect(DB_FILE)
                cur = conn.cursor()
                cur.execute("DELETE FROM emails WHERE message_id=?", (draft_id,))
                conn.commit()
                conn.close()
                
                self.refresh_tree_from_db()
                self.refresh_folder_tree()
                
                messagebox.showinfo("削除完了", "下書きを削除しました")
                win.destroy()
        
        btn_frame = tk.Frame(win)
        btn_frame.grid(row=4, column=1, pady=10)
        
        tk.Button(btn_frame, text="📤 送信", command=do_send, bg="#4CAF50", fg="white", width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="💾 更新", command=update_draft, bg="#FF9800", fg="white", width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="🗑️ 削除", command=delete_draft, bg="#f44336", fg="white", width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="❌ 閉じる", command=win.destroy, bg="#9E9E9E", fg="white", width=12).pack(side=tk.LEFT, padx=5)
    
    def reset_database(self):
        """データベース初期化"""
        if messagebox.askyesno("警告", "保存されている全てのメールデータを削除します。\nよろしいですか？\n(設定情報は消えません)"):
            self.db_mgr.reset_db()
            self.refresh_tree_from_db()
            self.refresh_folder_tree()
            messagebox.showinfo("完了", "データベースを初期化しました。\n「受信」ボタンを押して再取得してください。")
    
    def show_folder_context_menu(self, event):
        """フォルダツリー右クリックメニュー"""
        item = self.folder_tree.identify_row(event.y)
        if not item:
            return
        
        self.folder_tree.selection_set(item)
        
        # プロバイダかどうか判定
        if ":" in item:
            # サブフォルダ
            provider, folder_name = item.split(":", 1)
            
            # ゴミ箱の場合
            if folder_name == "__trash__":
                menu = tk.Menu(self.root, tearoff=0)
                menu.add_command(label="🗑️ ゴミ箱を空にする", command=lambda: self.empty_trash(provider))
                menu.add_command(label="↩️ 元に戻す", command=lambda: self.restore_from_trash(provider))
                menu.post(event.x_root, event.y_root)
                return
            
            # カスタムフォルダの場合
            if not folder_name.startswith("__"):
                menu = tk.Menu(self.root, tearoff=0)
                menu.add_command(label="✏️ 名前変更", command=lambda: self.rename_folder(provider, folder_name))
                menu.add_command(label="🗑️ 削除", command=lambda: self.delete_custom_folder(provider, folder_name))
                menu.post(event.x_root, event.y_root)
        else:
            # プロバイダノードまたはプロモ
            if item == "all":
                return  # 全件は編集不可
            
            menu = tk.Menu(self.root, tearoff=0)
            
            if item == "promo":
                # プロモ・ボックス用
                menu.add_command(label="➕ 新規フォルダ作成", command=lambda: self.create_custom_folder("__promo__"))
            else:
                # 通常プロバイダ用
                menu.add_command(label="➕ 新規フォルダ作成", command=lambda: self.create_custom_folder(item))
            
            menu.post(event.x_root, event.y_root)
    
    def show_mail_context_menu(self, event):
        """メールリスト右クリックメニュー"""
        item = self.tree.identify_row(event.y)
        if not item:
            return
        
        # 選択されたメールを取得
        selected = self.tree.selection()
        
        # 選択が空の場合は何もしない
        if not selected:
            return
        
        # 安全なIIDから元のMessage-IDリストを取得
        msg_ids = self.get_msgids_from_selection(selected)
        
        # 選択メールがゴミ箱にあるかチェック
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("SELECT folder FROM emails WHERE message_id=?", (msg_ids[0],))
        row = cur.fetchone()
        conn.close()
        
        is_in_trash = row and row[0] == "__trash__"
        
        menu = tk.Menu(self.root, tearoff=0)
        
        # ゴミ箱内のメールの場合
        if is_in_trash:
            menu.add_command(label="↩️ 元に戻す", command=lambda: self.restore_from_trash_single(msg_ids))
            menu.add_separator()
            menu.add_command(label="❌ 完全削除", command=lambda: self.permanently_delete_emails(msg_ids))
        else:
            # 通常メールの場合
            # 1件選択時: 返信・転送有効
            if len(selected) == 1:
                menu.add_command(label="↩️ 返信", command=self.open_reply_window)
                menu.add_command(label="↪️ 転送", command=self.open_forward_window)
            else:
                # 複数選択時: 返信・転送無効（灰色）
                menu.add_command(label="↩️ 返信", command=None, state=tk.DISABLED)
                menu.add_command(label="↪️ 転送", command=None, state=tk.DISABLED)
            
            menu.add_separator()
            
            # フォルダへ移動（サブメニュー化）
            move_menu = tk.Menu(menu, tearoff=0)
            
            if len(selected) == 1:
                # 1件選択: プロモ + 該当プロバイダ
                self.build_move_menu_single(move_menu, item)
            else:
                # 複数選択: プロモのみ
                self.build_move_menu_multiple(move_menu)
            
            menu.add_cascade(label="📂 フォルダへ移動", menu=move_menu)
            
            menu.add_separator()
            
            # 既読/未読切り替え
            menu.add_command(label="✉️ 未読にする", command=lambda: self.mark_as_unread(selected))
            menu.add_command(label="✅ 既読にする", command=lambda: self.mark_as_read(selected))
            
            menu.add_separator()
            
            # ゴミ箱へ移動
            menu.add_command(label="🗑️ ゴミ箱へ移動", command=lambda: self.move_to_trash(selected))
        
        menu.post(event.x_root, event.y_root)
    
    def restore_from_trash_single(self, message_ids):
        """ゴミ箱から元に戻す（個別メール用）"""
        if not message_ids:
            return
        
        result = messagebox.askyesno("確認", f"{len(message_ids)}件のメールを元に戻しますか？")
        
        if not result:
            return
        
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        
        for msg_id in message_ids:
            # folder を NULL に戻す（元のフォルダに復元）
            cur.execute("UPDATE emails SET folder=NULL WHERE message_id=?", (msg_id,))
        
        conn.commit()
        conn.close()
        
        messagebox.showinfo("完了", f"{len(message_ids)}件のメールを元に戻しました")
        self.refresh_tree_from_db()
        self.refresh_folder_tree()
    
    def permanently_delete_emails(self, message_ids):
        """メールを完全削除"""
        if not message_ids:
            return
        
        # 削除モード取得
        delete_mode = self.config_mgr.get("delete_mode") or "local_only"
        mode_text = "Mail Hubのみ" if delete_mode == "local_only" else "Gmail Cloudからも削除"
        
        result = messagebox.askyesno(
            "警告",
            f"{len(message_ids)}件のメールを完全削除します。\n\n"
            f"削除モード: {mode_text}\n\n"
            "この操作は取り消せません。よろしいですか？",
            icon='warning'
        )
        
        if not result:
            return
        
        # 完全削除処理
        for msg_id in message_ids:
            self.db_mgr.permanently_delete_email(msg_id, delete_mode, self.config_mgr)
        
        messagebox.showinfo("完了", f"{len(message_ids)}件のメールを完全削除しました")
        self.refresh_tree_from_db()
        self.refresh_folder_tree()
    
    def mark_as_read(self, message_ids):
        """既読にする"""
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        
        for msg_id in message_ids:
            cur.execute("UPDATE emails SET read_flag=1 WHERE message_id=?", (msg_id,))
        
        conn.commit()
        conn.close()
        
        self.refresh_tree_from_db()
    
    def mark_as_unread(self, message_ids):
        """未読にする"""
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        
        for msg_id in message_ids:
            cur.execute("UPDATE emails SET read_flag=0 WHERE message_id=?", (msg_id,))
        
        conn.commit()
        conn.close()
        
        self.refresh_tree_from_db()
    
    def move_to_trash(self, message_ids):
        """ゴミ箱へ移動"""
        if not message_ids:
            return
        
        result = messagebox.askyesno("確認", f"{len(message_ids)}件のメールをゴミ箱に移動しますか？")
        
        if not result:
            return
        
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        
        for msg_id in message_ids:
            # プロバイダを取得
            cur.execute("SELECT provider, is_promo FROM emails WHERE message_id=?", (msg_id,))
            row = cur.fetchone()
            
            if row:
                provider, is_promo = row
                
                # ゴミ箱へ移動（folderを__trash__に設定）
                cur.execute("UPDATE emails SET folder='__trash__' WHERE message_id=?", (msg_id,))
        
        conn.commit()
        conn.close()
        
        messagebox.showinfo("完了", f"{len(message_ids)}件のメールをゴミ箱に移動しました")
        self.refresh_tree_from_db()
        self.refresh_folder_tree()
    
    def build_move_menu_single(self, menu, message_id):
        """1件選択時のフォルダ移動メニュー構築"""
        # メールのプロバイダを取得
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("SELECT provider FROM emails WHERE message_id=?", (message_id,))
        row = cur.fetchone()
        conn.close()
        
        if not row:
            return
        
        provider = row[0]
        
        # プロモ・ボックス内にいる場合、「通常メールへ」を先頭に追加
        if self.current_promo_filter:
            menu.add_command(label="📧 通常メールへ", command=self.release_from_promo)
            menu.add_separator()
        
        # プロモ・ボックス
        menu.add_command(label="📂 プロモ・ボックス", command=self.move_to_promo)
        
        # プロモのゴミ箱
        menu.add_command(label="  🗑️ ゴミ箱", command=lambda: self.move_to_folder_direct("__promo__", "__trash__"))
        
        # プロモのカスタムフォルダ
        promo_folders = self.db_mgr.get_folders("__promo__")
        for folder_name, folder_type in promo_folders:
            if folder_type == 'custom':
                menu.add_command(label=f"  📂 {folder_name}", command=lambda fn=folder_name: self.move_to_folder_direct("__promo__", fn))
        
        # プロモに新規フォルダ作成
        menu.add_command(label="  ➕ 新規フォルダ作成", command=lambda: self.create_custom_folder("__promo__"))
        
        menu.add_separator()
        
        # 該当プロバイダ
        if provider:
            menu.add_command(label=f"📧 {provider}", command=lambda: self.move_to_folder_direct(provider, None))
            
            # プロバイダのゴミ箱
            menu.add_command(label="  🗑️ ゴミ箱", command=lambda: self.move_to_folder_direct(provider, "__trash__"))
            
            # プロバイダのカスタムフォルダ
            provider_folders = self.db_mgr.get_folders(provider)
            for folder_name, folder_type in provider_folders:
                if folder_type == 'custom':
                    menu.add_command(label=f"  📂 {folder_name}", command=lambda fn=folder_name: self.move_to_folder_direct(provider, fn))
            
            # プロバイダに新規フォルダ作成
            menu.add_command(label="  ➕ 新規フォルダ作成", command=lambda: self.create_custom_folder(provider))
    
    def build_move_menu_multiple(self, menu):
        """複数選択時のフォルダ移動メニュー構築"""
        # 選択されたメールのプロバイダを全取得
        selected = self.tree.selection()
        providers = set()
        
        # 安全なIIDから元のMessage-IDリストを取得
        msg_ids = self.get_msgids_from_selection(selected)
        
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        
        for msg_id in msg_ids:
            cur.execute("SELECT provider FROM emails WHERE message_id=?", (msg_id,))
            row = cur.fetchone()
            if row:
                providers.add(row[0])
        
        conn.close()
        
        # プロモ・ボックス内にいる場合、「通常メールへ」を先頭に追加
        if self.current_promo_filter:
            menu.add_command(label="📧 通常メールへ", command=self.release_from_promo)
            menu.add_separator()
        
        # プロモ・ボックス（常に表示）
        menu.add_command(label="📂 プロモ・ボックス", command=self.move_to_promo)
        
        # プロモのゴミ箱
        menu.add_command(label="  🗑️ ゴミ箱", command=lambda: self.move_to_folder_direct("__promo__", "__trash__"))
        
        # プロモのカスタムフォルダ
        promo_folders = self.db_mgr.get_folders("__promo__")
        for folder_name, folder_type in promo_folders:
            if folder_type == 'custom':
                menu.add_command(label=f"  📂 {folder_name}", command=lambda fn=folder_name: self.move_to_folder_direct("__promo__", fn))
        
        # プロモに新規フォルダ作成
        menu.add_command(label="  ➕ 新規フォルダ作成", command=lambda: self.create_custom_folder("__promo__"))
        
        # 単一プロバイダの場合のみ、該当プロバイダも表示
        if len(providers) == 1:
            provider = list(providers)[0]
            
            menu.add_separator()
            
            # 該当プロバイダ
            menu.add_command(label=f"📧 {provider}", command=lambda: self.move_to_folder_direct(provider, None))
            
            # プロバイダのゴミ箱
            menu.add_command(label="  🗑️ ゴミ箱", command=lambda: self.move_to_folder_direct(provider, "__trash__"))
            
            # プロバイダのカスタムフォルダ
            provider_folders = self.db_mgr.get_folders(provider)
            for folder_name, folder_type in provider_folders:
                if folder_type == 'custom':
                    menu.add_command(label=f"  📂 {folder_name}", command=lambda fn=folder_name, p=provider: self.move_to_folder_direct(p, fn))
            
            # プロバイダに新規フォルダ作成
            menu.add_command(label="  ➕ 新規フォルダ作成", command=lambda p=provider: self.create_custom_folder(p))
    
    def release_from_promo(self):
        """プロモ・ボックスから通常メールへ解放"""
        selected = self.tree.selection()
        if not selected:
            return
        
        # 安全なIIDから元のMessage-IDリストを取得
        msg_ids = self.get_msgids_from_selection(selected)
        
        # 選択されたメールの送信者ドメインを取得
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        
        domains = set()
        for msg_id in msg_ids:
            cur.execute("SELECT sender FROM emails WHERE message_id=?", (msg_id,))
            row = cur.fetchone()
            if row:
                sender = self.fetcher.clean_address(row[0])
                if "@" in sender:
                    domain = sender.split("@")[-1]
                    domains.add(domain)
        
        # 該当するプロモルールを検索
        patterns_to_delete = []
        for domain in domains:
            pattern = f"%@{domain}%"
            cur.execute("SELECT sender_pattern FROM promo_rules WHERE sender_pattern=?", (pattern,))
            if cur.fetchone():
                patterns_to_delete.append(pattern)
        
        conn.close()
        
        # 確認ダイアログ（ルール削除オプション付き）
        dialog = tk.Toplevel(self.root)
        dialog.title("プロモ解放確認")
        dialog.geometry("550x450")  # サイズ拡大
        dialog.transient(self.root)
        dialog.grab_set()
        
        # ヘッダー
        header_frame = tk.Frame(dialog, bg="#3498db")
        header_frame.pack(fill=tk.X)
        
        tk.Label(
            header_frame, 
            text=f"{len(selected)}件のメールをプロモ・ボックスから解放します", 
            font=("Arial", 12, "bold"),
            bg="#3498db",
            fg="white"
        ).pack(pady=15)
        
        # 警告メッセージ
        warning_frame = tk.Frame(dialog, bg="#fff3cd", relief=tk.GROOVE, bd=2)
        warning_frame.pack(fill=tk.X, padx=10, pady=10)
        
        warning_text = (
            "⚠️ 注意事項\n\n"
            "プロモルールが登録されている場合、次回メール取得時に\n"
            "再びプロモ・ボックスに振り分けられます。"
        )
        tk.Label(warning_frame, text=warning_text, font=("Arial", 9), 
                bg="#fff3cd", fg="#856404", justify=tk.LEFT).pack(padx=10, pady=10)
        
        # ルール削除オプション
        var_delete_rules = tk.BooleanVar(value=False)
        
        if patterns_to_delete:
            rule_frame = tk.LabelFrame(dialog, text="プロモルール削除オプション", font=("Arial", 10, "bold"))
            rule_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            chk = tk.Checkbutton(
                rule_frame, 
                text="プロモルールも削除する（今後このドメインからのメールはプロモに入りません）",
                variable=var_delete_rules,
                font=("Arial", 9)
            )
            chk.pack(anchor=tk.W, padx=10, pady=5)
            
            tk.Label(rule_frame, text="検出されたルール:", font=("Arial", 9, "bold")).pack(anchor=tk.W, padx=10, pady=(10, 5))
            
            listbox = tk.Listbox(rule_frame, height=5, font=("Arial", 9))
            listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
            
            for pattern in patterns_to_delete:
                listbox.insert(tk.END, f"  ・{pattern}")
        
        # ボタン（拡大＆改善）
        result = {"confirmed": False, "delete_rules": False}
        
        def on_ok():
            result["confirmed"] = True
            result["delete_rules"] = var_delete_rules.get()
            dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        btn_frame = tk.Frame(dialog, bg="#ecf0f1", relief=tk.RAISED, bd=2)
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=0)
        
        button_container = tk.Frame(btn_frame, bg="#ecf0f1")
        button_container.pack(pady=15)
        
        tk.Button(
            button_container, 
            text="✓ 解放する", 
            command=on_ok, 
            bg="#27ae60", 
            fg="white", 
            font=("Arial", 11, "bold"),
            width=18,
            height=2
        ).pack(side=tk.LEFT, padx=10)
        
        tk.Button(
            button_container, 
            text="✗ キャンセル", 
            command=on_cancel, 
            bg="#e74c3c", 
            fg="white", 
            font=("Arial", 11, "bold"),
            width=18,
            height=2
        ).pack(side=tk.LEFT, padx=10)
        
        # ダイアログを中央に配置
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        dialog.wait_window()
        
        if not result["confirmed"]:
            return
        
        # プロモ解放処理
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        
        for msg_id in msg_ids:
            cur.execute("UPDATE emails SET is_promo=0, folder=NULL WHERE message_id=?", (msg_id,))
        
        # プロモルール削除（オプション）
        if result["delete_rules"] and patterns_to_delete:
            for pattern in patterns_to_delete:
                cur.execute("DELETE FROM promo_rules WHERE sender_pattern=?", (pattern,))
        
        conn.commit()
        conn.close()
        
        # 完了メッセージ
        rules_msg = f"（プロモルール{len(patterns_to_delete)}件削除）" if result["delete_rules"] else ""
        messagebox.showinfo("完了", f"{len(selected)}件のメールを通常メールに移動しました{rules_msg}")
        
        self.refresh_tree_from_db()
        self.refresh_folder_tree()
    
    def move_to_folder_direct(self, provider, folder):
        """選択メールを指定フォルダへ直接移動"""
        selected = self.tree.selection()
        if not selected:
            return
        
        # プロモ移動の場合
        if provider == "__promo__":
            conn = sqlite3.connect(DB_FILE)
            cur = conn.cursor()
            
            # 送信者情報取得
            msg_ids = self.get_msgids_from_selection(selected)
            placeholders = ','.join(['?'] * len(msg_ids))
            cur.execute(f"SELECT DISTINCT sender FROM emails WHERE message_id IN ({placeholders})", msg_ids)
            senders_raw = [row[0] for row in cur.fetchall()]
            
            # 送信者をクリーン化してドメイン抽出
            senders_clean = []
            for sender in senders_raw:
                clean = self.fetcher.clean_address(sender)
                if "@" in clean:
                    domain = clean.split("@")[-1]
                    senders_clean.append((sender, domain))
            
            # ルール作成確認ダイアログ
            dialog = tk.Toplevel(self.root)
            dialog.title("プロモ・ボックスに移動")
            if folder:
                dialog_title = f"「{folder}」フォルダに {len(msg_ids)}件のメールを移動します"
            else:
                dialog_title = f"プロモ・ボックスに {len(msg_ids)}件のメールを移動します"
            dialog.geometry("550x400")
            dialog.transient(self.root)
            dialog.grab_set()
            
            tk.Label(dialog, text=dialog_title, font=("Arial", 12, "bold")).pack(pady=10)
            
            var_learn = tk.BooleanVar(value=True)
            chk = tk.Checkbutton(dialog, 
                                text="今後これらの送信者からのメールを自動的にこのフォルダに振り分ける", 
                                variable=var_learn, font=("Arial", 10))
            chk.pack(pady=10)
            
            tk.Label(dialog, text="検出された送信者:", font=("Arial", 10, "bold")).pack(anchor=tk.W, padx=20)
            
            sender_frame = tk.Frame(dialog)
            sender_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
            
            sender_text = tk.Text(sender_frame, height=8, width=60, wrap=tk.WORD)
            sender_scroll = tk.Scrollbar(sender_frame, command=sender_text.yview)
            sender_text.config(yscrollcommand=sender_scroll.set)
            
            sender_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            sender_scroll.pack(side=tk.RIGHT, fill=tk.Y)
            
            folder_display = f" → {folder}" if folder else " → プロモ・ボックス"
            for orig_sender, domain in senders_clean:
                sender_text.insert(tk.END, f"• {orig_sender}\n  ルール: %@{domain}%{folder_display}\n\n")
            sender_text.config(state=tk.DISABLED)
            
            result = {"confirmed": False, "learn": False}
            
            def on_confirm():
                result["confirmed"] = True
                result["learn"] = var_learn.get()
                dialog.destroy()
            
            def on_cancel():
                dialog.destroy()
            
            btn_frame = tk.Frame(dialog)
            btn_frame.pack(pady=15)
            
            tk.Button(btn_frame, text="移動＆学習", command=on_confirm, 
                     bg="#4CAF50", fg="white", width=15, font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
            tk.Button(btn_frame, text="キャンセル", command=on_cancel, 
                     bg="#f44336", fg="white", width=15).pack(side=tk.LEFT, padx=5)
            
            # ダイアログを中央に配置
            dialog.update_idletasks()
            x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
            y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
            dialog.geometry(f"+{x}+{y}")
            
            dialog.wait_window()
            
            if result["confirmed"]:
                # メールをフォルダに移動
                for msg_id in msg_ids:
                    cur.execute("UPDATE emails SET is_promo=1, folder=? WHERE message_id=?", (folder, msg_id))
                
                # 学習ルール追加（フォルダ情報も含む）
                if result["learn"]:
                    for sender, domain in senders_clean:
                        pattern = f"%@{domain}%"
                        try:
                            cur.execute("INSERT INTO promo_rules (sender_pattern, added_date, match_count, target_folder) VALUES (?, datetime('now'), 0, ?)", 
                                       (pattern, folder))
                        except sqlite3.IntegrityError:
                            # 既に存在する場合はフォルダ情報を更新
                            cur.execute("UPDATE promo_rules SET target_folder=? WHERE sender_pattern=?", (folder, pattern))
                
                conn.commit()
                conn.close()
                
                self.refresh_tree_from_db()
                self.refresh_folder_tree()
                
                learn_msg = "＆学習ルール追加" if result["learn"] else ""
                messagebox.showinfo("完了", f"{len(msg_ids)}件をフォルダに移動しました{learn_msg}")
                self.update_promo_button_state()
            else:
                conn.close()
        else:
            # 通常プロバイダ移動
            for msg_id in msg_ids:
                self.db_mgr.move_to_folder(msg_id, folder)
            
            messagebox.showinfo("完了", f"{len(selected)}件のメールを移動しました")
            self.refresh_tree_from_db()
            self.refresh_folder_tree()
    
    def empty_trash(self, provider):
        """ゴミ箱を空にする"""
        # ゴミ箱内のメール数を取得
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        
        if provider == "__promo__":
            cur.execute("SELECT COUNT(*) FROM emails WHERE is_promo=1 AND folder='__trash__'")
        else:
            cur.execute("SELECT COUNT(*) FROM emails WHERE provider=? AND folder='__trash__'", (provider,))
        
        count = cur.fetchone()[0]
        conn.close()
        
        if count == 0:
            messagebox.showinfo("情報", "ゴミ箱は空です")
            return
        
        # 確認ダイアログ
        delete_mode = self.config_mgr.get("delete_mode") or "local_only"
        mode_text = "Mail Hubのみ" if delete_mode == "local_only" else "Gmail Cloudからも削除"
        
        result = messagebox.askyesno(
            "確認",
            f"ゴミ箱内の{count}件のメールを完全削除します。\n\n"
            f"削除モード: {mode_text}\n\n"
            "この操作は取り消せません。よろしいですか？"
        )
        
        if not result:
            return
        
        # ゴミ箱内のメールIDを取得
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        
        if provider == "__promo__":
            cur.execute("SELECT message_id FROM emails WHERE is_promo=1 AND folder='__trash__'")
        else:
            cur.execute("SELECT message_id FROM emails WHERE provider=? AND folder='__trash__'", (provider,))
        
        message_ids = [row[0] for row in cur.fetchall()]
        conn.close()
        
        # 完全削除処理
        for msg_id in message_ids:
            self.db_mgr.permanently_delete_email(msg_id, delete_mode, self.config_mgr)
        
        messagebox.showinfo("完了", f"{count}件のメールを完全削除しました")
        self.refresh_tree_from_db()
        self.refresh_folder_tree()
    
    def restore_from_trash(self, provider):
        """ゴミ箱から元に戻す"""
        # ゴミ箱内のメール数を取得
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        
        if provider == "__promo__":
            cur.execute("SELECT COUNT(*) FROM emails WHERE is_promo=1 AND folder='__trash__'")
        else:
            cur.execute("SELECT COUNT(*) FROM emails WHERE provider=? AND folder='__trash__'", (provider,))
        
        count = cur.fetchone()[0]
        
        if count == 0:
            conn.close()
            messagebox.showinfo("情報", "ゴミ箱は空です")
            return
        
        # 確認ダイアログ
        result = messagebox.askyesno(
            "確認",
            f"ゴミ箱内の{count}件のメールを元に戻します。\n\n"
            "よろしいですか？"
        )
        
        if not result:
            conn.close()
            return
        
        # ゴミ箱内のメールを元に戻す（folder=NULL）
        if provider == "__promo__":
            cur.execute("UPDATE emails SET folder=NULL WHERE is_promo=1 AND folder='__trash__'")
        else:
            cur.execute("UPDATE emails SET folder=NULL WHERE provider=? AND folder='__trash__'", (provider,))
        
        conn.commit()
        conn.close()
        
        messagebox.showinfo("完了", f"{count}件のメールを元に戻しました")
        self.refresh_tree_from_db()
        self.refresh_folder_tree()
    
    def create_custom_folder(self, provider):
        """カスタムフォルダ作成"""
        folder_name = tk.simpledialog.askstring("新規フォルダ", f"{provider} に作成するフォルダ名:")
        if not folder_name:
            return
        
        if self.db_mgr.create_folder(provider, folder_name, 'custom'):
            messagebox.showinfo("成功", f"フォルダ「{folder_name}」を作成しました")
            self.refresh_folder_tree()
        else:
            messagebox.showerror("エラー", "同名のフォルダが既に存在します")
    
    def delete_custom_folder(self, provider, folder_name):
        """カスタムフォルダ削除"""
        result = messagebox.askyesno("確認", f"フォルダ「{folder_name}」を削除しますか？\n\n※ フォルダ内のメールは親フォルダに移動されます\n※ このフォルダへの自動振り分けルールも削除されます")
        if result:
            conn = sqlite3.connect(DB_FILE)
            cur = conn.cursor()
            
            # メールを親フォルダに移動
            if provider == "__promo__":
                cur.execute("UPDATE emails SET folder=NULL WHERE is_promo=1 AND folder=?", (folder_name,))
            else:
                cur.execute("UPDATE emails SET folder=NULL WHERE provider=? AND folder=?", (provider, folder_name))
            
            # プロモルールのフォルダ指定を解除
            if provider == "__promo__":
                cur.execute("UPDATE promo_rules SET target_folder=NULL WHERE target_folder=?", (folder_name,))
            
            # フォルダ定義削除（同じconnection内で実行）
            cur.execute("DELETE FROM folders WHERE provider=? AND folder_name=?", (provider, folder_name))
            
            conn.commit()
            conn.close()
            
            messagebox.showinfo("完了", f"フォルダを削除しました\n\nメールは親フォルダに移動されました")
            self.refresh_folder_tree()
            self.refresh_tree_from_db()
    
    def rename_folder(self, provider, old_name):
        """フォルダ名変更"""
        new_name = tk.simpledialog.askstring("名前変更", f"新しいフォルダ名:", initialvalue=old_name)
        if not new_name or new_name == old_name:
            return
        
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("UPDATE folders SET folder_name=? WHERE provider=? AND folder_name=?", (new_name, provider, old_name))
        cur.execute("UPDATE emails SET folder=? WHERE provider=? AND folder=?", (new_name, provider, old_name))
        conn.commit()
        conn.close()
        
        messagebox.showinfo("完了", "フォルダ名を変更しました")
        self.refresh_folder_tree()
    
    def show_move_to_folder_menu(self):
        """フォルダへ移動サブメニュー"""
        sel = self.tree.selection()
        if not sel:
            return
        
        item = sel[0]  # これはTreeview IID
        msg_id = self.get_msgid_from_selection(sel)  # 元のMessage-ID取得
        values = self.tree.item(item, "values")
        if not values:
            return
        
        if not msg_id:
            return
        
        # プロバイダ取得
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("SELECT provider FROM emails WHERE message_id=?", (msg_id,))
        row = cur.fetchone()
        conn.close()
        
        if not row:
            return
        
        provider = row[0]
        
        # フォルダ選択ダイアログ
        dialog = tk.Toplevel(self.root)
        dialog.title("フォルダへ移動")
        dialog.geometry("300x400")
        
        tk.Label(dialog, text="移動先フォルダを選択:", font=("Arial", 10, "bold")).pack(pady=10)
        
        listbox = tk.Listbox(dialog, font=("Arial", 10))
        listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # システムフォルダ
        listbox.insert(tk.END, "📤 送信済み")
        listbox.insert(tk.END, "📝 下書き")
        listbox.insert(tk.END, "🗑️ ゴミ箱")
        
        folder_map = {
            "📤 送信済み": "__sent__",
            "📝 下書き": "__drafts__",
            "🗑️ ゴミ箱": "__trash__",
        }
        
        # カスタムフォルダ
        folders = self.db_mgr.get_folders(provider)
        for folder_name, folder_type in folders:
            if folder_type == 'custom':
                listbox.insert(tk.END, f"📂 {folder_name}")
                folder_map[f"📂 {folder_name}"] = folder_name
        
        def do_move():
            sel_idx = listbox.curselection()
            if not sel_idx:
                return
            
            folder_label = listbox.get(sel_idx[0])
            folder_key = folder_map.get(folder_label, folder_label.replace("📂 ", ""))
            
            self.db_mgr.move_to_folder(msg_id, folder_key)
            messagebox.showinfo("完了", f"「{folder_label}」へ移動しました")
            dialog.destroy()
            self.refresh_tree_from_db()
        
        tk.Button(dialog, text="移動", command=do_move, bg="#2196F3", fg="white", width=15).pack(pady=10)
        tk.Button(dialog, text="キャンセル", command=dialog.destroy, width=15).pack(pady=5)
    
    def show_inbox(self):
        """受信箱表示"""
        self.view_config.pack_forget()
        self.view_inbox.pack(fill=tk.BOTH, expand=True)
        self.btn_inbox.config(relief=tk.SUNKEN, bg="white")
        self.btn_config.config(relief=tk.RAISED, bg="#eee")
    
    def show_config(self):
        """設定表示"""
        self.view_inbox.pack_forget()
        self.view_config.pack(fill=tk.BOTH, expand=True)
        self.btn_inbox.config(relief=tk.RAISED, bg="#eee")
        self.btn_config.config(relief=tk.SUNKEN, bg="white")
    
    def start_fetch_task(self):
        """メール受信開始"""
        if not self.config_mgr.get("password"):
            messagebox.showwarning("未設定", "設定タブでパスワードを設定してください")
            return
        
        self.btn_fetch.config(state=tk.DISABLED)
        self.lbl_status.config(text="受信中... (バックグラウンド)", fg="orange")
        self.root.update()
        
        thread = threading.Thread(target=self.run_fetch_logic)
        thread.start()
    
    def run_fetch_logic(self):
        """メール受信処理"""
        # プログレスバー表示
        self.root.after(0, self.show_progress)
        
        try:
            # 進捗コールバック関数
            def on_progress(current, total):
                self.root.after(0, self.update_progress, current, total)
            
            emails = self.fetcher.fetch_central(self.config_mgr, progress_callback=on_progress)
            new_count = self.db_mgr.save_emails(emails)
            msg = f"受信完了 (新規: {new_count}件)"
            is_error = False
        except Exception as e:
            msg = f"エラー: {e}"
            is_error = True
        
        # プログレスバー非表示
        self.root.after(0, self.hide_progress)
        self.root.after(0, self.on_fetch_complete, msg, is_error)
    
    def on_fetch_complete(self, msg, is_error):
        """受信完了処理"""
        self.lbl_status.config(text=msg, fg="red" if is_error else "green")
        self.btn_fetch.config(state=tk.NORMAL)
        
        if is_error:
            messagebox.showerror("受信エラー", msg)
        else:
            self.refresh_tree_from_db()
            self.refresh_folder_tree()
    
    def save_central_config(self):
        """集約アドレス設定保存（取得範囲・自動取得含む）"""
        email_val = self.ent_cen_email.get().strip()
        pass_val = self.ent_cen_pass.get().strip()
        
        if not email_val or not pass_val:
            messagebox.showwarning("不足", "メールアドレスとパスワードを入力してください")
            return
        
        try:
            self.root.config(cursor="watch")
            self.root.update()
            
            self.fetcher.test_connection_imap(self.config_mgr.get("imap_server"), email_val, pass_val)
            
            self.config_mgr.set("email", email_val)
            self.config_mgr.set("password", pass_val)
            
            # 取得範囲設定を保存（fetch_modeをfetch_rangeに変換）
            fetch_mode = self.fetch_mode_var.get()
            
            # fetch_mode → fetch_range変換
            fetch_range_map = {
                "latest_only": "latest",
                "last_week": "week",
                "last_month": "month",
                "last_3months": "3months",
                "last_year": "year",
                "all": "all",
                "custom": "custom"
            }
            
            fetch_range = fetch_range_map.get(fetch_mode, "week")
            
            self.config_mgr.set("fetch_mode", fetch_mode)
            self.config_mgr.set("fetch_range", fetch_range)  # fetch_rangeも保存
            self.config_mgr.set("custom_days", self.custom_days_var.get())
            
            # 自動取得設定を保存
            self.config_mgr.set("auto_fetch_on_startup", self.startup_var.get())
            self.config_mgr.set("auto_fetch_interval", self.interval_var.get())
            
            # 定期取得間隔を保存
            interval_text = self.interval_combo.get()
            if "15分" in interval_text:
                minutes = 15
            elif "30分" in interval_text:
                minutes = 30
            elif "1時間" in interval_text:
                minutes = 60
            elif "2時間" in interval_text:
                minutes = 120
            else:
                minutes = 30
            self.config_mgr.set("auto_fetch_interval_minutes", minutes)
            
            providers = self.config_mgr.get("providers") or []
            new_providers = [p for p in providers if p["email"] != email_val]
            new_providers.insert(0, {
                "email": email_val,
                "smtp_host": "smtp.gmail.com",
                "smtp_port": "587",
                "password": pass_val,
                "fallback_gmail": False
            })
            self.config_mgr.set("providers", new_providers)
            
            # SMTP送信用アカウント設定も保存
            smtp_accounts = self.config_mgr.get("smtp_accounts") or []
            new_smtp = [acc for acc in smtp_accounts if acc["email"] != email_val]
            new_smtp.append({
                "email": email_val,
                "password": pass_val,
                "smtp_server": "smtp.gmail.com",
                "smtp_port": 465  # SSL用
            })
            self.config_mgr.set("smtp_accounts", new_smtp)
            self.config_mgr.set("providers", new_providers)
            self.config_mgr.save()
            
            self.refresh_provider_list()
            self.root.config(cursor="")
            messagebox.showinfo("成功", "設定を保存しました")
        except Exception as e:
            self.root.config(cursor="")
            messagebox.showerror("失敗", str(e))
    
    def add_provider(self):
        """プロバイダ追加"""
        p_email = self.ent_prov_email.get().strip()
        p_host = self.ent_prov_host.get().strip()
        p_port = self.ent_prov_port.get().strip()
        p_pass = self.ent_prov_pass.get().strip()
        fallback = self.var_fallback.get()
        
        if not all([p_email, p_host, p_port, p_pass]):
            messagebox.showwarning("不足", "全ての項目を入力してください")
            return
        
        clean_email = self.fetcher.clean_address(p_email)
        
        try:
            self.root.config(cursor="watch")
            self.root.update()
            self.fetcher.test_connection_smtp(p_host, p_port, clean_email, p_pass)
            self.root.config(cursor="")
        except Exception as e:
            self.root.config(cursor="")
            if not messagebox.askyesno("接続失敗", f"SMTP接続に失敗しました:\n{e}\n\nそれでも設定を保存しますか？"):
                return
        
        provs = self.config_mgr.get("providers") or []
        provs = [p for p in provs if p["email"] != clean_email]
        provs.append({
            "email": clean_email,
            "smtp_host": p_host,
            "smtp_port": p_port,
            "password": p_pass,
            "fallback_gmail": fallback
        })
        self.config_mgr.set("providers", provs)
        
        # SMTP送信用アカウント設定も保存（ポート番号に応じて）
        smtp_accounts = self.config_mgr.get("smtp_accounts") or []
        new_smtp = [acc for acc in smtp_accounts if acc["email"] != clean_email]
        new_smtp.append({
            "email": clean_email,
            "password": p_pass,
            "smtp_server": p_host,
            "smtp_port": int(p_port)  # ポート番号をそのまま使用
        })
        self.config_mgr.set("smtp_accounts", new_smtp)
        
        self.config_mgr.save()
        
        self.refresh_provider_list()
        messagebox.showinfo("成功", f"{clean_email} を登録しました")
    
    def refresh_provider_list(self):
        """プロバイダリスト更新"""
        self.tree_prov.delete(*self.tree_prov.get_children())
        
        provs = self.config_mgr.get("providers") or []
        for p in provs:
            fb = "あり" if p.get("fallback_gmail") else "-"
            self.tree_prov.insert("", tk.END, values=(p["email"], p["smtp_host"], p["smtp_port"], fb))
    
    def delete_provider(self):
        """プロバイダ削除"""
        sel = self.tree_prov.selection()
        if not sel:
            return
        
        tgt = self.tree_prov.item(sel, "values")[0]
        
        if tgt == self.config_mgr.get("email"):
            messagebox.showwarning("不可", "集約アドレスは削除できません")
            return
        
        if messagebox.askyesno("削除", f"{tgt} を削除しますか？"):
            provs = [p for p in self.config_mgr.get("providers") if p["email"] != tgt]
            self.config_mgr.set("providers", provs)
            self.config_mgr.save()
            self.refresh_provider_list()
    
    def open_compose_window(self):
        """新規メール作成ウィンドウ"""
        # SMTPアカウント取得
        smtp_accounts = self.config_mgr.get("smtp_accounts") or []
        
        # smtp_accountsがない場合、providersから生成
        if not smtp_accounts:
            providers = self.config_mgr.get("providers") or []
            for p in providers:
                smtp_accounts.append({
                    "email": p.get("email", ""),
                    "password": p.get("password", ""),
                    "smtp_server": p.get("smtp_host", "smtp.gmail.com"),
                    "smtp_port": 465 if "gmail" in p.get("smtp_host", "") else int(p.get("smtp_port", 587))
                })
            
            # 保存
            if smtp_accounts:
                self.config_mgr.set("smtp_accounts", smtp_accounts)
                self.config_mgr.save()
        
        if not smtp_accounts:
            messagebox.showerror("エラー", "SMTP設定が登録されていません。\n設定画面で送信アカウントを登録してください。")
            return
        
        win = tk.Toplevel(self.root)
        win.title("✉️ 新規メール作成")
        win.geometry("800x600")
        
        # 送信元選択
        tk.Label(win, text="送信元:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        
        from_var = tk.StringVar()
        from_combo = ttk.Combobox(win, textvariable=from_var, width=50, state="readonly")
        from_combo['values'] = [acc['email'] for acc in smtp_accounts]
        from_combo.current(0)
        from_combo.grid(row=0, column=1, sticky="w", padx=10, pady=5)
        
        # 宛先
        tk.Label(win, text="宛先:", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky="w", padx=10, pady=5)
        to_entry = tk.Entry(win, width=50)
        to_entry.grid(row=1, column=1, sticky="w", padx=10, pady=5)
        
        # 件名
        tk.Label(win, text="件名:", font=("Arial", 10, "bold")).grid(row=2, column=0, sticky="w", padx=10, pady=5)
        subject_entry = tk.Entry(win, width=50)
        subject_entry.grid(row=2, column=1, sticky="w", padx=10, pady=5)
        
        # 本文
        tk.Label(win, text="本文:", font=("Arial", 10, "bold")).grid(row=3, column=0, sticky="nw", padx=10, pady=5)
        body_text = tk.Text(win, width=60, height=15, wrap=tk.WORD)
        body_text.grid(row=3, column=1, padx=10, pady=5)
        
        # 添付ファイル
        tk.Label(win, text="添付:", font=("Arial", 10, "bold")).grid(row=4, column=0, sticky="nw", padx=10, pady=5)
        
        attachment_frame = tk.Frame(win)
        attachment_frame.grid(row=4, column=1, sticky="w", padx=10, pady=5)
        
        attachments = []  # 添付ファイルパスのリスト
        
        attachment_listbox = tk.Listbox(attachment_frame, width=50, height=4)
        attachment_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        def add_attachment():
            """添付ファイル追加"""
            from tkinter import filedialog
            files = filedialog.askopenfilenames(title="添付ファイルを選択")
            for file_path in files:
                if file_path not in attachments:
                    attachments.append(file_path)
                    attachment_listbox.insert(tk.END, os.path.basename(file_path))
        
        def remove_attachment():
            """添付ファイル削除"""
            selection = attachment_listbox.curselection()
            if selection:
                index = selection[0]
                attachment_listbox.delete(index)
                attachments.pop(index)
        
        btn_attachment_frame = tk.Frame(attachment_frame)
        btn_attachment_frame.pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_attachment_frame, text="📎 追加", command=add_attachment, width=8).pack(pady=2)
        tk.Button(btn_attachment_frame, text="❌ 削除", command=remove_attachment, width=8).pack(pady=2)
        
        # 送信ボタン
        def do_send():
            to = to_entry.get().strip()
            subject = subject_entry.get().strip()
            body = body_text.get("1.0", tk.END).strip()
            from_email = from_var.get()
            
            if not to or not subject:
                messagebox.showerror("エラー", "宛先と件名は必須です")
                return
            
            # SMTP設定取得
            smtp_account = next((acc for acc in smtp_accounts if acc['email'] == from_email), None)
            
            if not smtp_account:
                messagebox.showerror("エラー", "送信アカウント情報が見つかりません")
                return
            
            try:
                # SMTP送信（添付ファイル対応）
                import smtplib
                from email.mime.text import MIMEText
                from email.mime.multipart import MIMEMultipart
                from email.mime.base import MIMEBase
                from email import encoders
                
                if attachments:
                    # 添付ファイルあり
                    msg = MIMEMultipart()
                    msg["Subject"] = subject
                    msg["From"] = from_email
                    msg["To"] = to
                    
                    # 本文
                    msg.attach(MIMEText(body, "plain", "utf-8"))
                    
                    # 添付ファイル
                    for file_path in attachments:
                        try:
                            with open(file_path, "rb") as f:
                                part = MIMEBase("application", "octet-stream")
                                part.set_payload(f.read())
                                encoders.encode_base64(part)
                                part.add_header(
                                    "Content-Disposition",
                                    f"attachment; filename= {os.path.basename(file_path)}",
                                )
                                msg.attach(part)
                        except Exception as e:
                            messagebox.showerror("添付エラー", f"ファイル添付に失敗しました:\n{file_path}\n{e}")
                            return
                else:
                    # 添付ファイルなし
                    msg = MIMEText(body, "plain", "utf-8")
                    msg["Subject"] = subject
                    msg["From"] = from_email
                    msg["To"] = to
                
                with smtplib.SMTP_SSL(smtp_account['smtp_server'], smtp_account.get('smtp_port', 465)) as server:
                    server.login(smtp_account['email'], smtp_account['password'])
                    server.send_message(msg)
                
                messagebox.showinfo("成功", "メールを送信しました")
                win.destroy()
                
            except Exception as e:
                messagebox.showerror("送信失敗", f"メール送信に失敗しました:\n{e}")
        
        btn_frame = tk.Frame(win)
        btn_frame.grid(row=5, column=1, pady=10)
        
        # 下書き保存機能
        def save_draft():
            to = to_entry.get().strip()
            subject = subject_entry.get().strip()
            body = body_text.get("1.0", tk.END).strip()
            from_email = from_var.get()
            
            if not subject:
                messagebox.showerror("エラー", "件名は必須です")
                return
            
            # 下書きとして保存
            from datetime import datetime
            import uuid
            
            draft_data = {
                "message_id": str(uuid.uuid4()),
                "original_to": to,
                "subject": subject,
                "sender": from_email,
                "date_disp": datetime.now().strftime("%Y/%m/%d %H:%M:%S"),
                "timestamp": datetime.now().isoformat(),
                "raw_data": body,
                "provider": from_email.split("@")[-1] if "@" in from_email else "unknown",
                "folder": "__drafts__"
            }
            
            conn = sqlite3.connect(DB_FILE)
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO emails (message_id, original_to, subject, sender, date_disp, timestamp, raw_data, provider, folder)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                draft_data["message_id"],
                draft_data["original_to"],
                draft_data["subject"],
                draft_data["sender"],
                draft_data["date_disp"],
                draft_data["timestamp"],
                draft_data["raw_data"],
                draft_data["provider"],
                draft_data["folder"]
            ))
            conn.commit()
            conn.close()
            
            self.refresh_tree_from_db()
            self.refresh_folder_tree()
            
            messagebox.showinfo("保存完了", "下書きを保存しました")
            win.destroy()
        
        tk.Button(btn_frame, text="📤 送信", command=do_send, bg="#4CAF50", fg="white", width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="📝 下書き保存", command=save_draft, bg="#FF9800", fg="white", width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="❌ キャンセル", command=win.destroy, bg="#f44336", fg="white", width=15).pack(side=tk.LEFT, padx=5)
    
    def open_forward_window(self):
        """転送ウィンドウ起動"""
        sel = self.tree.selection()
        if not sel:
            return
        
        # 安全なIIDから元のMessage-IDを取得
        msg_id = self.get_msgid_from_selection(sel)
        
        # 元メールの情報を取得
        item = self.tree.item(sel, "values")
        orig_to = item[0]
        orig_subject = item[1]
        orig_sender = item[2]
        target_account = self.fetcher.clean_address(orig_to.split()[0])
        
        # メール本文を取得
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("SELECT raw_data FROM emails WHERE message_id=?", (sel[0],))
        row = cur.fetchone()
        conn.close()
        
        orig_body = row[0] if row else ""
        
        # SMTPアカウント取得
        smtp_accounts = self.config_mgr.get("smtp_accounts") or []
        
        # smtp_accountsがない場合、providersから生成
        if not smtp_accounts:
            providers = self.config_mgr.get("providers") or []
            for p in providers:
                smtp_accounts.append({
                    "email": p.get("email", ""),
                    "password": p.get("password", ""),
                    "smtp_server": p.get("smtp_host", "smtp.gmail.com"),
                    "smtp_port": 465 if "gmail" in p.get("smtp_host", "") else int(p.get("smtp_port", 587))
                })
            
            # 保存
            if smtp_accounts:
                self.config_mgr.set("smtp_accounts", smtp_accounts)
                self.config_mgr.save()
        
        if not smtp_accounts:
            messagebox.showerror("エラー", "SMTP設定が登録されていません。\n設定画面で送信アカウントを登録してください。")
            return
        
        win = tk.Toplevel(self.root)
        win.title(f"↪️ 転送: {target_account}")
        win.geometry("800x600")
        
        # 送信元選択
        tk.Label(win, text="送信元:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        
        from_var = tk.StringVar()
        from_combo = ttk.Combobox(win, textvariable=from_var, width=50, state="readonly")
        from_combo['values'] = [acc['email'] for acc in smtp_accounts]
        
        # デフォルトで現在のアカウントを選択
        default_index = 0
        for i, acc in enumerate(smtp_accounts):
            if acc['email'] == target_account:
                default_index = i
                break
        from_combo.current(default_index)
        from_combo.grid(row=0, column=1, sticky="w", padx=10, pady=5)
        
        # 宛先
        tk.Label(win, text="宛先:", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky="w", padx=10, pady=5)
        to_entry = tk.Entry(win, width=50)
        to_entry.grid(row=1, column=1, sticky="w", padx=10, pady=5)
        
        # 件名
        tk.Label(win, text="件名:", font=("Arial", 10, "bold")).grid(row=2, column=0, sticky="w", padx=10, pady=5)
        subject_entry = tk.Entry(win, width=50)
        subject_entry.insert(0, f"Fwd: {orig_subject}")
        subject_entry.grid(row=2, column=1, sticky="w", padx=10, pady=5)
        
        # 本文（元メールを引用）
        tk.Label(win, text="本文:", font=("Arial", 10, "bold")).grid(row=3, column=0, sticky="nw", padx=10, pady=5)
        body_text = tk.Text(win, width=60, height=15, wrap=tk.WORD)
        
        # 転送メッセージを作成
        forward_body = f"""

---------- Forwarded message ---------
From: {orig_sender}
Subject: {orig_subject}

{orig_body}
"""
        body_text.insert("1.0", forward_body)
        body_text.grid(row=3, column=1, padx=10, pady=5)
        
        # 添付ファイル
        tk.Label(win, text="添付:", font=("Arial", 10, "bold")).grid(row=4, column=0, sticky="nw", padx=10, pady=5)
        
        attachment_frame = tk.Frame(win)
        attachment_frame.grid(row=4, column=1, sticky="w", padx=10, pady=5)
        
        attachments = []
        
        attachment_listbox = tk.Listbox(attachment_frame, width=50, height=4)
        attachment_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        def add_attachment():
            from tkinter import filedialog
            files = filedialog.askopenfilenames(title="添付ファイルを選択")
            for file_path in files:
                if file_path not in attachments:
                    attachments.append(file_path)
                    attachment_listbox.insert(tk.END, os.path.basename(file_path))
        
        def remove_attachment():
            selection = attachment_listbox.curselection()
            if selection:
                index = selection[0]
                attachment_listbox.delete(index)
                attachments.pop(index)
        
        btn_attachment_frame = tk.Frame(attachment_frame)
        btn_attachment_frame.pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_attachment_frame, text="📎 追加", command=add_attachment, width=8).pack(pady=2)
        tk.Button(btn_attachment_frame, text="❌ 削除", command=remove_attachment, width=8).pack(pady=2)
        
        # 送信ボタン
        def do_send():
            to = to_entry.get().strip()
            subject = subject_entry.get().strip()
            body = body_text.get("1.0", tk.END).strip()
            from_email = from_var.get()
            
            if not to or not subject:
                messagebox.showerror("エラー", "宛先と件名は必須です")
                return
            
            # SMTP設定取得
            smtp_account = next((acc for acc in smtp_accounts if acc['email'] == from_email), None)
            
            if not smtp_account:
                messagebox.showerror("エラー", "送信アカウント情報が見つかりません")
                return
            
            try:
                # SMTP送信（添付ファイル対応）
                import smtplib
                from email.mime.text import MIMEText
                from email.mime.multipart import MIMEMultipart
                from email.mime.base import MIMEBase
                from email import encoders
                
                if attachments:
                    msg = MIMEMultipart()
                    msg["Subject"] = subject
                    msg["From"] = from_email
                    msg["To"] = to
                    
                    msg.attach(MIMEText(body, "plain", "utf-8"))
                    
                    for file_path in attachments:
                        try:
                            with open(file_path, "rb") as f:
                                part = MIMEBase("application", "octet-stream")
                                part.set_payload(f.read())
                                encoders.encode_base64(part)
                                part.add_header(
                                    "Content-Disposition",
                                    f"attachment; filename= {os.path.basename(file_path)}",
                                )
                                msg.attach(part)
                        except Exception as e:
                            messagebox.showerror("添付エラー", f"ファイル添付に失敗しました:\n{file_path}\n{e}")
                            return
                else:
                    msg = MIMEText(body, "plain", "utf-8")
                    msg["Subject"] = subject
                    msg["From"] = from_email
                    msg["To"] = to
                
                with smtplib.SMTP_SSL(smtp_account['smtp_server'], smtp_account.get('smtp_port', 465)) as server:
                    server.login(smtp_account['email'], smtp_account['password'])
                    server.send_message(msg)
                
                # 送信済みとして保存
                from datetime import datetime
                import uuid
                
                sent_email_data = {
                    "message_id": str(uuid.uuid4()),
                    "original_to": to,
                    "subject": subject,
                    "sender": from_email,
                    "date_disp": datetime.now().strftime("%Y/%m/%d %H:%M:%S"),
                    "timestamp": datetime.now().isoformat(),
                    "raw_data": body,
                    "provider": from_email.split("@")[-1] if "@" in from_email else "unknown"
                }
                
                self.db_mgr.save_sent_email(sent_email_data)
                
                # 画面更新
                self.refresh_tree_from_db()
                self.refresh_folder_tree()
                
                messagebox.showinfo("成功", "メールを転送しました")
                win.destroy()
                
            except Exception as e:
                messagebox.showerror("送信失敗", f"メール転送に失敗しました:\n{e}")
        
        def save_draft_forward():
            """転送下書き保存"""
            to = to_entry.get().strip()
            subject = subject_entry.get().strip()
            body = body_text.get("1.0", tk.END).strip()
            from_email = from_var.get()
            
            if not subject:
                messagebox.showerror("エラー", "件名は必須です")
                return
            
            from datetime import datetime
            import uuid
            
            draft_data = {
                "message_id": str(uuid.uuid4()),
                "original_to": to,
                "subject": subject,
                "sender": from_email,
                "date_disp": datetime.now().strftime("%Y/%m/%d %H:%M:%S"),
                "timestamp": datetime.now().isoformat(),
                "raw_data": body,
                "provider": from_email.split("@")[-1] if "@" in from_email else "unknown",
                "folder": "__drafts__"
            }
            
            conn = sqlite3.connect(DB_FILE)
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO emails (message_id, original_to, subject, sender, date_disp, timestamp, raw_data, provider, folder)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                draft_data["message_id"],
                draft_data["original_to"],
                draft_data["subject"],
                draft_data["sender"],
                draft_data["date_disp"],
                draft_data["timestamp"],
                draft_data["raw_data"],
                draft_data["provider"],
                draft_data["folder"]
            ))
            conn.commit()
            conn.close()
            
            self.refresh_tree_from_db()
            self.refresh_folder_tree()
            
            messagebox.showinfo("保存完了", "下書きを保存しました")
            win.destroy()
        
        btn_frame = tk.Frame(win)
        btn_frame.grid(row=5, column=1, pady=10)
        
        tk.Button(btn_frame, text="📤 転送", command=do_send, bg="#4CAF50", fg="white", width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="📝 下書き保存", command=save_draft_forward, bg="#FF9800", fg="white", width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="❌ キャンセル", command=win.destroy, bg="#f44336", fg="white", width=15).pack(side=tk.LEFT, padx=5)
    
    def open_reply_window(self):
        """返信ウィンドウ起動"""
        sel = self.tree.selection()
        if not sel:
            return
        
        # 安全なIIDから元のMessage-IDを取得
        msg_id = self.get_msgid_from_selection(sel)
        
        item = self.tree.item(sel, "values")
        orig_to = item[0]
        orig_sender = item[2]
        target_account = self.fetcher.clean_address(orig_to.split()[0])
        
        win = tk.Toplevel(self.root)
        win.title(f"返信: {target_account} として")
        win.geometry("600x500")
        
        tk.Label(win, text="宛先:").pack(anchor=tk.W, padx=10)
        ent_to = tk.Entry(win, width=50)
        ent_to.pack(fill=tk.X, padx=10)
        ent_to.insert(0, orig_sender)
        
        tk.Label(win, text="件名:").pack(anchor=tk.W, padx=10)
        ent_sub = tk.Entry(win, width=50)
        ent_sub.pack(fill=tk.X, padx=10)
        ent_sub.insert(0, f"Re: {item[1]}")
        
        tk.Label(win, text="本文:").pack(anchor=tk.W, padx=10)
        txt_body = tk.Text(win)
        txt_body.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        providers = self.config_mgr.get("providers") or []
        my_conf = next((p for p in providers if p["email"] == target_account), None)
        
        lbl_info = tk.Label(win, text="", fg="blue")
        lbl_info.pack(pady=5)
        
        # Microsoftドメインのチェック
        microsoft_domains = ["live.jp", "outlook.jp", "outlook.com", "hotmail.co.jp", "hotmail.com", "msn.com"]
        is_microsoft = any(target_account.endswith("@" + domain) for domain in microsoft_domains)
        
        if my_conf:
            if is_microsoft:
                # Microsoftアカウントは外部SMTP拒否のため、Gmail経由で送信
                lbl_info.config(text=f"【情報】{target_account} として Gmail経由で送信します。", fg="orange")
            else:
                # その他のプロバイダはSMTPサーバーで送信
                lbl_info.config(text=f"【情報】{target_account} のSMTPサーバーで送信します。")
        else:
            lbl_info.config(text=f"【警告】設定が見つかりません。デフォルト(Gmail)から送信されます。", fg="red")
            my_conf = {"email": target_account, "fallback_gmail": True}
        
        def do_send():
            body = txt_body.get("1.0", tk.END)
            to = ent_to.get()
            sub = ent_sub.get()
            try:
                success, msg = self.fetcher.send_email(my_conf, self.config_mgr.config, to, sub, body)
                
                if success:
                    # 送信成功時の処理
                    
                    # 1. 元メールに返信済みフラグ
                    original_msg_id = msg_id  # 元のMessage-ID
                    self.db_mgr.mark_as_replied(original_msg_id)
                    
                    # 2. 送信メールをDBに保存
                    from datetime import datetime
                    import uuid
                    
                    sent_email_data = {
                        "message_id": str(uuid.uuid4()),  # 一意なID生成
                        "original_to": to,
                        "subject": sub,
                        "sender": target_account,
                        "date_disp": datetime.now().strftime("%Y/%m/%d %H:%M:%S"),
                        "timestamp": datetime.now().isoformat(),
                        "raw_data": body,
                        "provider": target_account.split("@")[-1] if "@" in target_account else "unknown"
                    }
                    
                    self.db_mgr.save_sent_email(sent_email_data)
                    
                    # 3. 画面更新
                    self.refresh_tree_from_db()
                    self.refresh_folder_tree()
                
                messagebox.showinfo("完了", msg)
                win.destroy()
            except Exception as e:
                messagebox.showerror("エラー", f"送信失敗: {e}")
        
        def save_draft_reply():
            """返信下書き保存"""
            to = ent_to.get().strip()
            subject = ent_sub.get().strip()
            body = txt_body.get("1.0", tk.END).strip()
            
            if not subject:
                messagebox.showerror("エラー", "件名は必須です")
                return
            
            from datetime import datetime
            import uuid
            
            draft_data = {
                "message_id": str(uuid.uuid4()),
                "original_to": to,
                "subject": subject,
                "sender": target_account,
                "date_disp": datetime.now().strftime("%Y/%m/%d %H:%M:%S"),
                "timestamp": datetime.now().isoformat(),
                "raw_data": body,
                "provider": target_account.split("@")[-1] if "@" in target_account else "unknown",
                "folder": "__drafts__"
            }
            
            conn = sqlite3.connect(DB_FILE)
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO emails (message_id, original_to, subject, sender, date_disp, timestamp, raw_data, provider, folder)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                draft_data["message_id"],
                draft_data["original_to"],
                draft_data["subject"],
                draft_data["sender"],
                draft_data["date_disp"],
                draft_data["timestamp"],
                draft_data["raw_data"],
                draft_data["provider"],
                draft_data["folder"]
            ))
            conn.commit()
            conn.close()
            
            self.refresh_tree_from_db()
            self.refresh_folder_tree()
            
            messagebox.showinfo("保存完了", "下書きを保存しました")
            win.destroy()
        
        btn_frame = tk.Frame(win)
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="📤 送信", command=do_send, bg="#2196F3", fg="white", width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="📝 下書き保存", command=save_draft_reply, bg="#FF9800", fg="white", width=15).pack(side=tk.LEFT, padx=5)
    
    def move_to_promo(self):
        """選択メールをプロモに移動"""
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("未選択", "プロモに移動するメールを選択してください")
            return
        
        # プロモフォルダ内にいる場合は警告
        if self.current_promo_filter:
            messagebox.showinfo("既にプロモ", "このメールは既にプロモ・ボックス内にあります")
            return
        
        # 選択メール情報取得
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        
        msg_ids = self.get_msgids_from_selection(sel)
        placeholders = ','.join(['?'] * len(msg_ids))
        cur.execute(f"SELECT DISTINCT sender FROM emails WHERE message_id IN ({placeholders})", msg_ids)
        senders_raw = [row[0] for row in cur.fetchall()]
        
        # 送信者をクリーン化してドメイン抽出
        senders_clean = []
        for sender in senders_raw:
            clean = self.fetcher.clean_address(sender)
            if "@" in clean:
                domain = clean.split("@")[-1]
                senders_clean.append((sender, domain))
        
        # 確認ダイアログ
        dialog = tk.Toplevel(self.root)
        dialog.title("プロモに移動")
        dialog.geometry("550x350")
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text=f"{len(msg_ids)}件のメールをプロモに移動します", 
                font=("Arial", 12, "bold")).pack(pady=10)
        
        var_learn = tk.BooleanVar(value=True)
        chk = tk.Checkbutton(dialog, 
                            text="今後これらの送信者からのメールを自動的にプロモに振り分ける", 
                            variable=var_learn, font=("Arial", 10))
        chk.pack(pady=10)
        
        tk.Label(dialog, text="検出された送信者:", font=("Arial", 10, "bold")).pack(anchor=tk.W, padx=20)
        
        sender_frame = tk.Frame(dialog)
        sender_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
        
        sender_text = tk.Text(sender_frame, height=8, width=60, wrap=tk.WORD)
        sender_scroll = tk.Scrollbar(sender_frame, command=sender_text.yview)
        sender_text.config(yscrollcommand=sender_scroll.set)
        
        sender_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sender_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        for orig_sender, domain in senders_clean:
            sender_text.insert(tk.END, f"• {orig_sender}\n  → ルール: %@{domain}%\n\n")
        sender_text.config(state=tk.DISABLED)
        
        result = {"confirmed": False, "learn": False}
        
        def on_confirm():
            result["confirmed"] = True
            result["learn"] = var_learn.get()
            dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=15)
        
        tk.Button(btn_frame, text="移動＆学習", command=on_confirm, 
                 bg="#4CAF50", fg="white", width=15, font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="キャンセル", command=on_cancel, 
                 bg="#f44336", fg="white", width=15).pack(side=tk.LEFT, padx=5)
        
        # ダイアログを中央に配置
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        dialog.wait_window()
        
        if result["confirmed"]:
            # プロモフラグ設定
            cur.execute(f"UPDATE emails SET is_promo=1 WHERE message_id IN ({placeholders})", msg_ids)
            
            # 学習ルール追加
            if result["learn"]:
                for sender, domain in senders_clean:
                    pattern = f"%@{domain}%"
                    try:
                        cur.execute("INSERT INTO promo_rules (sender_pattern, added_date, match_count, target_folder) VALUES (?, datetime('now'), 0, NULL)", (pattern,))
                    except sqlite3.IntegrityError:
                        # 既に存在する場合はスキップ
                        pass
            
            conn.commit()
            conn.close()
            
            self.refresh_tree_from_db()
            self.refresh_folder_tree()
            
            learn_msg = "＆学習ルール追加" if result["learn"] else ""
            messagebox.showinfo("完了", f"{len(msg_ids)}件をプロモに移動しました{learn_msg}")
            self.update_promo_button_state()
        else:
            conn.close()

# ==========================================
# Main
# ==========================================
if __name__ == "__main__":
    root = tk.Tk()
    app = MailHubApp(root)
    root.mainloop()