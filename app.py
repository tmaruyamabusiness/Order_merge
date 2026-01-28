"""
Flask Web Application for 手配発注マージシステム
"""

from flask import Flask, render_template, request, jsonify, send_file, session
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime, timedelta, timezone
import pandas as pd
import os
import sys
import tempfile
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment ,Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import TextBlock, CellRichText
import hashlib
import json
from werkzeug.utils import secure_filename
import re
import qrcode
import io
from io import BytesIO
import base64
import threading
import time
import shutil
import pyodbc
from pathlib import Path
import pytz
from datetime import datetime, timedelta, timezone
import subprocess
import win32com.client as win32
from threading import Thread
import pythoncom
from flask_cors import CORS
from openpyxl.worksheet.page import PageMargins
from openpyxl.chart import BarChart, Reference
import glob
from PIL import Image
from utils import Constants, DataUtils, MekkiUtils, ExcelStyler, generate_qr_code, create_gantt_chart_sheet, EmailSender, DeliveryUtils        

app = Flask(__name__)

# Load configuration
# config.pyがある場合はそちらを使用、なければデフォルト設定
try:
    from config import get_config
    app.config.from_object(get_config())
except ImportError:
    # デフォルト設定
    app.config['SECRET_KEY'] = 'your-secret-key-here-change-in-production'
    app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///order_management.db'
    app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
    app.config['UPLOAD_FOLDER'] = 'uploads'
    app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size for large Excel files
    
    # Network path configuration
    app.config['DEFAULT_EXCEL_PATH'] = r'\\server3\Share-data\Document\仕入れ\002_手配リスト\手配発注_ALL.xlsx'
    app.config['HISTORY_EXCEL_PATH'] = r'\\server3\Share-data\Document\仕入れ\002_手配リスト\手配発注マージリスト発行履歴.xlsx'
    app.config['SEIBAN_LIST_PATH'] = r'\\server3\share-data\Document\Acrossデータ\製番一覧表.xlsx'
    app.config['EXPORT_EXCEL_PATH'] = r'\\SERVER3\Share-data\Document\仕入れ\002_手配リスト\手配発注リスト'
    app.config['AUTO_REFRESH_INTERVAL'] = 3600  # 1時間ごとに自動更新
    app.config['USE_ODBC'] = False  # ODBCを使用する場合はTrue
    app.config['ODBC_CONNECTION_STRING'] = ''  # ODBC接続文字列（必要に応じて設定）

db = SQLAlchemy(app)

# Ensure upload directory exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs('exports', exist_ok=True)
os.makedirs('cache', exist_ok=True)

# Global variables for background tasks
last_refresh_time = None
refresh_thread = None
cached_file_info = {}

# Global variables for background tasks
last_refresh_time = None
refresh_thread = None
cached_file_info = {}

# 🔥 発注リストの高速検索用キャッシュ
order_all_cache = {}
order_all_cache_time = None
CACHE_EXPIRY_SECONDS = 28800  # 8時間キャッシュ


# Database Models
class Order(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    seiban = db.Column(db.String(50), nullable=False)
    unit = db.Column(db.String(100))
    status = db.Column(db.String(50), default='受入準備前')
    location = db.Column(db.String(50), default='未定')
    remarks = db.Column(db.Text)
    product_name = db.Column(db.String(200))
    customer_abbr = db.Column(db.String(100))
    memo2 = db.Column(db.String(200))
    pallet_number = db.Column(db.String(50))  # ← 追加
    floor = db.Column(db.String(10))  # ← 追加
    image_path = db.Column(db.String(500))  # 画像パス
    is_archived = db.Column(db.Boolean, default=False)
    archived_at = db.Column(db.DateTime)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    updated_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc), onupdate=lambda: datetime.now(timezone.utc))
    
class OrderDetail(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    order_id = db.Column(db.Integer, db.ForeignKey('order.id'), nullable=False)
    delivery_date = db.Column(db.String(20))
    supplier = db.Column(db.String(100))
    supplier_cd = db.Column(db.String(50))
    order_number = db.Column(db.String(50))
    quantity = db.Column(db.Integer)
    unit_measure = db.Column(db.String(20))
    item_name = db.Column(db.String(200))
    spec1 = db.Column(db.String(200))
    spec2 = db.Column(db.String(200))
    item_code = db.Column(db.String(50))
    order_type_code = db.Column(db.String(20))
    order_type = db.Column(db.String(50))
    maker = db.Column(db.String(100))
    remarks = db.Column(db.Text)
    member_count = db.Column(db.Integer)
    required_count = db.Column(db.Integer)
    seiban = db.Column(db.String(50))
    material = db.Column(db.String(100))
    is_received = db.Column(db.Boolean, default=False)
    received_at = db.Column(db.DateTime)
    has_internal_processing = db.Column(db.Boolean, default=False)  # 社内加工フラグ
    parent_id = db.Column(db.Integer, db.ForeignKey('order_detail.id'), nullable=True)# 🔥 親子関係フィールド
    part_number = db.Column(db.String(50))
    page_number = db.Column(db.String(20))
    row_number = db.Column(db.String(20))
    hierarchy = db.Column(db.Integer)
    # (children関係）
    children = db.relationship('OrderDetail', 
                            backref=db.backref('parent', remote_side=[id]),
                            lazy='dynamic')
    
    order = db.relationship('Order', backref=db.backref('details', lazy=True))


class ReceivedHistory(db.Model):
    """受入履歴テーブル - 発注番号をキーに受入情報を永続保存"""
    id = db.Column(db.Integer, primary_key=True)
    order_number = db.Column(db.String(50), nullable=False, index=True)  # 発注番号（キー）
    item_name = db.Column(db.String(200))  # 品名
    spec1 = db.Column(db.String(200))  # 仕様1
    quantity = db.Column(db.Integer)  # 数量
    is_received = db.Column(db.Boolean, default=True)  # 受入状態（True=受入、False=キャンセル）
    received_at = db.Column(db.DateTime)  # 受入日時
    cancelled_at = db.Column(db.DateTime)  # キャンセル日時
    received_by = db.Column(db.String(100))  # 受入者（IPアドレス）
    cancelled_by = db.Column(db.String(100))  # キャンセル者（IPアドレス）
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    updated_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc), onupdate=lambda: datetime.now(timezone.utc))

    @classmethod
    def record_receive(cls, order_number, item_name, spec1, quantity, client_ip):
        """受入を記録"""
        # 既存レコードを検索（発注番号+品名+仕様1+数量で一意）
        existing = cls.query.filter_by(
            order_number=order_number,
            item_name=item_name,
            spec1=spec1,
            quantity=quantity
        ).first()

        if existing:
            # 既存レコードを更新
            existing.is_received = True
            existing.received_at = datetime.now(timezone.utc)
            existing.received_by = client_ip
            existing.cancelled_at = None
            existing.cancelled_by = None
        else:
            # 新規レコードを作成
            history = cls(
                order_number=order_number,
                item_name=item_name,
                spec1=spec1,
                quantity=quantity,
                is_received=True,
                received_at=datetime.now(timezone.utc),
                received_by=client_ip
            )
            db.session.add(history)

        db.session.commit()

    @classmethod
    def record_cancel(cls, order_number, item_name, spec1, quantity, client_ip):
        """受入キャンセルを記録"""
        existing = cls.query.filter_by(
            order_number=order_number,
            item_name=item_name,
            spec1=spec1,
            quantity=quantity
        ).first()

        if existing:
            existing.is_received = False
            existing.cancelled_at = datetime.now(timezone.utc)
            existing.cancelled_by = client_ip
            db.session.commit()

    @classmethod
    def is_received_in_history(cls, order_number, item_name, spec1, quantity):
        """履歴から受入状態を確認"""
        existing = cls.query.filter_by(
            order_number=order_number,
            item_name=item_name,
            spec1=spec1,
            quantity=quantity,
            is_received=True
        ).first()
        return existing is not None

    @classmethod
    def get_received_info(cls, order_number, item_name, spec1, quantity):
        """履歴から受入情報を取得"""
        return cls.query.filter_by(
            order_number=order_number,
            item_name=item_name,
            spec1=spec1,
            quantity=quantity,
            is_received=True
        ).first()


class EditLog(db.Model):
    """編集ログテーブル"""
    id = db.Column(db.Integer, primary_key=True)
    detail_id = db.Column(db.Integer, db.ForeignKey('order_detail.id'))
    action = db.Column(db.String(50))  # 'receive', 'unreceive'
    ip_address = db.Column(db.String(45))
    timestamp = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    user_agent = db.Column(db.String(500))

class ProcessingHistory(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    serial_no = db.Column(db.Integer)
    issue_date = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    filename = db.Column(db.String(200))
    file_size_kb = db.Column(db.Float)
    seiban = db.Column(db.String(50))

# Initialize database
with app.app_context():
    db.create_all()

def to_jst(utc_dt):
    """UTC時刻をJSTに変換"""
    if utc_dt is None:
        return None
    if utc_dt.tzinfo is None:
        # naive datetimeの場合、UTCとして扱う
        utc_dt = pytz.utc.localize(utc_dt)
    jst = pytz.timezone('Asia/Tokyo')
    return utc_dt.astimezone(jst)

# Utility Functions
def check_network_file_access():
    """Check if network file is accessible"""
    try:
        network_path = Path(app.config['DEFAULT_EXCEL_PATH'])
        print(f"Checking path: {network_path}")  # デバッグログ
        
        if network_path.exists():
            file_stats = network_path.stat()
            file_size_mb = file_stats.st_size / (1024 * 1024)
            modified_time = datetime.fromtimestamp(file_stats.st_mtime)
            
            result = {
                'accessible': True,
                'path': str(network_path),
                'size_mb': round(file_size_mb, 2),
                'modified': modified_time.isoformat(),
                'filename': network_path.name
            }
            print(f"File info: {result}")  # デバッグログ
            return result
        else:
            print(f"File not found: {network_path}")  # デバッグログ
            return {
                'accessible': False,
                'error': f'ファイルが見つかりません: {network_path}'
            }
    except Exception as e:
        print(f"Access error: {str(e)}")  # デバッグログ
        return {
            'accessible': False,
            'error': f'アクセスエラー: {str(e)}'
        }

def copy_network_file_to_local():
    """Copy network file to local cache"""
    try:
        # パスをそのまま使用（既に正しいUNCパス形式）
        network_path = Path(app.config['DEFAULT_EXCEL_PATH'])
        if not network_path.exists():
            return None, f"ネットワークファイルが見つかりません: {network_path}"
        
        # Create cache filename with timestamp
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        cache_filename = f'cache/cached_{timestamp}_手配発注_ALL.xlsx'
        
        # Copy file to cache
        shutil.copy2(str(network_path), cache_filename)
        
        global last_refresh_time, cached_file_info
        last_refresh_time = datetime.now()
        cached_file_info = check_network_file_access()
        
        return cache_filename, None
    except Exception as e:
        return None, f"コピーエラー: {str(e)}"
    
def get_cad_file_info(spec1):
    """仕様1からCADファイル情報を取得"""
    if not spec1 or not spec1.startswith('N'):
        return None
    
    # NKA-00437-00-00 → K を抽出
    parts = spec1.split('-')
    if len(parts) < 2 or len(parts[0]) < 2:
        return None
    
    # 2文字目のアルファベットを取得
    folder_letter = parts[0][1].upper()
    
    # CADフォルダパス
    cad_folder = f"\\\\SERVER3\\Share-data\\CadData\\Parts\\{folder_letter}"
    
    # ファイルを検索（ワイルドカード）
    # 例: NKA-00437-00-00*.mx2 または NKA-00437-00-00*.pdf
    mx2_pattern = os.path.join(cad_folder, f"{spec1}*.mx2")
    pdf_pattern = os.path.join(cad_folder, f"{spec1}*.pdf")
    
    mx2_files = glob.glob(mx2_pattern)
    pdf_files = glob.glob(pdf_pattern)
    
    if not mx2_files and not pdf_files:
        return None
    
    return {
        'folder': cad_folder,
        'letter': folder_letter,
        'spec1': spec1,
        'mx2_files': mx2_files,
        'pdf_files': pdf_files,
        'has_mx2': len(mx2_files) > 0,
        'has_pdf': len(pdf_files) > 0
    }

def load_seiban_info():
    """製番一覧表から品名、得意先略称、メモ2を読み込む"""
    try:
        seiban_file = app.config.get('SEIBAN_LIST_PATH', r'\\server3\share-data\Document\Acrossデータ\製番一覧表.xlsx')
        seiban_path = Path(seiban_file)
        
        if not seiban_path.exists():
            print(f"製番一覧表が見つかりません: {seiban_path}")
            return {}
        
        # Excelファイルを読み込み
        df = pd.read_excel(str(seiban_path), sheet_name='製番')
        
        # 製番と情報の辞書を作成
        seiban_info = {}
        for _, row in df.iterrows():
            if pd.notna(row.get('製番')):
                seiban_info[str(row['製番'])] = {
                    'product_name': str(row.get('品名', '')) if pd.notna(row.get('品名')) else '',
                    'customer_abbr': str(row.get('得意先略称', '')) if pd.notna(row.get('得意先略称')) else '',
                    'memo2': str(row.get('メモ２', '')) if pd.notna(row.get('メモ２')) else ''
                }
        
        return seiban_info
    except Exception as e:
        print(f"製番一覧表読み込みエラー: {str(e)}")
        return {}
    

def check_file_update():
    """ファイルの更新をチェック"""
    global cached_file_info
    
    try:
        current_info = check_network_file_access()
        if not current_info['accessible']:
            return False, None
        
        if not cached_file_info:
            return True, "初回読み込み"
        
        # 最終更新時刻を比較
        cached_time = datetime.fromisoformat(cached_file_info.get('modified', ''))
        current_time = datetime.fromisoformat(current_info['modified'])
        
        if current_time > cached_time:
            return True, f"ファイルが更新されました（{current_time.strftime('%Y-%m-%d %H:%M:%S')}）"
        
        return False, None
    except Exception as e:
        return False, str(e)

def load_order_all_cache():
    """発注_ALLシートをメモリにキャッシュ（高速検索用）"""
    global order_all_cache, order_all_cache_time
    
    try:
        # キャッシュが有効か確認
        if order_all_cache_time:
            elapsed = (datetime.now(timezone.utc) - order_all_cache_time).total_seconds()
            if elapsed < CACHE_EXPIRY_SECONDS:
                print(f"✅ キャッシュ有効（残り{int(CACHE_EXPIRY_SECONDS - elapsed)}秒）")
                return True
        
        print("🔄 発注_ALLシートを読み込み中...")
        
        excel_path = Path(app.config['DEFAULT_EXCEL_PATH'])
        if not excel_path.exists():
            print(f"❌ ファイルが見つかりません: {excel_path}")
            return False
        
        # read_only=Trueで高速化、data_only=Trueで数式を評価
        df = pd.read_excel(
            str(excel_path), 
            sheet_name='発注_ALL',
            dtype={
                '発注番号': str,
                '納期': str,
                '製番': str,
                '材質': str,
                '品名': str,
                '仕様１': str,
                '仕入先略称': str,
                '発注数': str
            }
        )
        
        # 発注番号をキーとした辞書に変換（重複対応）
        order_all_cache.clear()
        
        # 🔥 サンプルキーをログ出力
        sample_keys = []
        
        for idx, row in df.iterrows():
            order_num = DataUtils.safe_str(row.get('発注番号', ''))
            if not order_num or order_num == '':
                continue
            
            # 🔥 最初の10件のキーをサンプル収集
            if len(sample_keys) < 10:
                sample_keys.append(f"元の値: '{order_num}'")
            
            # 発注番号を正規化（浮動小数点対策）
            order_num = DataUtils.normalize_order_number(order_num)
            
            # 🔥 正規化後のキーもサンプル収集
            if len(sample_keys) < 20:
                sample_keys.append(f"正規化後: '{order_num}'")
            
            # 同一発注番号が複数ある場合はリスト化
            if order_num not in order_all_cache:
                order_all_cache[order_num] = []
            
            order_all_cache[order_num].append({
                'delivery_date': DataUtils.safe_str(row.get('納期', '')),
                'seiban': DataUtils.safe_str(row.get('製番', '')),
                'material': DataUtils.safe_str(row.get('材質', '')),
                'item_name': DataUtils.safe_str(row.get('品名', '')),
                'spec1': DataUtils.safe_str(row.get('仕様１', '')),
                'supplier': DataUtils.safe_str(row.get('仕入先略称', '')),
                'quantity': DataUtils.safe_int(row.get('発注数', 0)),
                'unit_measure': DataUtils.safe_str(row.get('単位', '')),
                'staff': DataUtils.safe_str(row.get('担当者', ''))
            })
        
        order_all_cache_time = datetime.now(timezone.utc)
        
    except Exception as e:
        print(f"❌ キャッシュ読み込みエラー: {e}")
        import traceback
        traceback.print_exc()
        return False

def search_order_from_cache(order_number):
    """キャッシュから発注番号を検索"""
    if not load_order_all_cache():
        return None
    
    # 発注番号を正規化
    search_key = DataUtils.normalize_order_number(order_number)
    
    # 🔥 デバッグログ追加
    print(f"🔍 検索: 元の値='{order_number}' → 正規化後='{search_key}'")
    
    if search_key in order_all_cache:
        print(f"  ✅ キャッシュHIT: {len(order_all_cache[search_key])}件")
        return order_all_cache[search_key]
    else:
        # 🔥 部分一致検索を試行
        print(f"  ❌ キャッシュMISS")
        print(f"  🔎 類似キー検索中...")
        similar_keys = [k for k in list(order_all_cache.keys())[:50] if search_key in k or k in search_key]
        if similar_keys:
            print(f"    類似キー（最大5件）: {similar_keys[:5]}")
        else:
            print(f"    類似キーなし")
    
    return None

def auto_refresh_network_file():
    """Background task to refresh network file periodically"""
    while True:
        time.sleep(app.config['AUTO_REFRESH_INTERVAL'])
        try:
            cache_file, error = copy_network_file_to_local()
            if cache_file:
                print(f"自動更新完了: {datetime.now()}")
            else:
                print(f"自動更新エラー: {error}")
        except Exception as e:
            print(f"自動更新例外: {str(e)}")

# Start background refresh thread
def start_auto_refresh():
    global refresh_thread
    if not refresh_thread or not refresh_thread.is_alive():
        refresh_thread = threading.Thread(target=auto_refresh_network_file, daemon=True)
        refresh_thread.start()

def extract_seiban_from_filename(filename):
    """Extract seiban (MHTxxxx) from filename"""
    pattern = r'(MHT\d{4})'
    match = re.search(pattern, filename)
    return match.group(1) if match else None

def process_excel_file(file_path, sheet1_name, sheet2_name, seiban_prefix, order_date_from=None, order_date_to=None):
    """Process Excel file and merge data"""
    try:
        df1 = pd.read_excel(file_path, sheet_name=sheet1_name, header=0)
        df2 = pd.read_excel(file_path, sheet_name=sheet2_name, header=0)
        
        # デバッグ情報
        print(f"=== デバッグ情報 ===")
        print(f"検索製番: {seiban_prefix}")
        print(f"シート1名: {sheet1_name}, 件数: {len(df1)}件")
        print(f"シート2名: {sheet2_name}, 件数: {len(df2)}件")
        
        # 製番の前処理（空白除去）
        df1['製番'] = df1['製番'].astype(str).str.strip()
        df2['製番'] = df2['製番'].astype(str).str.strip()
        
        # 製番でフィルタリング
        df1 = df1[df1['製番'].str.startswith(seiban_prefix, na=False)]
        df2 = df2[df2['製番'].str.startswith(seiban_prefix, na=False)]
        
        print(f"製番フィルタ({seiban_prefix})後: シート1={len(df1)}件, シート2={len(df2)}件")
        print(f"===================")
        
        return process_excel_file_from_dataframes(df1, df2, seiban_prefix, order_date_from, order_date_to)
        
    except Exception as e:
        print(f"エラー詳細: {e}")
        import traceback
        traceback.print_exc()
        raise Exception(f"Error processing Excel: {str(e)}")

def process_excel_file_from_dataframes(df1, df2, seiban_prefix, order_date_from=None, order_date_to=None):
    """Process dataframes and merge data"""
    try:
        df1 = df1.fillna('')
        df2 = df2.fillna('')
        
        if (order_date_from or order_date_to) and '発注日' in df2.columns:
            print(f"発注日フィルタ適用:")
            df2['発注日'] = pd.to_datetime(df2['発注日'], errors='coerce')
            before_count = len(df2)
            
            if order_date_from:
                filter_date_from = pd.to_datetime(order_date_from)
                df2 = df2[df2['発注日'] >= filter_date_from]
                print(f"  開始日: {order_date_from}以降")
            
            if order_date_to:
                filter_date_to = pd.to_datetime(order_date_to)
                df2 = df2[df2['発注日'] <= filter_date_to]
                print(f"  終了日: {order_date_to}まで")
            
            after_count = len(df2)
            print(f"発注日フィルタ: {before_count}件 → {after_count}件 ({before_count - after_count}件除外)")
        
        if '発注番号' in df1.columns:
            df1['発注番号'] = df1['発注番号'].apply(lambda x: str(int(float(x))) if isinstance(x, (int, float)) and x != '' else str(x))
        if '発注番号' in df2.columns:
            df2['発注番号'] = df2['発注番号'].apply(lambda x: str(int(float(x))) if isinstance(x, (int, float)) and x != '' else str(x))
        
        if '手配区分CD' in df1.columns:
            df1['手配区分CD'] = df1['手配区分CD'].apply(lambda x: str(int(float(x))) if isinstance(x, (int, float)) and x != '' else str(x))
        if '手配区分CD' in df2.columns:
            df2['手配区分CD'] = df2['手配区分CD'].apply(lambda x: str(int(float(x))) if isinstance(x, (int, float)) and x != '' else str(x))
        
        # 仕入先CDも文字列に変換
        if '仕入先CD' in df2.columns:
            df2['仕入先CD'] = df2['仕入先CD'].apply(lambda x: str(int(float(x))) if isinstance(x, (int, float)) and x != '' else str(x))
        
        if '納期' in df2.columns:
            df2.loc[df2['納期'] != '', '納期'] = pd.to_datetime(
                df2.loc[df2['納期'] != '', '納期'], 
                errors='coerce'
            )
        
        # 仕入先CD用の列を初期化
        if '仕入先CD' not in df1.columns:
            df1['仕入先CD'] = ''
        
        # Merge data
        for i, row in df1.iterrows():
            matched = False
            
            if all(col in row for col in ['材質', '仕様１', '製番']):
                if row['材質'] and row['仕様１'] and row['製番']:
                    cond = ((df2['材質'] == row['材質']) &
                           (df2['仕様１'] == row['仕様１']) &
                           (df2['製番'] == row['製番']))
                    match = df2[cond]
                    if not match.empty:
                        matched = True
                        if '発注番号' in match.columns:
                            df1.at[i, '発注番号'] = str(match['発注番号'].iloc[0])
                        if '仕入先略称' in match.columns:
                            df1.at[i, '仕入先略称'] = match['仕入先略称'].iloc[0]
                        if '仕入先CD' in match.columns:
                            df1.at[i, '仕入先CD'] = str(match['仕入先CD'].iloc[0])
                        if '納期' in match.columns:
                            try:
                                if pd.notna(match['納期'].iloc[0]) and match['納期'].iloc[0] != '':
                                    df1.at[i, '納期'] = match['納期'].dt.strftime('%y/%m/%d').iloc[0]
                            except:
                                df1.at[i, '納期'] = str(match['納期'].iloc[0])
            
            if not matched and all(col in row for col in ['製番', '仕様１', '手配区分']):
                if row['製番'] and row['仕様１']:
                    cond = ((df2['製番'] == row['製番']) &
                           (df2['仕様１'] == row['仕様１']))
                    if row['手配区分'] and '手配区分' in df2.columns:
                        cond = cond & (df2['手配区分'] == row['手配区分'])
                    
                    match = df2[cond]
                    if not match.empty:
                        if '発注番号' in match.columns:
                            df1.at[i, '発注番号'] = str(match['発注番号'].iloc[0])
                        if '仕入先略称' in match.columns:
                            df1.at[i, '仕入先略称'] = match['仕入先略称'].iloc[0]
                        if '仕入先CD' in match.columns:
                            df1.at[i, '仕入先CD'] = str(match['仕入先CD'].iloc[0])
                        if '納期' in match.columns:
                            try:
                                if pd.notna(match['納期'].iloc[0]) and match['納期'].iloc[0] != '':
                                    df1.at[i, '納期'] = match['納期'].dt.strftime('%y/%m/%d').iloc[0]
                            except:
                                df1.at[i, '納期'] = str(match['納期'].iloc[0])
        
        # Reorder columns（仕入先CDを含める - DB保存用）
        cols = ['納期', '仕入先略称', '仕入先CD', '発注番号', '手配数', '単位', '品名', '仕様１', '仕様２',
                '品目CD', '手配区分CD', '手配区分', 'メーカー', '備考', '員数', '必要数', '製番', '材質']
        cols = [c for c in cols if c in df1.columns]
        df1 = df1[cols]
        
        if '発注番号' in df1.columns:
            df1_with_order = df1[df1['発注番号'] != '']
            df1_without_order = df1[df1['発注番号'] == '']
            
            if not df1_with_order.empty:
                df1_with_order = df1_with_order.sort_values('発注番号')
            
            df1 = pd.concat([df1_with_order, df1_without_order], ignore_index=True)
        
        return df1
    except Exception as e:
        raise Exception(f"Error processing dataframes: {str(e)}")
    
def create_order_detail_with_parts(row, order, all_received_items, safe_str, safe_int):
    """OrderDetail作成"""
    import pandas as pd
    
    order_type = safe_str(row.get('手配区分', ''))
    has_internal = '社内加工' in order_type or '追加工' in order_type
    
    # 仕入先CDを取得
    supplier_cd = row.get('仕入先CD', '')
    if isinstance(supplier_cd, (int, float)) and not pd.isna(supplier_cd):
        supplier_cd = str(int(supplier_cd))
    else:
        supplier_cd = safe_str(supplier_cd)
    
    detail = OrderDetail(
        order_id=order.id,
        delivery_date=safe_str(row.get('納期', '')),
        supplier=safe_str(row.get('仕入先略称', '')),
        supplier_cd=supplier_cd,  # 追加
        order_number=DataUtils.normalize_order_number(row.get('発注番号', '')),
        quantity=safe_int(row.get('手配数', 0)),
        unit_measure=safe_str(row.get('単位', '')),
        item_name=safe_str(row.get('品名', '')),
        spec1=safe_str(row.get('仕様１', '')),
        spec2=safe_str(row.get('仕様２', '')),
        item_code=safe_str(row.get('品目CD', '')),
        order_type_code=DataUtils.normalize_order_number(row.get('手配区分CD', '')),
        order_type=order_type,
        maker=safe_str(row.get('メーカー', '')),
        remarks=safe_str(row.get('備考', '')),
        member_count=safe_int(row.get('員数', 0)),
        required_count=safe_int(row.get('必要数', 0)),
        seiban=safe_str(row.get('製番', '')),
        material=safe_str(row.get('材質', '')).replace('-', ''),
        has_internal_processing=has_internal,
        part_number=safe_str(row.get('部品No', '')),
        page_number=safe_str(row.get('ページNo', '')),
        row_number=safe_str(row.get('行No', '')),
        hierarchy=safe_int(row.get('階層', 0))
    )
    
    _restore_received_status(detail, all_received_items)
    return detail

def _restore_received_status(detail, all_received_items):
    """受入状態復元（既存データ優先、なければReceivedHistoryから復元）"""
    restored = False

    # 1. まず既存データ（同じ製番内）から復元を試みる
    if detail.order_number and detail.order_number in all_received_items:
        for received in all_received_items[detail.order_number]:
            if (received['item_name'] == detail.item_name and
                received['spec1'] == detail.spec1 and
                received['quantity'] == detail.quantity):
                detail.is_received = True
                detail.received_at = received['received_at']
                restored = True
                break

    # 2. 既存データで復元できなかった場合、ReceivedHistoryから復元
    if not restored and detail.order_number:
        history = ReceivedHistory.get_received_info(
            order_number=detail.order_number,
            item_name=detail.item_name,
            spec1=detail.spec1,
            quantity=detail.quantity
        )
        if history:
            detail.is_received = True
            detail.received_at = history.received_at
            print(f"✅ 受入履歴から復元: 発注番号={detail.order_number}, 品名={detail.item_name}")
            
def update_order_status(order):
    """注文ステータス更新"""
    if not order.details:
        return
    
    all_received = all(d.is_received for d in order.details)
    any_received = any(d.is_received for d in order.details)
    
    if all_received:
        order.status = Constants.STATUS_COMPLETED
    elif any_received:
        order.status = Constants.STATUS_IN_PROGRESS
    else:
        order.status = Constants.STATUS_BEFORE
    
    order.updated_at = datetime.now(timezone.utc)

def save_order_to_excel(order, filepath):
    """注文をExcelファイルに保存（一時ファイル経由）"""
    import tempfile
    import shutil
    
    try:
        unit_display = order.unit if order.unit else 'ユニット名無し'
        sheet_name = f"{order.seiban}_{unit_display}"
        sheet_name = re.sub(r'[\\\/\?\*\[\]:]', '', sheet_name)[:31]
        
        temp_fd, temp_path = tempfile.mkstemp(suffix='.xlsx')
        os.close(temp_fd)
        
        try:
            if Path(filepath).exists():
                try:
                    wb = load_workbook(filepath)
                except PermissionError:
                    return False, "ファイルが他のユーザーによって使用中です"
                
                # 既存シートを削除
                if sheet_name in wb.sheetnames:
                    del wb[sheet_name]
                
                # ガントチャートシートを削除して再作成
                if "納期ガントチャート" in wb.sheetnames:
                    del wb["納期ガントチャート"]
            else:
                wb = Workbook()
                wb.remove(wb.active)
            
            # 🔥 全ユニットを取得してガントチャートを再作成
            orders = Order.query.filter_by(seiban=order.seiban, is_archived=False).all()
            create_gantt_chart_sheet(wb, order.seiban, orders)
            
            # 新しいシートを作成
            ws = wb.create_sheet(sheet_name)
            create_order_sheet(ws, order, sheet_name)
            
            wb.save(temp_path)
            wb.close()
            
            try:
                shutil.move(temp_path, filepath)
                return True, None
            except PermissionError:
                backup_path = filepath.replace('.xlsx', f'_backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
                shutil.move(temp_path, backup_path)
                print(f"⚠️  バックアップ保存: {backup_path}")
                return False, f"元ファイルが使用中のため、バックアップを作成しました: {backup_path}"
                
        finally:
            if Path(temp_path).exists():
                try:
                    os.remove(temp_path)
                except:
                    pass
        
    except Exception as e:
        return False, str(e)
    
def get_order_excel_path(seiban, product_name=None, customer_abbr=None):
    """製番に対応するExcelファイルパスを取得（品名・客先名付き）"""
    export_dir = Path(app.config['EXPORT_EXCEL_PATH'])
    export_dir.mkdir(parents=True, exist_ok=True)
    
    # 品名と客先名をファイル名に含める（Windowsファイル名禁止文字を除去）
    if product_name:
        safe_product_name = re.sub(r'[\\/:*?"<>|]', '', product_name)
        if customer_abbr:
            safe_customer_abbr = re.sub(r'[\\/:*?"<>|]', '', customer_abbr)
            filename = f"{seiban}_{safe_product_name}_{safe_customer_abbr}_手配発注リスト.xlsx"
        else:
            filename = f"{seiban}_{safe_product_name}_手配発注リスト.xlsx"
    else:
        filename = f"{seiban}_手配発注リスト.xlsx"
    
    return str(export_dir / filename)
    
def update_order_excel(order_id):
    """注文IDに対応するExcelファイルを更新"""
    try:
        order = db.session.get(Order, order_id)
        if not order:
            return False, "注文が見つかりません"
        
        # 品名と客先名を渡す
        filepath = get_order_excel_path(order.seiban, order.product_name, order.customer_abbr)
        success, error = save_order_to_excel(order, filepath)
        
        if success:
            print(f"✅ Excel更新成功: {filepath}")
        else:
            print(f"❌ Excel更新失敗: {error}")
        
        return success, error
    except Exception as e:
        return False, str(e)
    
def save_to_database(df, seiban_prefix):
    """Save processed data to database"""
    try:
        seiban_info = load_seiban_info()
        info = seiban_info.get(seiban_prefix, {})
        product_name = info.get('product_name', '')  # 品名を取得
        customer_abbr = info.get('customer_abbr', '')  # 客先名を取得
        
        df['材質'] = df['材質'].replace('', '-')
        materials = df['材質'].unique()
        
        cols_to_keep = ['部品No', 'ページNo', '行No', '階層']
        for col in cols_to_keep:
            if col not in df.columns:
                df[col] = ''
        
        all_received_items = {}
        existing_orders = Order.query.filter_by(seiban=seiban_prefix).all()
        for existing_order in existing_orders:
            for detail in existing_order.details:
                if detail.is_received and detail.order_number:
                    key = str(detail.order_number)
                    if key not in all_received_items:
                        all_received_items[key] = []
                    all_received_items[key].append({
                        'item_name': detail.item_name,
                        'spec1': detail.spec1,
                        'quantity': detail.quantity,
                        'is_received': True,
                        'received_at': detail.received_at
                    })
        
        def safe_str(value):
            if pd.isna(value) or value is None:
                return ''
            if isinstance(value, float) and value == value:
                try:
                    return str(int(value))
                except:
                    return str(value)
            return str(value)
        
        def safe_int(value, default=0):
            if pd.isna(value) or value is None:
                return default
            try:
                return int(float(value))
            except (ValueError, TypeError):
                return default
        
        for material in materials:
            material_df = df[df['材質'] == material]
            
            # 🔥 ユニット名を正規化（検索と作成で統一）
            unit_name = material if material and material != '-' else ''
            
            # 🔥 正規化されたユニット名で検索
            order = Order.query.filter_by(seiban=seiban_prefix, unit=unit_name).first()
            if not order:
                # 🔥 新規作成
                order = Order(
                    seiban=seiban_prefix, 
                    unit=unit_name,  # 正規化された値を使用
                    product_name=product_name,
                    customer_abbr=info.get('customer_abbr', ''),
                    memo2=info.get('memo2', '')
                )
                db.session.add(order)
                db.session.flush()
                print(f"✅ 新規ユニット作成: {seiban_prefix} - {unit_name or 'ユニット名無し'}")
            else:
                # 🔥 既存レコードを更新
                order.product_name = product_name
                order.customer_abbr = info.get('customer_abbr', '')
                order.memo2 = info.get('memo2', '')
                print(f"🔄 既存ユニット更新: {seiban_prefix} - {unit_name or 'ユニット名無し'} (ID: {order.id})")
            
            # 既存の詳細を削除して再作成
            OrderDetail.query.filter_by(order_id=order.id).delete()
            
            part_groups = {}
            
            for _, row in material_df.iterrows():
                part_no = safe_str(row.get('部品No', ''))
                page_no = safe_str(row.get('ページNo', ''))
                material_key = safe_str(row.get('材質', ''))
                
                group_key = (part_no, page_no, material_key)
                
                if group_key not in part_groups:
                    part_groups[group_key] = []
                part_groups[group_key].append(row)
            
            for group_key, rows in part_groups.items():
                part_no, page_no, material_key = group_key
                
                blanks = []
                processed = []
                others = []
                
                for row in rows:
                    order_type_code = safe_str(row.get('手配区分CD', ''))

                    # 手配区分CDが空欄のものは除外
                    if not order_type_code or order_type_code.strip() == '':
                        item_name = safe_str(row.get('品名', ''))
                        print(f"除外: {item_name} - 手配区分CDが空欄")
                        continue

                    if order_type_code == '13':
                        blanks.append(row)
                    elif order_type_code == '11':
                        processed.append(row)
                    else:
                        others.append(row)
                
                blanks.sort(key=lambda r: safe_int(r.get('行No', 0)))
                processed.sort(key=lambda r: safe_int(r.get('行No', 0)))
                
                if blanks or processed:
                    print(f"\nグループ: 部品No={part_no}, ページNo={page_no}")
                    print(f"ブランク（親）候補: {len(blanks)}個, 追加工（子）候補: {len(processed)}個")
                
                used_processed = set()
                
                for blank_row in blanks:
                    blank_row_no = safe_int(blank_row.get('行No', 0))
                    
                    closest_processed = None
                    min_diff = float('inf')
                    
                    for i, proc_row in enumerate(processed):
                        if i in used_processed:
                            continue
                        
                        proc_row_no = safe_int(proc_row.get('行No', 0))
                        diff = abs(blank_row_no - proc_row_no)
                        
                        if diff < min_diff:
                            min_diff = diff
                            closest_processed = (i, proc_row)
                    
                    parent_detail = create_order_detail_with_parts(
                        blank_row, order, all_received_items, safe_str, safe_int
                    )
                    db.session.add(parent_detail)
                    db.session.flush()
                    
                    blank_name = safe_str(blank_row.get('品名', ''))
                    
                    if closest_processed is not None:
                        proc_idx, proc_row = closest_processed
                        used_processed.add(proc_idx)
                        
                        child_detail = create_order_detail_with_parts(
                            proc_row, order, all_received_items, safe_str, safe_int
                        )
                        child_detail.parent_id = parent_detail.id
                        db.session.add(child_detail)
                        
                        proc_name = safe_str(proc_row.get('品名', ''))
                        proc_row_no = safe_int(proc_row.get('行No', 0))
                        
                        print(f"親子設定: 親({blank_name[:15]}, 行No={blank_row_no}) "
                              f"→ 子({proc_name[:15]}, 行No={proc_row_no}, 差={min_diff})")
                    else:
                        print(f"親のみ: {blank_name[:15]} (行No={blank_row_no}) - 対応する子なし")
                
                for i, proc_row in enumerate(processed):
                    if i not in used_processed:
                        order_type_code = safe_str(proc_row.get('手配区分CD', ''))
                        spec1 = safe_str(proc_row.get('仕様１', ''))
                        
                        if (order_type_code == '15' and spec1 and spec1.strip() and re.match(r'^M\d', spec1)) or \
                        (not spec1 or not spec1.strip()):
                            if not spec1 or not spec1.strip():
                                print(f"除外: {proc_name} (仕様1空欄) - 仕様1が未入力")
                            else:
                                print(f"除外: {proc_name} ({spec1}) - 在庫部品のM+数値")
                            continue
                        
                        proc_detail = create_order_detail_with_parts(
                            proc_row, order, all_received_items, safe_str, safe_int
                        )
                        db.session.add(proc_detail)
                        
                        proc_name = safe_str(proc_row.get('品名', ''))
                        proc_row_no = safe_int(proc_row.get('行No', 0))
                        print(f"子のみ: {proc_name[:15]} (行No={proc_row_no}) - 対応する親なし")
                
                for row in others:
                    order_type_code = safe_str(row.get('手配区分CD', ''))
                    spec1 = safe_str(row.get('仕様１', ''))
                    
                    if order_type_code == '15' and re.match(r'^M\d', spec1):
                        item_name = safe_str(row.get('品名', ''))
                        print(f"除外: {item_name} ({spec1}) - 在庫部品のM+数値")
                        continue
                    
                    detail = create_order_detail_with_parts(
                        row, order, all_received_items, safe_str, safe_int
                    )
                    db.session.add(detail)
        
        for order in Order.query.filter_by(seiban=seiban_prefix).all():
            update_order_status(order)

        db.session.commit()
                
        try:
            filepath = get_order_excel_path(seiban_prefix, product_name, customer_abbr)
            
            if filepath:
                orders = Order.query.filter_by(seiban=seiban_prefix, is_archived=False).all()
                
                if orders:
                    wb = Workbook()
                    wb.remove(wb.active)  # デフォルトシートを削除
                    
                    # 🔥 1シート目: ガントチャート
                    create_gantt_chart_sheet(wb, seiban_prefix, orders)
                    
                    # 2シート目以降: 各ユニットの手配リスト
                    for order in orders:
                        unit = order.unit if order.unit else 'ユニット名無し'
                        sheet_name = f"{seiban_prefix}_{unit}"[:31]
                        ws = wb.create_sheet(title=sheet_name)
                        create_order_sheet(ws, order, sheet_name)
                    
                    wb.save(filepath)
                    wb.close()
                    print(f"✅ Excel自動出力成功（ガントチャート付き）: {filepath}")
            else:
                print("⚠️  Excel出力パスの取得失敗")
        except Exception as excel_error:
            print(f"⚠️  Excel出力エラー（処理は継続）: {excel_error}")
            import traceback
            traceback.print_exc()
        
        return True
        
    except Exception as e:
        db.session.rollback()
        raise Exception(f"Database error: {str(e)}")

def get_server_url():
    """サーバーのURLを取得（IP + ポート）"""
    try:
        import socket
        # ホスト名からIPアドレスを取得
        hostname = socket.gethostname()
        ip_address = socket.gethostbyname(hostname)
        
        # 設定からHTTPSの使用状況を確認
        config_obj = get_config() if hasattr(app, 'config') else None
        use_https = False
        port = 8080
        
        if config_obj:
            use_https = getattr(config_obj, 'USE_HTTPS', False)
            port = getattr(config_obj, 'PORT', 8080)
        
        protocol = 'https' if use_https else 'http'
        return f"{protocol}://{ip_address}:{port}"
    except Exception as e:
        print(f"サーバーURL取得エラー: {e}")
        # フォールバック
        return "http://localhost:8080"
    
def create_order_sheet(ws, order, sheet_name=None):
    """ワークシート作成（縦向き印刷、QRコードH列配置）"""
    from openpyxl.worksheet.page import PageMargins
    from openpyxl.drawing.image import Image
    from io import BytesIO
    import qrcode
    
    if sheet_name:
        ws.title = sheet_name
    
    # ヘッダー情報
    unit_display = order.unit if order.unit else 'ユニット名無し'
    customer = order.customer_abbr if order.customer_abbr else ''
    memo = order.memo2 if order.memo2 else ''
    product_name = order.product_name if order.product_name else ''
    
    # 🔥 QRコード生成（受入専用ページURL）
    try:
        server_url = get_server_url()
        receive_url = f"{server_url}/receive/{order.id}"
        
        # QRコード画像を生成
        qr = qrcode.QRCode(
            version=1,
            box_size=8,
            border=3
        )
        qr.add_data(receive_url)
        qr.make(fit=True)
        
        qr_img = qr.make_image(fill_color="black", back_color="white")
        
        # BytesIOに保存
        qr_buffer = BytesIO()
        qr_img.save(qr_buffer, format='PNG')
        qr_buffer.seek(0)
        
        # Excelに画像を挿入
        img = Image(qr_buffer)
        img.width = 100
        img.height = 100
        
        # 🔥 QRコードをH1セルに配置
        ws.add_image(img, 'I1')
        
        # 🔥 URLテキストとラベルをJ列に配置（QRコードの右側）
        ws['K1'] = '💻️ 受入確認専用ページ(社内LANよりアクセス)'
        ws['K1'].font = Font(size=9, bold=True)
        ws['K1'].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

        # 🔥 URLテキストとラベルをM列に配置（QRコードの右側）
        ws['M1'] = '💻️ 受入確認専用ページ(社内LANよりアクセス)'
        ws['M1'].font = Font(size=9, bold=True)
        ws['M1'].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

        ws['M2'] = receive_url
        ws['M2'].font = Font(size=8, color='0000FF', underline='single')
        ws['M2'].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        
    except Exception as e:
        print(f"⚠️ QRコード生成エラー: {e}")
    
    # 🔥 1行目: 4列に情報を配置
    # A1: 製番 + 品名 + 得意先 + メモ
    a1_text_parts = [order.seiban]
    if product_name:
        a1_text_parts.append(product_name)
    if customer:
        a1_text_parts.append(customer)
    if memo:
        a1_text_parts.append(memo)
    
    ws['A1'] = ' '.join(a1_text_parts)
    ws['A1'].font = Font(size=14, bold=True)
    ws['A1'].alignment = Alignment(horizontal='left', vertical='center')
    
    # A2: ユニット名
    ws['A2'] = unit_display
    ws['A2'].font = Font(size=14, bold=True)
    ws['A2'].alignment = Alignment(horizontal='left', vertical='center')

    # A3: 注意書き（赤字）
    ws['A3'] = '※ピンク塗は受入済 製番外の持ち出しは必ず記録を残すこと データは保存先にて随時更新'
    ws['A3'].font = Font(size=9, bold=True, color=Constants.COLOR_RED)
    ws['A3'].alignment = Alignment(horizontal='left', vertical='center')

    # A4: ネットワークパス（赤字）
    ws['A4'] = r'\\SERVER3\Share-data\Document\仕入れ\002_手配リスト\手配発注リスト'
    ws['A4'].font = Font(size=9, bold=True, color=Constants.COLOR_RED)
    ws['A4'].alignment = Alignment(horizontal='left', vertical='center')

    # 🔥 行の高さ調整
    ws.row_dimensions[1].height = 35
    
    # 🔥 ヘッダー行（6行目）
    headers = Constants.EXCEL_COLUMNS
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=6, column=col_idx)
        cell.value = header
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.fill = PatternFill(start_color=Constants.COLOR_HEADER, 
                               end_color=Constants.COLOR_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # 🔥 列幅設定（縦向き印刷用に最適化）
    column_widths = {
        'A': 9,   # 納入日（新規）
        'B': 6,   # 納入数（新規）
        'C': 9,   # 納期
        'D': 11,  # 仕入先略称
        'E': 9,   # 発注番号
        'F': 5,   # 手配数
        'G': 4,   # 単位
        'H': 18,  # 品名
        'I': 15,  # 仕様１
        'J': 12,  # 仕様２
        'K': 10,  # 手配区分
        'L': 8,   # メーカー
        'M': 12   # 備考
    }

    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width

    # 🔥 検収データを読み込み
    delivery_dict = DeliveryUtils.load_delivery_data()

    # 🔥 データ行を書き込む（7行目から開始）
    row_idx = 7
    parent_details = [d for d in order.details if d.parent_id is None]

    for detail in parent_details:
        row_idx = _write_detail_row(ws, detail, row_idx, is_parent=True, delivery_dict=delivery_dict)

        # 子アイテム
        children = [d for d in order.details if d.parent_id == detail.id]
        for child in children:
            row_idx = _write_detail_row(ws, child, row_idx, is_parent=False, delivery_dict=delivery_dict)
    
    # 🔥 ページ設定（縦向き印刷）
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    
    # 🔥 余白を最小化
    ws.page_margins = PageMargins(
        left=0.25,
        right=0.25,
        top=0.3,
        bottom=0.5,
        header=0.15,
        footer=0.2
    )
    
    # 🔥 印刷タイトル行（ヘッダーを毎ページ印刷）
    ws.print_title_rows = '1:6'
    ws.print_area = f'A1:M{row_idx - 1}'
    
    # 🔥 フッター設定（フォントサイズ10に縮小）
    footer_parts = []
    
    # 製番は必須
    footer_parts.append(order.seiban)
    
    # ユニット名
    if unit_display and unit_display != 'ユニット名無し':
        footer_parts.append(unit_display)
    
    # 品名（長すぎる場合は省略）
    if product_name:
        display_product_name = product_name if len(product_name) <= 20 else product_name[:20] + '...'
        footer_parts.append(display_product_name)
    
    # 得意先略称
    if customer:
        footer_parts.append(customer)
    
    # メモ２
    if memo:
        footer_parts.append(memo)
    
    footer_text = '_'.join(footer_parts)
    
    # 🔥 フッター設定（フォントサイズを10に変更）
    for footer in [ws.oddFooter, ws.evenFooter, ws.firstFooter]:
        footer.left.text = f"&10&B{footer_text}"  # &10でフォントサイズ10
        footer.center.text = "&P / &N"
        footer.right.text = f"&10&B{footer_text}"
    
    # 🔥 改ページプレビュー表示
    ws.sheet_view.view = 'pageBreakPreview'
    
    return ws

def _setup_page_settings(ws):
    """ページ設定"""
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = False
    ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.2, bottom=1.2, header=0.2, footer=0.1)

def _create_header_rows(ws, order, unit_display, customer, memo):
    """ヘッダー行作成"""
    header_parts = [f"{order.seiban}({order.product_name})"]
    if unit_display:
        header_parts.append(unit_display)
    if customer:
        header_parts.append(customer)
    if memo:
        header_parts.append(memo)
    
    ws['A1'] = '_'.join(header_parts) + "_"
    ws['A1'].font = Font(size=26, bold=True)
    ws['A1'].alignment = Alignment(horizontal='left', vertical='center')
    
    ws['A2'] = '※ピンク塗は受入済　製番外の持ち出しは必ず記録を残すこと　データは保存先にて随時更新: \\\\SERVER3\\Share-data\\Document\\仕入れ\\002_手配リスト\\手配発注リスト'
    ws['A2'].font = Font(size=11, bold=True, color=Constants.COLOR_RED)
    ws['A2'].alignment = Alignment(horizontal='left', vertical='center')
    
    ws.merge_cells('A1:K1')
    ws.merge_cells('A2:K2')
    ws.row_dimensions[1].height = 35
    ws.row_dimensions[2].height = 20


def _create_column_headers(ws):
    """列ヘッダー作成"""
    for col, header in enumerate(Constants.EXCEL_COLUMNS, 1):
        cell = ws.cell(row=3, column=col, value=header)
        cell.font = Font(bold=True, color=Constants.COLOR_WHITE)
        cell.fill = PatternFill(start_color=Constants.COLOR_HEADER, 
                               end_color=Constants.COLOR_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal='center', vertical='center')


def _create_data_rows(ws, order):
    """データ行作成"""
    row_idx = 4
    parent_details = [d for d in order.details if d.parent_id is None]

    # 検収データを読み込み
    delivery_dict = DeliveryUtils.load_delivery_data()

    for detail in parent_details:
        row_idx = _write_detail_row(ws, detail, row_idx, is_parent=True, delivery_dict=delivery_dict)

        # 子アイテム
        children = [d for d in order.details if d.parent_id == detail.id]
        for child in children:
            row_idx = _write_detail_row(ws, child, row_idx, is_parent=False, delivery_dict=delivery_dict)

    return row_idx


def _write_detail_row(ws, detail, row_idx, is_parent=True, delivery_dict=None):
    """詳細行を出力"""
    is_blank = '加工用ブランク' in str(detail.order_type)
    supplier_cd = getattr(detail, 'supplier_cd', None)
    spec1_value = detail.spec1 or ''
    spec2_value = detail.spec2 or ''
    is_mekki = MekkiUtils.is_mekki_target(supplier_cd, spec2_value, spec1_value)

    remarks = MekkiUtils.add_mekki_alert(detail.remarks) if is_mekki else (detail.remarks or '')

    # 検収データから納入日・納入数を取得
    delivery_info = DeliveryUtils.get_delivery_info(detail.order_number, delivery_dict)
    delivery_date = delivery_info.get('納入日', '')
    delivery_qty = delivery_info.get('納入数', 0)
    # 納入数が0の場合は空欄表示
    delivery_qty_display = delivery_qty if delivery_qty > 0 else ''

    data = [
        detail.received_at.strftime('%Y-%m-%d %H:%M:%S') if detail.received_at else '',  # 検収日
        '受入済' if detail.is_received else '未受入',  # 検収数（状態表示）
        detail.delivery_date, detail.supplier, detail.order_number,
        detail.quantity, detail.unit_measure, detail.item_name,
        detail.spec1, spec2_value, detail.order_type, detail.maker, remarks
    ]

    row_fill = ExcelStyler.get_fill(detail.is_received, row_idx % 2 == 0, not is_parent)
    cell_font = ExcelStyler.get_font(is_blank, False)

    for col, value in enumerate(data, 1):
        cell = ws.cell(row=row_idx, column=col, value=value)
        cell.fill = row_fill
        cell.alignment = Alignment(vertical='center')

        if col == 10 and is_mekki:  # 仕様２のカラムがJ(10)に変更
            cell.font = ExcelStyler.get_font(False, True)
        elif cell_font:
            cell.font = cell_font

        if col == 8 and not is_parent:  # 品名のカラムがH(8)に変更
            cell.value = f"  └ {value}"

    ws.row_dimensions[row_idx].height = 27
    return row_idx + 1

def _setup_print_settings(ws, row_idx, order, unit_display, customer, memo):
    """印刷設定"""
    ws.print_title_rows = '1:3'
    ws.print_area = f'A1:M{row_idx - 1}'
    
    footer_parts = [order.seiban]
    if unit_display:
        footer_parts.append(unit_display)
    if customer:
        footer_parts.append(customer)
    if memo:
        footer_parts.append(memo)
    
    footer_text = f"&24&B{'_'.join(footer_parts)}"
    
    for footer in [ws.oddFooter, ws.evenFooter, ws.firstFooter]:
        footer.center.text = "&P / &N"
        footer.left.text = footer_text
        footer.right.text = footer_text
    
    ws.sheet_view.view = 'pageBreakPreview'
    
# Excelファイル更新用の関数を追加
def refresh_excel_file():
    """Excelファイルの更新処理"""
    excel = None
    try:
        # COMを初期化（重要）
        pythoncom.CoInitialize()
        
        excel_path = app.config['DEFAULT_EXCEL_PATH']
        
        # Excel COMオブジェクトを使用
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # ファイルを開く（リンクを自動更新）
        wb = excel.Workbooks.Open(excel_path, UpdateLinks=3)
        
        # 全接続を更新
        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
        
        # 保存して閉じる
        wb.Save()
        wb.Close(False)
        excel.Quit()
        
        return True, "Excelファイルを更新しました"
        
    except Exception as e:
        return False, f"更新エラー: {str(e)}"
        
    finally:
        # 必ずクリーンアップ
        try:
            if excel:
                excel.Quit()
        except:
            pass
        
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        
def detect_seibans_from_excel(file_path, sheet_name, min_seiban='MHT0600'):
    """Excelから製番を自動検出"""
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)
        
        if '製番' not in df.columns:
            return []
        
        # 製番列から一意の値を取得
        seibans = df['製番'].dropna().unique()
        
        # MHTで始まり、指定番号以降のものをフィルター
        filtered_seibans = []
        for seiban in seibans:
            seiban_str = str(seiban).strip()
            if seiban_str.startswith('MHT'):
                # MHT0600 -> 600として比較
                try:
                    seiban_num = int(seiban_str.replace('MHT', ''))
                    min_num = int(min_seiban.replace('MHT', ''))
                    if seiban_num >= min_num:
                        filtered_seibans.append(seiban_str)
                except ValueError:
                    continue
        
        return sorted(filtered_seibans)
    except Exception as e:
        print(f"製番検出エラー: {e}")
        return []
    
# グローバル変数に追加
previous_seiban_counts = {}

def get_seiban_counts(file_path, sheet_name='手配リスト_ALL'):
    """製番ごとの件数を取得"""
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)
        if '製番' not in df.columns:
            return {}
        
        counts = df['製番'].value_counts().to_dict()
        return {str(k): int(v) for k, v in counts.items()}
    except Exception as e:
        print(f"件数取得エラー: {e}")
        return {}
    
@app.route('/api/order/<int:order_id>/send-completion-email', methods=['POST'])
def send_completion_email(order_id):
    """納品完了メールを作成してメーラーを起動"""
    try:
        order = Order.query.get_or_404(order_id)
        
        # 🔥 実際のExcelファイルパスを取得（get_order_excel_path()を使用）
        excel_path = get_order_excel_path(
            seiban=order.seiban,
            product_name=order.product_name,
            customer_abbr=order.customer_abbr
        )
        
        # メール送信
        success = EmailSender.send_completion_notification(
            seiban=order.seiban,
            product_name=order.product_name or '',
            customer_abbr=order.customer_abbr or '',
            unit=order.unit or '',
            memo2=order.memo2 or '',
            floor=order.floor or '',
            pallet_number=order.pallet_number or '',
            excel_path=excel_path,  # 🔥 実際のファイルパスを渡す
            sender_name='丸山'
        )
        
        if success:
            return jsonify({
                'success': True,
                'message': 'メーラーを起動しました'
            })
        else:
            return jsonify({
                'success': False,
                'error': 'メーラーの起動に失敗しました'
            }), 500
            
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

@app.route('/api/check-network-file-with-diff')
def check_network_file_with_diff():
    """ネットワークファイルの存在確認と差分検出"""
    try:
        network_file = app.config['DEFAULT_EXCEL_PATH']
        network_path = Path(network_file)
        
        if not network_path.exists():
            return jsonify({
                'accessible': False,
                'error': 'ファイルが見つかりません'
            })
        
        stat = network_path.stat()
        
        # 🔥 製番一覧表から情報を読み込み
        seiban_info_dict = load_seiban_info()
        
        # Excelから全シートの製番を読み込み
        try:
            wb = load_workbook(str(network_path), read_only=True, data_only=True)
            current_seiban_data = {}
            
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row and row[0]:  # 製番列
                        seiban = str(row[0]).strip()
                        if seiban and not seiban.startswith('#'):
                            if seiban not in current_seiban_data:
                                current_seiban_data[seiban] = 0
                            current_seiban_data[seiban] += 1
            
            wb.close()
            
            # 前回のキャッシュと比較
            cache_file = Path(app.config['UPLOAD_FOLDER']) / 'seiban_cache.json'
            diff_list = []
            
            if cache_file.exists():
                with open(cache_file, 'r', encoding='utf-8') as f:
                    cached_data = json.load(f)
                
                # 差分を検出
                for seiban, count in current_seiban_data.items():
                    old_count = cached_data.get(seiban, 0)
                    if count > old_count:
                        added = count - old_count
                        
                        # 🔥 製番一覧表から追加情報を取得
                        seiban_details = seiban_info_dict.get(seiban, {})
                        
                        diff_list.append({
                            'seiban': seiban,
                            'added': added,
                            'total': count,
                            'product_name': seiban_details.get('product_name', ''),
                            'customer_abbr': seiban_details.get('customer_abbr', ''),
                            'memo2': seiban_details.get('memo2', '')
                        })
                
                # 差分を追加件数でソート
                diff_list.sort(key=lambda x: x['added'], reverse=True)
            
            # 新しいデータをキャッシュ
            with open(cache_file, 'w', encoding='utf-8') as f:
                json.dump(current_seiban_data, f, ensure_ascii=False, indent=2)
            
            return jsonify({
                'accessible': True,
                'filename': network_path.name,
                'size_mb': round(stat.st_size / (1024 * 1024), 2),
                'modified': stat.st_mtime,
                'total_seibans': len(current_seiban_data),
                'diff': diff_list
            })
            
        except Exception as e:
            return jsonify({
                'accessible': True,
                'filename': network_path.name,
                'size_mb': round(stat.st_size / (1024 * 1024), 2),
                'modified': stat.st_mtime,
                'error': f'データ読み込みエラー: {str(e)}'
            })
    
    except Exception as e:
        return jsonify({
            'accessible': False,
            'error': str(e)
        })

@app.route('/api/detect-seibans', methods=['POST'])
def detect_seibans():
    """製番を自動検出"""
    try:
        data = request.json
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        min_seiban = data.get('min_seiban', 'MHT0600')
        
        if not filepath or not sheet_name:
            return jsonify({'error': 'ファイルパスとシート名が必要です'}), 400
        
        seibans = detect_seibans_from_excel(filepath, sheet_name, min_seiban)
        
        return jsonify({
            'success': True,
            'seibans': seibans,
            'count': len(seibans)
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/process', methods=['POST'])
def process_file_endpoint():
    """Excelファイル処理のメインエンドポイント"""
    try:
        data = request.json
        filepath = data['filepath']
        sheet1 = data['sheet1']
        sheet2 = data['sheet2']
        seiban = data['seiban']
        order_date_from = data.get('order_date_from')
        order_date_to = data.get('order_date_to')
        
        df_merged = process_excel_file(
            filepath, 
            sheet1, 
            sheet2, 
            seiban, 
            order_date_from, 
            order_date_to
        )
        
        if df_merged is None or len(df_merged) == 0:
            return jsonify({
                'success': False, 
                'error': f'製番 {seiban} のデータが見つかりません'
            })
        
        save_to_database(df_merged, seiban)
        
        return jsonify({
            'success': True,
            'message': f'{seiban} の処理が完了しました（{len(df_merged)}件）'
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})

# Routes
@app.route('/api/refresh-excel', methods=['POST'])
def refresh_excel_endpoint():
    """Excelファイルを更新するエンドポイント"""
    try:
        result = {'success': False, 'message': ''}
        
        def run_refresh():
            result['success'], result['message'] = refresh_excel_file()
        
        thread = Thread(target=run_refresh)
        thread.start()
        thread.join(timeout=60)
        
        if thread.is_alive():
            return jsonify({
                'success': False,
                'error': 'タイムアウト（60秒）'
            }), 500
        
        if result['success']:
            # キャッシュをクリア
            global cached_file_info, last_refresh_time
            last_refresh_time = datetime.now()
            
            # ファイル情報を取得
            file_info = check_network_file_access()
            cached_file_info = file_info
            
            # file_infoが正常に取得できているか確認
            if not file_info or not file_info.get('accessible'):
                # フォールバック: 基本情報のみ返す
                return jsonify({
                    'success': True,
                    'message': result['message'],
                    'file_info': {
                        'accessible': False,
                        'filename': 'Excel更新完了',
                        'size_mb': 0,
                        'modified': datetime.now().isoformat()
                    }
                })
            
            return jsonify({
                'success': True,
                'message': result['message'],
                'file_info': file_info
            })
        else:
            return jsonify({
                'success': False,
                'error': result['message']
            }), 500
            
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

# Pythonスクリプトを実行するエンドポイントを修正
@app.route('/api/run-refresh-script', methods=['POST'])
def run_refresh_script():
    """refresh_order_list.pyを実行"""
    try:
        # スクリプトパスのリスト（優先順位順）
        script_paths = [
            r"C:\Users\t.maruyama\order_merge\refresh_order_list.py",  # アプリと同じフォルダ
            r"C:\Users\t.maruyama\refresh_order_list.py",  # 元のパス
            os.path.join(os.path.dirname(__file__), "refresh_order_list.py")  # 相対パス
        ]
        
        # 存在するスクリプトを探す
        script_path = None
        for path in script_paths:
            if os.path.exists(path):
                script_path = path
                break
        
        if not script_path:
            return jsonify({
                'success': False,
                'error': f'スクリプトが見つかりません。以下のパスを確認してください:\n' + '\n'.join(script_paths)
            }), 404
        
        # Pythonスクリプトを実行（タイムアウトを延長）
        result = subprocess.run(
            [sys.executable, script_path],
            capture_output=True,
            text=True,
            timeout=120,  # 120秒に延長
            encoding='utf-8',  # 文字エンコーディングを指定
            errors='replace'  # エンコーディングエラーを回避
        )
        
        if result.returncode == 0:
            return jsonify({
                'success': True,
                'message': 'スクリプト実行成功',
                'output': result.stdout
            })
        else:
            return jsonify({
                'success': False,
                'error': result.stderr or 'スクリプト実行失敗',
                'output': result.stdout  # エラー時も出力を表示
            }), 500
            
    except subprocess.TimeoutExpired:
        return jsonify({
            'success': False,
            'error': 'タイムアウト（120秒）- スクリプトが長時間実行されています'
        }), 500
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@app.route('/')
def index():
    """Main page"""
    # Start auto refresh on first load
    start_auto_refresh()
    
    # デバッグ: 設定されているパスを確認
    print("=" * 50)
    print("設定されているパス:")
    print(f"DEFAULT_EXCEL_PATH: {app.config['DEFAULT_EXCEL_PATH']}")
    print(f"HISTORY_EXCEL_PATH: {app.config['HISTORY_EXCEL_PATH']}")
    print(f"SEIBAN_LIST_PATH: {app.config.get('SEIBAN_LIST_PATH', 'Not set')}")
    print("=" * 50)
    
    return render_template('index.html')

@app.route('/api/debug-paths')
def debug_paths():
    """パスの接続状態をデバッグ"""
    import os
    
    debug_info = {
        'configured_paths': {
            'excel': app.config['DEFAULT_EXCEL_PATH'],
            'history': app.config['HISTORY_EXCEL_PATH'],
            'seiban': app.config.get('SEIBAN_LIST_PATH', 'Not configured')
        },
        'path_checks': {}
    }
    
    # 各パスの存在確認
    for name, path in debug_info['configured_paths'].items():
        if path != 'Not configured':
            # os.path.exists での確認
            exists_os = os.path.exists(path)
            
            # Path オブジェクトでの確認
            path_obj = Path(path)
            exists_path = path_obj.exists()
            
            debug_info['path_checks'][name] = {
                'path': path,
                'os_exists': exists_os,
                'path_exists': exists_path,
                'is_file': path_obj.is_file() if exists_path else False,
                'parent_exists': path_obj.parent.exists()
            }
            
            # ファイルサイズの取得
            if exists_os:
                try:
                    stats = os.stat(path)
                    debug_info['path_checks'][name]['size_mb'] = round(stats.st_size / (1024 * 1024), 2)
                    debug_info['path_checks'][name]['readable'] = os.access(path, os.R_OK)
                except Exception as e:
                    debug_info['path_checks'][name]['error'] = str(e)
    
    # サーバーへの接続確認
    server_path = r'\\server3'
    debug_info['server_connection'] = {
        'path': server_path,
        'connected': os.path.exists(server_path)
    }
    
    if debug_info['server_connection']['connected']:
        try:
            shares = os.listdir(server_path)
            debug_info['server_connection']['shares'] = shares[:10]  # 最初の10個
        except Exception as e:
            debug_info['server_connection']['error'] = str(e)
    
    # Python環境情報
    debug_info['environment'] = {
        'python_version': sys.version,
        'platform': sys.platform,
        'cwd': os.getcwd(),
        'flask_env': app.config.get('ENV', 'not set')
    }
    
    return jsonify(debug_info)

@app.route('/api/check-network-file')
def check_network_file():
    """Check if network file is accessible"""
    file_info = check_network_file_access()
    return jsonify(file_info)

@app.route('/api/load-network-file', methods=['POST'])
def load_network_file():
    """Load file from network location"""
    try:
        # Try to copy network file to local cache
        cache_file, error = copy_network_file_to_local()
        
        if error:
            return jsonify({
                'success': False,
                'error': error,
                'suggest_upload': True
            }), 400
        
        # Load sheet names from cached file
        wb = load_workbook(cache_file, read_only=True)
        sheet_names = wb.sheetnames
        wb.close()
        
        return jsonify({
            'success': True,
            'filepath': cache_file,
            'sheet_names': sheet_names,
            'file_info': cached_file_info
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e),
            'suggest_upload': True
        }), 500

@app.route('/api/load-from-odbc', methods=['POST'])
def load_from_odbc_endpoint():
    """Load data directly from ODBC"""
    try:
        data = request.json
        seiban = data.get('seiban', '')
        
        if not seiban:
            return jsonify({'error': '製番を入力してください'}), 400
        
        # Load from ODBC
        odbc_data, error = load_from_odbc()
        
        if error:
            return jsonify({
                'success': False,
                'error': error
            }), 500
        
        # Process the data
        df1 = odbc_data['手配リスト']
        df2 = odbc_data['発注リスト']
        
        # Filter by seiban
        df1 = df1[df1['製番'].astype(str).str.startswith(seiban)]
        df2 = df2[df2['製番'].astype(str).str.startswith(seiban)]
        
        # Merge and save to database
        df_merged = process_excel_file_from_dataframes(df1, df2, seiban)
        save_to_database(df_merged, seiban)
        
        return jsonify({
            'success': True,
            'message': f'ODBC経由で{len(df_merged)}件のデータを処理しました'
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/get-system-status')
def get_system_status():
    """Get system status including cache and refresh info"""
    try:
        status = {
            'last_refresh': last_refresh_time.isoformat() if last_refresh_time else None,
            'cached_file': cached_file_info,
            'auto_refresh_enabled': refresh_thread and refresh_thread.is_alive() if refresh_thread else False,
            'refresh_interval_minutes': app.config['AUTO_REFRESH_INTERVAL'] / 60,
            'odbc_enabled': app.config['USE_ODBC'],
            'network_path': app.config['DEFAULT_EXCEL_PATH']
        }
        return jsonify(status)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/upload', methods=['POST'])
def upload_file():
    """Handle file upload and processing"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Get sheet names
        wb = load_workbook(filepath, read_only=True)
        sheet_names = wb.sheetnames
        wb.close()
        
        return jsonify({
            'success': True,
            'filepath': filepath,
            'sheet_names': sheet_names
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/orders')
def get_orders():
    """Get all active orders"""
    try:
        from sqlalchemy import func, case
        
        # ユニット名無しを最後にソート
        results = db.session.query(
            Order,
            func.count(OrderDetail.id).label('detail_count'),
            func.sum(OrderDetail.is_received.cast(db.Integer)).label('received_count'),
            func.max(OrderDetail.has_internal_processing.cast(db.Integer)).label('has_internal_processing')
        ).outerjoin(OrderDetail).filter(
            Order.is_archived == False
        ).group_by(Order.id).order_by(
            Order.seiban.desc(),
            case(
                (Order.unit == '', 1),  # ユニット名無しを最後に
                (Order.unit == None, 1),
                else_=0
            ),
            Order.unit
        ).all()
        
        orders = []
        for order, detail_count, received_count, has_internal in results:
            orders.append({
                'id': order.id,
                'seiban': order.seiban,
                'unit': order.unit or '',
                'product_name': order.product_name or '',
                'customer_abbr': order.customer_abbr or '',
                'memo2': order.memo2 or '',
                'location': order.location,
                'status': order.status,
                'remarks': order.remarks or '',
                'created_at': to_jst(order.created_at).strftime('%Y-%m-%d %H:%M:%S'),
                'updated_at': to_jst(order.updated_at).strftime('%Y-%m-%d %H:%M:%S'),
                'detail_count': detail_count or 0,
                'received_count': received_count or 0,
                'has_internal_processing': bool(has_internal)
            })
        
        return jsonify(orders)

    except Exception as e:
        import traceback
        print(f"Error getting orders: {str(e)}")
        print(traceback.print_exc())
        return jsonify({'error': str(e)}), 500

@app.route('/api/orders/gantt-data')
def get_gantt_data():
    """ガントチャート用に最適化された納期データを一括取得"""
    try:
        from sqlalchemy import func

        # 全アクティブ注文と詳細を一度に取得
        orders = Order.query.filter_by(is_archived=False).all()

        gantt_data = []
        for order in orders:
            # 各注文の納期を取得
            delivery_dates = [
                d.delivery_date for d in order.details
                if d.delivery_date and d.delivery_date.strip() and d.delivery_date != '-'
            ]

            if delivery_dates:
                # 進捗計算
                total_details = len(order.details)
                received_details = sum(1 for d in order.details if d.is_received)
                progress = (received_details / total_details * 100) if total_details > 0 else 0

                gantt_data.append({
                    'id': order.id,
                    'seiban': order.seiban,
                    'unit': order.unit or 'ユニット名無し',
                    'status': order.status,
                    'progress': progress,
                    'delivery_dates': delivery_dates  # 納期のリスト
                })

        return jsonify(gantt_data)

    except Exception as e:
        import traceback
        print(f"Error getting gantt data: {str(e)}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/archived-orders')
def get_archived_orders():
    """Get archived orders"""
    try:
        orders = Order.query.filter_by(is_archived=True).order_by(Order.archived_at.desc()).all()
        result = []
        for order in orders:
            # unitが空の場合は'-'として表示
            unit_display = order.unit if order.unit else '-'
            
            result.append({
                'id': order.id,
                'seiban': order.seiban,
                'unit': unit_display,
                'product_name': order.product_name or '',
                'status': order.status,
                'location': order.location,
                'remarks': order.remarks,
                'archived_at': order.archived_at.isoformat() if order.archived_at else None,
                'detail_count': len(order.details),
                'received_count': sum(1 for d in order.details if d.is_received)
            })
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/order/<int:order_id>/archive', methods=['POST'])
def archive_order(order_id):
    """Archive an order"""
    try:
        order = Order.query.get_or_404(order_id)
        
        # 納品完了チェック
        if order.status != '納品完了':
            return jsonify({
                'success': False,
                'error': '納品完了の注文のみアーカイブできます'
            }), 400
        
        order.is_archived = True
        order.archived_at = datetime.now(timezone.utc)
        db.session.commit()
        
        return jsonify({
            'success': True,
            'message': f'製番 {order.seiban} をアーカイブしました'
        })
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500

@app.route('/api/order/<int:order_id>/unarchive', methods=['POST'])
def unarchive_order(order_id):
    """Unarchive an order"""
    try:
        order = Order.query.get_or_404(order_id)
        
        order.is_archived = False
        order.archived_at = None
        db.session.commit()
        
        return jsonify({
            'success': True,
            'message': f'製番 {order.seiban} を復元しました'
        })
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500
    
@app.route('/receive/<int:order_id>')
def receive_page(order_id):
    """受入専用ページ（スマートフォン用）"""
    try:
        order = Order.query.get_or_404(order_id)

        # 🔥 検収データを読み込み
        delivery_dict = DeliveryUtils.load_delivery_data()

        # 詳細リストを取得
        details = []
        for detail in order.details:
            # 検収データから納入日・納入数を取得
            delivery_info = DeliveryUtils.get_delivery_info(detail.order_number, delivery_dict)

            details.append({
                'id': detail.id,
                'delivery_date': detail.delivery_date,
                'supplier': detail.supplier,
                'order_number': detail.order_number,
                'quantity': detail.quantity,
                'unit_measure': detail.unit_measure,
                'item_name': detail.item_name,
                'spec1': detail.spec1,
                'spec2': detail.spec2,
                'order_type': detail.order_type,
                'remarks': detail.remarks,
                'is_received': detail.is_received,
                'parent_id': detail.parent_id,
                'has_internal_processing': detail.has_internal_processing,
                # 🔥 検収データを追加
                'received_delivery_date': delivery_info.get('納入日', ''),
                'received_delivery_qty': delivery_info.get('納入数', 0)
            })
        
        # スマートフォン用のシンプルなHTMLを返す
        html = f"""
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>受入 - {order.seiban}</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            background: #f5f5f5;
            padding: 10px;
        }}
        .header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 10px;
            margin-bottom: 15px;
            text-align: center;
        }}
        .header h1 {{
            font-size: 1.5em;
            margin-bottom: 5px;
        }}
        .info-box {{
            background: white;
            padding: 15px;
            border-radius: 10px;
            margin-bottom: 15px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }}
        .info-row {{
            display: flex;
            justify-content: space-between;
            padding: 8px 0;
            border-bottom: 1px solid #eee;
        }}
        .info-row:last-child {{
            border-bottom: none;
        }}
        .label {{
            font-weight: bold;
            color: #666;
        }}
        .detail-item {{
            background: white;
            padding: 15px;
            border-radius: 10px;
            margin-bottom: 10px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }}
        .detail-item.received {{
            background: #d4edda;
            border-left: 4px solid #28a745;
        }}
        .detail-item.child {{
            margin-left: 20px;
            background: #f8f9fa;
        }}
        .detail-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 10px;
        }}
        .item-name {{
            font-weight: bold;
            font-size: 1.1em;
            color: #333;
        }}
        .btn {{
            padding: 10px 20px;
            border: none;
            border-radius: 5px;
            font-size: 1em;
            cursor: pointer;
            width: 100%;
            margin-top: 10px;
        }}
        .btn-primary {{
            background: #667eea;
            color: white;
        }}
        .btn-success {{
            background: #28a745;
            color: white;
        }}
        .btn-warning {{
            background: #ffc107;
            color: #212529;
        }}
        .status-badge {{
            padding: 5px 10px;
            border-radius: 15px;
            font-size: 0.85em;
            font-weight: bold;
        }}
        .badge-success {{
            background: #28a745;
            color: white;
        }}
        .badge-warning {{
            background: #ffc107;
            color: #212529;
        }}
        .detail-row {{
            padding: 5px 0;
            font-size: 0.9em;
        }}
        .toast {{
            position: fixed;
            top: 20px;
            right: 20px;
            background: #28a745;
            color: white;
            padding: 15px 20px;
            border-radius: 5px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.3);
            z-index: 1000;
            display: none;
        }}
        .toast.show {{
            display: block;
            animation: slideIn 0.3s;
        }}
        @keyframes slideIn {{
            from {{ transform: translateX(100%); }}
            to {{ transform: translateX(0); }}
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>📦 {order.seiban}</h1>
        <div>{order.unit or ''}</div>
    </div>
    
    <div class="info-box">
        <div class="info-row">
            <span class="label">ステータス:</span>
            <span>{order.status}</span>
        </div>
        <div class="info-row">
            <span class="label">品名:</span>
            <span>{order.product_name or ''}</span>
        </div>
        <div class="info-row">
            <span class="label">得意先:</span>
            <span>{order.customer_abbr or ''}</span>
        </div>
    </div>

    <!-- 🔥 場所・パレット番号編集セクション -->
    <div class="info-box" style="background: #e7f3ff; border-left: 4px solid #007bff;">
        <div style="margin-bottom: 10px; font-weight: bold; color: #004085;">📍 保管場所</div>
        <div style="margin-bottom: 10px;">
            <label style="display: block; font-size: 0.9em; color: #666; margin-bottom: 5px;">場所</label>
            <select id="floorInput" style="width: 100%; padding: 10px; border: 1px solid #007bff; border-radius: 5px; font-size: 0.95em;">
                <option value="">未設定</option>
                <option value="1F" {'selected' if order.floor == '1F' else ''}>1F</option>
                <option value="2F" {'selected' if order.floor == '2F' else ''}>2F</option>
            </select>
        </div>
        <div style="margin-bottom: 10px;">
            <label style="display: block; font-size: 0.9em; color: #666; margin-bottom: 5px;">パレット（棚）番号</label>
            <select id="palletInput" style="width: 100%; padding: 10px; border: 1px solid #007bff; border-radius: 5px; font-size: 0.95em;">
                <option value="">未設定</option>
                <!-- パレット -->
                <option value="P001" {'selected' if order.pallet_number == 'P001' else ''}>P001(パレット)</option>
                <option value="P002" {'selected' if order.pallet_number == 'P002' else ''}>P002(パレット)</option>
                <option value="P003" {'selected' if order.pallet_number == 'P003' else ''}>P003(パレット)</option>
                <option value="P004" {'selected' if order.pallet_number == 'P004' else ''}>P004(パレット)</option>
                <option value="P005" {'selected' if order.pallet_number == 'P005' else ''}>P005(パレット)</option>
                <option value="P006" {'selected' if order.pallet_number == 'P006' else ''}>P006(パレット)</option>
                <option value="P007" {'selected' if order.pallet_number == 'P007' else ''}>P007(パレット)</option>
                <option value="P008" {'selected' if order.pallet_number == 'P008' else ''}>P008(パレット)</option>
                <option value="P009" {'selected' if order.pallet_number == 'P009' else ''}>P009(パレット)</option>
                <option value="P010" {'selected' if order.pallet_number == 'P010' else ''}>P010(パレット)</option>
                <!-- 台車 -->
                <option value="D001" {'selected' if order.pallet_number == 'D001' else ''}>D001（台車）</option>
                <option value="D002" {'selected' if order.pallet_number == 'D002' else ''}>D002（台車）</option>
                <option value="D003" {'selected' if order.pallet_number == 'D003' else ''}>D003（台車）</option>
                <!-- 棚 -->
                <option value="T001" {'selected' if order.pallet_number == 'T001' else ''}>T001(棚)</option>
                <option value="T002" {'selected' if order.pallet_number == 'T002' else ''}>T002(棚)</option>
                <option value="T003" {'selected' if order.pallet_number == 'T003' else ''}>T003(棚)</option>
                <option value="T004" {'selected' if order.pallet_number == 'T004' else ''}>T004(棚)</option>
                <option value="T005" {'selected' if order.pallet_number == 'T005' else ''}>T005(棚)</option>
            </select>
        </div>
    </div>

    <!-- 🔥 備考セクション追加 -->
    <div class="info-box" style="background: #fff3cd;">
        <div style="margin-bottom: 10px; font-weight: bold; color: #856404;">📝 備考</div>
        <textarea id="remarksInput" style="width: 100%; min-height: 80px; padding: 10px; border: 1px solid #ffc107; border-radius: 5px; font-size: 0.95em; resize: vertical;">{order.remarks or ''}</textarea>
    </div>

    <!-- 📷 画像セクション -->
    <div class="info-box" style="background: #e8f5e9; border-left: 4px solid #4caf50;">
        <div style="margin-bottom: 10px; font-weight: bold; color: #2e7d32;">📷 画像</div>
        <div id="imagePreviewArea" style="margin: 10px 0; text-align: center;">
            <img id="orderImage" src="/api/order/{order.id}/image"
                 style="max-width: 100%; max-height: 250px; border-radius: 8px; display: none; cursor: pointer;"
                 onclick="openImageFullscreen(this.src)"
                 onerror="this.style.display='none'; document.getElementById('noImageText').style.display='block';"
                 onload="this.style.display='block'; document.getElementById('noImageText').style.display='none';">
            <p id="noImageText" style="color: #888; font-style: italic;">画像なし</p>
        </div>
        <div style="display: flex; gap: 8px; flex-wrap: wrap;">
            <label style="flex: 1; min-width: 120px;">
                <div class="btn btn-primary" style="text-align: center; margin: 0;">📤 ファイル選択</div>
                <input type="file" id="imageUploadFile" accept="image/*"
                       style="display: none;" onchange="uploadOrderImage({order.id})">
            </label>
            <label style="flex: 1; min-width: 120px;">
                <div class="btn btn-success" style="text-align: center; margin: 0;">📸 カメラ撮影</div>
                <input type="file" id="imageUploadCamera" accept="image/*" capture="environment"
                       style="display: none;" onchange="uploadOrderImageFromCamera({order.id})">
            </label>
        </div>
        <button class="btn" style="background: #dc3545; color: white; margin-top: 8px;" onclick="deleteOrderImage({order.id})">🗑️ 画像を削除</button>
        <p style="font-size: 0.75em; color: #666; margin-top: 8px; text-align: center;">※FullHD (1920x1080) に自動圧縮されます</p>
    </div>

    <!-- 🔥 統合保存ボタン -->
    <button class="btn btn-primary" onclick="saveAll()" style="width: 100%; padding: 15px; font-size: 1.1em; margin-top: 10px;">💾 保存</button>
    
    <h3 style="margin: 20px 0 10px 5px;">詳細リスト</h3>
    <div id="detailsList">
        {''.join([create_detail_html(d, details) for d in details if not d['parent_id']])}
    </div>
    
    <div id="toast" class="toast"></div>
    
    <script>
        // 🔥 ページ読み込み後にイベントリスナーを設定
        document.addEventListener('DOMContentLoaded', function() {{
            // CADリンクにイベントリスナーを追加
            document.querySelectorAll('.cad-link').forEach(function(link) {{
                link.addEventListener('click', function(e) {{
                    e.preventDefault();
                    const detailId = this.getAttribute('data-detail-id');
                    openCadFile(detailId);
                }});
            }});
        }});
        
        // CADファイルを開く関数
        async function openCadFile(detailId) {{
            try {{
                const response = await fetch('/api/open-cad/' + detailId);
                
                // JSONレスポンスの場合（ローカルで直接起動した場合）
                if (response.headers.get('content-type')?.includes('application/json')) {{
                    const data = await response.json();
                    
                    if (data.success) {{
                        if (data.opened_locally) {{
                            showToast('🔧 ' + data.message + ': ' + data.file_name, 'success');
                        }} else {{
                            showToast('📄 ' + data.message, 'success');
                        }}
                    }} else {{
                        showToast('❌ ' + data.error, 'error');
                    }}
                }} else {{
                    // ファイルダウンロード/表示の場合
                    const contentType = response.headers.get('content-type');
                    
                    if (contentType.includes('application/pdf')) {{
                        // PDFは新しいタブで開く
                        const url = `/api/open-cad/${{detailId}}`;
                        window.open(url, '_blank');
                        showToast('📄 PDF図面を開きました', 'success');
                    }} else {{
                        // MX2ファイルをダウンロード
                        const blob = await response.blob();
                        const url = window.URL.createObjectURL(blob);
                        const a = document.createElement('a');
                        a.href = url;
                        a.download = response.headers.get('content-disposition')?.split('filename=')[1] || 'file.mx2';
                        document.body.appendChild(a);
                        a.click();
                        a.remove();
                        window.URL.revokeObjectURL(url);
                        showToast('💾 MX2ファイルをダウンロードしました', 'info');
                    }}
                }}
                
            }} catch (error) {{
                showToast('❌ ファイルを開けませんでした: ' + error, 'error');
            }}
        }}
        
        // 🔥 統合保存関数（保管場所と備考を一度に保存）
        async function saveAll() {{
            const floor = document.getElementById('floorInput').value;
            const palletNumber = document.getElementById('palletInput').value;
            const remarks = document.getElementById('remarksInput').value;

            try {{
                const response = await fetch('/api/order/{order.id}/update', {{
                    method: 'POST',
                    headers: {{ 'Content-Type': 'application/json' }},
                    body: JSON.stringify({{
                        floor: floor,
                        pallet_number: palletNumber,
                        remarks: remarks
                    }})
                }});

                const data = await response.json();

                if (data.success) {{
                    showToast('✅ 保存しました', 'success');
                }} else {{
                    showToast('❌ エラー: ' + data.error, 'error');
                }}
            }} catch (error) {{
                showToast('❌ 保存エラー: ' + error, 'error');
            }}
        }}
        
        // 受入切替関数
        async function toggleReceive(detailId, setReceived, orderNumber, itemName, spec1, quantity) {{
            const action = setReceived ? '受入' : '受入取消';
            
            const confirmMessage = 'このアイテムを' + action + 'しますか？\\n\\n' +
                '発注番号: ' + (orderNumber || '未設定') + '\\n' +
                '品名: ' + (itemName || '未設定') + '\\n' +
                '仕様１: ' + (spec1 || '未設定') + '\\n' +
                '数量: ' + (quantity || '未設定');
            
            if (!confirm(confirmMessage)) {{
                return;
            }}
            
            try {{
                const response = await fetch('/api/detail/' + detailId + '/receive', {{
                    method: 'POST',
                    headers: {{ 'Content-Type': 'application/json' }},
                    body: JSON.stringify({{ is_received: setReceived }})
                }});
                
                if (response.ok) {{
                    showToast(setReceived ? '✅ 受入しました' : '⚠️ 受入を取り消しました');
                    setTimeout(function() {{ location.reload(); }}, 1000);
                }} else {{
                    const errorData = await response.json();
                    showToast('❌ エラー: ' + (errorData.error || '不明なエラー'), 'error');
                }}
            }} catch (error) {{
                showToast('❌ ネットワークエラー: ' + error, 'error');
                console.error('Error:', error);
            }}
        }}
        
        // トースト表示関数
        function showToast(message, type) {{
            type = type || 'success';
            const toast = document.getElementById('toast');
            toast.textContent = message;

            if (type === 'error') {{
                toast.style.background = '#dc3545';
            }} else if (type === 'info') {{
                toast.style.background = '#17a2b8';
            }} else {{
                toast.style.background = '#28a745';
            }}

            toast.classList.add('show');
            setTimeout(function() {{
                toast.classList.remove('show');
            }}, 3000);
        }}

        // 📷 画像アップロード（ファイル選択）
        async function uploadOrderImage(orderId) {{
            const fileInput = document.getElementById('imageUploadFile');
            await processImageUpload(orderId, fileInput);
        }}

        // 📸 画像アップロード（カメラ撮影）
        async function uploadOrderImageFromCamera(orderId) {{
            const fileInput = document.getElementById('imageUploadCamera');
            await processImageUpload(orderId, fileInput);
        }}

        // 画像アップロード処理
        async function processImageUpload(orderId, fileInput) {{
            const file = fileInput.files[0];

            if (!file) {{
                return;
            }}

            showToast('📤 アップロード中...', 'info');

            const formData = new FormData();
            formData.append('image', file);

            try {{
                const response = await fetch('/api/order/' + orderId + '/upload-image', {{
                    method: 'POST',
                    body: formData
                }});

                const data = await response.json();

                if (data.success) {{
                    showToast('✅ 画像をアップロードしました', 'success');
                    // 画像を再読み込み
                    const img = document.getElementById('orderImage');
                    img.src = '/api/order/' + orderId + '/image?t=' + Date.now();
                }} else {{
                    showToast('❌ エラー: ' + data.error, 'error');
                }}
            }} catch (error) {{
                showToast('❌ アップロードエラー: ' + error, 'error');
            }}

            // ファイル選択をリセット
            fileInput.value = '';
        }}

        // 🗑️ 画像削除
        async function deleteOrderImage(orderId) {{
            if (!confirm('画像を削除しますか？')) {{
                return;
            }}

            try {{
                const response = await fetch('/api/order/' + orderId + '/delete-image', {{
                    method: 'DELETE'
                }});

                const data = await response.json();

                if (data.success) {{
                    showToast('✅ 画像を削除しました', 'success');
                    document.getElementById('orderImage').style.display = 'none';
                    document.getElementById('noImageText').style.display = 'block';
                }} else {{
                    showToast('❌ エラー: ' + data.error, 'error');
                }}
            }} catch (error) {{
                showToast('❌ 削除エラー: ' + error, 'error');
            }}
        }}

        // 🔍 画像をフルスクリーンで表示
        function openImageFullscreen(src) {{
            const overlay = document.createElement('div');
            overlay.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.95); display: flex; justify-content: center; align-items: center; z-index: 10000; cursor: pointer;';

            const img = document.createElement('img');
            img.src = src;
            img.style.cssText = 'max-width: 95%; max-height: 95%; object-fit: contain;';

            const closeBtn = document.createElement('div');
            closeBtn.innerHTML = '✕';
            closeBtn.style.cssText = 'position: absolute; top: 15px; right: 20px; color: white; font-size: 2em; cursor: pointer;';
            closeBtn.onclick = function() {{ overlay.remove(); }};

            overlay.appendChild(img);
            overlay.appendChild(closeBtn);
            overlay.onclick = function(e) {{ if (e.target === overlay) overlay.remove(); }};
            document.body.appendChild(overlay);
        }}
    </script>
</body>
</html>
"""
        return html
    except Exception as e:
        import traceback
        traceback.print_exc()
        return f"<html><body><h1>エラー</h1><p>{str(e)}</p></body></html>", 500

@app.route('/api/order/<int:order_id>/update-remarks', methods=['POST'])
def update_order_remarks(order_id):
    """備考のみを更新するAPI"""
    try:
        data = request.json
        order = Order.query.get_or_404(order_id)
        
        order.remarks = data.get('remarks', '')
        
        db.session.commit()
        
        return jsonify({
            'success': True,
            'message': '備考を更新しました'
        })
    except Exception as e:
        db.session.rollback()
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

def create_detail_html(detail, all_details):
    """詳細アイテムのHTML生成"""
    is_received = detail['is_received']
    has_children = any(d['parent_id'] == detail['id'] for d in all_details)
    
    def escape_js(text):
        if not text:
            return ''
        return str(text).replace("'", "\\'").replace('"', '\\"').replace('\n', '\\n')
    
    order_number = escape_js(detail.get('order_number', ''))
    item_name = escape_js(detail.get('item_name', ''))
    spec1 = escape_js(detail.get('spec1', ''))
    quantity_str = f"{detail.get('quantity', '')} {detail.get('unit_measure', '')}".strip()
    
    # CAD図面情報を取得
    cad_info = get_cad_file_info(detail.get('spec1', ''))
    spec1_display = detail.get('spec1', '-')
    
    # 🔥 仕様1の表示（data属性を使用）
    if cad_info:
        if cad_info['has_pdf']:
            file_info = f"📄 PDF有 ({len(cad_info['pdf_files'])}件)"
            spec1_html = f'''
            <div>
                <strong>仕様１:</strong> 
                <a href="#" class="cad-link" data-detail-id="{detail['id']}" 
                   style="color: #007bff; text-decoration: underline; cursor: pointer;">
                    {spec1_display}
                </a>
                <span style="font-size: 0.8em; color: #28a745; margin-left: 5px;">{file_info}</span>
            </div>
            '''
        elif cad_info['has_mx2']:
            file_info = f"🔧 mx2のみ ({len(cad_info['mx2_files'])}件)"
            spec1_html = f'''
            <div>
                <strong>仕様１:</strong> 
                <a href="#" class="cad-link" data-detail-id="{detail['id']}" 
                   style="color: #007bff; text-decoration: underline; cursor: pointer;">
                    {spec1_display}
                </a>
                <div style="font-size: 0.75em; color: #856404; margin-top: 3px;">
                    {file_info}<br>
                    ⚠️ iCAD MX導入PCのみ(ダウンロードファイル)
                </div>
            </div>
            '''
        else:
            spec1_html = f'<div><strong>仕様１:</strong> {spec1_display}</div>'
    else:
        spec1_html = f'<div><strong>仕様１:</strong> {spec1_display}</div>'
    
    # 親アイテムのHTML
    html = f"""
    <div class="detail-item {'received' if is_received else ''}">
        <div class="detail-header">
            <div class="item-name">{detail['item_name'] or '-'}</div>
            {f'<span class="status-badge badge-success">✅ 受入済</span>' if is_received else ''}
        </div>
        
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 5px 15px; font-size: 0.9em; margin: 10px 0;">
            <div><strong>発注番号:</strong> {detail['order_number'] or '-'}</div>
            <div><strong>納期:</strong> {detail['delivery_date'] or '-'}</div>
            {spec1_html}
            <div><strong>数量:</strong> {detail['quantity'] or ''} {detail['unit_measure'] or ''}</div>
            <div><strong>仕入先:</strong> {detail['supplier'] or '-'}</div>
            <div><strong>手配区分:</strong> {detail['order_type'] or '-'}</div>
        </div>

        {f'<div style="background: #e3f2fd; padding: 8px; border-radius: 5px; margin: 10px 0; font-size: 0.85em; border-left: 3px solid #2196f3;"><strong>📦 検収:</strong> {detail.get("received_delivery_date", "-")} / {int(detail.get("received_delivery_qty", 0)) if detail.get("received_delivery_qty") else "-"}個</div>' if detail.get('received_delivery_qty') else ''}
        
        {f'<div style="background: #fff3cd; padding: 8px; border-radius: 5px; margin: 10px 0; font-size: 0.9em;"><strong>備考:</strong> {detail["remarks"]}</div>' if detail.get('remarks') else ''}
        {f'<span class="status-badge badge-warning">追加工有</span>' if has_children else ''}
        
        <button class="btn {'btn-warning' if is_received else 'btn-primary'}" 
                onclick="toggleReceive({detail['id']}, {str(not is_received).lower()}, '{order_number}', '{item_name}', '{spec1}', '{quantity_str}')">
            {('受入取消' if is_received else '受入')}
        </button>
    </div>
    """
    
    # 子アイテムも同様に処理
    children = [d for d in all_details if d['parent_id'] == detail['id']]
    for child in children:
        child_received = child['is_received']
        
        child_order_number = escape_js(child.get('order_number', ''))
        child_item_name = escape_js(child.get('item_name', ''))
        child_spec1 = escape_js(child.get('spec1', ''))
        child_quantity_str = f"{child.get('quantity', '')} {child.get('unit_measure', '')}".strip()
        
        child_cad_info = get_cad_file_info(child.get('spec1', ''))
        child_spec1_display = child.get('spec1', '-')
        
        if child_cad_info:
            if child_cad_info['has_pdf']:
                child_file_info = f"📄 PDF有"
                child_spec1_html = f'''
                <div>
                    <strong>仕様１:</strong> 
                    <a href="#" class="cad-link" data-detail-id="{child['id']}" 
                       style="color: #007bff; text-decoration: underline; cursor: pointer;">
                        {child_spec1_display}
                    </a>
                    <span style="font-size: 0.8em; color: #28a745; margin-left: 5px;">{child_file_info}</span>
                </div>
                '''
            elif child_cad_info['has_mx2']:
                child_file_info = f"🔧 mx2のみ"
                child_spec1_html = f'''
                <div>
                    <strong>仕様１:</strong> 
                    <a href="#" class="cad-link" data-detail-id="{child['id']}" 
                       style="color: #007bff; text-decoration: underline; cursor: pointer;">
                        {child_spec1_display}
                    </a>
                    <div style="font-size: 0.75em; color: #856404; margin-top: 3px;">
                        {child_file_info}<br>
                        ⚠️ iCAD MX導入PCのみ
                    </div>
                </div>
                '''
            else:
                child_spec1_html = f'<div><strong>仕様１:</strong> {child_spec1_display}</div>'
        else:
            child_spec1_html = f'<div><strong>仕様１:</strong> {child_spec1_display}</div>'
        
        html += f"""
    <div class="detail-item child {'received' if child_received else ''}">
        <div class="detail-header">
            <div class="item-name">└─ {child['item_name'] or '-'}</div>
            {f'<span class="status-badge badge-success">✅ 受入済</span>' if child_received else ''}
        </div>
        
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 5px 15px; font-size: 0.9em; margin: 10px 0;">
            <div><strong>発注番号:</strong> {child['order_number'] or '-'}</div>
            <div><strong>納期:</strong> {child['delivery_date'] or '-'}</div>
            {child_spec1_html}
            <div><strong>数量:</strong> {child['quantity'] or ''} {child['unit_measure'] or ''}</div>
            <div><strong>仕入先:</strong> {child['supplier'] or '-'}</div>
            <div><strong>手配区分:</strong> {child['order_type'] or '-'}</div>
        </div>
        
        <button class="btn {'btn-warning' if child_received else 'btn-primary'}" 
                onclick="toggleReceive({child['id']}, {str(not child_received).lower()}, '{child_order_number}', '{child_item_name}', '{child_spec1}', '{child_quantity_str}')">
            {'受入取消' if child_received else '受入'}
        </button>
    </div>
        """
    
    return html

@app.route('/api/order/<int:order_id>')
def get_order_details(order_id):
    """Get order details"""
    try:
        order = Order.query.get_or_404(order_id)

        # 🔥 検収データを読み込み
        delivery_dict = DeliveryUtils.load_delivery_data()

        details = []
        for detail in order.details:
            # 🔥 CAD図面情報を取得
            cad_info = get_cad_file_info(detail.spec1)

            # 🔥 検収データを取得
            delivery_info = DeliveryUtils.get_delivery_info(detail.order_number, delivery_dict)

            detail_dict = {
                'id': detail.id,
                'delivery_date': detail.delivery_date,
                'supplier': detail.supplier,
                'order_number': detail.order_number,
                'quantity': detail.quantity,
                'unit_measure': detail.unit_measure,
                'item_name': detail.item_name,
                'spec1': detail.spec1,
                'spec2': detail.spec2,
                'order_type': detail.order_type,
                'remarks': detail.remarks,
                'is_received': detail.is_received,
                'received_at': detail.received_at.isoformat() if detail.received_at else None,
                'has_internal_processing': detail.has_internal_processing,
                'parent_id': detail.parent_id,  # 🔥 親子関係を追加
                # 🔥 検収データを追加
                'received_delivery_date': delivery_info.get('納入日', ''),
                'received_delivery_qty': delivery_info.get('納入数', 0)
            }

            # 🔥 CAD情報を追加
            if cad_info:
                detail_dict['cad_info'] = {
                    'has_pdf': cad_info['has_pdf'],
                    'has_mx2': cad_info['has_mx2'],
                    'pdf_count': len(cad_info['pdf_files']),
                    'mx2_count': len(cad_info['mx2_files'])
                }
            else:
                detail_dict['cad_info'] = None

            details.append(detail_dict)
        
        return jsonify({
            'order': {
                'id': order.id,
                'seiban': order.seiban,
                'created_at': to_jst(order.created_at).strftime('%Y-%m-%d %H:%M:%S'),
                'updated_at': to_jst(order.updated_at).strftime('%Y-%m-%d %H:%M:%S'),
                'unit': order.unit or '',  # 空の場合は空文字列を返す
                'product_name': order.product_name or '',
                'customer_abbr': order.customer_abbr or '',
                'pallet_number': order.pallet_number,
                'floor': order.floor,
                'memo2': order.memo2 or '',
                'status': order.status,
                'location': order.location,
                'remarks': order.remarks
            },
            'details': details,
            'qr_code': generate_qr_code(f"{get_server_url()}/receive/{order.id}")
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/order/<int:order_id>/update', methods=['POST'])
def update_order(order_id):
    """Update order status, location, pallet and floor"""
    try:
        order = Order.query.get_or_404(order_id)
        data = request.json
        
        # 🔥 ステータス変更前の値を保存
        old_status = order.status
        new_status = data.get('status', order.status)
        
        # 更新処理
        if 'status' in data:
            order.status = data['status']
        if 'location' in data:
            order.location = data['location']
        if 'remarks' in data:
            order.remarks = data['remarks']
        if 'pallet_number' in data:
            order.pallet_number = data['pallet_number']
        if 'floor' in data:
            order.floor = data['floor']
        
        order.updated_at = datetime.now(timezone.utc)
        db.session.commit()
        
        # 🔥 納品完了になった場合の処理
        response_data = {
            'success': True,
            'message': 'Order updated successfully'
        }
        
        if old_status != '納品完了' and new_status == '納品完了':
            # メール送信を促す
            response_data['show_email_prompt'] = True
            response_data['order_info'] = {
                'seiban': order.seiban,
                'product_name': order.product_name or '',
                'customer_abbr': order.customer_abbr or '',
                'floor': order.floor or '',
                'pallet_number': order.pallet_number or '',
                'order_id': order.id
            }
        
        return jsonify(response_data)
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/detail/<int:detail_id>/toggle-receive', methods=['POST'])
def toggle_receive_detail(detail_id):
    """Toggle receive status for a detail item"""
    try:
        detail = OrderDetail.query.get_or_404(detail_id)

        # 現在の状態を取得
        was_received = detail.is_received
        action = 'unreceive' if was_received else 'receive'

        # ステータスをトグル
        detail.is_received = not was_received
        detail.received_at = None if not detail.is_received else datetime.now(timezone.utc)

        # クライアントIPを取得
        client_ip = request.headers.get('X-Forwarded-For', request.remote_addr)
        if ',' in client_ip:
            client_ip = client_ip.split(',')[0].strip()

        # 🔥 受入履歴を記録（発注番号がある場合のみ）
        if detail.order_number:
            if detail.is_received:
                ReceivedHistory.record_receive(
                    order_number=detail.order_number,
                    item_name=detail.item_name,
                    spec1=detail.spec1,
                    quantity=detail.quantity,
                    client_ip=client_ip
                )
            else:
                ReceivedHistory.record_cancel(
                    order_number=detail.order_number,
                    item_name=detail.item_name,
                    spec1=detail.spec1,
                    quantity=detail.quantity,
                    client_ip=client_ip
                )

        # 編集ログを記録
        log = EditLog(
            detail_id=detail_id,
            action=action,
            ip_address=request.remote_addr,
            user_agent=request.user_agent.string if request.user_agent else 'Unknown'
        )
        db.session.add(log)
        
        # 注文全体のステータスを更新
        order = detail.order
        all_received = all(d.is_received for d in order.details)
        any_received = any(d.is_received for d in order.details)
        
        if all_received:
            order.status = '納品完了'
        elif any_received:
            order.status = '納品中'
        else:
            order.status = '受入準備前'
        
        order.updated_at = datetime.now(timezone.utc)
        
        db.session.commit()

        # 🔥 Excelファイルを自動更新
        update_order_excel(order.id)
        
        # メッセージ作成（詳細情報を含む）
        if detail.is_received:
            message = f'✅ 受入完了\n'
        else:
            message = f'❌ 受入取消\n'
    
        # 社内加工の警告
        has_internal = False
        if detail.has_internal_processing:
            message += '\n\n⚠️ 注意: 社内加工/追加工品です'
            has_internal = True
        
        return jsonify({
            'success': True,
            'message': message,
            'is_received': detail.is_received,
            'order_status': order.status,
            'has_internal_processing': has_internal
        })
        
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500

@app.route('/api/detail/<int:detail_id>/receive', methods=['POST'])
def receive_detail(detail_id):
    """Mark detail as received (deprecated - use toggle-receive instead)"""
    return toggle_receive_detail(detail_id)

@app.route('/api/detail/<int:detail_id>/logs')
def get_detail_logs(detail_id):
    """Get edit logs for a specific detail"""
    try:
        logs = EditLog.query.filter_by(detail_id=detail_id).order_by(EditLog.timestamp.desc()).all()
        
        log_data = []
        for log in logs:
            log_data.append({
                'id': log.id,
                'action': '受入' if log.action == 'receive' else '取消',
                'ip_address': log.ip_address,
                'timestamp': log.timestamp.strftime('%Y-%m-%d %H:%M:%S'),
                'user_agent': log.user_agent[:50] if log.user_agent else 'Unknown'
            })
        
        return jsonify({
            'success': True,
            'logs': log_data,
            'total': len(log_data)
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/check-update')
def check_update():
    """ファイルの更新をチェック"""
    has_update, message = check_file_update()
    return jsonify({
        'has_update': has_update,
        'message': message,
        'current_info': cached_file_info
    })

@app.route('/api/load-history')
def load_history():
    """発行履歴を読み込み"""
    try:
        # パスをそのまま使用（既に正しいUNCパス形式）
        history_path = Path(app.config['HISTORY_EXCEL_PATH'])
        
        if not history_path.exists():
            return jsonify({
                'success': False,
                'error': f'履歴ファイルが見つかりません: {history_path}'
            }), 404
        
        # 履歴ファイルを読み込み
        df = pd.read_excel(str(history_path))
        
        # データを辞書形式に変換
        history_data = []
        for _, row in df.iterrows():
            filename = row.get('ファイル名', '')
            seiban = extract_seiban_from_filename(filename)
            
            history_data.append({
                'no': row.get('No.', ''),
                'issue_date': str(row.get('発行日', '')),
                'filename': filename,
                'size_kb': row.get('容量(KB)', 0),
                'seiban': seiban or '不明'
            })
        
        return jsonify({
            'success': True,
            'data': history_data,
            'total': len(history_data)
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

@app.route('/api/export/<int:order_id>')
def export_order(order_id):
    """注文データをExcelファイルとしてエクスポート"""
    try:
        order = Order.query.get(order_id)
        if not order:
            return jsonify({'success': False, 'error': '注文が見つかりません'}), 404
        
        wb = Workbook()
        ws = wb.active
        ws.title = f"{order.seiban}_{order.unit}"
        
        # ヘッダー
        headers = ['製番', 'ユニット', '品名', '仕様１', '仕様２', '数量', '単位', 
                   '納期', '手配区分', '発注番号', '仕入先', '仕入先CD', '備考', '検収日', '検収数']
        ws.append(headers)
        
        # データ
        for detail in order.details:
            row = [
                order.seiban,
                order.unit,
                detail.item_name,
                detail.spec1,
                detail.spec2,
                detail.quantity,
                detail.unit_measure,
                detail.delivery_date,
                detail.order_type,
                detail.order_number,
                detail.supplier,
                detail.supplier_cd,
                detail.remarks,
                detail.received_at.strftime('%Y-%m-%d %H:%M:%S') if detail.received_at else '',
                '受入済' if detail.is_received else '未受入'
            ]
            ws.append(row)
        
        # メモリ上に保存
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        wb.close()
        
        filename = f"{order.seiban}_{order.unit}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/order/<int:order_id>/delete', methods=['DELETE'])
def delete_order(order_id):
    """Delete an order and its details"""
    try:
        order = Order.query.get_or_404(order_id)
        
        # Delete all details first
        OrderDetail.query.filter_by(order_id=order_id).delete()
        
        # Delete the order
        db.session.delete(order)
        db.session.commit()
        
        return jsonify({'success': True, 'message': 'Order deleted successfully'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500

@app.route('/api/import-history', methods=['POST'])
def import_history():
    """Import order history from Excel file"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        # Read Excel file
        df = pd.read_excel(file)
        
        imported_count = 0
        for _, row in df.iterrows():
            filename = row.get('ファイル名', '')
            seiban = extract_seiban_from_filename(filename)
            
            if seiban:
                # Check if already exists
                existing = ProcessingHistory.query.filter_by(filename=filename).first()
                if not existing:
                    history = ProcessingHistory(
                        serial_no=row.get('No.', 0),
                        issue_date=pd.to_datetime(row.get('発行日', datetime.now())),
                        filename=filename,
                        file_size_kb=row.get('容量(KB)', 0),
                        seiban=seiban
                    )
                    db.session.add(history)
                    imported_count += 1
        
        db.session.commit()
        
        return jsonify({
            'success': True,
            'message': f'Imported {imported_count} records successfully'
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/search-by-spec1/<spec1>')
def search_by_spec1(spec1):
    """仕様１で検索（マージ済み + 未マージ対応）"""
    try:
        print(f"\n{'='*60}")
        print(f"🔍 仕様１検索: {spec1}")
        print(f"{'='*60}")
        
        # 1. マージ済みデータから検索
        details = OrderDetail.query.filter(
            OrderDetail.spec1.contains(spec1)
        ).all()
        
        print(f"  マージ済みデータ: {len(details)}件")
        
        result_list = []
        
        # マージ済みデータを結果に追加
        for detail in details:
            result_list.append({
                'id': detail.id,
                'order_id': detail.order_id,
                'seiban': detail.seiban,
                'unit': detail.material,
                'item_name': detail.item_name,
                'spec1': detail.spec1,
                'order_number': detail.order_number,
                'quantity': detail.quantity,
                'unit_measure': detail.unit_measure,
                'is_received': detail.is_received,
                'delivery_date': detail.delivery_date,
                'supplier': detail.supplier,
                'staff': '',
                'source': 'merged'
            })

        # 🔥 キャッシュが存在するか、または読み込みが必要かをフラグで返す
        cache_needs_loading = False
        if not order_all_cache_time:
            cache_needs_loading = True
        else:
            elapsed = (datetime.now(timezone.utc) - order_all_cache_time).total_seconds()
            if elapsed >= CACHE_EXPIRY_SECONDS:
                cache_needs_loading = True
        
        # 🔥 2. 未マージデータを検索（発注_ALLシートから）
        if not load_order_all_cache():
            print(f"  ⚠️  キャッシュ読み込み失敗")
        else:
            print(f"  キャッシュから検索中...")
            matched_count = 0
            
            # 全キャッシュから仕様１で検索
            for order_num, items in order_all_cache.items():
                for item in items:
                    item_spec1 = item.get('spec1', '')
                    
                    # 🔥 部分一致検索（大文字小文字無視）
                    if spec1.upper() in item_spec1.upper():
                        matched_count += 1
                        
                        # 🔥 マージ済みデータと重複していないか確認
                        is_duplicate = any(
                            r['seiban'] == item['seiban'] and 
                            r['spec1'] == item_spec1 and
                            r['item_name'] == item['item_name']
                            for r in result_list
                        )
                        
                        if not is_duplicate:
                            result_list.append({
                                'id': None,
                                'order_id': None,
                                'seiban': item['seiban'],
                                'unit': item['material'],
                                'item_name': item['item_name'],
                                'spec1': item_spec1,
                                'order_number': order_num,
                                'quantity': item['quantity'],
                                'unit_measure': item['unit_measure'],
                                'is_received': False,
                                'delivery_date': item['delivery_date'],
                                'supplier': item['supplier'],
                                'staff': item.get('staff', ''),
                                'source': 'order_all'
                            })
            
            print(f"  キャッシュから: {matched_count}件ヒット（重複除外後: {len([r for r in result_list if r['source'] == 'order_all'])}件）")
        
        print(f"  合計結果: {len(result_list)}件")
        print(f"{'='*60}\n")
        
        if not result_list:
            return jsonify({
                'found': False,
                'message': f'仕様１ "{spec1}" が見つかりません（マージ済み・未マージ両方を検索しました）',
                'cache_needs_loading': cache_needs_loading 
            }), 404
        
        return jsonify({
            'found': True,
            'count': len(result_list),
            'details': result_list,
            'has_unmerged': any(r['source'] == 'order_all' for r in result_list),
            'cache_needs_loading': cache_needs_loading
        })
    except Exception as e:
        print(f"❌ 検索エラー: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/search-by-purchase-order/<purchase_order_number>')
def search_by_purchase_order(purchase_order_number):
    """発注番号で検索（浮動小数点対応 + 未マージデータ対応）"""
    try:
        # 🔥 デバッグログ
        print(f"\n{'='*60}")
        print(f"🔍 発注番号検索: {purchase_order_number}")
        print(f"{'='*60}")
        
        # 浮動小数点数として入力された場合の対策
        search_number = purchase_order_number
        if '.' in search_number and search_number.endswith('.0'):
            search_number = search_number.replace('.0', '')
        
        print(f"  正規化後: {search_number}")

        cache_needs_loading = False
        if not order_all_cache_time:
            cache_needs_loading = True
        else:
            elapsed = (datetime.now(timezone.utc) - order_all_cache_time).total_seconds()
            if elapsed >= CACHE_EXPIRY_SECONDS:
                cache_needs_loading = True
        
        # 1. マージ済みデータから検索
        details = OrderDetail.query.filter(
            db.or_(
                OrderDetail.order_number == search_number,
                OrderDetail.order_number == purchase_order_number
            )
        ).all()
        
        print(f"  マージ済みデータ: {len(details)}件")
        
        result_list = []
        
        # マージ済みデータを結果に追加
        for detail in details:
            result_list.append({
                'id': detail.id,
                'order_id': detail.order_id,
                'seiban': detail.seiban,
                'unit': detail.material,
                'item_name': detail.item_name,
                'spec1': detail.spec1,
                'quantity': detail.quantity,
                'unit_measure': detail.unit_measure,
                'is_received': detail.is_received,
                'delivery_date': detail.delivery_date,
                'supplier': detail.supplier,
                'source': 'merged',
                'staff': '-'
            })
        
        # 2. 未マージデータを検索（発注_ALLシートから）
        cache_results = search_order_from_cache(search_number)
        
        if cache_results:
            print(f"  キャッシュから: {len(cache_results)}件")
            for item in cache_results:
                # マージ済みデータと重複していないか確認
                is_duplicate = any(
                    r['seiban'] == item['seiban'] and 
                    r['spec1'] == item['spec1'] and
                    r['item_name'] == item['item_name']
                    for r in result_list
                )
                
                if not is_duplicate:
                    result_list.append({
                        'id': None,
                        'order_id': None,
                        'seiban': item['seiban'],
                        'unit': item['material'],
                        'item_name': item['item_name'],
                        'spec1': item['spec1'],
                        'quantity': item['quantity'],
                        'unit_measure': item['unit_measure'],
                        'is_received': False,
                        'delivery_date': item['delivery_date'],
                        'staff': item.get('staff', ''),
                        'supplier': item['supplier'],
                        'source': 'order_all'
                    })
        else:
            print(f"  キャッシュから: 0件")
        
        print(f"  合計結果: {len(result_list)}件")
        print(f"{'='*60}\n")
        
        if not result_list:
            return jsonify({
                'found': False,
                'message': f'発注番号 {purchase_order_number} が見つかりません（マージ済み・未マージ両方を検索しました）',
                'cache_needs_loading': cache_needs_loading
            }), 404
        
        return jsonify({
            'found': True,
            'count': len(result_list),
            'details': result_list,
            'has_unmerged': any(r['source'] == 'order_all' for r in result_list),
            'cache_needs_loading': cache_needs_loading
        })
    except Exception as e:
        print(f"❌ 検索エラー: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/receive-by-purchase-order', methods=['POST'])
def receive_by_purchase_order():
    """発注番号で一括受入"""
    try:
        data = request.json
        purchase_order_number = data.get('purchase_order_number')
        
        if not purchase_order_number:
            return jsonify({'error': '発注番号を入力してください'}), 400
        
        # 発注番号に紐づく全アイテムを取得
        details = OrderDetail.query.filter_by(
            order_number=purchase_order_number
        ).all()
        
        if not details:
            return jsonify({
                'error': f'発注番号 {purchase_order_number} が見つかりません'
            }), 404
        
        # 社内加工チェック
        has_internal_processing = False
        received_count = 0
        
        for detail in details:
            if not detail.is_received:
                detail.is_received = True
                detail.received_at = datetime.now(timezone.utc)
                received_count += 1
                
                if detail.has_internal_processing:
                    has_internal_processing = True
                
                # ログを記録
                log = EditLog(
                    detail_id=detail.id,
                    action='receive',
                    ip_address=request.remote_addr,
                    user_agent=request.user_agent.string if request.user_agent else 'Unknown'
                )
                db.session.add(log)
        
        # 注文のステータス更新
        orders_to_update = set()
        for detail in details:
            orders_to_update.add(detail.order)
        
        for order in orders_to_update:
            all_received = all(d.is_received for d in order.details)
            if all_received:
                order.status = '納品完了'
            elif any(d.is_received for d in order.details):
                order.status = '納品中'
        
        db.session.commit()
        
        message = f'発注番号 {purchase_order_number} の {received_count} 件を受入しました'
        if has_internal_processing:
            message += '\n⚠️ 注意: 社内加工/追加工品が含まれています'
        
        return jsonify({
            'success': True,
            'message': message,
            'received_count': received_count,
            'has_internal_processing': has_internal_processing
        })
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500

@app.route('/api/purchase-order-stats')
def purchase_order_stats():
    """発注番号の統計情報を取得"""
    try:
        from sqlalchemy import func
        
        stats = db.session.query(
            OrderDetail.order_number,
            func.count(OrderDetail.id).label('total_items'),
            func.sum(OrderDetail.quantity).label('total_quantity'),
            func.sum(OrderDetail.is_received.cast(db.Integer)).label('received_items'),
            func.min(OrderDetail.spec1).label('spec1'),
            func.min(OrderDetail.item_name).label('item_name')  # 品名を追加
        ).filter(
            OrderDetail.order_number != '',
            OrderDetail.order_number != None
        ).group_by(
            OrderDetail.order_number
        ).all()
        
        result = []
        for stat in stats:
            order_number = stat.order_number
            if order_number and '.0' in order_number:
                order_number = order_number.replace('.0', '')
            
            completion_rate = (stat.received_items / stat.total_items * 100) if stat.total_items > 0 else 0
            result.append({
                'purchase_order_number': order_number,
                'total_items': stat.total_items,
                'total_quantity': stat.total_quantity or 0,
                'received_items': stat.received_items or 0,
                'completion_rate': round(completion_rate, 1),
                'spec1': stat.spec1 or '-',
                'item_name': stat.item_name or '-' 
            })
        
        result.sort(key=lambda x: x['completion_rate'])
        
        return jsonify({
            'total_orders': len(result),
            'stats': result
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
@app.route('/api/export-seiban/<seiban>')
def export_seiban(seiban):
    """製番全体をExcelファイルとしてエクスポート"""
    try:
        orders = Order.query.filter_by(seiban=seiban).all()
        
        if not orders:
            return jsonify({'success': False, 'error': '製番が見つかりません'}), 404
        
        wb = Workbook()
        ws = wb.active
        ws.title = seiban
        
        # ヘッダー
        headers = ['製番', 'ユニット', '品名', '仕様１', '仕様２', '数量', '単位', 
                   '納期', '手配区分', '発注番号', '仕入先', '仕入先CD', '備考', '検収日', '検収数']
        ws.append(headers)
        
        # 全ユニットのデータを出力
        for order in orders:
            for detail in order.details:
                row = [
                    order.seiban,
                    order.unit,
                    detail.item_name,
                    detail.spec1,
                    detail.spec2,
                    detail.quantity,
                    detail.unit_measure,
                    detail.delivery_date,
                    detail.order_type,
                    detail.order_number,
                    detail.supplier,
                    detail.supplier_cd,
                    detail.remarks,
                    detail.received_at.strftime('%Y-%m-%d %H:%M:%S') if detail.received_at else '',
                    '受入済' if detail.is_received else '未受入'
                ]
                ws.append(row)
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        wb.close()
        
        filename = f"{seiban}_all_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500
    
@app.route('/api/orders/delete-multiple', methods=['POST'])
def delete_multiple_orders():
    """複数の注文を一括削除"""
    try:
        data = request.json
        order_ids = data.get('order_ids', [])
        
        if not order_ids:
            return jsonify({'error': '削除する注文を選択してください'}), 400
        
        # 削除対象を取得
        orders = Order.query.filter(Order.id.in_(order_ids)).all()
        
        if not orders:
            return jsonify({'error': '対象の注文が見つかりません'}), 404
        
        deleted_count = 0
        seibans = []
        
        for order in orders:
            # 詳細を削除
            OrderDetail.query.filter_by(order_id=order.id).delete()
            
            # 注文を削除
            seibans.append(f"{order.seiban}({order.unit or '-'})")
            db.session.delete(order)
            deleted_count += 1
        
        db.session.commit()
        
        return jsonify({
            'success': True,
            'message': f'{deleted_count}件の注文を削除しました',
            'deleted': seibans
        })
        
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500
    
@app.route('/api/pallets/list')
def get_pallets_list():
    """パレット一覧を取得（品名付き）"""
    try:
        from sqlalchemy import func
        
        # 製番一覧表から品名を読み込み
        seiban_info = load_seiban_info()
        
        # パレット番号でグループ化して取得
        pallets = db.session.query(
            Order.pallet_number,
            Order.floor,
            func.count(Order.id).label('order_count')
        ).filter(
            Order.pallet_number != None,
            Order.pallet_number != '',
            Order.is_archived == False
        ).group_by(
            Order.pallet_number,
            Order.floor
        ).all()
        
        result = []
        for pallet in pallets:
            # そのパレットに含まれる注文を取得
            orders = Order.query.filter_by(
                pallet_number=pallet.pallet_number,
                is_archived=False
            ).all()
            
            orders_data = []
            for order in orders:
                # 製番一覧表から品名を取得（既にDBに品名がある場合はそちらを優先）
                product_name = order.product_name
                if not product_name and order.seiban in seiban_info:
                    product_name = seiban_info[order.seiban].get('product_name', '')
                
                # 製番一覧表から得意先略称を取得（既にDBに得意先略称がある場合はそちらを優先）
                customer_abbr = order.customer_abbr
                if not customer_abbr and order.seiban in seiban_info:
                    customer_abbr = seiban_info[order.seiban].get('customer_abbr', '')
                
                orders_data.append({
                    'id': order.id,
                    'seiban': order.seiban,
                    'unit': order.unit,
                    'status': order.status,
                    'product_name': product_name,
                    'customer_abbr': customer_abbr  # ← 製番一覧表からも取得するように修正
                })
            
            result.append({
                'pallet_number': pallet.pallet_number,
                'floor': pallet.floor,
                'order_count': pallet.order_count,
                'orders': orders_data
            })
        
        # パレット番号でソート
        result.sort(key=lambda x: x['pallet_number'])
        
        return jsonify({
            'success': True,
            'pallets': result,
            'total_pallets': len(result)
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/pallets/search')
def search_pallet():
    """製番、品名、または得意先略称でパレットを検索"""
    try:
        search_query = request.args.get('query', '')  # queryパラメータに統一
        
        if not search_query:
            return jsonify({'error': '検索キーワードを入力してください'}), 400
        
        # 製番一覧表から品名と得意先略称を読み込み
        seiban_info = load_seiban_info()  # ← load_seiban_data() から変更
        
        # 1. 製番で検索（部分一致）
        orders_by_seiban = Order.query.filter(
            Order.seiban.like(f'%{search_query}%'),
            Order.is_archived == False
        ).all()
        
        # 2. DBの得意先略称で検索（部分一致）
        orders_by_customer = Order.query.filter(
            Order.customer_abbr.like(f'%{search_query}%'),
            Order.is_archived == False
        ).all()
        
        # 3. 製番一覧表から品名と得意先略称で検索
        matching_seibans = []
        for seiban, info in seiban_info.items():
            product_name = info.get('product_name', '')
            customer_abbr = info.get('customer_abbr', '')
            
            # 品名または得意先略称に検索キーワードが含まれる場合
            if (search_query.lower() in product_name.lower() or 
                search_query.lower() in customer_abbr.lower()):
                matching_seibans.append(seiban)
        
        # 品名または得意先略称で見つかった製番の注文を取得
        orders_by_info = []
        if matching_seibans:
            orders_by_info = Order.query.filter(
                Order.seiban.in_(matching_seibans),
                Order.is_archived == False
            ).all()
        
        # 重複を除いて結合
        all_orders = list(set(orders_by_seiban + orders_by_customer + orders_by_info))
        
        if not all_orders:
            return jsonify({
                'success': False,
                'error': f'「{search_query}」に一致する製番、品名、または得意先略称が見つかりません'
            }), 404
        
        results = []
        for order in all_orders:
            # 製番一覧表から品名を取得（既にDBに品名がある場合はそちらを優先）
            product_name = order.product_name
            if not product_name and order.seiban in seiban_info:
                product_name = seiban_info[order.seiban].get('product_name', '')
            
            # 製番一覧表から得意先略称を取得（既にDBに得意先略称がある場合はそちらを優先）
            customer_abbr = order.customer_abbr
            if not customer_abbr and order.seiban in seiban_info:
                customer_abbr = seiban_info[order.seiban].get('customer_abbr', '')
            
            results.append({
                'id': order.id,
                'seiban': order.seiban,
                'unit': order.unit,
                'status': order.status,
                'pallet_number': order.pallet_number or '未設定',
                'floor': order.floor or '未設定',
                'product_name': product_name,
                'customer_abbr': customer_abbr  # ← 製番一覧表からも取得した値を使用
            })
        
        return jsonify({
            'success': True,
            'results': results
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500



@app.route('/api/pallets/<pallet_number>/label')
def get_pallet_label(pallet_number):
    """パレットラベルを生成して返す"""
    try:
        # そのパレット番号の注文を取得
        orders = Order.query.filter_by(
            pallet_number=pallet_number,
            is_archived=False
        ).all()
        
        if not orders:
            return jsonify({'error': 'パレットが見つかりません'}), 404
        
        # QRコード生成
        qr = qrcode.QRCode(version=1, box_size=10, border=4)
        qr.add_data(f'PALLET:{pallet_number}')
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        
        # 画像をBase64エンコード
        buffered = BytesIO()
        img.save(buffered, format="PNG")
        qr_base64 = base64.b64encode(buffered.getvalue()).decode()
        
        # 注文情報をまとめる
        orders_info = []
        for order in orders:
            orders_info.append({
                'seiban': order.seiban,
                'unit': order.unit,
                'status': order.status,
                'product_name': order.product_name
            })
        
        return jsonify({
            'success': True,
            'pallet_number': pallet_number,
            'floor': orders[0].floor if orders else '未設定',
            'qr_code': qr_base64,
            'orders': orders_info,
            'order_count': len(orders)
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/pallets/stats')
def get_pallet_stats():
    """パレット統計情報を取得"""
    try:
        from sqlalchemy import func
        
        # パレット別の統計
        stats = db.session.query(
            Order.pallet_number,
            Order.floor,
            func.count(Order.id).label('total_orders'),
            func.sum(case((Order.status == '納品完了', 1), else_=0)).label('completed_orders'),
            func.sum(case((Order.status == '納品中', 1), else_=0)).label('in_progress_orders')
        ).filter(
            Order.pallet_number != None,
            Order.pallet_number != '',
            Order.is_archived == False
        ).group_by(
            Order.pallet_number,
            Order.floor
        ).all()
        
        result = []
        for stat in stats:
            completion_rate = (stat.completed_orders / stat.total_orders * 100) if stat.total_orders > 0 else 0
            result.append({
                'pallet_number': stat.pallet_number,
                'floor': stat.floor or '未設定',
                'total_orders': stat.total_orders,
                'completed_orders': stat.completed_orders,
                'in_progress_orders': stat.in_progress_orders,
                'completion_rate': round(completion_rate, 1)
            })
        
        result.sort(key=lambda x: x['pallet_number'])
        
        return jsonify({
            'success': True,
            'stats': result
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

import subprocess
import os

@app.route('/api/open-cad/<int:detail_id>')
def open_cad_file(detail_id):
    """CADファイルを開く（ローカルは直接起動、リモートはダウンロード）"""
    try:
        detail = OrderDetail.query.get_or_404(detail_id)
        cad_info = get_cad_file_info(detail.spec1)
        
        if not cad_info:
            return jsonify({
                'success': False,
                'error': 'CADファイルが見つかりません'
            }), 404
        
        # 優先順位: PDF → mx2
        if cad_info['has_pdf']:
            file_path = cad_info['pdf_files'][0]
            file_type = 'PDF'
            mimetype = 'application/pdf'
        elif cad_info['has_mx2']:
            file_path = cad_info['mx2_files'][0]
            file_type = 'MX2'
            mimetype = 'application/octet-stream'
        else:
            return jsonify({
                'success': False,
                'error': 'ファイルが見つかりません'
            }), 404
        
        # 🔥 リクエスト元のIPアドレスを取得
        client_ip = request.remote_addr
        
        # 🔥 ローカルホストまたはサーバー自身からのアクセスか判定
        is_local = client_ip in ['127.0.0.1', '::1', 'localhost'] or \
                   client_ip == request.host.split(':')[0]  # サーバーのIPアドレス
        
        print(f"🔍 CADファイルアクセス: IP={client_ip}, ローカル={is_local}, ファイル={file_type}")
        
        # 🔥 MX2ファイルかつローカルアクセスの場合のみ直接起動
        if file_type == 'MX2' and is_local:
            try:
                # サーバー側でiCAD MXを起動
                os.startfile(file_path)
                
                return jsonify({
                    'success': True,
                    'file_type': file_type,
                    'file_name': os.path.basename(file_path),
                    'message': 'iCAD MXで図面を開きました',
                    'opened_locally': True
                })
            except Exception as e:
                return jsonify({
                    'success': False,
                    'error': f'ファイルを開けませんでした: {str(e)}'
                }), 500
        
        # 🔥 それ以外（リモートアクセスまたはPDF）はダウンロード/表示
        try:
            return send_file(
                file_path,
                mimetype=mimetype,
                as_attachment=(file_type == 'MX2'),  # MX2はダウンロード、PDFは表示
                download_name=os.path.basename(file_path)
            )
            
        except Exception as e:
            return jsonify({
                'success': False,
                'error': f'ファイルを送信できませんでした: {str(e)}'
            }), 500
        
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


# ==================== 画像アップロード/表示機能 ====================

# 画像保存ディレクトリ
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads', 'images')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# FullHD解像度
FULLHD_WIDTH = 1920
FULLHD_HEIGHT = 1080


def compress_to_fullhd(image_data):
    """画像をFullHD（1920x1080）以下に圧縮"""
    img = Image.open(io.BytesIO(image_data))

    # EXIF情報に基づいて回転を修正
    try:
        from PIL import ExifTags
        for orientation in ExifTags.TAGS.keys():
            if ExifTags.TAGS[orientation] == 'Orientation':
                break
        exif = img._getexif()
        if exif is not None:
            orientation_value = exif.get(orientation)
            if orientation_value == 3:
                img = img.rotate(180, expand=True)
            elif orientation_value == 6:
                img = img.rotate(270, expand=True)
            elif orientation_value == 8:
                img = img.rotate(90, expand=True)
    except (AttributeError, KeyError, IndexError):
        pass

    # 元のサイズ
    original_width, original_height = img.size

    # リサイズが必要かチェック
    if original_width <= FULLHD_WIDTH and original_height <= FULLHD_HEIGHT:
        # リサイズ不要、でもJPEGに変換して圧縮
        output = io.BytesIO()
        if img.mode in ('RGBA', 'P'):
            img = img.convert('RGB')
        img.save(output, format='JPEG', quality=85, optimize=True)
        return output.getvalue()

    # アスペクト比を維持してリサイズ
    ratio = min(FULLHD_WIDTH / original_width, FULLHD_HEIGHT / original_height)
    new_width = int(original_width * ratio)
    new_height = int(original_height * ratio)

    # リサイズ
    img = img.resize((new_width, new_height), Image.LANCZOS)

    # JPEG形式で保存
    output = io.BytesIO()
    if img.mode in ('RGBA', 'P'):
        img = img.convert('RGB')
    img.save(output, format='JPEG', quality=85, optimize=True)

    return output.getvalue()


@app.route('/api/order/<int:order_id>/upload-image', methods=['POST'])
def upload_order_image(order_id):
    """注文に画像をアップロード（FullHD圧縮）"""
    try:
        order = Order.query.get(order_id)
        if not order:
            return jsonify({'success': False, 'error': '注文が見つかりません'}), 404

        if 'image' not in request.files:
            return jsonify({'success': False, 'error': '画像ファイルがありません'}), 400

        file = request.files['image']
        if file.filename == '':
            return jsonify({'success': False, 'error': 'ファイルが選択されていません'}), 400

        # 拡張子チェック
        allowed_extensions = {'png', 'jpg', 'jpeg', 'gif', 'webp'}
        ext = file.filename.rsplit('.', 1)[-1].lower() if '.' in file.filename else ''
        if ext not in allowed_extensions:
            return jsonify({'success': False, 'error': '許可されていないファイル形式です'}), 400

        # 画像データを読み込み
        image_data = file.read()

        # FullHDに圧縮
        compressed_data = compress_to_fullhd(image_data)

        # ファイル名を生成（order_id + タイムスタンプ）
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"order_{order_id}_{timestamp}.jpg"
        filepath = os.path.join(UPLOAD_FOLDER, filename)

        # 古い画像があれば削除
        if order.image_path and os.path.exists(order.image_path):
            try:
                os.remove(order.image_path)
            except:
                pass

        # 保存
        with open(filepath, 'wb') as f:
            f.write(compressed_data)

        # DBに保存
        order.image_path = filepath
        db.session.commit()

        return jsonify({
            'success': True,
            'message': '画像をアップロードしました',
            'image_url': f'/api/order/{order_id}/image'
        })

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/order/<int:order_id>/image')
def get_order_image(order_id):
    """注文の画像を取得"""
    try:
        order = Order.query.get(order_id)
        if not order:
            return jsonify({'error': '注文が見つかりません'}), 404

        if not order.image_path or not os.path.exists(order.image_path):
            return jsonify({'error': '画像がありません'}), 404

        return send_file(
            order.image_path,
            mimetype='image/jpeg'
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/order/<int:order_id>/delete-image', methods=['DELETE'])
def delete_order_image(order_id):
    """注文の画像を削除"""
    try:
        order = Order.query.get(order_id)
        if not order:
            return jsonify({'success': False, 'error': '注文が見つかりません'}), 404

        if order.image_path and os.path.exists(order.image_path):
            try:
                os.remove(order.image_path)
            except:
                pass

        order.image_path = None
        db.session.commit()

        return jsonify({'success': True, 'message': '画像を削除しました'})

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


if __name__ == '__main__':

    # 設定を取得
    config_obj = get_config()
    
    # SSL/TLSコンテキストを取得
    ssl_context = None
    if hasattr(config_obj, 'USE_HTTPS') and config_obj.USE_HTTPS:
        from config import get_ssl_context
        ssl_context = get_ssl_context(config_obj)
        
        if ssl_context:
            print("🔒 HTTPS有効")
            if ssl_context == 'adhoc':
                print("⚠️  自己署名証明書を使用（開発用）")
                print("📱 QRコードスキャナーが使用可能です")
        else:
            print("⚠️  HTTPS無効 - QRコードスキャナーは使用できません")
    
    # サーバー起動
    app.run(
        debug=config_obj.DEBUG,
        host='0.0.0.0',
        port=8080,
        ssl_context=ssl_context
    )