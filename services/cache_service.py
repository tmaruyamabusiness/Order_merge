"""
キャッシュ・マスタデータ管理 - cache_service.py
発注_ALLシートのメモリキャッシュ、製番一覧表の読み込み
"""
import shutil
from datetime import datetime, timezone
from pathlib import Path
import pandas as pd
from flask import current_app
from utils import DataUtils

# グローバルキャッシュ変数
order_all_cache = {}
order_all_cache_time = None
CACHE_EXPIRY_SECONDS = 28800  # 8時間

last_refresh_time = None
cached_file_info = {}


def check_network_file_access():
    """ネットワークファイルのアクセスチェック"""
    try:
        network_path = Path(current_app.config['DEFAULT_EXCEL_PATH'])
        print(f"Checking path: {network_path}")

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
            print(f"File info: {result}")
            return result
        else:
            print(f"File not found: {network_path}")
            return {
                'accessible': False,
                'error': f'ファイルが見つかりません: {network_path}'
            }
    except Exception as e:
        print(f"Access error: {str(e)}")
        return {
            'accessible': False,
            'error': f'アクセスエラー: {str(e)}'
        }


def copy_network_file_to_local():
    """ネットワークファイルをローカルキャッシュにコピー"""
    global last_refresh_time, cached_file_info
    try:
        network_path = Path(current_app.config['DEFAULT_EXCEL_PATH'])
        if not network_path.exists():
            return None, f"ネットワークファイルが見つかりません: {network_path}"

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        cache_filename = f'cache/cached_{timestamp}_手配発注_ALL.xlsx'

        shutil.copy2(str(network_path), cache_filename)

        last_refresh_time = datetime.now()
        cached_file_info = check_network_file_access()

        return cache_filename, None
    except Exception as e:
        return None, f"コピーエラー: {str(e)}"


def check_file_update():
    """ファイルの更新をチェック"""
    global cached_file_info

    try:
        current_info = check_network_file_access()
        if not current_info['accessible']:
            return False, None

        if not cached_file_info:
            return True, "初回読み込み"

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
        if order_all_cache_time:
            elapsed = (datetime.now(timezone.utc) - order_all_cache_time).total_seconds()
            if elapsed < CACHE_EXPIRY_SECONDS:
                print(f"✅ キャッシュ有効（残り{int(CACHE_EXPIRY_SECONDS - elapsed)}秒）")
                return True

        print("🔄 発注_ALLシートを読み込み中...")

        excel_path = Path(current_app.config['DEFAULT_EXCEL_PATH'])
        if not excel_path.exists():
            print(f"❌ ファイルが見つかりません: {excel_path}")
            return False

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

        order_all_cache.clear()

        sample_keys = []

        for idx, row in df.iterrows():
            order_num = DataUtils.safe_str(row.get('発注番号', ''))
            if not order_num or order_num == '':
                continue

            if len(sample_keys) < 10:
                sample_keys.append(f"元の値: '{order_num}'")

            order_num = DataUtils.normalize_order_number(order_num)

            if len(sample_keys) < 20:
                sample_keys.append(f"正規化後: '{order_num}'")

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

    search_key = DataUtils.normalize_order_number(order_number)

    print(f"🔍 検索: 元の値='{order_number}' → 正規化後='{search_key}'")

    if search_key in order_all_cache:
        print(f"  ✅ キャッシュHIT: {len(order_all_cache[search_key])}件")
        return order_all_cache[search_key]
    else:
        print(f"  ❌ キャッシュMISS")
        print(f"  🔎 類似キー検索中...")
        similar_keys = [k for k in list(order_all_cache.keys())[:50] if search_key in k or k in search_key]
        if similar_keys:
            print(f"    類似キー（最大5件）: {similar_keys[:5]}")
        else:
            print(f"    類似キーなし")

    return None


def load_seiban_info():
    """製番一覧表から品名、得意先略称、メモ2を読み込む"""
    try:
        seiban_file = current_app.config.get('SEIBAN_LIST_PATH', r'\\server3\share-data\Document\Acrossデータ\製番一覧表.xlsx')
        seiban_path = Path(seiban_file)

        if not seiban_path.exists():
            print(f"製番一覧表が見つかりません: {seiban_path}")
            return {}

        df = pd.read_excel(str(seiban_path), sheet_name='製番')

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
