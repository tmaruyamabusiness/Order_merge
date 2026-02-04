"""
Across DB 直接クエリモジュール
V_D系ビュー（読み取り専用）への安全なアクセスを提供
"""

import pyodbc
from decimal import Decimal
from datetime import datetime, date


# 利用可能なビュー一覧
AVAILABLE_VIEWS = {
    'V_D発注': '発注データ（発注番号・仕入先・納期など）',
    'V_D発注残': '発注残データ（納入済数を含む）',
    'V_D手配リスト': '手配リスト（BOM・部品表）',
    'V_D仕入': '仕入データ（納入実績）',
}

# DSN接続設定
DSN_CONNECTION = "DSN=Across;"


def get_connection():
    """読み取り専用でAcross DBに接続"""
    conn = pyodbc.connect(DSN_CONNECTION, readonly=True)
    cursor = conn.cursor()
    cursor.execute("USE acrossDB;")
    return conn, cursor


def format_value(value):
    """値を表示用に整形"""
    if value is None:
        return None
    elif isinstance(value, Decimal):
        if value == int(value):
            return int(value)
        return float(value)
    elif isinstance(value, (datetime, date)):
        return value.strftime('%Y-%m-%d')
    else:
        return str(value).strip()


def query_view(view_name, where_clause=None, params=None, limit=100):
    """
    指定ビューからデータ取得

    Args:
        view_name: ビュー名 (AVAILABLE_VIEWS のキー)
        where_clause: WHERE条件文字列 (例: "製番 = ?")
        params: パラメータリスト (例: ['MHT0620'])
        limit: 最大取得件数

    Returns:
        dict: { columns: [...], rows: [[...], ...], count: int }
    """
    if view_name not in AVAILABLE_VIEWS:
        raise ValueError(f"不正なビュー名: {view_name}")

    conn = None
    try:
        conn, cursor = get_connection()

        sql = f"SELECT TOP {int(limit)} * FROM dbo.[{view_name}]"
        if where_clause:
            sql += f" WHERE {where_clause}"

        cursor.execute(sql, params or [])

        columns = [col[0] for col in cursor.description]
        raw_rows = cursor.fetchall()

        rows = []
        for row in raw_rows:
            rows.append([format_value(v) for v in row])

        return {
            'columns': columns,
            'rows': rows,
            'count': len(rows)
        }
    finally:
        if conn:
            conn.close()


def search_order(order_number):
    """
    発注番号で検索（ゼロパディング対応）

    Args:
        order_number: 発注番号 (例: '89074' or '00089074')

    Returns:
        dict: 検索結果
    """
    # 8桁ゼロパディング
    padded = str(order_number).strip().zfill(8)
    return query_view('V_D発注', '発注番号 = ?', [padded])


def search_by_seiban(view_name, seiban):
    """
    製番で検索

    Args:
        view_name: ビュー名
        seiban: 製番 (例: 'MHT0620')

    Returns:
        dict: 検索結果
    """
    return query_view(view_name, '製番 = ?', [seiban.strip()])


def search_order_remaining(order_number):
    """
    発注残（納入状況）を検索

    Args:
        order_number: 発注番号

    Returns:
        dict: 検索結果（納入済数を含む）
    """
    padded = str(order_number).strip().zfill(8)
    return query_view('V_D発注残', '発注番号 = ?', [padded])


def search_receipts(order_number):
    """
    仕入（納入実績）を検索

    Args:
        order_number: 発注番号

    Returns:
        dict: 検索結果（納入日・納入数を含む）
    """
    padded = str(order_number).strip().zfill(8)
    return query_view('V_D仕入', '発注番号 = ?', [padded])


def get_view_columns(view_name):
    """ビューのカラム一覧を取得"""
    if view_name not in AVAILABLE_VIEWS:
        raise ValueError(f"不正なビュー名: {view_name}")

    conn = None
    try:
        conn, cursor = get_connection()
        cursor.execute(f"SELECT TOP 0 * FROM dbo.[{view_name}]")
        columns = [col[0] for col in cursor.description]
        return columns
    finally:
        if conn:
            conn.close()


def test_connection():
    """接続テスト"""
    conn = None
    try:
        conn, cursor = get_connection()
        cursor.execute("SELECT 1")
        return {'success': True, 'message': 'Across DB接続OK'}
    except Exception as e:
        return {'success': False, 'message': str(e)}
    finally:
        if conn:
            conn.close()
