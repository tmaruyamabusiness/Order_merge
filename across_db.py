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


def merge_test_by_seiban(seiban):
    """
    製番でV_D手配リストとV_D発注をマージテスト
    現在のExcelマージ処理と同等のロジックをDB直接クエリで再現

    Args:
        seiban: 製番 (例: 'MHT0620')

    Returns:
        dict: マージ結果と統計情報
    """
    seiban = seiban.strip()
    conn = None
    try:
        conn, cursor = get_connection()

        # 1. V_D手配リスト（BOM）を取得
        cursor.execute("""
            SELECT 製番, 担当者, ページNo, 行No, 部品No, 階層, 品目CD,
                   品名, 仕様１, 仕様２, 手配区分CD, 手配区分, メーカー,
                   材質, 員数, 必要数, 手配数, 単位, 備考, 日付
            FROM dbo.[V_D手配リスト]
            WHERE 製番 = ?
        """, seiban)
        tehai_cols = [col[0] for col in cursor.description]
        tehai_rows = cursor.fetchall()

        # 2. V_D発注（発注データ）を取得
        cursor.execute("""
            SELECT 発注番号, 製番, 品名, 仕様１, 仕様２, 手配区分CD, 手配区分,
                   材質, 仕入先CD, 仕入先名, 仕入先略称, 発注数, 単位,
                   発注単価, 発注金額, 発注日, 納期, 回答納期, 備考
            FROM dbo.[V_D発注]
            WHERE 製番 = ?
        """, seiban)
        hatchu_cols = [col[0] for col in cursor.description]
        hatchu_rows = cursor.fetchall()

        # 3. V_D発注残（納入状況）を取得
        cursor.execute("""
            SELECT 発注番号, 製番, 品名, 仕様１, 発注数, 納入済数, 納入済金額
            FROM dbo.[V_D発注残]
            WHERE 製番 = ?
        """, seiban)
        remaining_cols = [col[0] for col in cursor.description]
        remaining_rows = cursor.fetchall()

        # 発注データを辞書化（マージ用）
        hatchu_list = []
        for row in hatchu_rows:
            rec = {}
            for j, col in enumerate(hatchu_cols):
                rec[col] = format_value(row[j])
            hatchu_list.append(rec)

        # 発注残データを辞書化（発注番号→納入済数）
        remaining_map = {}
        for row in remaining_rows:
            rec = {}
            for j, col in enumerate(remaining_cols):
                rec[col] = format_value(row[j])
            order_num = rec.get('発注番号', '')
            remaining_map[order_num] = rec

        # 4. マージ実行（現行ロジックと同等）
        merged_results = []
        match_count = 0
        unmatch_count = 0

        for row in tehai_rows:
            tehai_rec = {}
            for j, col in enumerate(tehai_cols):
                tehai_rec[col] = format_value(row[j])

            # マージ結果の初期値
            merged = dict(tehai_rec)
            merged['発注番号'] = ''
            merged['仕入先略称'] = ''
            merged['仕入先CD'] = ''
            merged['納期'] = ''
            merged['納入済数'] = ''
            merged['match_type'] = ''

            material = str(tehai_rec.get('材質', '') or '').strip()
            spec1 = str(tehai_rec.get('仕様１', '') or '').strip()
            order_type = str(tehai_rec.get('手配区分', '') or '').strip()

            matched = False

            # Primary match: 材質 + 仕様１ + 製番
            if material and spec1:
                for h in hatchu_list:
                    h_material = str(h.get('材質', '') or '').strip()
                    h_spec1 = str(h.get('仕様１', '') or '').strip()
                    h_seiban = str(h.get('製番', '') or '').strip()
                    if h_material == material and h_spec1 == spec1 and h_seiban == seiban:
                        merged['発注番号'] = h.get('発注番号', '')
                        merged['仕入先略称'] = h.get('仕入先略称', '')
                        merged['仕入先CD'] = h.get('仕入先CD', '')
                        merged['納期'] = h.get('納期', '')
                        merged['match_type'] = '材質+仕様１'
                        matched = True
                        break

            # Fallback match: 製番 + 仕様１ (+手配区分)
            if not matched and spec1:
                for h in hatchu_list:
                    h_spec1 = str(h.get('仕様１', '') or '').strip()
                    h_seiban = str(h.get('製番', '') or '').strip()
                    h_type = str(h.get('手配区分', '') or '').strip()
                    if h_seiban == seiban and h_spec1 == spec1:
                        if order_type and h_type and order_type != h_type:
                            continue
                        merged['発注番号'] = h.get('発注番号', '')
                        merged['仕入先略称'] = h.get('仕入先略称', '')
                        merged['仕入先CD'] = h.get('仕入先CD', '')
                        merged['納期'] = h.get('納期', '')
                        merged['match_type'] = '仕様１(+区分)'
                        matched = True
                        break

            # 納入済数を追加
            order_num = merged.get('発注番号', '')
            if order_num and order_num in remaining_map:
                rem = remaining_map[order_num]
                merged['納入済数'] = rem.get('納入済数', '')

            if matched:
                match_count += 1
            else:
                unmatch_count += 1

            merged_results.append(merged)

        # 5. ユニット（材質）別にグループ化
        units = {}
        for m in merged_results:
            unit = str(m.get('材質', '') or '').strip()
            if unit not in units:
                units[unit] = []
            units[unit].append(m)

        # 結果カラム
        result_columns = [
            '品名', '仕様１', '仕様２', '手配区分', '材質',
            '手配数', '単位', '発注番号', '仕入先略称', '納期',
            '納入済数', 'match_type', '階層', '部品No', '員数', '必要数'
        ]

        # ユニット別にフラット化
        result_rows = []
        for m in merged_results:
            result_rows.append([m.get(c, '') for c in result_columns])

        return {
            'success': True,
            'seiban': seiban,
            'columns': result_columns,
            'rows': result_rows,
            'stats': {
                'tehai_count': len(tehai_rows),
                'hatchu_count': len(hatchu_rows),
                'match_count': match_count,
                'unmatch_count': unmatch_count,
                'match_rate': round(match_count / len(tehai_rows) * 100, 1) if tehai_rows else 0,
                'unit_count': len(units),
                'units': list(units.keys())
            }
        }
    except Exception as e:
        return {'success': False, 'error': str(e)}
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
