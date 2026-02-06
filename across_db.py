"""
Across DB 直接クエリモジュール
V_D系ビュー（読み取り専用）への安全なアクセスを提供
"""

import pyodbc
import pandas as pd
from decimal import Decimal
from datetime import datetime, date


# 利用可能なビュー一覧
AVAILABLE_VIEWS = {
    'V_D発注': '発注データ（発注番号・仕入先・納期など）',
    'V_D発注残': '発注残データ（納入済数を含む）',
    'V_D手配リスト': '手配リスト（BOM・部品表）',
    'V_D仕入': '仕入データ（納入実績）',
    'V_D未発注': '未発注データ（社内加工品含む）',
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
        return value.strftime('%y/%m/%d')
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


def merge_from_db(seiban, order_date_from=None, order_date_to=None):
    """
    製番でV_D手配リストとV_D発注をマージし、save_to_database()互換のDataFrameを返す
    Excel経由の process_excel_file_from_dataframes() を完全に置き換える

    Args:
        seiban: 製番 (例: 'MHT0620')
        order_date_from: 発注日フィルタ開始日 (例: '2026-01-01')
        order_date_to: 発注日フィルタ終了日 (例: '2026-12-31')

    Returns:
        pandas.DataFrame: save_to_database()に渡せる形式のDataFrame
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

        if not tehai_rows:
            return None

        # 2. V_D発注（発注データ）を取得 - 回答納期を含む
        hatchu_sql = """
            SELECT 発注番号, 製番, 品名, 仕様１, 仕様２, 手配区分CD, 手配区分,
                   材質, 仕入先CD, 仕入先名, 仕入先略称, 発注数, 単位,
                   発注単価, 発注金額, 発注日, 納期, 回答納期, 備考
            FROM dbo.[V_D発注]
            WHERE 製番 = ?
        """
        hatchu_params = [seiban]

        # 発注日フィルタ
        if order_date_from:
            hatchu_sql += " AND 発注日 >= ?"
            hatchu_params.append(order_date_from)
        if order_date_to:
            hatchu_sql += " AND 発注日 <= ?"
            hatchu_params.append(order_date_to)

        cursor.execute(hatchu_sql, hatchu_params)
        hatchu_cols = [col[0] for col in cursor.description]
        hatchu_rows = cursor.fetchall()

        # 発注データを辞書リスト化
        hatchu_list = []
        for row in hatchu_rows:
            rec = {}
            for j, col in enumerate(hatchu_cols):
                rec[col] = format_value(row[j])
            hatchu_list.append(rec)

        # 3. マージ実行 → DataFrame行リスト構築
        merged_rows = []

        for row in tehai_rows:
            tehai_rec = {}
            for j, col in enumerate(tehai_cols):
                tehai_rec[col] = format_value(row[j])

            # 基本カラム（手配リストから）
            merged = {
                '納期': '',
                '回答納期': '',
                '仕入先略称': '',
                '仕入先CD': '',
                '発注番号': '',
                '手配数': tehai_rec.get('手配数', 0) or 0,
                '単位': tehai_rec.get('単位', '') or '',
                '品名': tehai_rec.get('品名', '') or '',
                '仕様１': tehai_rec.get('仕様１', '') or '',
                '仕様２': tehai_rec.get('仕様２', '') or '',
                '品目CD': tehai_rec.get('品目CD', '') or '',
                '手配区分CD': tehai_rec.get('手配区分CD', '') or '',
                '手配区分': tehai_rec.get('手配区分', '') or '',
                'メーカー': tehai_rec.get('メーカー', '') or '',
                '備考': tehai_rec.get('備考', '') or '',
                '員数': tehai_rec.get('員数', 0) or 0,
                '必要数': tehai_rec.get('必要数', 0) or 0,
                '製番': tehai_rec.get('製番', '') or '',
                '材質': tehai_rec.get('材質', '') or '',
                '部品No': tehai_rec.get('部品No', '') or '',
                'ページNo': tehai_rec.get('ページNo', '') or '',
                '行No': tehai_rec.get('行No', '') or '',
                '階層': tehai_rec.get('階層', 0) or 0,
            }

            material = str(merged['材質']).strip()
            spec1 = str(merged['仕様１']).strip()
            order_type = str(merged['手配区分']).strip()

            matched = False

            # Primary match: 材質 + 仕様１ + 製番
            if material and spec1:
                for h in hatchu_list:
                    h_material = str(h.get('材質', '') or '').strip()
                    h_spec1 = str(h.get('仕様１', '') or '').strip()
                    h_seiban = str(h.get('製番', '') or '').strip()
                    if h_material == material and h_spec1 == spec1 and h_seiban == seiban:
                        merged['発注番号'] = str(h.get('発注番号', '') or '')
                        merged['仕入先略称'] = str(h.get('仕入先略称', '') or '')
                        merged['仕入先CD'] = str(h.get('仕入先CD', '') or '')
                        merged['納期'] = str(h.get('納期', '') or '')
                        merged['回答納期'] = str(h.get('回答納期', '') or '')
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
                        merged['発注番号'] = str(h.get('発注番号', '') or '')
                        merged['仕入先略称'] = str(h.get('仕入先略称', '') or '')
                        merged['仕入先CD'] = str(h.get('仕入先CD', '') or '')
                        merged['納期'] = str(h.get('納期', '') or '')
                        merged['回答納期'] = str(h.get('回答納期', '') or '')
                        matched = True
                        break

            merged_rows.append(merged)

        # DataFrameに変換
        df = pd.DataFrame(merged_rows)

        # None値を空文字に
        df = df.fillna('')

        # 手配区分CDを文字列に正規化
        df['手配区分CD'] = df['手配区分CD'].apply(
            lambda x: str(int(float(x))) if x and x != '' else str(x)
        )

        # 発注番号あり/なしでソート
        if '発注番号' in df.columns:
            df_with = df[df['発注番号'] != '']
            df_without = df[df['発注番号'] == '']
            if not df_with.empty:
                df_with = df_with.sort_values('発注番号')
            df = pd.concat([df_with, df_without], ignore_index=True)

        return df

    finally:
        if conn:
            conn.close()


def search_mihatchu(seiban, supplier_cd=None, order_type_cd=None):
    """
    V_D未発注から検索（社内加工品の発注漏れ確認用）

    Args:
        seiban: 製番
        supplier_cd: 仕入先CD (例: 'MHT')
        order_type_cd: 手配区分CD (例: '11')
    """
    where_parts = ['製番 = ?']
    params = [seiban.strip()]
    if supplier_cd:
        where_parts.append('仕入先CD = ?')
        params.append(supplier_cd)
    if order_type_cd:
        where_parts.append('手配区分CD = ?')
        params.append(order_type_cd)
    return query_view('V_D未発注', ' AND '.join(where_parts), params, 500)


def merge_from_db_with_mihatchu(seiban, order_date_from=None, order_date_to=None):
    """
    merge_from_db + V_D未発注の社内加工品(MHT+11)を統合
    """
    # 通常マージ
    df = merge_from_db(seiban, order_date_from, order_date_to)

    # V_D未発注から社内加工品を取得
    conn = None
    try:
        conn, cursor = get_connection()
        cursor.execute("""
            SELECT 製番, 品名, 仕様１, 仕様２, 手配区分CD, 手配区分, メーカー,
                   材質, 仕入先CD, 仕入先略称, 発注数, 単位, 納期, 備考,
                   ページNo, 行No, 階層
            FROM dbo.[V_D未発注]
            WHERE 製番 = ? AND 仕入先CD = 'MHT' AND 手配区分CD = '11'
        """, seiban.strip())
        mihatchu_cols = [col[0] for col in cursor.description]
        mihatchu_rows = cursor.fetchall()

        if not mihatchu_rows:
            return df

        # 未発注データをDataFrame行に変換
        mihatchu_merged = []
        for row in mihatchu_rows:
            rec = {}
            for j, col in enumerate(mihatchu_cols):
                rec[col] = format_value(row[j])

            merged = {
                '納期': str(rec.get('納期', '') or ''),
                '回答納期': '',
                '仕入先略称': str(rec.get('仕入先略称', '') or ''),
                '仕入先CD': str(rec.get('仕入先CD', '') or ''),
                '発注番号': '',  # 未発注なので空
                '手配数': rec.get('発注数', 0) or 0,
                '単位': str(rec.get('単位', '') or ''),
                '品名': str(rec.get('品名', '') or ''),
                '仕様１': str(rec.get('仕様１', '') or ''),
                '仕様２': str(rec.get('仕様２', '') or ''),
                '品目CD': '',
                '手配区分CD': str(rec.get('手配区分CD', '') or ''),
                '手配区分': str(rec.get('手配区分', '') or ''),
                'メーカー': str(rec.get('メーカー', '') or ''),
                '備考': str(rec.get('備考', '') or ''),
                '員数': 0,
                '必要数': 0,
                '製番': seiban.strip(),
                '材質': str(rec.get('材質', '') or ''),
                '部品No': '',
                'ページNo': str(rec.get('ページNo', '') or ''),
                '行No': str(rec.get('行No', '') or ''),
                '階層': rec.get('階層', 0) or 0,
            }
            mihatchu_merged.append(merged)

        if not mihatchu_merged:
            return df

        df_mihatchu = pd.DataFrame(mihatchu_merged).fillna('')

        # 既にマージ済みdfに含まれている仕様１は除外（重複防止）
        if df is not None and not df.empty:
            existing_specs = set(df['仕様１'].astype(str).str.strip())
            df_mihatchu = df_mihatchu[~df_mihatchu['仕様１'].astype(str).str.strip().isin(existing_specs)]

        if df_mihatchu.empty:
            return df

        if df is not None and not df.empty:
            df = pd.concat([df, df_mihatchu], ignore_index=True)
        else:
            df = df_mihatchu

        return df

    finally:
        if conn:
            conn.close()


def check_db_updates(seibans):
    """
    指定製番リストについてDBの最新件数と更新状況をチェック

    Args:
        seibans: 製番リスト

    Returns:
        dict: { seiban: { tehai_count, hatchu_count, mihatchu_count, has_new } }
    """
    conn = None
    try:
        conn, cursor = get_connection()
        results = {}

        for seiban in seibans:
            s = seiban.strip()
            # 手配リスト件数
            cursor.execute("SELECT COUNT(*) FROM dbo.[V_D手配リスト] WHERE 製番 = ?", s)
            tehai_count = cursor.fetchone()[0]

            # 発注データ件数
            cursor.execute("SELECT COUNT(*) FROM dbo.[V_D発注] WHERE 製番 = ?", s)
            hatchu_count = cursor.fetchone()[0]

            # 未発注件数（社内加工品）
            cursor.execute("""
                SELECT COUNT(*) FROM dbo.[V_D未発注]
                WHERE 製番 = ? AND 仕入先CD = 'MHT' AND 手配区分CD = '11'
            """, s)
            mihatchu_count = cursor.fetchone()[0]

            results[s] = {
                'tehai_count': tehai_count,
                'hatchu_count': hatchu_count,
                'mihatchu_count': mihatchu_count,
            }

        return {'success': True, 'results': results}
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
