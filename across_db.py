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
    'V_D受注': '受注データ（製番・品名・顧客名など）',
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


def search_zaiko_buhin(seibans=None):
    """
    在庫部品（手配区分CD='15'）を検索
    在庫から集めて仕分けが必要な部品を抽出

    Args:
        seibans: 製番リスト（指定時はその製番のみ、None時は全製番）

    Returns:
        dict: { columns, rows, count, by_seiban }
    """
    conn = None
    try:
        conn, cursor = get_connection()

        if seibans and len(seibans) > 0:
            placeholders = ','.join(['?' for _ in seibans])
            cursor.execute(f"""
                SELECT 製番, 手配数, 品名, 仕様１, 仕様２, 手配区分, 備考, 材質, 日付
                FROM dbo.[V_D手配リスト]
                WHERE 手配区分CD = '15' AND 製番 IN ({placeholders})
                ORDER BY 製番, 品名
            """, seibans)
        else:
            cursor.execute("""
                SELECT 製番, 手配数, 品名, 仕様１, 仕様２, 手配区分, 備考, 材質, 日付
                FROM dbo.[V_D手配リスト]
                WHERE 手配区分CD = '15'
                ORDER BY 製番, 品名
            """)

        columns = [col[0] for col in cursor.description]
        raw_rows = cursor.fetchall()

        rows = []
        by_seiban = {}
        for row in raw_rows:
            formatted = [format_value(v) for v in row]
            rows.append(formatted)
            seiban = formatted[0] or ''
            if seiban not in by_seiban:
                by_seiban[seiban] = []
            by_seiban[seiban].append(formatted)

        return {
            'success': True,
            'columns': columns,
            'rows': rows,
            'count': len(rows),
            'by_seiban': by_seiban,
            'seiban_count': len(by_seiban)
        }
    except Exception as e:
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            conn.close()


def search_0zaiko_tehai():
    """
    0ZAIKO（在庫品発注用製番）の手配リストを検索
    在庫補充用に発注された部品を表示

    Returns:
        dict: { columns, rows, count }
    """
    conn = None
    try:
        conn, cursor = get_connection()

        cursor.execute("""
            SELECT 製番, 手配数, 品名, 仕様１, 仕様２, 手配区分, 備考, メーカー, 材質, 日付
            FROM dbo.[V_D手配リスト]
            WHERE 製番 = '0ZAIKO'
            ORDER BY 日付 DESC, 品名
        """)

        columns = [col[0] for col in cursor.description]
        raw_rows = cursor.fetchall()

        rows = []
        for row in raw_rows:
            rows.append([format_value(v) for v in row])

        return {
            'success': True,
            'columns': columns,
            'rows': rows,
            'count': len(rows)
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


# ========================================
# DB更新検知機能
# ========================================

# 前回のスナップショット保存用
_db_snapshot = {
    'tehai': {'count': 0, 'seibans': set()},
    'hacchu': {'count': 0, 'latest_numbers': set()},
    'last_check': None
}


def get_db_status():
    """
    DBの現在の状態を取得（更新検知用）

    Returns:
        dict: {
            'tehai': {'count': int, 'seibans': set, 'new_seibans': list},
            'hacchu': {'count': int, 'new_count': int},
            'timestamp': datetime
        }
    """
    conn = None
    try:
        conn, cursor = get_connection()

        # 手配リストの製番一覧と件数を取得
        cursor.execute("""
            SELECT 製番, COUNT(*) as cnt
            FROM dbo.[V_D手配リスト]
            GROUP BY 製番
        """)
        tehai_rows = cursor.fetchall()
        tehai_seibans = {row[0].strip() if row[0] else '' for row in tehai_rows}
        tehai_count = sum(row[1] for row in tehai_rows)

        # 発注リストの件数を取得
        cursor.execute("SELECT COUNT(*) FROM dbo.[V_D発注]")
        hacchu_count = cursor.fetchone()[0]

        # 最新の発注番号（直近100件）を取得
        cursor.execute("""
            SELECT TOP 100 発注番号
            FROM dbo.[V_D発注]
            ORDER BY 発注番号 DESC
        """)
        hacchu_rows = cursor.fetchall()
        latest_hacchu = {str(row[0]).strip() if row[0] else '' for row in hacchu_rows}

        return {
            'success': True,
            'tehai': {
                'count': tehai_count,
                'seibans': tehai_seibans,
                'seiban_count': len(tehai_seibans)
            },
            'hacchu': {
                'count': hacchu_count,
                'latest_numbers': latest_hacchu
            },
            'timestamp': datetime.now()
        }
    except Exception as e:
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            conn.close()


def check_db_updates(stored_snapshot=None):
    """
    DBの更新をチェック

    Args:
        stored_snapshot: 前回のスナップショット（Noneの場合は内部キャッシュを使用）

    Returns:
        dict: {
            'has_updates': bool,
            'tehai_changes': {'new_seibans': [...], 'count_diff': int},
            'hacchu_changes': {'new_orders': [...], 'count_diff': int},
            'current_snapshot': {...},
            'message': str
        }
    """
    global _db_snapshot

    # 現在の状態を取得
    current = get_db_status()
    if not current.get('success'):
        return {
            'success': False,
            'has_updates': False,
            'error': current.get('error', '接続エラー')
        }

    # 比較対象のスナップショット
    if stored_snapshot:
        prev = stored_snapshot
    else:
        prev = _db_snapshot

    # 初回チェックの場合
    if prev['last_check'] is None:
        _db_snapshot = {
            'tehai': {
                'count': current['tehai']['count'],
                'seibans': current['tehai']['seibans']
            },
            'hacchu': {
                'count': current['hacchu']['count'],
                'latest_numbers': current['hacchu']['latest_numbers']
            },
            'last_check': current['timestamp']
        }
        return {
            'success': True,
            'has_updates': False,
            'is_first_check': True,
            'tehai_changes': {'new_seibans': [], 'count_diff': 0},
            'hacchu_changes': {'new_orders': [], 'count_diff': 0},
            'message': '初回チェック完了（ベースライン設定）',
            'current_snapshot': {
                'tehai_count': current['tehai']['count'],
                'tehai_seiban_count': current['tehai']['seiban_count'],
                'hacchu_count': current['hacchu']['count'],
                'timestamp': current['timestamp'].isoformat()
            }
        }

    # 変更を検出
    has_updates = False
    messages = []

    # 手配リストの変更検出
    prev_seibans = prev['tehai'].get('seibans', set())
    curr_seibans = current['tehai']['seibans']
    new_seibans = list(curr_seibans - prev_seibans)
    tehai_count_diff = current['tehai']['count'] - prev['tehai'].get('count', 0)

    if new_seibans:
        has_updates = True
        messages.append(f"新規製番: {', '.join(sorted(new_seibans)[:5])}")
        if len(new_seibans) > 5:
            messages[-1] += f" 他{len(new_seibans)-5}件"

    if tehai_count_diff > 0:
        has_updates = True
        messages.append(f"手配リスト: +{tehai_count_diff}件")
    elif tehai_count_diff < 0:
        messages.append(f"手配リスト: {tehai_count_diff}件")

    # 発注リストの変更検出
    prev_hacchu = prev['hacchu'].get('latest_numbers', set())
    curr_hacchu = current['hacchu']['latest_numbers']
    new_orders = list(curr_hacchu - prev_hacchu)
    hacchu_count_diff = current['hacchu']['count'] - prev['hacchu'].get('count', 0)

    # 新規発注の詳細を取得
    new_order_details = []
    if new_orders and len(new_orders) <= 50:  # 50件以下なら詳細取得
        new_order_details = get_new_order_details(new_orders[:20])

    if hacchu_count_diff > 0:
        has_updates = True
        messages.append(f"発注リスト: +{hacchu_count_diff}件")
    elif hacchu_count_diff < 0:
        messages.append(f"発注リスト: {hacchu_count_diff}件")

    # スナップショット更新
    _db_snapshot = {
        'tehai': {
            'count': current['tehai']['count'],
            'seibans': current['tehai']['seibans']
        },
        'hacchu': {
            'count': current['hacchu']['count'],
            'latest_numbers': current['hacchu']['latest_numbers']
        },
        'last_check': current['timestamp']
    }

    return {
        'success': True,
        'has_updates': has_updates,
        'is_first_check': False,
        'tehai_changes': {
            'new_seibans': sorted(new_seibans),
            'count_diff': tehai_count_diff
        },
        'hacchu_changes': {
            'new_orders': sorted(new_orders)[:20],  # 最大20件
            'new_order_details': new_order_details,  # 詳細情報
            'count_diff': hacchu_count_diff
        },
        'message': ' / '.join(messages) if messages else '変更なし',
        'current_snapshot': {
            'tehai_count': current['tehai']['count'],
            'tehai_seiban_count': current['tehai']['seiban_count'],
            'hacchu_count': current['hacchu']['count'],
            'timestamp': current['timestamp'].isoformat()
        }
    }


def get_new_order_details(order_numbers):
    """
    発注番号リストから発注詳細を取得

    Args:
        order_numbers: 発注番号のリスト

    Returns:
        list: [{'発注番号': str, '製番': str, '品名': str, '仕様１': str, '発注数': int, '仕入先': str, '納期': str}, ...]
    """
    if not order_numbers:
        return []

    conn = None
    try:
        conn, cursor = get_connection()

        # IN句用のプレースホルダー作成
        placeholders = ','.join(['?' for _ in order_numbers])
        # 8桁ゼロパディング
        padded_numbers = [str(n).zfill(8) for n in order_numbers]

        sql = f"""
            SELECT 発注番号, 製番, 品名, 仕様１, 発注数, 単位, 仕入先略称, 納期
            FROM dbo.[V_D発注]
            WHERE 発注番号 IN ({placeholders})
            ORDER BY 発注番号 DESC
        """

        cursor.execute(sql, padded_numbers)
        rows = cursor.fetchall()

        details = []
        for row in rows:
            details.append({
                '発注番号': format_value(row[0]),
                '製番': format_value(row[1]),
                '品名': format_value(row[2]),
                '仕様１': format_value(row[3]),
                '発注数': format_value(row[4]),
                '単位': format_value(row[5]),
                '仕入先': format_value(row[6]),
                '納期': format_value(row[7])
            })

        return details
    except Exception as e:
        print(f"発注詳細取得エラー: {e}")
        return []
    finally:
        if conn:
            conn.close()


def get_delivery_schedule_from_db(start_date=None, days=7, seibans=None):
    """
    発注DBから納品予定を取得

    Args:
        start_date: 開始日 (YYYY-MM-DD形式、Noneで今日)
        days: 取得日数 (デフォルト7日)
        seibans: 製番リスト（指定時はその製番のみ）

    Returns:
        dict: {
            'success': bool,
            'items': [{'製番': str, '品名': str, ...}, ...],
            'days': {date: [items]},
            'total': int
        }
    """
    conn = None
    try:
        conn, cursor = get_connection()

        from datetime import datetime, date, timedelta

        # 日付範囲
        if start_date:
            start = datetime.strptime(start_date, '%Y-%m-%d').date()
        else:
            start = date.today()

        end = start + timedelta(days=days)

        # SQLクエリ構築
        sql = """
            SELECT
                発注番号, 製番, 品名, 仕様１, 発注数, 単位,
                仕入先略称, 納期, 手配区分
            FROM dbo.[V_D発注残]
            WHERE 納期 >= ? AND 納期 <= ?
        """
        params = [start.strftime('%Y-%m-%d'), end.strftime('%Y-%m-%d')]

        # 製番指定がある場合
        if seibans and len(seibans) > 0:
            placeholders = ','.join(['?' for _ in seibans])
            sql += f" AND 製番 IN ({placeholders})"
            params.extend(seibans)

        sql += " ORDER BY 納期, 製番, 品名"

        cursor.execute(sql, params)
        rows = cursor.fetchall()

        items = []
        days_dict = {}

        for row in rows:
            raw_date = row[7]  # 納期
            delivery_date_display = format_value(raw_date)  # 表示用 (YY/MM/DD)

            # 日付キー用 (YYYY-MM-DD形式)
            date_key = None
            if raw_date:
                if isinstance(raw_date, (datetime, date)):
                    date_key = raw_date.strftime('%Y-%m-%d')
                else:
                    date_key = str(raw_date)

            item = {
                '発注番号': format_value(row[0]),
                '製番': format_value(row[1]),
                '品名': format_value(row[2]),
                '仕様１': format_value(row[3]),
                '発注数': format_value(row[4]),
                '単位': format_value(row[5]),
                '仕入先': format_value(row[6]),
                '納期': delivery_date_display,
                '手配区分': format_value(row[8])
            }
            items.append(item)

            # 日別グループ化 (YYYY-MM-DD形式のキー)
            if date_key:
                if date_key not in days_dict:
                    days_dict[date_key] = []
                days_dict[date_key].append(item)

        return {
            'success': True,
            'items': items,
            'days': days_dict,
            'total': len(items),
            'start_date': start.strftime('%Y-%m-%d'),
            'end_date': end.strftime('%Y-%m-%d')
        }
    except Exception as e:
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            conn.close()


def get_seiban_updates(seiban):
    """
    特定製番の手配・発注の更新状況を取得

    Args:
        seiban: 製番

    Returns:
        dict: {'tehai_count': int, 'hacchu_count': int, 'mihatchu_count': int}
    """
    conn = None
    try:
        conn, cursor = get_connection()

        result = {'seiban': seiban}

        # 手配リスト件数
        cursor.execute(
            "SELECT COUNT(*) FROM dbo.[V_D手配リスト] WHERE 製番 = ?",
            [seiban]
        )
        result['tehai_count'] = cursor.fetchone()[0]

        # 発注件数
        cursor.execute(
            "SELECT COUNT(*) FROM dbo.[V_D発注] WHERE 製番 = ?",
            [seiban]
        )
        result['hacchu_count'] = cursor.fetchone()[0]

        # 未発注件数
        cursor.execute(
            "SELECT COUNT(*) FROM dbo.[V_D未発注] WHERE 製番 = ?",
            [seiban]
        )
        result['mihatchu_count'] = cursor.fetchone()[0]

        # 発注残件数
        cursor.execute(
            "SELECT COUNT(*) FROM dbo.[V_D発注残] WHERE 製番 = ?",
            [seiban]
        )
        result['hacchuzan_count'] = cursor.fetchone()[0]

        result['success'] = True
        return result
    except Exception as e:
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            conn.close()


def get_seiban_list_from_db(min_seiban=None):
    """
    V_D受注から製番一覧を取得（製番選択ドロップダウン用）

    Args:
        min_seiban: 最小製番（この番号以降を取得、例: 'MHT0600'）

    Returns:
        dict: {
            'success': bool,
            'items': [{'seiban': str, 'product_name': str, 'customer_name': str, 'memo2': str}, ...],
            'count': int
        }
    """
    conn = None
    try:
        conn, cursor = get_connection()

        # V_D受注から製番・品名・得意先略称・メモ２を取得
        # 製番ごとに1件のみ（重複除去）
        sql = """
            SELECT DISTINCT
                製番,
                品名,
                得意先略称,
                メモ２
            FROM dbo.[V_D受注]
            WHERE 製番 IS NOT NULL AND 製番 <> ''
        """
        params = []

        # 最小製番フィルタ
        if min_seiban:
            sql += " AND 製番 >= ?"
            params.append(min_seiban.strip())

        sql += " ORDER BY 製番 DESC"

        cursor.execute(sql, params)
        rows = cursor.fetchall()

        items = []
        for row in rows:
            seiban = format_value(row[0])
            if not seiban:
                continue
            items.append({
                'seiban': seiban,
                'product_name': format_value(row[1]) or '',
                'customer_name': format_value(row[2]) or '',
                'memo2': format_value(row[3]) or ''
            })

        return {
            'success': True,
            'items': items,
            'count': len(items)
        }
    except Exception as e:
        return {'success': False, 'error': str(e), 'items': []}
    finally:
        if conn:
            conn.close()
