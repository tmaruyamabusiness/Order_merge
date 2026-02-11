"""
キャッシュ・マスタデータ管理 - cache_service.py
製番情報の読み込み（V_D受注DBから取得、フォールバックでExcel）
"""
from pathlib import Path
import pandas as pd
from flask import current_app


def load_seiban_info():
    """製番情報を取得（V_D受注DBから取得、フォールバックでExcel）"""
    # まずDBから取得を試みる
    try:
        import across_db
        result = across_db.get_seiban_list_from_db()
        if result.get('success') and result.get('items'):
            # DB結果を辞書形式に変換
            seiban_info = {}
            for item in result['items']:
                seiban = item.get('seiban', '')
                if seiban:
                    seiban_info[seiban] = {
                        'product_name': item.get('product_name', ''),
                        'customer_abbr': item.get('customer_name', ''),  # customer_name → customer_abbr
                        'memo2': item.get('memo2', '')
                    }
            print(f"製番情報をDBから取得: {len(seiban_info)}件")
            return seiban_info
    except Exception as e:
        print(f"DB取得エラー、Excelにフォールバック: {str(e)}")

    # DBが使えない場合はExcelから読み込み（フォールバック）
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

        print(f"製番情報をExcelから取得: {len(seiban_info)}件")
        return seiban_info
    except Exception as e:
        print(f"製番一覧表読み込みエラー: {str(e)}")
        return {}
