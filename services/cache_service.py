"""
キャッシュ・マスタデータ管理 - cache_service.py
製番一覧表の読み込み
"""
from pathlib import Path
import pandas as pd
from flask import current_app


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
