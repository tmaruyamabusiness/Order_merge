"""
検収データ（DV_仕入.xlsx）読み込みユーティリティ
"""

import pandas as pd
from datetime import datetime
import os


class DeliveryUtils:
    """検収データユーティリティ"""

    # DV_仕入.xlsxのパス
    DV_SHIIRE_PATH = r'\\SERVER3\Share-data\Document\仕入れ\002_手配リスト\DV_仕入.xlsx'
    SHEET_NAME = '仕入_価格確認用'

    # キャッシュ（同一セッション内での再読み込みを防ぐ）
    _cache = None
    _cache_time = None
    CACHE_DURATION = 300  # 5分間キャッシュ

    @classmethod
    def load_delivery_data(cls, force_reload=False):
        """
        DV_仕入.xlsxから検収データを読み込み、発注番号をキーとした辞書を返す

        Returns:
            dict: {発注番号: {'納入日': date, '納入数': int}, ...}
        """
        # キャッシュチェック
        if not force_reload and cls._cache is not None:
            if cls._cache_time and (datetime.now() - cls._cache_time).seconds < cls.CACHE_DURATION:
                return cls._cache

        delivery_dict = {}

        try:
            # ファイル存在チェック
            if not os.path.exists(cls.DV_SHIIRE_PATH):
                print(f"⚠️ DV_仕入.xlsx が見つかりません: {cls.DV_SHIIRE_PATH}")
                return delivery_dict

            # Excelファイル読み込み
            df = pd.read_excel(
                cls.DV_SHIIRE_PATH,
                sheet_name=cls.SHEET_NAME,
                engine='openpyxl'
            )

            print(f"✅ DV_仕入.xlsx 読み込み成功: {len(df)}件")

            # 必要なカラムが存在するかチェック
            required_cols = ['発注番号', '納入日', '納入数']
            for col in required_cols:
                if col not in df.columns:
                    print(f"⚠️ 必須カラム '{col}' が見つかりません")
                    return delivery_dict

            # 発注番号でグループ化して集計
            for _, row in df.iterrows():
                order_number = row.get('発注番号')

                # 発注番号がNaNまたは空の場合はスキップ
                if pd.isna(order_number) or order_number == '':
                    continue

                # 発注番号を文字列に変換（整数の場合）
                if isinstance(order_number, (int, float)):
                    order_number = str(int(order_number))
                else:
                    order_number = str(order_number).strip()

                # 納入日を取得
                delivery_date = row.get('納入日')
                if pd.isna(delivery_date):
                    delivery_date = None
                elif isinstance(delivery_date, datetime):
                    delivery_date = delivery_date.strftime('%y/%m/%d')
                else:
                    delivery_date = str(delivery_date)

                # 納入数を取得
                delivery_qty = row.get('納入数', 0)
                if pd.isna(delivery_qty):
                    delivery_qty = 0
                else:
                    delivery_qty = float(delivery_qty)

                # 既存のエントリがあれば納入数を合計
                if order_number in delivery_dict:
                    delivery_dict[order_number]['納入数'] += delivery_qty
                    # 納入日は最新のものを使用
                    if delivery_date:
                        delivery_dict[order_number]['納入日'] = delivery_date
                else:
                    delivery_dict[order_number] = {
                        '納入日': delivery_date,
                        '納入数': delivery_qty
                    }

            # キャッシュ更新
            cls._cache = delivery_dict
            cls._cache_time = datetime.now()

            print(f"✅ 検収データ集計完了: {len(delivery_dict)}件の発注番号")

        except Exception as e:
            print(f"❌ DV_仕入.xlsx 読み込みエラー: {e}")
            import traceback
            traceback.print_exc()

        return delivery_dict

    @classmethod
    def get_delivery_info(cls, order_number, delivery_dict=None):
        """
        発注番号から納入情報を取得

        Args:
            order_number: 発注番号
            delivery_dict: 検収データ辞書（Noneの場合は自動読み込み）

        Returns:
            dict: {'納入日': str or None, '納入数': float or 0}
        """
        if delivery_dict is None:
            delivery_dict = cls.load_delivery_data()

        # 発注番号を文字列に正規化
        if order_number is None:
            return {'納入日': None, '納入数': 0}

        if isinstance(order_number, (int, float)):
            order_number = str(int(order_number))
        else:
            order_number = str(order_number).strip()

        return delivery_dict.get(order_number, {'納入日': None, '納入数': 0})

    @classmethod
    def clear_cache(cls):
        """キャッシュをクリア"""
        cls._cache = None
        cls._cache_time = None


# テスト用コード
if __name__ == '__main__':
    print("=== DeliveryUtils テスト ===")

    # データ読み込みテスト
    data = DeliveryUtils.load_delivery_data()

    # サンプル表示
    print("\n--- サンプルデータ（先頭5件） ---")
    for i, (order_num, info) in enumerate(data.items()):
        if i >= 5:
            break
        print(f"発注番号: {order_num}, 納入日: {info['納入日']}, 納入数: {info['納入数']}")
