"""
検収データユーティリティ（DB直接クエリに移行済み）
※ DV_仕入.xlsx のファイル読み込みは廃止
"""


class DeliveryUtils:
    """検収データユーティリティ（互換性維持用スタブ）"""

    @classmethod
    def load_delivery_data(cls, force_reload=False):
        """
        検収データを返す（DB直接クエリに移行済みのため空辞書を返す）

        Returns:
            dict: 空の辞書（互換性維持用）
        """
        return {}

    @classmethod
    def get_delivery_info(cls, order_number, delivery_dict=None):
        """
        発注番号から納入情報を取得（互換性維持用）

        Returns:
            dict: {'納入日': None, '納入数': 0}
        """
        return {'納入日': None, '納入数': 0}

    @classmethod
    def clear_cache(cls):
        """キャッシュをクリア（互換性維持用 - 何もしない）"""
        pass
