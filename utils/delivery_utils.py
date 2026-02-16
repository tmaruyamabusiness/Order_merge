"""
検収データユーティリティ（DB直接クエリで VD_仕入 から取得）
"""

# キャッシュ用辞書（発注番号 -> 納入情報）
_delivery_cache = {}


class DeliveryUtils:
    """検収データユーティリティ（VD_仕入からDB直接クエリ）"""

    @classmethod
    def load_delivery_data(cls, force_reload=False):
        """
        検収データを返す（互換性維持用 - キャッシュをクリアするトリガーとして使用）

        Returns:
            dict: 空の辞書（個別クエリで取得するため）
        """
        if force_reload:
            cls.clear_cache()
        return {}

    @classmethod
    def get_delivery_info(cls, order_number, delivery_dict=None):
        """
        発注番号から納入情報を取得（VD_仕入からDB直接クエリ）

        Args:
            order_number: 発注番号
            delivery_dict: 互換性維持用（使用しない）

        Returns:
            dict: {'納入日': str, '納入数': float}
        """
        global _delivery_cache

        if not order_number:
            return {'納入日': None, '納入数': 0}

        # キャッシュにあればそれを返す
        order_number_str = str(order_number).strip()
        if order_number_str in _delivery_cache:
            return _delivery_cache[order_number_str]

        try:
            import across_db
            result = across_db.search_receipts(order_number_str)

            if result['count'] == 0:
                info = {'納入日': None, '納入数': 0}
                _delivery_cache[order_number_str] = info
                return info

            columns = result['columns']
            rows = result['rows']

            # カラムインデックスを取得（カラム名のパターンマッチ）
            date_idx = None
            qty_idx = None
            for i, col in enumerate(columns):
                if not col:
                    continue
                col_str = str(col)
                # 納入日/仕入日/入荷日/伝票日付 など
                if any(pattern in col_str for pattern in ['納入日', '仕入日', '入荷日', '検収日', '伝票日付', '日付']):
                    if date_idx is None:  # 最初に見つかったものを使用
                        date_idx = i
                # 納入数/仕入数/数量 など
                if any(pattern in col_str for pattern in ['納入数', '仕入数', '入荷数', '数量']):
                    if qty_idx is None:  # 最初に見つかったものを使用
                        qty_idx = i

            # カラムが見つからなかった場合のログ
            if date_idx is None or qty_idx is None:
                print(f"[DeliveryUtils] VD_仕入カラム: {columns}")
                print(f"[DeliveryUtils] 日付カラムidx={date_idx}, 数量カラムidx={qty_idx}")

            # 複数行の場合（分納）: 最新日付を使用、数量は合計
            latest_date = None
            total_qty = 0

            for row in rows:
                if date_idx is not None and row[date_idx]:
                    date_val = row[date_idx]
                    if latest_date is None or (date_val and date_val > latest_date):
                        latest_date = date_val
                if qty_idx is not None and row[qty_idx]:
                    try:
                        total_qty += float(row[qty_idx])
                    except (ValueError, TypeError):
                        pass

            info = {
                '納入日': latest_date,
                '納入数': total_qty
            }
            _delivery_cache[order_number_str] = info
            return info

        except Exception as e:
            print(f"[DeliveryUtils] VD_仕入クエリエラー (発注番号={order_number}): {e}")
            return {'納入日': None, '納入数': 0}

    @classmethod
    def clear_cache(cls):
        """キャッシュをクリア"""
        global _delivery_cache
        _delivery_cache = {}
