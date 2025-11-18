"""
データ処理ユーティリティモジュール
Untitled.py の126行目～171行目から抽出
"""

import pandas as pd


class DataUtils:
    """データ処理ユーティリティ"""
    
    @staticmethod
    def safe_str(value):
        """
        安全な文字列変換
        
        Args:
            value: 変換する値
            
        Returns:
            str: 変換された文字列（None/NaNの場合は空文字列）
        """
        if pd.isna(value) or value is None:
            return ''
        if isinstance(value, float) and value == value:
            try:
                return str(int(value))
            except:
                return str(value)
        return str(value)
    
    @staticmethod
    def safe_int(value, default=0):
        """
        安全な整数変換
        
        Args:
            value: 変換する値
            default: デフォルト値
            
        Returns:
            int: 変換された整数（変換失敗時はdefault）
        """
        if pd.isna(value) or value is None:
            return default
        try:
            return int(float(value))
        except (ValueError, TypeError):
            return default
    
    @staticmethod
    def normalize_order_number(order_number):
        """
        発注番号正規化（ゼロパディング + 浮動小数点対策）
        
        Args:
            order_number: 発注番号
            
        Returns:
            str: 正規化された発注番号
            
        Examples:
            >>> DataUtils.normalize_order_number("00086922")
            '86922'
            >>> DataUtils.normalize_order_number(86922.0)
            '86922'
        """
        if order_number is None:
            return ''
        
        # 文字列に変換
        order_str = str(order_number).strip()
        
        # 空文字列チェック
        if not order_str or order_str == 'nan':
            return ''
        
        try:
            # ゼロパディングを削除して数値化
            # "00086922" → 86922 → "86922"
            order_int = int(float(order_str))
            return str(order_int)
        except (ValueError, TypeError):
            return order_str


# テスト用コード
if __name__ == '__main__':
    # safe_strのテスト
    print("=== safe_str テスト ===")
    print(f"None → '{DataUtils.safe_str(None)}'")
    print(f"123.0 → '{DataUtils.safe_str(123.0)}'")
    print(f"'test' → '{DataUtils.safe_str('test')}'")
    
    # safe_intのテスト
    print("\n=== safe_int テスト ===")
    print(f"None → {DataUtils.safe_int(None)}")
    print(f"'123' → {DataUtils.safe_int('123')}")
    print(f"'abc' → {DataUtils.safe_int('abc', default=-1)}")
    
    # normalize_order_numberのテスト
    print("\n=== normalize_order_number テスト ===")
    print(f"'00086922' → '{DataUtils.normalize_order_number('00086922')}'")
    print(f"86922.0 → '{DataUtils.normalize_order_number(86922.0)}'")
    print(f"None → '{DataUtils.normalize_order_number(None)}'")
