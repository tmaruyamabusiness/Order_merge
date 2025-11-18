"""
メッキ判定ユーティリティモジュール
Untitled.py の172行目～197行目から抽出
"""

import re
from .constants import Constants


class MekkiUtils:
    """メッキ判定ユーティリティ"""
    
    @staticmethod
    def is_mekki_target(supplier_cd, spec2):
        """
        メッキ出対象判定
        
        Args:
            supplier_cd: 仕入先コード
            spec2: 仕様2（メッキパターンを含む可能性のある文字列）
            
        Returns:
            bool: メッキ出対象の場合True
        """
        if not spec2 or not isinstance(spec2, str):
            return False
        
        # 仕入先コードがメッキ仕入先でない場合はFalse
        if supplier_cd != Constants.MEKKI_SUPPLIER_CD and supplier_cd != str(Constants.MEKKI_SUPPLIER_CD):
            return False
        
        # メッキパターンに一致するかチェック
        for pattern in Constants.MEKKI_PATTERNS:
            if re.search(pattern, spec2, re.IGNORECASE):
                return True
        return False
    
    @staticmethod
    def add_mekki_alert(remarks):
        """
        備考にメッキ出を追加
        
        Args:
            remarks: 既存の備考
            
        Returns:
            str: メッキ出アラートを追加した備考
        """
        remarks = remarks or ''
        if Constants.MEKKI_ALERT_TEXT not in remarks:
            return f"{remarks} {Constants.MEKKI_ALERT_TEXT}" if remarks else Constants.MEKKI_ALERT_TEXT
        return remarks


# テスト用コード
if __name__ == '__main__':
    print("=== MekkiUtils テスト ===")
    
    # is_mekki_targetのテスト
    print("\n--- is_mekki_target テスト ---")
    print(f"supplier_cd=116, spec2='/Ni-P' → {MekkiUtils.is_mekki_target(116, '/Ni-P')}")
    print(f"supplier_cd=116, spec2='/NiCr' → {MekkiUtils.is_mekki_target(116, '/NiCr')}")
    print(f"supplier_cd=100, spec2='/Ni-P' → {MekkiUtils.is_mekki_target(100, '/Ni-P')}")
    print(f"supplier_cd=116, spec2='通常' → {MekkiUtils.is_mekki_target(116, '通常')}")
    
    # add_mekki_alertのテスト
    print("\n--- add_mekki_alert テスト ---")
    print(f"remarks='' → '{MekkiUtils.add_mekki_alert('')}'")
    print(f"remarks='注意事項' → '{MekkiUtils.add_mekki_alert('注意事項')}'")
    print(f"remarks='◆メッキ出' → '{MekkiUtils.add_mekki_alert('◆メッキ出')}'")
