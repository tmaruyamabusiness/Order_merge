"""
定数定義モジュール
Untitled.py の68行目～103行目から抽出
"""

class Constants:
    """アプリケーション定数"""
    
    # メッキ関連
    MEKKI_SUPPLIER_CD = 116
    MEKKI_PATTERNS = [
        r'/Ni-P', r'/NiCr', r'／Ni-P', r'／NiCr',
        r'Ｎｉ－Ｐ', r'Ｃｒ', r'ＮｉＣｒ'
    ]
    # 仕入先に関係なくメッキ出となる仕様１パターン
    MEKKI_SPEC1_CODES = ['NMA-00017-00-00']
    MEKKI_ALERT_TEXT = '⚠️メッキ出'
    
    # 手配区分CD
    ORDER_TYPE_BLANK = '13'  # 加工用ブランク
    ORDER_TYPE_PROCESSED = '11'  # 追加工
    ORDER_TYPE_STOCK = '15'  # 在庫部品
    
    # Excel出力
    EXCEL_COLUMNS = ['納入日', '納入数', '納期', '仕入先略称', '発注番号', '手配数', '単位', '品名',
                     '仕様１', '仕様２', '手配区分', 'メーカー', '備考']
    COLUMN_WIDTHS = {
        'A': 10, 'B': 7, 'C': 10, 'D': 12, 'E': 10, 'F': 6, 'G': 6, 'H': 25,
        'I': 20, 'J': 15, 'K': 12, 'L': 10, 'M': 15
    }
    
    # 色設定
    COLOR_GRAY = "DADADA"
    COLOR_WHITE = "FFFFFF"
    COLOR_ACCEPT_LIGHT = "FEC0D7"  # 薄いピンク
    COLOR_ACCEPT_DARK = "FFABC9"   # 濃いピンク
    COLOR_RED = "FF0000"
    COLOR_HEADER = "4F4F4F"
    COLOR_CHILD = "F0F0F0"
    
    # ステータス
    STATUS_BEFORE = '受入準備前'
    STATUS_IN_PROGRESS = '納品中'
    STATUS_COMPLETED = '納品完了'