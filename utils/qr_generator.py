"""
QRコード生成ユーティリティモジュール
Untitled.py の1178行目～1190行目から抽出
"""

import qrcode
import base64
from io import BytesIO


def generate_qr_code(data):
    """
    QRコードを生成してBase64エンコードされた文字列を返す
    
    Args:
        data: QRコードに埋め込むデータ（文字列）
        
    Returns:
        str: Base64エンコードされたQRコード画像データ
        
    Examples:
        >>> qr_data = generate_qr_code("ORDER:12345")
        >>> print(f"QRコード長: {len(qr_data)}文字")
    """
    qr = qrcode.QRCode(version=1, box_size=10, border=5)
    qr.add_data(data)
    qr.make(fit=True)
    
    img = qr.make_image(fill_color="black", back_color="white")
    buffer = BytesIO()
    img.save(buffer, format='PNG')
    buffer.seek(0)
    
    return base64.b64encode(buffer.getvalue()).decode()


# テスト用コード
if __name__ == '__main__':
    print("=== QRコード生成テスト ===")
    
    # テストデータ
    test_data = [
        "ORDER:12345",
        "MHT0614",
        "https://example.com/order/12345"
    ]
    
    for data in test_data:
        qr_code = generate_qr_code(data)
        print(f"\nデータ: {data}")
        print(f"QRコード長: {len(qr_code)}文字")
        print(f"先頭30文字: {qr_code[:30]}...")
        
        # HTMLで表示する場合の例
        print(f"HTML: <img src='data:image/png;base64,{qr_code[:30]}...' />")
