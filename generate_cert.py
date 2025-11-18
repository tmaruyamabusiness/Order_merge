"""
開発用SSL証明書生成スクリプト
"""

import os
from datetime import datetime, timedelta
from cryptography import x509
from cryptography.x509.oid import NameOID
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.hazmat.primitives import serialization

def generate_self_signed_cert():
    """自己署名証明書を生成"""
    
    # 秘密鍵を生成
    print("🔐 秘密鍵を生成中...")
    private_key = rsa.generate_private_key(
        public_exponent=65537,
        key_size=2048,
    )
    
    # 証明書の詳細情報
    subject = issuer = x509.Name([
        x509.NameAttribute(NameOID.COUNTRY_NAME, "JP"),
        x509.NameAttribute(NameOID.STATE_OR_PROVINCE_NAME, "Tokyo"),
        x509.NameAttribute(NameOID.LOCALITY_NAME, "Tokyo"),
        x509.NameAttribute(NameOID.ORGANIZATION_NAME, "Order Management System"),
        x509.NameAttribute(NameOID.COMMON_NAME, "localhost"),
    ])
    
    # 証明書を生成
    print("📜 証明書を生成中...")
    cert = x509.CertificateBuilder().subject_name(
        subject
    ).issuer_name(
        issuer
    ).public_key(
        private_key.public_key()
    ).serial_number(
        x509.random_serial_number()
    ).not_valid_before(
        datetime.utcnow()
    ).not_valid_after(
        datetime.utcnow() + timedelta(days=365)
    ).add_extension(
        x509.SubjectAlternativeName([
            x509.DNSName("localhost"),
            x509.DNSName("127.0.0.1"),
            x509.DNSName("::1"),
        ]),
        critical=False,
    ).sign(private_key, hashes.SHA256())
    
    # 秘密鍵をファイルに保存
    print("💾 key.pemを保存中...")
    with open("key.pem", "wb") as f:
        f.write(private_key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.TraditionalOpenSSL,
            encryption_algorithm=serialization.NoEncryption()
        ))
    
    # 証明書をファイルに保存
    print("💾 cert.pemを保存中...")
    with open("cert.pem", "wb") as f:
        f.write(cert.public_bytes(serialization.Encoding.PEM))
    
    print("✅ 証明書の生成が完了しました！")
    print("📁 生成されたファイル:")
    print("   - cert.pem (証明書)")
    print("   - key.pem (秘密鍵)")
    print("\n🚀 アプリケーションを起動できます: python app.py")

if __name__ == "__main__":
    try:
        # 必要なパッケージのインストール確認
        try:
            import cryptography
        except ImportError:
            print("⚠️  cryptographyパッケージが必要です")
            print("インストール: pip install cryptography")
            exit(1)
        
        generate_self_signed_cert()
        
    except Exception as e:
        print(f"❌ エラーが発生しました: {e}")
        print("\n代替方法を使用してください（方法2を参照）")