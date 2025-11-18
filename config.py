"""
設定ファイル - config.py
環境に応じて設定を変更してください
"""

import os
import ssl
from pathlib import Path

class Config:
    """基本設定"""
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'dev-secret-key-change-in-production'
    SQLALCHEMY_DATABASE_URI = 'sqlite:///order_management.db'
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    
    # ファイルアップロード設定
    UPLOAD_FOLDER = 'uploads'
    MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB
    
    # ネットワークパス設定（Windows UNCパス形式）
    DEFAULT_EXCEL_PATH = r'\\server3\Share-data\Document\仕入れ\002_手配リスト\手配発注_ALL.xlsx'
    HISTORY_EXCEL_PATH = r'\\server3\Share-data\Document\仕入れ\002_手配リスト\手配発注マージリスト発行履歴.xlsx'
    SEIBAN_LIST_PATH = r'\\server3\share-data\Document\Acrossデータ\製番一覧表.xlsx'
    EXPORT_EXCEL_PATH = r'\\SERVER3\Share-data\Document\仕入れ\002_手配リスト\手配発注リスト'  # 🔥 追加
    
    # 自動更新設定
    AUTO_REFRESH_INTERVAL = 3600  # 秒単位（3600 = 1時間）
    AUTO_REFRESH_ENABLED = True
    
    # ODBC設定
    USE_ODBC = False  # ODBCを使用する場合はTrue
    ODBC_CONNECTION_STRING = ''
    
    # HTTPS/SSL設定
    USE_HTTPS = False
    SSL_CERT_PATH = None
    SSL_KEY_PATH = None

class DevelopmentConfig(Config):
    """開発環境設定"""
    DEBUG = True
    
    # 開発環境用HTTPS設定（自己署名証明書）
    USE_HTTPS = True
    SSL_CERT_PATH = 'cert.pem'
    SSL_KEY_PATH = 'key.pem'
    # または adhoc を使用（pyopenssl必要）
    # SSL_CONTEXT = 'adhoc'
    
    # テスト用のローカルパス
    DEFAULT_EXCEL_PATH = r'\\server3\Share-data\Document\仕入れ\002_手配リスト\手配発注_ALL.xlsx'
    HISTORY_EXCEL_PATH = r'\\server3\Share-data\Document\仕入れ\002_手配リスト\手配発注マージリスト発行履歴.xlsx'
    AUTO_REFRESH_INTERVAL = 300  # 5分（開発時は短めに）

class ProductionConfig(Config):
    """本番環境設定"""
    DEBUG = False
    SECRET_KEY = os.environ.get('SECRET_KEY')  # 環境変数から取得
    
    # 本番環境用HTTPS設定
    USE_HTTPS = True
    SSL_CERT_PATH = os.environ.get('SSL_CERT_PATH', '/etc/ssl/certs/cert.pem')
    SSL_KEY_PATH = os.environ.get('SSL_KEY_PATH', '/etc/ssl/private/key.pem')
    
    # 本番環境のネットワークパス
    DEFAULT_EXCEL_PATH = r'\\server3\Share-data\Document\仕入れ\002_手配リスト\手配発注_ALL.xlsx'
    HISTORY_EXCEL_PATH = r'\\server3\Share-data\Document\仕入れ\002_手配リスト\手配発注マージリスト発行履歴.xlsx'
    
    # ODBC接続（Excel）
    # USE_ODBC = True
    # ODBC_CONNECTION_STRING = '''
    # DRIVER={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};
    # DBQ=\\server3\\Share-data\\Document\\仕入れ\\002_手配リスト\\手配発注_ALL.xlsx;
    # ReadOnly=1;
    # '''
    
    # Windows認証を使用する場合（SQL Server）
    USE_ODBC = True  # 必要に応じてTrueに変更
    ODBC_CONNECTION_STRING = '''
        DRIVER={SQL Server};
        SERVER=SERVER3;
        DATABASE=Across;
        Trusted_Connection=yes;
    '''

class TestConfig(Config):
    """テスト環境設定"""
    TESTING = True
    SQLALCHEMY_DATABASE_URI = 'sqlite:///test_order_management.db'
    WTF_CSRF_ENABLED = False
    USE_HTTPS = False  # テスト時はHTTPS不要

# 環境変数で設定を切り替え
config = {
    'development': DevelopmentConfig,
    'production': ProductionConfig,
    'testing': TestConfig,
    'default': DevelopmentConfig
}

def get_config():
    """現在の環境設定を取得"""
    env = os.environ.get('FLASK_ENV', 'development')
    return config.get(env, config['default'])

def get_ssl_context(config_obj):
    """SSL/TLSコンテキストを取得"""
    if not config_obj.USE_HTTPS:
        return None
    
    # adhoc証明書（開発用、pyopensslが必要）
    if hasattr(config_obj, 'SSL_CONTEXT') and config_obj.SSL_CONTEXT == 'adhoc':
        return 'adhoc'
    
    # 証明書ファイルが存在する場合
    if config_obj.SSL_CERT_PATH and config_obj.SSL_KEY_PATH:
        cert_path = Path(config_obj.SSL_CERT_PATH)
        key_path = Path(config_obj.SSL_KEY_PATH)
        
        if cert_path.exists() and key_path.exists():
            return (str(cert_path), str(key_path))
        else:
            print(f"警告: SSL証明書が見つかりません")
            print(f"  証明書: {cert_path}")
            print(f"  秘密鍵: {key_path}")
            return None
    
    # デフォルトでadhocを使用（開発環境）
    if config_obj.DEBUG:
        return 'adhoc'
    
    return None