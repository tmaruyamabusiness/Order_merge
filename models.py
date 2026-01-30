"""
データベースモデル定義 - models.py
app.py からインポートして外部モジュールに提供
"""
# app.pyの循環インポートを避けるため、遅延インポートを使用
def get_db():
    from app import db
    return db

def get_models():
    from app import Order, OrderDetail, ReceivedHistory, EditLog, ProcessingHistory
    return Order, OrderDetail, ReceivedHistory, EditLog, ProcessingHistory
