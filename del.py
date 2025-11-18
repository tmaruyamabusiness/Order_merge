# Pythonコンソールまたは新しいエンドポイントで実行
from app import app, db, Order
from sqlalchemy import func

with app.app_context():
    # 同じ製番・ユニットの組み合わせで重複しているレコードを検索
    duplicates = db.session.query(
        Order.seiban,
        Order.unit,
        func.count(Order.id).label('count')
    ).filter(
        Order.is_archived == False
    ).group_by(
        Order.seiban,
        Order.unit
    ).having(
        func.count(Order.id) > 1
    ).all()
    
    print(f"🔍 重複検出: {len(duplicates)}件")
    
    for dup in duplicates:
        seiban, unit, count = dup
        print(f"\n📦 製番: {seiban}, ユニット: {unit or 'ユニット名無し'}, 重複数: {count}")
        
        # 同じ製番・ユニットのレコードを全て取得
        orders = Order.query.filter_by(
            seiban=seiban,
            unit=unit,
            is_archived=False
        ).order_by(Order.id.asc()).all()
        
        # 最初のレコードを残して、残りを削除
        keep_order = orders[0]
        print(f"  ✅ 保持: ID={keep_order.id}, 詳細数={len(keep_order.details)}")
        
        for order in orders[1:]:
            print(f"  🗑️  削除: ID={order.id}, 詳細数={len(order.details)}")
            # 詳細も一緒に削除される（cascade設定による）
            db.session.delete(order)
    
    db.session.commit()
    print("\n✅ クリーンアップ完了")
