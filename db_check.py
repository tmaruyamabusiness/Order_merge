from app import app, db, Order, OrderDetail

def check_parent_child_relationship(seiban, unit):
    """親子関係を確認"""
    with app.app_context():
        order = Order.query.filter_by(seiban=seiban, unit=unit).first()
        
        if not order:
            print(f"❌ {seiban} - {unit} が見つかりません")
            return
        
        print(f"\n✅ 注文ID: {order.id}, 製番: {order.seiban}, ユニット: {order.unit}")
        print("=" * 110)
        print(f"{'ID':4} {'parent_id':10} {'品名':30} {'仕様１':35} {'手配区分':20}")
        print("=" * 110)
        
        # 親子関係を確認
        parent_child_pairs = []
        
        for detail in order.details:
            parent_id_str = str(detail.parent_id) if detail.parent_id else 'None'
            item_name = (detail.item_name[:28] + '..') if len(detail.item_name) > 28 else detail.item_name
            spec1 = (detail.spec1[:33] + '..') if len(detail.spec1) > 33 else detail.spec1
            order_type = (detail.order_type[:18] + '..') if len(detail.order_type) > 18 else detail.order_type
            
            print(f"{detail.id:<4} {parent_id_str:<10} {item_name:<30} {spec1:<35} {order_type:<20}")
            
            # 親子ペアを記録
            if detail.parent_id:
                parent_child_pairs.append((detail.parent_id, detail.id))
        
        print("=" * 110)
        
        # 親子関係のサマリー
        if parent_child_pairs:
            print(f"\n🔗 親子関係が見つかりました: {len(parent_child_pairs)}組")
            for parent_id, child_id in parent_child_pairs:
                parent = OrderDetail.query.get(parent_id)
                child = OrderDetail.query.get(child_id)
                print(f"  親({parent_id}): {parent.item_name} → 子({child_id}): {child.item_name}")
        else:
            print(f"\n⚠️  親子関係が設定されていません")
        
        print()

if __name__ == '__main__':
    # MHT0614のカッターを確認
    check_parent_child_relationship('MHT0614', 'カッター')