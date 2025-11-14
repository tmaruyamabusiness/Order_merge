"""
既存データベースの浮動小数点数を修正するスクリプト
発注番号と手配区分CDの .0 を除去
"""

from app import app, db, OrderDetail
from sqlalchemy import text

def fix_float_numbers():
    """既存データの浮動小数点数を修正"""
    
    with app.app_context():
        try:
            print("=" * 50)
            print("浮動小数点数の修正を開始します")
            print("=" * 50)
            
            # 発注番号の修正
            try:
                # .0 で終わる発注番号を検索
                details_with_float = OrderDetail.query.filter(
                    OrderDetail.order_number.like('%.0')
                ).all()
                
                fixed_count = 0
                for detail in details_with_float:
                    old_number = detail.order_number
                    # .0 を除去
                    new_number = old_number.replace('.0', '')
                    detail.order_number = new_number
                    fixed_count += 1
                    print(f"発注番号修正: {old_number} → {new_number}")
                
                if fixed_count > 0:
                    db.session.commit()
                    print(f"✅ {fixed_count}件の発注番号を修正しました")
                else:
                    print("ℹ️ 修正が必要な発注番号はありません")
                    
            except Exception as e:
                print(f"⚠️ 発注番号修正エラー: {e}")
                db.session.rollback()
            
            # 手配区分CDの修正
            try:
                # .0 で終わる手配区分CDを検索
                details_with_float_cd = OrderDetail.query.filter(
                    OrderDetail.order_type_code.like('%.0')
                ).all()
                
                fixed_count = 0
                for detail in details_with_float_cd:
                    old_code = detail.order_type_code
                    # .0 を除去
                    new_code = old_code.replace('.0', '')
                    detail.order_type_code = new_code
                    fixed_count += 1
                    print(f"手配区分CD修正: {old_code} → {new_code}")
                
                if fixed_count > 0:
                    db.session.commit()
                    print(f"✅ {fixed_count}件の手配区分CDを修正しました")
                else:
                    print("ℹ️ 修正が必要な手配区分CDはありません")
                    
            except Exception as e:
                print(f"⚠️ 手配区分CD修正エラー: {e}")
                db.session.rollback()
            
            print("\n✅ 浮動小数点数の修正が完了しました")
            
        except Exception as e:
            print(f"❌ 修正エラー: {e}")
            return False
    
    return True

def check_current_data():
    """現在のデータ状態を確認"""
    
    with app.app_context():
        try:
            print("\n" + "=" * 50)
            print("現在のデータ状態")
            print("=" * 50)
            
            # 発注番号の確認
            with db.engine.connect() as conn:
                result = conn.execute(text("""
                    SELECT DISTINCT order_number 
                    FROM order_detail 
                    WHERE order_number LIKE '%.0' 
                    OR order_number LIKE '%.%'
                    LIMIT 10
                """))
                
                float_numbers = result.fetchall()
                if float_numbers:
                    print("\n⚠️ 浮動小数点形式の発注番号:")
                    for row in float_numbers:
                        print(f"  - {row[0]}")
                else:
                    print("\n✅ 浮動小数点形式の発注番号はありません")
                
                # 手配区分CDの確認
                result = conn.execute(text("""
                    SELECT DISTINCT order_type_code 
                    FROM order_detail 
                    WHERE order_type_code LIKE '%.0' 
                    OR order_type_code LIKE '%.%'
                    LIMIT 10
                """))
                
                float_codes = result.fetchall()
                if float_codes:
                    print("\n⚠️ 浮動小数点形式の手配区分CD:")
                    for row in float_codes:
                        print(f"  - {row[0]}")
                else:
                    print("\n✅ 浮動小数点形式の手配区分CDはありません")
                
                # 統計情報
                result = conn.execute(text("""
                    SELECT COUNT(DISTINCT order_number) as count
                    FROM order_detail
                    WHERE order_number IS NOT NULL AND order_number != ''
                """))
                
                count = result.fetchone()[0]
                print(f"\n📊 総発注番号数: {count}件")
                
        except Exception as e:
            print(f"❌ 確認エラー: {e}")

if __name__ == "__main__":
    print("浮動小数点数修正ツール")
    print("=" * 50)
    
    # 現在の状態を確認
    check_current_data()
    
    # 修正を実行
    response = input("\n修正を実行しますか？ (y/n): ")
    
    if response.lower() == 'y':
        if fix_float_numbers():
            print("\n修正後の状態:")
            check_current_data()
    else:
        print("修正をキャンセルしました")