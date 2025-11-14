import sqlite3

# データベースに接続
conn = sqlite3.connect('order_management.db')
cursor = conn.cursor()

print("🔄 データベースマイグレーション開始...")

# カラムを追加
try:
    cursor.execute('ALTER TABLE "order" ADD COLUMN pallet_number VARCHAR(50)')
    print("✅ pallet_numberカラムを追加しました")
except Exception as e:
    print(f"⚠️  pallet_number: {e}")

try:
    cursor.execute('ALTER TABLE "order" ADD COLUMN floor VARCHAR(10)')
    print("✅ floorカラムを追加しました")
except Exception as e:
    print(f"⚠️  floor: {e}")

# インデックスを作成
try:
    cursor.execute('CREATE INDEX idx_order_pallet_number ON "order"(pallet_number)')
    print("✅ pallet_numberインデックスを作成しました")
except Exception as e:
    print(f"⚠️  インデックス: {e}")

try:
    cursor.execute('CREATE INDEX idx_order_floor ON "order"(floor)')
    print("✅ floorインデックスを作成しました")
except Exception as e:
    print(f"⚠️  インデックス: {e}")

# コミットして閉じる
conn.commit()
conn.close()

print("\n✅ マイグレーション完了！")