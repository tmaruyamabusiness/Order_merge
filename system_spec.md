# 手配発注マージシステム - システム仕様書

## 1. システム概要

製造業向けの発注管理Webアプリケーション。製番（製造番号）単位で部品の手配・発注・受入を管理し、Excel帳票出力やDB連携を提供する。

### 1.1 技術スタック

| カテゴリ | 技術 |
|---------|------|
| バックエンド | Python 3.x, Flask |
| データベース | SQLite (ローカル), SQL Server (Across DB/ODBC) |
| フロントエンド | HTML5, CSS3, JavaScript (vanilla) |
| Excel処理 | openpyxl, pandas |
| その他 | QRコード生成 (qrcode), Chart.js, win32com |

### 1.2 主要機能

1. **データ処理**: Excel/DB からのデータ取込・マージ
2. **受注管理**: 製番・ユニット単位での発注明細管理
3. **受入管理**: QRコード・バーコードによる受入確認
4. **Excel出力**: 手配発注リスト、ガントチャート出力
5. **パレット管理**: ユニット保管場所の追跡
6. **DB連携**: Across DB (SQL Server) への直接クエリ
7. **メール通知**: 納品完了通知メール作成

---

## 2. データベース設計

### 2.1 ローカルDB (SQLite)

#### Order テーブル（製番・ユニット単位）

| カラム名 | 型 | 説明 |
|---------|-----|------|
| id | INTEGER | 主キー |
| seiban | VARCHAR(50) | 製番 (例: MHT0620) |
| unit | VARCHAR(100) | ユニット名（材質から派生） |
| status | VARCHAR(50) | ステータス（受入準備前/納品中/納品完了） |
| location | VARCHAR(50) | 保管場所 |
| remarks | TEXT | 備考 |
| product_name | VARCHAR(200) | 品名 |
| customer_abbr | VARCHAR(100) | 得意先略称 |
| memo2 | VARCHAR(200) | メモ２ |
| pallet_number | VARCHAR(50) | パレット番号 |
| floor | VARCHAR(10) | フロア（1F/2F等） |
| image_path | VARCHAR(500) | 画像パス |
| is_archived | BOOLEAN | アーカイブ済フラグ |
| archived_at | DATETIME | アーカイブ日時 |
| created_at | DATETIME | 作成日時 |
| updated_at | DATETIME | 更新日時 |

#### OrderDetail テーブル（発注明細）

| カラム名 | 型 | 説明 |
|---------|-----|------|
| id | INTEGER | 主キー |
| order_id | INTEGER | Order.id への外部キー |
| delivery_date | VARCHAR(20) | 納期 |
| supplier | VARCHAR(100) | 仕入先名 |
| supplier_cd | VARCHAR(50) | 仕入先コード |
| order_number | VARCHAR(50) | 発注番号 |
| quantity | INTEGER | 手配数 |
| unit_measure | VARCHAR(20) | 単位 |
| item_name | VARCHAR(200) | 品名 |
| spec1 | VARCHAR(200) | 仕様１（部品コード） |
| spec2 | VARCHAR(200) | 仕様２ |
| item_code | VARCHAR(50) | 品目コード |
| order_type_code | VARCHAR(20) | 手配区分CD |
| order_type | VARCHAR(50) | 手配区分 |
| maker | VARCHAR(100) | メーカー |
| remarks | TEXT | 備考 |
| member_count | INTEGER | 員数 |
| required_count | INTEGER | 必要数 |
| seiban | VARCHAR(50) | 製番 |
| material | VARCHAR(100) | 材質 |
| is_received | BOOLEAN | 受入済フラグ |
| received_at | DATETIME | 受入日時 |
| has_internal_processing | BOOLEAN | 社内加工フラグ |
| parent_id | INTEGER | 親明細ID（階層構造用） |
| part_number | VARCHAR(50) | 部品No |
| page_number | VARCHAR(20) | ページNo |
| row_number | VARCHAR(20) | 行No |
| hierarchy | INTEGER | 階層 |
| reply_delivery_date | VARCHAR(20) | 回答納期 |

#### ReceivedHistory テーブル（受入履歴）

発注番号をキーとして受入状態を永続化。

| カラム名 | 型 | 説明 |
|---------|-----|------|
| id | INTEGER | 主キー |
| order_number | VARCHAR(50) | 発注番号 |
| item_name | VARCHAR(200) | 品名 |
| spec1 | VARCHAR(200) | 仕様１ |
| quantity | INTEGER | 数量 |
| is_received | BOOLEAN | 受入状態 |
| received_at | DATETIME | 受入日時 |
| received_by | VARCHAR(100) | 受入者（IPアドレス） |
| cancelled_at | DATETIME | キャンセル日時 |
| cancelled_by | VARCHAR(100) | キャンセル者 |

#### EditLog テーブル（編集ログ）

| カラム名 | 型 | 説明 |
|---------|-----|------|
| id | INTEGER | 主キー |
| detail_id | INTEGER | OrderDetail.id への外部キー |
| action | VARCHAR(50) | アクション（receive/unreceive） |
| ip_address | VARCHAR(45) | 実行者IP |
| timestamp | DATETIME | タイムスタンプ |
| user_agent | VARCHAR(500) | ユーザーエージェント |

#### ProcessingHistory テーブル（処理履歴）

| カラム名 | 型 | 説明 |
|---------|-----|------|
| id | INTEGER | 主キー |
| serial_no | INTEGER | 連番 |
| issue_date | DATETIME | 発行日時 |
| filename | VARCHAR(200) | ファイル名 |
| file_size_kb | FLOAT | ファイルサイズ(KB) |
| seiban | VARCHAR(50) | 製番 |

### 2.2 Across DB (SQL Server / ODBC)

読み取り専用で以下のビューにアクセス:

| ビュー名 | 説明 |
|---------|------|
| V_D発注 | 発注データ（発注番号・仕入先・納期など） |
| V_D発注残 | 発注残データ（納入済数を含む） |
| V_D手配リスト | 手配リスト（BOM・部品表） |
| V_D仕入 | 仕入データ（納入実績） |
| V_D未発注 | 未発注データ（社内加工品含む） |
| V_D受注 | 受注データ（製番・品名・顧客名など） |

---

## 3. APIエンドポイント

### 3.1 注文管理

| メソッド | パス | 説明 |
|---------|------|------|
| GET | `/api/orders` | 注文一覧取得 |
| GET | `/api/order/<id>` | 注文詳細取得 |
| POST | `/api/order/<id>/update` | 注文更新 |
| DELETE | `/api/order/<id>/delete` | 注文削除 |
| POST | `/api/orders/delete-multiple` | 複数注文削除 |
| POST | `/api/order/<id>/archive` | アーカイブ |
| POST | `/api/order/<id>/unarchive` | アーカイブ解除 |
| GET | `/api/archived-orders` | アーカイブ一覧 |
| POST | `/api/order/<id>/update-remarks` | 備考更新 |
| POST | `/api/order/<id>/upload-image` | 画像アップロード |
| GET | `/api/order/<id>/image` | 画像取得 |
| DELETE | `/api/order/<id>/delete-image` | 画像削除 |

### 3.2 データ処理

| メソッド | パス | 説明 |
|---------|------|------|
| POST | `/api/upload` | Excelアップロード |
| POST | `/api/process` | データ処理（マージ実行） |
| GET | `/api/export/<id>` | Excel出力 |
| GET | `/api/export-seiban/<seiban>` | 製番単位Excel出力 |
| POST | `/api/refresh-excel` | Excel更新 |
| POST | `/api/run-refresh-script` | 更新スクリプト実行 |
| POST | `/api/generate-labels` | ラベル生成 |

### 3.3 受入管理

| メソッド | パス | 説明 |
|---------|------|------|
| GET | `/receive/<seiban>/<unit>` | 受入確認ページ |
| POST | `/api/detail/<id>/toggle-receive` | 受入トグル |
| POST | `/api/detail/<id>/receive` | 受入処理 |
| GET | `/api/detail/<id>/logs` | 編集ログ取得 |
| GET | `/api/detail/<id>/cad-info` | CAD情報取得 |
| POST | `/api/receive-by-purchase-order` | 発注番号で受入 |

### 3.4 検索

| メソッド | パス | 説明 |
|---------|------|------|
| GET | `/api/search-by-spec1/<spec1>` | 仕様１検索 |
| GET | `/api/search-by-purchase-order/<num>` | 発注番号検索 |
| GET | `/api/purchase-order-stats` | 発注統計 |

### 3.5 パレット管理

| メソッド | パス | 説明 |
|---------|------|------|
| GET | `/api/pallets/list` | パレット一覧 |
| GET | `/api/pallets/search` | パレット検索 |
| GET | `/api/pallets/<num>/label` | パレットラベル |
| GET | `/api/pallets/stats` | パレット統計 |

### 3.6 Across DB連携

| メソッド | パス | 説明 |
|---------|------|------|
| GET | `/api/across-db/test` | 接続テスト |
| GET | `/api/across-db/check-updates` | 更新チェック |
| GET | `/api/across-db/status` | DB状態取得 |
| GET | `/api/across-db/seiban-status/<seiban>` | 製番状態 |
| GET | `/api/across-db/delivery-schedule` | 納品スケジュール |
| GET | `/api/across-db/columns` | カラム一覧 |
| POST | `/api/across-db/query` | クエリ実行 |
| GET | `/api/across-db/order-detail` | 発注詳細 |
| POST | `/api/across-db/process` | DB直接処理 |
| GET | `/api/across-db/merge-test` | マージテスト |
| GET | `/api/across-db/mihatchu` | 未発注検索 |
| POST | `/api/across-db/zaiko-buhin` | 在庫部品検索 |
| GET | `/api/across-db/0zaiko` | 0ZAIKO検索 |

### 3.7 その他

| メソッド | パス | 説明 |
|---------|------|------|
| GET | `/` | メインページ |
| GET | `/api/debug-paths` | パスデバッグ |
| GET | `/api/seiban-list` | 製番一覧 |
| POST | `/api/detect-seibans` | 製番検出 |
| POST | `/api/refresh-seiban` | 製番更新 |
| GET | `/api/check-network-file` | ネットワークファイル確認 |
| POST | `/api/load-network-file` | ネットワークファイル読込 |
| POST | `/api/load-from-odbc` | ODBC読込 |
| GET | `/api/delivery-schedule` | 納品スケジュール |
| GET | `/api/orders/gantt-data` | ガントデータ |
| GET | `/api/check-update` | 更新確認 |
| GET | `/api/load-history` | 履歴読込 |
| POST | `/api/import-history` | 履歴インポート |
| GET | `/api/get-system-status` | システム状態 |
| GET | `/api/open-cad/<id>` | CADファイル開く |
| GET | `/api/open-cad-by-spec/<spec1>` | 仕様１でCAD開く |
| POST | `/api/order/<id>/send-completion-email` | 完了メール送信 |

---

## 4. ユーティリティモジュール

### 4.1 utils/constants.py - 定数定義

```python
class Constants:
    # メッキ関連
    MEKKI_SUPPLIER_CD = 116
    MEKKI_PATTERNS = [r'/Ni-P', r'/NiCr', ...]
    MEKKI_SPEC1_CODES = ['NMA-00017-00-00']

    # 手配区分CD
    ORDER_TYPE_BLANK = '13'      # 加工用ブランク
    ORDER_TYPE_PROCESSED = '11'  # 追加工
    ORDER_TYPE_STOCK = '15'      # 在庫部品

    # ステータス
    STATUS_BEFORE = '受入準備前'
    STATUS_IN_PROGRESS = '納品中'
    STATUS_COMPLETED = '納品完了'
```

### 4.2 utils/data_utils.py - データ処理

- `safe_str(value)`: 安全な文字列変換
- `safe_int(value, default)`: 安全な整数変換
- `normalize_order_number(order_number)`: 発注番号正規化（ゼロパディング除去）

### 4.3 utils/mekki_utils.py - メッキ判定

- `is_mekki_target(supplier_cd, spec2, spec1)`: メッキ出対象判定
- `add_mekki_alert(remarks)`: 備考にメッキ出アラート追加

### 4.4 utils/email_sender.py - メール送信

納品完了メール作成・メーラー起動

### 4.5 utils/excel_gantt_chart.py - ガントチャート

Excelシートにセルベースのガントチャートを描画

### 4.6 utils/delivery_utils.py - 検収データ

DB直接クエリに移行済み（互換性維持用スタブ）

---

## 5. サービスモジュール

### 5.1 services/cad_service.py - CADファイル操作

仕様１コード（例: NKA-00437-00-00）からCADファイル（mx2/pdf）を検索

パス規則:
```
\\SERVER3\Share-data\CadData\Parts\{アルファベット}\{仕様１}*.{mx2|pdf}
```

### 5.2 services/excel_export.py - Excel出力

- 手配発注リストExcel生成
- QRコード埋込
- ガントチャートシート作成
- COM経由Excel更新

### 5.3 services/cache_service.py - キャッシュ

ファイル情報キャッシュ

---

## 6. 外部連携

### 6.1 ネットワークパス

| 設定キー | パス | 用途 |
|---------|------|------|
| HISTORY_EXCEL_PATH | `\\server3\Share-data\Document\仕入れ\002_手配リスト\手配発注マージリスト発行履歴.xlsx` | 発行履歴 |
| SEIBAN_LIST_PATH | `\\server3\share-data\Document\Acrossデータ\製番一覧表.xlsx` | 製番一覧 |
| EXPORT_EXCEL_PATH | `\\SERVER3\Share-data\Document\仕入れ\002_手配リスト\手配発注リスト` | 出力先 |

### 6.2 CADファイルパス

```
\\SERVER3\Share-data\CadData\Parts\{A-Z}\
```

### 6.3 Across DB接続

```
DSN=Across;
USE acrossDB;
```

---

## 7. UI構成

### 7.1 メインタブ

1. **受注リスト** - 製番一覧、ガントチャート、ステータス管理
2. **データ処理** - Excel/DBからのデータ取込
3. **入荷管理** - QRスキャン、発注番号検索による受入
4. **ユニット保管場所** - パレット管理
5. **アーカイブ** - 完了済み製番
6. **発行履歴** - 処理履歴

### 7.2 外部ライブラリ

- Html5Qrcode: QRコード・バーコードスキャン
- Chart.js: ガントチャート描画
- chartjs-adapter-date-fns: 日付アダプタ
- chartjs-plugin-annotation: 今日線表示

---

## 8. データフロー

### 8.1 データ取込フロー

```
[Excel/Across DB]
      ↓
[/api/process] マージ処理
      ↓
[OrderDetail作成] 発注明細DB保存
      ↓
[Order作成/更新] 製番・ユニット単位で集約
      ↓
[Excel出力] 手配発注リスト生成
```

### 8.2 受入フロー

```
[QRスキャン or 発注番号入力]
      ↓
[/receive/<seiban>/<unit>] 受入ページ表示
      ↓
[/api/detail/<id>/toggle-receive] 受入トグル
      ↓
[ReceivedHistory記録] 履歴永続化
      ↓
[EditLog記録] 操作ログ
      ↓
[Excel更新] 手配発注リスト自動更新
```

### 8.3 マージロジック

1. **V_D手配リスト**（BOM）を基準に
2. **V_D発注**とマッチング
   - Primary: 材質 + 仕様１ + 製番
   - Fallback: 製番 + 仕様１ (+手配区分)
3. **V_D未発注**から社内加工品(MHT+11)を追加

---

## 9. 設定

### 9.1 config.py

```python
class Config:
    SECRET_KEY = 'dev-secret-key'
    SQLALCHEMY_DATABASE_URI = 'sqlite:///order_management.db'
    UPLOAD_FOLDER = 'uploads'
    MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB
    USE_ODBC = False
    USE_HTTPS = False

class DevelopmentConfig(Config):
    DEBUG = True
    USE_HTTPS = True
    SSL_CERT_PATH = 'cert.pem'
    SSL_KEY_PATH = 'key.pem'

class ProductionConfig(Config):
    DEBUG = False
    USE_ODBC = True
```

### 9.2 環境変数

- `FLASK_ENV`: development / production / testing
- `SECRET_KEY`: セッション暗号化キー
- `SSL_CERT_PATH`: SSL証明書パス
- `SSL_KEY_PATH`: SSL秘密鍵パス

---

## 10. ファイル構成

```
Order_merge/
├── app.py                    # メインアプリケーション
├── config.py                 # 設定ファイル
├── models.py                 # DBモデル（app.pyからの参照用）
├── across_db.py              # Across DBクエリモジュール
├── label_maker.py            # ラベル作成
├── utils/
│   ├── __init__.py
│   ├── constants.py          # 定数定義
│   ├── data_utils.py         # データ処理
│   ├── mekki_utils.py        # メッキ判定
│   ├── email_sender.py       # メール送信
│   ├── excel_styler.py       # Excelスタイル
│   ├── excel_gantt_chart.py  # ガントチャート
│   ├── qr_generator.py       # QRコード生成
│   └── delivery_utils.py     # 検収データ（スタブ）
├── services/
│   ├── __init__.py
│   ├── cad_service.py        # CADファイル操作
│   ├── excel_export.py       # Excel出力
│   └── cache_service.py      # キャッシュ
├── templates/
│   └── index.html            # メインテンプレート
├── static/
│   ├── styles.css
│   ├── gantt-chart.js
│   ├── qr-scanner.js
│   ├── pallet-manager.js
│   ├── delivery-schedule.js
│   └── across-db.js
├── instance/
│   └── order_management.db   # SQLiteデータベース
├── uploads/                  # アップロードファイル
├── exports/                  # 出力ファイル
├── cache/                    # キャッシュ
└── labels/                   # ラベル出力
```

---

## 11. 手配区分コード

| コード | 名称 | 説明 |
|--------|------|------|
| 11 | 追加工 | 社内加工品 |
| 13 | 加工用ブランク | ブランク材 |
| 15 | 在庫部品 | 在庫から払出 |

---

## 12. 特記事項

### 12.1 メッキ出判定

以下の条件でメッキ出対象と判定:
1. 仕入先CD = 116 かつ 仕様２にメッキパターン含む
2. 仕様１が `NMA-00017-00-00` の場合（仕入先に関係なく）

メッキパターン: `/Ni-P`, `/NiCr`, `Ｎｉ－Ｐ`, `Ｃｒ`, `ＮｉＣｒ`

### 12.2 発注番号正規化

- Excelからの読込時に浮動小数点化される問題への対処
- ゼロパディング（8桁）の除去/追加
- 例: `00086922` → `86922` → `00086922`（DB検索時）

### 12.3 階層構造

OrderDetailは`parent_id`による親子関係をサポート。
BOMの構成部品を階層表示可能。

### 12.4 受入履歴永続化

`ReceivedHistory`テーブルで発注番号ベースの受入状態を永続化。
Orderが削除・再作成されても受入状態を維持。
