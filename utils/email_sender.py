"""
Email sender module for order completion notifications
納品完了メール送信モジュール
"""

import urllib.parse
import webbrowser
from typing import Optional

class EmailSender:
    """メール送信クラス"""
    
    # デフォルト宛先
    DEFAULT_TO_RECIPIENTS = [
        'y_takahashi@dangan-v.com',
        'a_tatsumi@dangan-v.com',
        'k_horie@dangan-v.com',
        'k_nakahara@dangan-v.com'
    ]
    
    DEFAULT_CC_RECIPIENTS = [
        'y_maruyama@dangan-v.com'  # CCに追加

    ]
    
    @staticmethod
    def create_completion_email(
        seiban: str,
        product_name: str = '',
        customer_abbr: str = '',
        unit: str = '',
        memo2: str = '',
        floor: str = '',
        pallet_number: str = '',
        excel_path: str = '',
        sender_name: str = '丸山'
    ) -> dict:
        """
        納品完了メールのテンプレートを作成
        
        Args:
            seiban: 製番
            product_name: 品名
            customer_abbr: 得意先略称
            unit: ユニット名
            memo2: メモ２
            floor: フロア情報
            pallet_number: パレット番号
            excel_path: 実際のExcelファイルパス（app.pyから取得）
            sender_name: 送信者名（デフォルト: 丸山）
        
        Returns:
            dict: {subject: 件名, body: 本文, to: 宛先}
        """
        
        # 🔥 件名を簡潔に（製番のみ）
        subject = f"【納品完了】{seiban}_{unit if unit else '（ユニット名無し）'}({customer_abbr})"
        
        # 🔥 概要作成（product_nameを優先、なければmemo2）
        overview = product_name if product_name else (memo2 if memo2 else '（情報なし）')
        
        # 🔥 場所情報（改行付き）
        location_parts = []
        if floor:
            location_parts.append(floor)
        if pallet_number:
            location_parts.append(f"パレット{pallet_number}")
        
        location_text = ''.join(location_parts) if location_parts else '未設定'
        
        # パレットラベルの注意書き
        pallet_note = '\n※パレット側面にラベル貼付有' if location_parts else ''
        
        # 🔥 excel_pathが渡されていない場合のフォールバック
        if not excel_path:
            excel_path = r'\\SERVER3\Share-data\Document\仕入れ\002_手配リスト\手配発注リスト'
        
        # 🔥 本文生成（箇条書き形式、改行を整理）
        body = f"""各位
お疲れ様です、{sender_name}です

下記製番のユニットの納品が完了しましたのでご連絡します。

製番：{seiban}
概要：{overview}
ユニット：{unit if unit else '（ユニット名指定なし）'}
客先：{customer_abbr if customer_abbr else '（指定なし）'}

📦️場所：{location_text}{pallet_note}
※組立方法や組立開始時期については各担当者との相談をお願いします。

💻 受入内容確認ページ
{excel_path}

【リストの見方】
🔴ピンク色系で塗られているのものは受入済み、灰色か白で塗られているものは未受入です。
ただし、グルア室外にある在庫部品は各自で必要に応じて在庫棚よりピッキングをお願いします。

⚠️閲覧後は必ず閉じてください（データ更新ができなくなります)

以上、よろしくお願いします"""
        
        return {
            'subject': subject,
            'body': body,
            'to': ','.join(EmailSender.DEFAULT_TO_RECIPIENTS),  # 🔥 TOのみ
            'cc': ','.join(EmailSender.DEFAULT_CC_RECIPIENTS)   # 🔥 CC追加
        }
    
    @staticmethod
    def open_email_client(subject: str, body: str, to: str, cc: str = '') -> bool:
        """
        デフォルトメーラーを起動してメールを作成
        
        Args:
            subject: 件名
            body: 本文
            to: 宛先（カンマ区切り）
            cc: CC（カンマ区切り）
        
        Returns:
            bool: 成功したかどうか
        """
        try:
            # URLエンコード（改行を%0D%0Aに変換）
            encoded_subject = urllib.parse.quote(subject)
            encoded_body = urllib.parse.quote(body)
            encoded_to = urllib.parse.quote(to)
            
            # 🔥 CC用のURLパラメータを追加
            mailto_url = f"mailto:{encoded_to}?subject={encoded_subject}&body={encoded_body}"
            if cc:
                encoded_cc = urllib.parse.quote(cc)
                mailto_url += f"&cc={encoded_cc}"
            
            # デフォルトメーラーを起動
            webbrowser.open(mailto_url)
            
            return True
        except Exception as e:
            print(f"❌ メーラー起動エラー: {e}")
            return False
    
    @staticmethod
    def send_completion_notification(
        seiban: str,
        product_name: str = '',
        customer_abbr: str = '',
        unit: str = '',
        memo2: str = '',
        floor: str = '',
        pallet_number: str = '',
        excel_path: str = '',
        sender_name: str = '丸山'
    ) -> bool:
        """
        納品完了メールを作成してメーラーを起動（オールインワン関数）
        
        Args:
            seiban: 製番
            product_name: 品名
            customer_abbr: 得意先略称
            unit: ユニット名
            memo2: メモ２
            floor: フロア情報
            pallet_number: パレット番号
            excel_path: 実際のExcelファイルパス（app.pyから取得）
            sender_name: 送信者名（デフォルト: 丸山）
        
        Returns:
            bool: 成功したかどうか
        """
        # メールテンプレート作成
        email_data = EmailSender.create_completion_email(
            seiban=seiban,
            product_name=product_name,
            customer_abbr=customer_abbr,
            unit=unit,
            memo2=memo2,
            floor=floor,
            pallet_number=pallet_number,
            excel_path=excel_path,
            sender_name=sender_name
        )
        
        # メーラー起動
        return EmailSender.open_email_client(
            subject=email_data['subject'],
            body=email_data['body'],
            to=email_data['to'],
            cc=email_data['cc']
        )


# テスト用コード
if __name__ == '__main__':
    # テスト実行
    test_path = r'\\SERVER3\Share-data\Document\仕入れ\002_手配リスト\手配発注リスト\MHT0614_カッターユニット MHT483代替え_ナリコマ_神戸_手配発注リスト.xlsx'
    
    success = EmailSender.send_completion_notification(
        seiban='MHT0614',
        product_name='カッター',
        customer_abbr='ナリコマ',
        unit='カッタ',
        memo2='カッターユニット　MHT483代替え',
        floor='1F',
        pallet_number='P001',
        excel_path=test_path,
        sender_name='丸山'
    )
    
    if success:
        print("✅ メーラーを起動しました")
        print(f"\nファイルパス: {test_path}")
    else:
        print("❌ メーラー起動に失敗しました")