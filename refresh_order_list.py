# refresh_order_list.py
import time
import win32com.client as win32
from datetime import datetime
import os
import sys
import schedule

BOOK_PATH = r"\\server3\Share-data\Document\仕入れ\002_手配リスト\手配発注_ALL.xlsx"
LOG_PATH = r"\\server3\Share-data\Document\仕入れ\002_手配リスト\更新ログ_手配発注_ALL.txt"

# 前回の件数を保持する辞書
previous_counts = {"sheet1": None, "sheet2": None}

def get_next_serial_no(log_path):
    """次のシリアル番号を取得"""
    if not os.path.exists(log_path):
        return "0001"
    
    try:
        with open(log_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            if lines:
                # 最後の行から現在のシリアル番号を取得
                last_line = lines[-1].strip()
                if last_line:
                    current_serial = int(last_line.split('|')[0])
                    return f"{current_serial + 1:04d}"
    except:
        pass
    
    return "0001"

def get_previous_counts(log_path):
    """前回の件数を取得"""
    if not os.path.exists(log_path):
        return None, None
    
    try:
        with open(log_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            if lines:
                last_line = lines[-1].strip()
                if last_line:
                    parts = last_line.split('|')
                    if len(parts) >= 4:
                        sheet1_count = int(parts[2])
                        sheet2_count = int(parts[3])
                        return sheet1_count, sheet2_count
    except:
        pass
    
    return None, None

def get_last_person(ws, person_column, last_row):
    """指定されたシートの最終行の担当者を取得"""
    try:
        if last_row > 1:  # ヘッダー行を除いてデータがある場合
            person = ws.Cells(last_row, person_column).Value
            if person:
                return str(person).strip()
    except:
        pass
    return None

def refresh_excel():
    """Excelファイルを更新する関数"""
    global previous_counts
    
    try:
        excel = win32.DispatchEx("Excel.Application")   # 新規インスタンス
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(BOOK_PATH, UpdateLinks=3)  # 3=外部リンク自動更新
        
        # 更新前のデータ件数を取得
        sheet1_count = 0
        sheet2_count = 0
        
        try:
            # シート1のデータ件数を取得
            ws1 = wb.Worksheets(1)
            last_row1 = ws1.Cells(ws1.Rows.Count, 1).End(-4162).Row  # -4162 = xlUp
            sheet1_count = max(0, last_row1 - 1)  # ヘッダー行を除く
        except:
            sheet1_count = 0
        
        try:
            # シート2のデータ件数を取得
            ws2 = wb.Worksheets(2)
            last_row2 = ws2.Cells(ws2.Rows.Count, 1).End(-4162).Row  # -4162 = xlUp
            sheet2_count = max(0, last_row2 - 1)  # ヘッダー行を除く
        except:
            sheet2_count = 0
        
        wb.RefreshAll()                                      # 全接続を更新
        excel.CalculateUntilAsyncQueriesDone()               # 非同期クエリ完了待ち
        
        # 更新後のデータ件数と担当者を取得
        sheet1_last_person = None
        sheet2_last_person = None
        
        try:
            # シート1の更新後データ件数と担当者
            last_row1_after = ws1.Cells(ws1.Rows.Count, 1).End(-4162).Row
            sheet1_count_after = max(0, last_row1_after - 1)
            sheet1_last_person = get_last_person(ws1, 3, last_row1_after)  # C列 = 3
        except:
            sheet1_count_after = sheet1_count
        
        try:
            # シート2の更新後データ件数と担当者
            last_row2_after = ws2.Cells(ws2.Rows.Count, 1).End(-4162).Row
            sheet2_count_after = max(0, last_row2_after - 1)
            sheet2_last_person = get_last_person(ws2, 4, last_row2_after)  # D列 = 4
        except:
            sheet2_count_after = sheet2_count
        
        wb.Save()                                            # 上書き保存
        wb.Close(False)
        excel.Quit()
        
        # 前回の件数を取得（初回実行時）
        if previous_counts["sheet1"] is None or previous_counts["sheet2"] is None:
            prev_sheet1, prev_sheet2 = get_previous_counts(LOG_PATH)
            if prev_sheet1 is not None:
                previous_counts["sheet1"] = prev_sheet1
                previous_counts["sheet2"] = prev_sheet2
        
        # 件数の変化をチェック
        sheet1_diff = 0
        sheet2_diff = 0
        if previous_counts["sheet1"] is not None and previous_counts["sheet2"] is not None:
            sheet1_diff = sheet1_count_after - previous_counts["sheet1"]
            sheet2_diff = sheet2_count_after - previous_counts["sheet2"]
        
        # 増加した場合の担当者を決定
        last_person = None
        if sheet1_diff > 0 and sheet1_last_person:
            last_person = sheet1_last_person
        elif sheet2_diff > 0 and sheet2_last_person:
            last_person = sheet2_last_person
        
        # ログを記録
        serial_no = get_next_serial_no(LOG_PATH)
        update_time = datetime.now().strftime('%y/%m/%d-%H:%M')
        log_entry = f"{serial_no}|{update_time}|{sheet1_count_after}|{sheet2_count_after}"
        
        # 件数が増加した場合は担当者情報を追加
        if last_person and (sheet1_diff > 0 or sheet2_diff > 0):
            log_entry += f"|{last_person}"
        
        with open(LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(log_entry + '\n')
        
        print(" No | 日付 |手配件数|発注件数|担当者")
        print(f"更新完了: {log_entry}")
        
        # 件数の変化をチェックして通知
        if previous_counts["sheet1"] is not None and previous_counts["sheet2"] is not None:
            if sheet1_diff != 0 or sheet2_diff != 0:
                print("\n" + "="*50)
                print("【件数変化通知】")
                
                if sheet1_diff != 0:
                    change_type = "増加" if sheet1_diff > 0 else "減少"
                    print(f"  手配: {abs(sheet1_diff)}件{change_type} "
                          f"({previous_counts['sheet1']}件 → {sheet1_count_after}件)")
                    if sheet1_diff > 0 and sheet1_last_person:
                        print(f"  末尾にある担当者: {sheet1_last_person}")
                
                if sheet2_diff != 0:
                    change_type = "増加" if sheet2_diff > 0 else "減少"
                    print(f"  発注: {abs(sheet2_diff)}件{change_type} "
                          f"({previous_counts['sheet2']}件 → {sheet2_count_after}件)")
                    if sheet2_diff > 0 and sheet2_last_person:
                        print(f"  末尾にある担当者: {sheet2_last_person}")
                
                print("="*50 + "\n")
        
        # 現在の件数を保存
        previous_counts["sheet1"] = sheet1_count_after
        previous_counts["sheet2"] = sheet2_count_after
        
        return True
        
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        # Excelが開いている場合は終了
        try:
            excel.Quit()
        except:
            pass
        return False

def run_scheduled():
    """スケジュール実行用の関数"""
    print("自動更新を開始します。1時間ごとに更新を実行します。")
    print("終了するには Ctrl+C を押してください。")
    
    # 初回実行
    refresh_excel()
    
    # 1時間ごとにスケジュール
    schedule.every(1).hours.do(refresh_excel)
    
    while True:
        try:
            schedule.run_pending()
            time.sleep(60)  # 1分ごとにチェック
        except KeyboardInterrupt:
            print("\n自動更新を終了します。")
            break

def main():
    """メイン関数"""
    if len(sys.argv) > 1 and sys.argv[1] == "--auto":
        # 自動実行モード
        run_scheduled()
    else:
        # 単発実行
        refresh_excel()

if __name__ == "__main__":
    main()
