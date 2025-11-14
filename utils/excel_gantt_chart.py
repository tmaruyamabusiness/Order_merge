"""
ガントチャート作成関数
- Y軸左側にユニット名を表示
- X軸下部に日数を表示し、チャート右側に日付対応表を配置
"""

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime, timedelta

def create_gantt_chart_sheet(wb, seiban, orders):
    """
    ガントチャートシートを作成
    
    Parameters:
    - wb: openpyxl Workbook オブジェクト
    - seiban: 製番（例: MHT0614）
    - orders: Order オブジェクトのリスト
    """
    try:
        ws = wb.create_sheet("納期ガントチャート", 0)
        
        # ========================================
        # 1. ヘッダー設定
        # ========================================
        ws['A1'] = f"{seiban} 納期ガントチャート"
        ws['A1'].font = Font(size=20, bold=True, color="1F4E78")
        ws.merge_cells('A1:H1')
        ws.row_dimensions[1].height = 35
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        
        # ========================================
        # 2. データ収集
        # ========================================
        unit_data = []
        base_date = None
        
        for order in orders:
            unit_name = order.unit if order.unit else 'ユニット名無し'
            
            print(f"  処理中: {seiban}_{unit_name}, details数={len(order.details) if order.details else 0}")
            
            if not order.details:
                print(f"    ⚠️  スキップ: detailsが空")
                continue
            
            # 納期を持つ詳細のみ抽出
            dates = []
            for detail in order.details:
                if detail.delivery_date:
                    try:
                        date_str = detail.delivery_date.strip()
                        
                        if not date_str or date_str == '-':
                            continue
                        
                        if '/' in date_str:
                            parts = date_str.split('/')
                            if len(parts) == 3:
                                year = int(parts[0])
                                if year < 100:
                                    year = 2000 + year if year < 50 else 1900 + year
                                month = int(parts[1])
                                day = int(parts[2])
                                date_obj = datetime(year, month, day)
                                dates.append(date_obj)
                    except Exception as e:
                        print(f"    ⚠️  日付パースエラー ({date_str}): {e}")
                        continue
            
            print(f"    有効な納期: {len(dates)}件")
            
            if dates:
                min_date = min(dates)
                max_date = max(dates)
                
                if base_date is None:
                    base_date = min_date
                else:
                    base_date = min(base_date, min_date)
                
                unit_data.append({
                    'unit': unit_name,
                    'min_date': min_date,
                    'max_date': max_date,
                    'count': len(order.details)
                })
                print(f"    ✅ ガントチャートに追加: {min_date.strftime('%Y/%m/%d')} ～ {max_date.strftime('%Y/%m/%d')}")
            else:
                print(f"    ⚠️  スキップ: 有効な納期データなし")
        
        if not unit_data or base_date is None:
            ws['A5'] = '納期データがありません'
            ws['A5'].font = Font(size=14, color="FF0000", bold=True)
            print(f"⚠️  {seiban}: ガントチャート作成不可（データなし）")
            return
        
        # ========================================
        # 3. 基準日情報を表示
        # ========================================
        max_days = max((data['max_date'] - base_date).days for data in unit_data)
        end_date = base_date + timedelta(days=max_days)
        
        ws['A2'] = f"基準日: {base_date.strftime('%Y/%m/%d')}"
        ws['A2'].font = Font(size=12, bold=True, color="FF0000")
        
        ws['D2'] = f"期間: {base_date.strftime('%Y/%m/%d')} ～ {end_date.strftime('%Y/%m/%d')} （{max_days + 1}日間）"
        ws['D2'].font = Font(size=11, italic=True, color="1F4E78")
        ws.merge_cells('D2:G2')
        
        # ========================================
        # 4. 列ヘッダー
        # ========================================
        headers = ['ユニット', '最早納期', '最遅納期', '期間(日)', '開始日数', '件数']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col, value=header)
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin', color="000000"),
                right=Side(style='thin', color="000000"),
                top=Side(style='thin', color="000000"),
                bottom=Side(style='thin', color="000000")
            )
        
        # ========================================
        # 5. データ行を書き込み
        # ========================================
        row_idx = 5
        for data in sorted(unit_data, key=lambda x: x['min_date']):
            start_days = (data['min_date'] - base_date).days
            duration = max(1, (data['max_date'] - data['min_date']).days + 1)
            
            row_values = [
                data['unit'],
                data['min_date'].strftime('%Y/%m/%d'),
                data['max_date'].strftime('%Y/%m/%d'),
                duration,
                start_days,
                data['count']
            ]
            
            for col, value in enumerate(row_values, 1):
                cell = ws.cell(row=row_idx, column=col, value=value)
                cell.alignment = Alignment(
                    horizontal='center' if col > 1 else 'left',
                    vertical='center'
                )
                cell.border = Border(
                    left=Side(style='thin', color="CCCCCC"),
                    right=Side(style='thin', color="CCCCCC"),
                    top=Side(style='thin', color="CCCCCC"),
                    bottom=Side(style='thin', color="CCCCCC")
                )
                
                # 期間が長い場合は背景色を変更
                if col == 4 and value > 10:
                    cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            
            row_idx += 1
        
        # ========================================
        # 6. ガントチャート作成
        # ========================================
        chart = BarChart()
        chart.type = "bar"
        chart.style = 12
        chart.title = f"{seiban} 納期ガントチャート"
        
        # Y軸（ユニット名）の設定
        chart.y_axis.title = 'ユニット'
        chart.y_axis.delete = False
        
        # X軸（日数）の設定
        chart.x_axis.title = f'日数（基準日: {base_date.strftime("%Y/%m/%d")}）'
        chart.x_axis.delete = False
        
        # カテゴリ軸（ユニット名）を設定
        cats_ref = Reference(ws, min_col=1, min_row=5, max_row=row_idx-1)
        chart.set_categories(cats_ref)
        
        # 開始日数のデータ（透明バー）
        start_ref = Reference(ws, min_col=5, min_row=4, max_row=row_idx-1)
        chart.add_data(start_ref, titles_from_data=True)
        
        # 期間のデータ（色付きバー）
        duration_ref = Reference(ws, min_col=4, min_row=4, max_row=row_idx-1)
        chart.add_data(duration_ref, titles_from_data=True)
        
        # スタック表示
        chart.grouping = "stacked"
        chart.overlap = 100
        
        # 系列の色設定
        if len(chart.series) >= 1:
            chart.series[0].graphicalProperties.solidFill = "FFFFFF"
            chart.series[0].graphicalProperties.line.noFill = True
        
        if len(chart.series) >= 2:
            chart.series[1].graphicalProperties.solidFill = "4472C4"
        
        chart.legend = None
        
        # チャートサイズ調整
        chart.width = 22
        chart.height = max(12, len(unit_data) * 1.5)
        
        # グリッド線を設定
        chart.x_axis.majorGridlines = None
        chart.y_axis.majorGridlines = None
        
        # チャートを配置
        ws.add_chart(chart, "H5")
        
        # ========================================
        # 7. 日付対応表をチャート右側に追加
        # ========================================
        from openpyxl.utils import get_column_letter
        
        date_table_col = 26  # Z列
        date_table_row = 5
        
        # ヘッダー
        ws.cell(row=date_table_row, column=date_table_col, value="日数→日付").font = Font(bold=True, size=11, color="1F4E78")
        ws.cell(row=date_table_row + 1, column=date_table_col, value="日数").font = Font(bold=True, size=10)
        ws.cell(row=date_table_row + 1, column=date_table_col + 1, value="日付").font = Font(bold=True, size=10)
        
        # 5日ごとに日付を表示
        table_row = date_table_row + 2
        for i in range(0, max_days + 1, 5):
            current_date = base_date + timedelta(days=i)
            ws.cell(row=table_row, column=date_table_col, value=i)
            ws.cell(row=table_row, column=date_table_col + 1, value=current_date.strftime('%m/%d'))
            table_row += 1
        
        # 列幅調整（get_column_letterを使用）
        ws.column_dimensions[get_column_letter(date_table_col)].width = 8
        ws.column_dimensions[get_column_letter(date_table_col + 1)].width = 10
        
        # ========================================
        # 8. 列幅調整
        # ========================================
        ws.column_dimensions['A'].width = 22
        ws.column_dimensions['B'].width = 12
        ws.column_dimensions['C'].width = 12
        ws.column_dimensions['D'].width = 10
        ws.column_dimensions['E'].width = 10
        ws.column_dimensions['F'].width = 8
        
        print(f"✅ ガントチャート作成完了: {len(unit_data)}ユニット、基準日={base_date.strftime('%Y/%m/%d')}")
        
    except Exception as e:
        print(f"⚠️  ガントチャート作成エラー: {e}")
        import traceback
        traceback.print_exc()


# ========================================
# テスト用コード
# ========================================
if __name__ == '__main__':
    class DummyDetail:
        def __init__(self, delivery_date):
            self.delivery_date = delivery_date
    
    class DummyOrder:
        def __init__(self, unit, dates):
            self.unit = unit
            self.details = [DummyDetail(d) for d in dates]
    
    # ダミーデータ作成
    orders = [
        DummyOrder('カッター', ['25/10/14', '25/10/20', '25/10/30']),
        DummyOrder('架台補修', ['25/10/27']),
        DummyOrder('カッター昇降', ['25/10/31'])
    ]
    
    # Excelファイル作成
    wb = Workbook()
    wb.remove(wb.active)
    
    create_gantt_chart_sheet(wb, 'MHT0614', orders)
    
    wb.save(r'\\SERVER3\Share-data\Document\仕入れ\002_手配リスト\手配発注リスト')
    print("✅ テストファイル作成完了: gantt_final.xlsx")