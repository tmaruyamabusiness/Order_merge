"""
CADファイル操作 - cad_service.py
仕様1からCADファイル情報を取得
"""
import os
import glob


def get_cad_file_info(spec1):
    """仕様1からCADファイル情報を取得"""
    if not spec1 or not spec1.startswith('N'):
        return None

    # NKA-00437-00-00 → K を抽出
    parts = spec1.split('-')
    if len(parts) < 2 or len(parts[0]) < 2:
        return None

    # 2文字目のアルファベットを取得
    folder_letter = parts[0][1].upper()

    # CADフォルダパス
    cad_folder = f"\\\\SERVER3\\Share-data\\CadData\\Parts\\{folder_letter}"

    # ファイルを検索（ワイルドカード）
    mx2_pattern = os.path.join(cad_folder, f"{spec1}*.mx2")
    pdf_pattern = os.path.join(cad_folder, f"{spec1}*.pdf")

    mx2_files = glob.glob(mx2_pattern)
    pdf_files = glob.glob(pdf_pattern)

    if not mx2_files and not pdf_files:
        return None

    return {
        'folder': cad_folder,
        'letter': folder_letter,
        'spec1': spec1,
        'mx2_files': mx2_files,
        'pdf_files': pdf_files,
        'has_mx2': len(mx2_files) > 0,
        'has_pdf': len(pdf_files) > 0
    }
