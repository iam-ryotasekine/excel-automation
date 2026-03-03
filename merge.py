import pandas as pd
import glob
import os

# --- 設定エリア ---
# 読み込むフォルダ名と、出来上がるファイル名
target_folder = "sales_data"
output_file = "完成版_マスターデータ.xlsx"

print("🔄 Excelファイルの合体を開始します...")

# 1. フォルダの中にある「.xlsx」ファイルを全部探し出す
file_paths = glob.glob(os.path.join(target_folder, "*.xlsx"))

# 2. 見つけたファイルを順番に読み込んで、1つの箱にまとめる
all_data = []
for file in file_paths:
    # Excelを読み込む
    df = pd.read_excel(file)
    
    # オマケ機能：誰のデータか分かるように、ファイル名（Aさん_売上.xlsxなど）を新しい列に追加
    df['元ファイル名'] = os.path.basename(file)
    
    all_data.append(df)

# 3. バラバラのデータを縦にガッチャンコ！（合体）
merged_df = pd.concat(all_data, ignore_index=True)

# 4. 新しいExcelファイルとして書き出す
merged_df.to_excel(output_file, index=False)

print(f"✨ 完了！ {len(file_paths)}個のファイルを一瞬で合体させて '{output_file}' を作成しました！")