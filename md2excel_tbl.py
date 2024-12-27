"""
このスクリプトは、Markdownファイルを読み込み、Excelのようなグリッド形式に変換し、結果をMarkdown形式で出力します。
環境変数 'FILES_OUTMD_DIR' と 'FILES_OUTTBL_DIR' を使用して、入力ディレクトリと出力ディレクトリを指定します。
"""

import os
import pandas as pd
from pathlib import Path

# Markdownの内容を解析してテーブルに変換する関数を定義
def parse_markdown_to_table(markdown_content):
    lines = markdown_content.split('\n')
    tables = {}
    current_sheet = None
    current_data = []

    for line in lines:
        if line.startswith('## '):
            if current_sheet and current_data:
                tables[current_sheet] = pd.DataFrame(current_data, columns=["Cell Position", "Cell Data"])
            current_sheet = line.replace('## ', '').strip()
            current_data = []
        elif line.startswith('### ') and current_sheet:
            cell_info = line.replace('### ', '').split(':', 1)
            if len(cell_info) == 2:
                cell_position = cell_info[0].strip()
                cell_data = cell_info[1].strip()
                current_data.append([cell_position, cell_data])

    if current_sheet and current_data:
        tables[current_sheet] = pd.DataFrame(current_data, columns=["Cell Position", "Cell Data"])
    return tables

# Excelのようなグリッドを作成する関数を定義
def create_excel_like_grid(tables):
    grids = {}

    for sheet_name, df in tables.items():
        positions = df["Cell Position"].str.extract(r'([A-Z]+)(\d+)')
        df["Column"] = positions[0].apply(lambda col: sum((ord(c) - 64) * (26 ** i) for i, c in enumerate(reversed(col))))
        df["Row"] = positions[1].astype(int)

        grid = pd.DataFrame(index=range(1, df["Row"].max() + 1), columns=range(1, df["Column"].max() + 1))
        for _, row in df.iterrows():
            grid.loc[row["Row"], row["Column"]] = row["Cell Data"]

        grids[sheet_name] = grid.fillna("")
    return grids

# ディレクトリを処理するメイン関数
def process_markdown_directory(input_dir, output_dir):
    input_path = Path(input_dir)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    for markdown_file in input_path.glob("*.md"):
        with open(markdown_file, 'r', encoding='utf-8') as file:
            content = file.read()

        tables = parse_markdown_to_table(content)
        grids = create_excel_like_grid(tables)

        output_file = output_path / markdown_file.name
        with open(output_file, 'w', encoding='utf-8') as outfile:
            if "シート構成サマリ" in tables:
                outfile.write("## シート構成サマリ\n")
                outfile.write(tables["シート構成サマリ"].to_markdown(index=False))
                outfile.write("\n\n")

            for sheet_name, grid in grids.items():
                outfile.write(f"## {sheet_name}\n")
                outfile.write(grid.to_markdown(index=True))
                outfile.write("\n\n")

# 環境変数を使用してスクリプトを実行
if __name__ == "__main__":
    input_dir = os.getenv("FILES_OUTMD_DIR")
    output_dir = os.getenv("FILES_OUTTBL_DIR")

    if not input_dir or not output_dir:
        print("Error: Please set the environment variables 'FILES_OUTMD_DIR' and 'FILES_OUTTBL_DIR'.")
    else:
        process_markdown_directory(input_dir, output_dir)
