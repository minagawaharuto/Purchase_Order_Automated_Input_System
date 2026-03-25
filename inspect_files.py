import pandas as pd
import sys

def inspect_excel(file_path):
    print(f"\n--- Inspecting {file_path} ---")
    try:
        xl = pd.ExcelFile(file_path)
        print(f"Sheets: {xl.sheet_names}")
        for sheet in xl.sheet_names:
            print(f"\n[Sheet: {sheet}]")
            df = pd.read_excel(file_path, sheet_name=sheet)
            print(f"Columns: {df.columns.tolist()}")
            print(df.head(10))
    except Exception as e:
        print(f"Error reading {file_path}: {e}")

if __name__ == "__main__":
    try:
        import pandas as pd
        import openpyxl
    except ImportError as e:
        print(f"Missing library: {e}")
        sys.exit(1)

    inspect_excel('CSVデータ.xlsx')
    inspect_excel('G2828-02　きなこがでろーんTシャツ.xlsx')
