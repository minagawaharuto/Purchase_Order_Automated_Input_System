import pandas as pd

def inspect_full(file_path, sheet_name):
    print(f"\n--- Detailed Inspection of {file_path} [Sheet: {sheet_name}] ---")
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    # Print first 50 rows to see the layout
    for i, row in df.iterrows():
        print(f"Row {i:2}: {row.tolist()}")
        if i >= 60:
            break

if __name__ == "__main__":
    inspect_full('G2828-02　きなこがでろーんTシャツ.xlsx', 'サンプル発注書')
