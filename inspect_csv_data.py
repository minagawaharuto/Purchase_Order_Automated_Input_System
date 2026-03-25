import pandas as pd

def inspect_data(file_path):
    print(f"\n--- Detailed Inspection of {file_path} ---")
    df = pd.read_excel(file_path)
    cols = ['Name', 'Created at', 'Lineitem name', 'Lineitem quantity', 'Lineitem sku']
    print(df[cols].head(20))

if __name__ == "__main__":
    inspect_data('CSVデータ.xlsx')
