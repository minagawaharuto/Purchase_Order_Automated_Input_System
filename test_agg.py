import pandas as pd

def aggregate_product(file_path, product_base_name):
    df = pd.read_excel(file_path)
    # Filter by product name (ignoring variants)
    target_df = df[df['Lineitem name'].str.contains(product_base_name, na=False)].copy()
    
    # Extract Color and Size
    # Format: "Name - Color / Size"
    def parse_variant(name):
        if ' - ' in name:
            parts = name.split(' - ')
            base = parts[0]
            variant = parts[1]
            if ' / ' in variant:
                v_parts = variant.split(' / ')
                return v_parts[0].strip(), v_parts[1].strip()
        return None, None

    target_df[['Color', 'Size']] = target_df['Lineitem name'].apply(lambda x: pd.Series(parse_variant(x)))
    
    # Aggregate
    agg = target_df.groupby(['Color', 'Size'])['Lineitem quantity'].sum().unstack(fill_value=0)
    print(f"\n--- Aggregation for '{product_base_name}' ---")
    print(agg)

if __name__ == "__main__":
    aggregate_product('CSVデータ.xlsx', 'きなこがでろーんTシャツ')
    aggregate_product('CSVデータ.xlsx', 'でっかく前にナノときなこTシャツ')
