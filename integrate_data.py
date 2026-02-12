import pandas as pd
import glob
import os
import datetime
import re

# Configuration
INPUT_DIR = 'csvData'
TODAY = datetime.date.today().strftime('%Y-%m-%d')
OUTPUT_CSV = f'学習記録_統合_{TODAY}.csv'
OUTPUT_EXCEL = f'学習記録_統合_{TODAY}.xlsx'

def main():
    # 1. Gather all CSV files
    all_files = glob.glob(os.path.join(INPUT_DIR, '**', '*.csv'), recursive=True)
    
    if not all_files:
        print("No CSV files found.")
        return

    df_list = []
    print(f"Found {len(all_files)} CSV files. Processing...")
    
    for filename in all_files:
        try:
            # Try reading with different encodings if necessary
            try:
                df = pd.read_csv(filename)
            except UnicodeDecodeError:
                df = pd.read_csv(filename, encoding='cp932')
            
            df_list.append(df)
        except Exception as e:
            print(f"Error reading {filename}: {e}")

    if not df_list:
        print("No valid dataframes found.")
        return

    # 2. Combine DataFrames
    full_df = pd.concat(df_list, ignore_index=True)
    
    # 3. Data Cleaning & Normalization
    numeric_cols = ['問題数', '正答数']
    for col in numeric_cols:
        full_df[col] = pd.to_numeric(full_df[col], errors='coerce').fillna(0)
    
    full_df['日付'] = pd.to_datetime(full_df['日付'], errors='coerce')

    # Normalize Student ID (first 6 chars)
    full_df['学籍番号_norm'] = full_df['学籍番号'].astype(str).str[:6]

    # Normalize Name (remove whitespace)
    full_df['氏名_norm'] = full_df['氏名'].astype(str).str.replace(r'\s+', '', regex=True)

    # 4. Name Unification
    # Create valid ID -> Name map
    # Logic: For each normalized ID, find the "best" name.
    # "Best" = strictly contains Kanji if possible, otherwise most frequent.
    
    # Helper to check if string has Kanji
    def has_kanji(text):
        return bool(re.search(r'[\u4e00-\u9faf]', text))

    # Get unique (ID, Name) pairs count
    name_counts = full_df.groupby(['学籍番号_norm', '氏名_norm']).size().reset_index(name='count')
    
    id_name_map = {}
    for pid in name_counts['学籍番号_norm'].unique():
        candidates = name_counts[name_counts['学籍番号_norm'] == pid]
        
        # Filter for Kanji names
        kanji_candidates = candidates[candidates['氏名_norm'].apply(has_kanji)]
        
        if not kanji_candidates.empty:
            # Pick most frequent among Kanji names
            best_name = kanji_candidates.sort_values('count', ascending=False).iloc[0]['氏名_norm']
        else:
            # Pick most frequent among all
            best_name = candidates.sort_values('count', ascending=False).iloc[0]['氏名_norm']
            
        id_name_map[pid] = best_name

    # Apply unified name
    full_df['氏名'] = full_df['学籍番号_norm'].map(id_name_map)
    full_df['学籍番号'] = full_df['学籍番号_norm'] # Use normalized ID as the official ID

    # 5. Aggregation (Combine records for same student, day, field)
    # Group by relevant columns
    group_cols = ['学籍番号', '氏名', '日付', '分野']
    
    aggregated_df = full_df.groupby(group_cols)[numeric_cols].sum().reset_index()
    
    # Recalculate percentage
    aggregated_df['正答率(%)'] = (aggregated_df['正答数'] / aggregated_df['問題数'] * 100).fillna(0).round().astype(int)
    
    # 6. Export Flat CSV (Detail Data)
    # Format date for CSV output
    output_flat = aggregated_df.copy()
    output_flat['日付'] = output_flat['日付'].dt.strftime('%Y/%-m/%-d')
    output_flat.to_csv(OUTPUT_CSV, index=False, encoding='utf-8-sig')
    print(f"Exported flat CSV to {OUTPUT_CSV}")

    # 7. Create Matrix for Excel
    pivot_df = aggregated_df.pivot_table(
        index=['学籍番号', '氏名'],
        columns=['日付', '分野'],
        values=['問題数', '正答数', '正答率(%)'],
        aggfunc='sum'
    )
    
    # Swap levels to get Date, Field, ValueType
    pivot_df.columns = pivot_df.columns.swaplevel(0, 1) # Date, ValueType, Field
    pivot_df.columns = pivot_df.columns.swaplevel(1, 2) # Date, Field, ValueType
    
    # Rename ValueTypes
    rename_map = {'問題数': '問', '正答数': '正', '正答率(%)': '率'}
    pivot_df.rename(columns=rename_map, level=2, inplace=True)
    
    # Sort columns
    pivot_df.sort_index(axis=1, level=[0, 1], inplace=True)

    # 8. Create Total Sheet (New Requirement)
    # Group by Student only
    total_df = aggregated_df.groupby(['学籍番号', '氏名'])[numeric_cols].sum().reset_index()
    total_df['正答率(%)'] = (total_df['正答数'] / total_df['問題数'] * 100).fillna(0).round().astype(int)
    
    # Sort by Correct Rate Ascending
    total_df = total_df.sort_values('正答率(%)', ascending=True)

    # 9. Export to Excel
    with pd.ExcelWriter(OUTPUT_EXCEL) as writer:
        pivot_df.to_excel(writer, sheet_name='サマリー')
        output_flat.to_excel(writer, sheet_name='詳細データ', index=False)
        total_df.to_excel(writer, sheet_name='総合データ', index=False)
    
    print(f"Exported Excel to {OUTPUT_EXCEL}")

if __name__ == "__main__":
    main()
