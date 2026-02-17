import pandas as pd
import os
import glob

def load_and_normalize_data(current_dir):
    """
    Loads all CSV files, concatenates them, and normalizes Student IDs so that
    each Name has a unique ID (handling cases where one student has multiple IDs).
    """
    csv_files = glob.glob(os.path.join(current_dir, "*.csv"))
    
    if not csv_files:
        print("No CSV files found in the directory.")
        return None

    print(f"Found {len(csv_files)} CSV files.")
    
    # Read and concatenate all CSVs
    df_list = []
    for file in csv_files:
        try:
            df = pd.read_csv(file)
            df_list.append(df)
            print(f"Loaded: {os.path.basename(file)}")
        except Exception as e:
            print(f"Error reading {file}: {e}")
    
    if not df_list:
        print("No data loaded.")
        return None

    all_data = pd.concat(df_list, ignore_index=True)
    
    # Drop exact duplicates first
    initial_count = len(all_data)
    all_data.drop_duplicates(inplace=True)
    print(f"Removed {initial_count - len(all_data)} duplicate rows.")
    
    # Normalize Student IDs based on Name
    # Rule: If a student has multiple IDs, use the most frequent one (mode) or max.
    if '氏名' in all_data.columns and '学籍番号' in all_data.columns:
        # Create a mapping: Name -> Unifying ID
        def get_representative_id(ids):
            mode = ids.mode()
            if not mode.empty:
                return mode.iloc[0]
            return ids.max()
            
        name_to_id = all_data.groupby('氏名')['学籍番号'].apply(get_representative_id).to_dict()
        
        # Apply mapping
        all_data['学籍番号'] = all_data['氏名'].map(name_to_id)
        print("Normalized Student IDs based on Name (merged multiple IDs for same name).")
    
    return all_data

def organize_student_data(current_dir, all_data):
    """
    Groups data by Student ID and saves individual CSV files.
    """
    # Create output directory
    output_dir = os.path.join(current_dir, "student_data")
    os.makedirs(output_dir, exist_ok=True)
    
    if '学籍番号' not in all_data.columns or '氏名' not in all_data.columns:
        print("Error: Required columns '学籍番号' or '氏名' not found.")
        return

    grouped = all_data.groupby('学籍番号')
    
    print(f"Processing data for {len(grouped)} students...")
    
    for student_id, group in grouped:
        # Get student name
        student_name = group['氏名'].iloc[0]
        
        # Sanitize name for filename
        safe_name = str(student_name).replace(" ", "_").replace("　", "_")
        safe_id = str(student_id).replace(" ", "")
        
        filename = f"{safe_name}_{safe_id}.csv"
        filepath = os.path.join(output_dir, filename)
        
        # Save to CSV
        group.to_csv(filepath, index=False)

    print(f"Successfully organized individual CSVs into {output_dir}")

def create_aggregated_excel_summary(current_dir, all_data):
    """
    Aggregates data by Student/Field and saves to Excel.
    """
    # 1. Ensure numeric columns
    numeric_cols = ['問題数', '正答数']
    for col in numeric_cols:
        if col in all_data.columns:
            all_data[col] = pd.to_numeric(all_data[col], errors='coerce').fillna(0)
        else:
            all_data[col] = 0

    # 2. Group by ID, Name, Field and Sum
    grouped = all_data.groupby(['学籍番号', '氏名', '分野'])[numeric_cols].sum().reset_index()

    # 3. Pivot
    # Index: ID, Name
    # Columns: Field
    # Values: Questions, Correct
    pivoted = grouped.pivot(index=['学籍番号', '氏名'], columns='分野', values=['問題数', '正答数'])
    
    # 4. Flatten columns
    # The columns will be a MultiIndex like ('問題数', 'FieldA'), ('正答数', 'FieldA')
    pivoted.columns = [f"{col[1]}_{col[0]}" for col in pivoted.columns]
    pivoted.reset_index(inplace=True)

    # 5. Save to Excel
    output_file = os.path.join(current_dir, "summary_aggregated_20260216-0220.xlsx")
    try:
        pivoted.to_excel(output_file, index=False)
        print(f"Successfully created aggregated Excel summary: {output_file}")
    except Exception as e:
        print(f"Error saving Excel file: {e}")

if __name__ == "__main__":
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    data = load_and_normalize_data(current_dir)
    
    if data is not None:
        organize_student_data(current_dir, data)
        create_aggregated_excel_summary(current_dir, data)
