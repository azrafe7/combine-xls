import argparse
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def get_column(df, column_spec):
    if column_spec.isdigit():
        return df.columns[int(column_spec)]
    return column_spec

def combine_excel_files(file_a, file_b, column_a, column_b, output_file, case_sensitive=True, like_comparison=False, debug=False):
    # Read Excel files
    df_a = pd.read_excel(file_a)
    df_b = pd.read_excel(file_b)

    # Get actual column names
    col_a = get_column(df_a, column_a)
    col_b = get_column(df_b, column_b)

    # Prepare columns for comparison
    if not case_sensitive:
        df_a[col_a] = df_a[col_a].astype(str).str.lower()
        df_b[col_b] = df_b[col_b].astype(str).str.lower()

    # Perform merge based on comparison type
    if like_comparison:
        merged_df = pd.merge(df_a, df_b, left_on=df_a[col_a].str.contains,
                             right_on=df_b[col_b], how='inner', suffixes=('', '_B'))
    else:
        merged_df = pd.merge(df_a, df_b, left_on=col_a, right_on=col_b, how='inner', suffixes=('', '_B'))

    # Remove duplicate columns and use header from file A
    columns_to_drop = []
    for col in merged_df.columns:
        if col.endswith('_B') and col[:-2] in merged_df.columns:
            if merged_df[col].equals(merged_df[col[:-2]]):
                columns_to_drop.append(col)
            else:
                merged_df = merged_df.rename(columns={col: f"{col[:-2]}_B"})
    merged_df = merged_df.drop(columns=columns_to_drop)

    # Save the merged dataframe to a new Excel file
    merged_df.to_excel(output_file, index=False)
    
    if debug:
        # Load the workbook and get the active sheet
        wb = load_workbook(output_file)
        ws = wb.active

        # Define fill colors
        fill_match = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")  # Light Blue
        fill_both = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")  # Green
        fill_a = PatternFill(start_color="E0FFE0", end_color="E0FFE0", fill_type="solid")  # Light Green
        fill_b = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")  # Light Yellow

        # Create sets of columns from each file
        cols_a = set(df_a.columns)
        cols_b = set(df_b.columns)

        # Apply fill colors
        for col in range(1, ws.max_column + 1):
            col_name = ws.cell(row=1, column=col).value
            if col_name == col_a or col_name == col_b:
                for row in range(2, ws.max_row + 1):
                    ws.cell(row=row, column=col).fill = fill_match
            elif col_name in cols_a and col_name in cols_b:
                for row in range(2, ws.max_row + 1):
                    ws.cell(row=row, column=col).fill = fill_both
            elif col_name in cols_a:
                for row in range(2, ws.max_row + 1):
                    ws.cell(row=row, column=col).fill = fill_a
            elif col_name in cols_b:
                for row in range(2, ws.max_row + 1):
                    ws.cell(row=row, column=col).fill = fill_b

        # Save the workbook
        wb.save(output_file)

    print(f"Combined file saved as {output_file}")
    print(f"Number of rows in file A: {len(df_a)}")
    print(f"Number of rows in file B: {len(df_b)}")
    print(f"Number of rows in merged file: {len(merged_df)}")

def main():
    parser = argparse.ArgumentParser(description="Combine two Excel files based on a specified column.")
    parser.add_argument("file_a", help="Path to the first Excel file")
    parser.add_argument("file_b", help="Path to the second Excel file")
    parser.add_argument("--column_a", required=True, help="Column name or index in the first file for comparison")
    parser.add_argument("--column_b", required=True, help="Column name or index in the second file for comparison")
    parser.add_argument("--output", default="merged.xlsx", help="Path to save the combined Excel file (default: merged.xlsx)")
    parser.add_argument("--case-insensitive", action="store_true", help="Perform case-insensitive comparison")
    parser.add_argument("--like", action="store_true", help="Use 'LIKE' comparison instead of exact match")
    parser.add_argument("--debug", action="store_true", help="Highlight cells based on their source file")

    args = parser.parse_args()

    combine_excel_files(args.file_a, args.file_b, args.column_a, args.column_b, 
                        args.output, not args.case_insensitive, args.like, args.debug)

if __name__ == "__main__":
    main()