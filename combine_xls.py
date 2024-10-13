import argparse
import pandas as pd

def combine_excel_files(file_a, file_b, column_a, column_b, output_file, case_sensitive=True, like_comparison=False):
    # Read Excel files
    df_a = pd.read_excel(file_a)
    df_b = pd.read_excel(file_b)

    # Prepare columns for comparison
    if not case_sensitive:
        df_a[column_a] = df_a[column_a].str.lower()
        df_b[column_b] = df_b[column_b].str.lower()

    # Perform merge based on comparison type
    if like_comparison:
        merged_df = pd.merge(df_a, df_b, left_on=df_a[column_a].str.contains,
                             right_on=df_b[column_b], how='inner', suffixes=('_A', '_B'))
    else:
        merged_df = pd.merge(df_a, df_b, left_on=column_a, right_on=column_b, how='inner', suffixes=('_A', '_B'))

    # Save the merged dataframe to a new Excel file
    merged_df.to_excel(output_file, index=False)
    print(f"Combined file saved as {output_file}")
    print(f"Number of rows in file A: {len(df_a)}")
    print(f"Number of rows in file B: {len(df_b)}")
    print(f"Number of rows in merged file: {len(merged_df)}")

def main():
    parser = argparse.ArgumentParser(description="Combine two Excel files based on a specified column.")
    parser.add_argument("file_a", help="Path to the first Excel file")
    parser.add_argument("file_b", help="Path to the second Excel file")
    parser.add_argument("--column_a", required=True, help="Column name in the first file for comparison")
    parser.add_argument("--column_b", required=True, help="Column name in the second file for comparison")
    parser.add_argument("--output", default="merged.xlsx", help="Path to save the combined Excel file (default: merged.xlsx)")
    parser.add_argument("--case-insensitive", action="store_true", help="Perform case-insensitive comparison")
    parser.add_argument("--like", action="store_true", help="Use 'LIKE' comparison instead of exact match")

    args = parser.parse_args()

    combine_excel_files(args.file_a, args.file_b, args.column_a, args.column_b, 
                        args.output, not args.case_insensitive, args.like)

if __name__ == "__main__":
    main()