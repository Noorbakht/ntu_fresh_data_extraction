#!/usr/bin/env python3
"""
Test script to demonstrate how to combine ComBaseExport Excel files.
This script will combine all ComBaseExport Excel files in the Downloads directory
into a single Excel file with two tabs: Data Records and Logs.
"""

from ntu_fresh_selenium_bs import combine_excel_files

if __name__ == "__main__":
    print("Combining ComBaseExport Excel files...")
    output_file = "ComBaseCombined.xlsx"
    
    if combine_excel_files(output_file):
        print(f"Excel files combined successfully into {output_file} in your Downloads folder.")
        print("The combined file contains two tabs:")
        print("1. Data Records - Contains all data records from the exported files")
        print("2. Logs - Contains all logs from the exported files")
    else:
        print("Failed to combine Excel files.")
        print("Make sure you have ComBaseExport Excel files in your Downloads directory.")
