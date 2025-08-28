#!/usr/bin/env python3
"""
TRC Log Cleaner - Remove \\N entries from pipe-delimited log files
"""

import os
import sys
from pathlib import Path

def clean_log_line(line):
    r"""
    Clean a single log line by removing \N entries

    Args:
        line (str): The log line to clean

    Returns:
        str: Cleaned line with \N entries removed
    """
    if not line.strip():
        return line

    # Split by pipe delimiter
    parts = line.strip().split('|')

    # Filter out \N entries and empty strings
    cleaned_parts = [part for part in parts if part != '\\N' and part.strip() != '']

    # Rejoin with pipe delimiter
    return '|'.join(cleaned_parts)

def clean_log_file(input_file, output_file=None):
    r"""
    Clean a log file by removing \N entries from all lines

    Args:
        input_file (str): Path to input log file
        output_file (str, optional): Path to output file. If None, creates _cleaned suffix

    Returns:
        str: Path to the cleaned output file
    """
    input_path = Path(input_file)

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")

    # Generate output filename if not provided
    if output_file is None:
        output_file = input_path.parent / f"{input_path.stem}_cleaned{input_path.suffix}"

    output_path = Path(output_file)

    print(f"Cleaning log file: {input_file}")
    print(f"Output will be saved to: {output_file}")

    total_lines = 0
    cleaned_lines = 0

    try:
        with open(input_path, 'r', encoding='utf-8', errors='ignore') as infile, \
             open(output_path, 'w', encoding='utf-8') as outfile:

            for line_num, line in enumerate(infile, 1):
                total_lines += 1

                # Clean the line
                cleaned_line = clean_log_line(line)

                # Only write non-empty lines after cleaning
                if cleaned_line.strip():
                    outfile.write(cleaned_line + '\n')
                    cleaned_lines += 1

                # Progress indicator for large files
                if line_num % 1000 == 0:
                    print(f"Processed {line_num} lines...")

    except Exception as e:
        print(f"Error processing file: {e}")
        return None

    print(f"Processing complete!")
    print(f"Total lines processed: {total_lines}")
    print(f"Lines with data after cleaning: {cleaned_lines}")

    return str(output_path)

def main():
    """Main function to handle command line arguments and drag & drop"""
    if len(sys.argv) < 2:
        print("TRC Log Cleaner v1.1 - Drag & Drop Support")
        print("=" * 50)
        print("Usage: python trc_log_cleaner.py <input_file(s)>")
        print("\nFeatures:")
        print("  • Drag and drop multiple files")
        print("  • Automatic output naming (_cleaned suffix)")
        print("  • Batch processing support")
        print("\nExamples:")
        print("  python trc_log_cleaner.py WorldSvr_01_01_250828.GameLog")
        print("  python trc_log_cleaner.py file1.log file2.log file3.log")
        print("\nDrag and drop files onto this script in Windows Explorer!")
        sys.exit(1)

    # Get all input files (supporting drag & drop of multiple files)
    input_files = sys.argv[1:]

    print("TRC Log Cleaner v1.1")
    print("=" * 30)
    print(f"Processing {len(input_files)} file(s)...")
    print()

    success_count = 0
    failed_files = []

    for i, input_file in enumerate(input_files, 1):
        print(f"[{i}/{len(input_files)}] Processing: {input_file}")

        try:
            result_file = clean_log_file(input_file)
            if result_file:
                print(f"  ✅ Success: {result_file}")
                success_count += 1
            else:
                print("  ❌ Failed to clean file")
                failed_files.append(input_file)
        except Exception as e:
            print(f"  ❌ Error: {e}")
            failed_files.append(input_file)

        print()

    # Summary
    print("=" * 30)
    print("Processing Summary:")
    print(f"  Total files: {len(input_files)}")
    print(f"  Successful: {success_count}")
    print(f"  Failed: {len(failed_files)}")

    if failed_files:
        print("\nFailed files:")
        for failed_file in failed_files:
            print(f"  • {failed_file}")

    if success_count > 0:
        print("\n🎉 All cleaned files are ready!")
        print("   You can now drag them onto TRC_Filter_Excel_3.py for Excel reports.")
    if len(sys.argv) == 2 and success_count == 1:
        # If only one file was processed successfully, suggest next step
        result_file = clean_log_file(sys.argv[1])
        if result_file:
            print(f"\n💡 Next: python TRC_Filter_Excel_3.py \"{result_file}\"")

    sys.exit(0 if success_count > 0 else 1)

if __name__ == "__main__":
    main()
