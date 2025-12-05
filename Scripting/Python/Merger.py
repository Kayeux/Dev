#!/usr/bin/env python3

import os

def merge_files_in_folder(folder_path, output_file="merged_file.txt"):
    # Get a list of all items in the folder, sorted alphabetically
    items = sorted(os.listdir(folder_path))

    with open(output_file, "w", encoding="utf-8") as outfile:
        for item in items:
            file_path = os.path.join(folder_path, item)
            
            # Check only files (skip directories)
            if os.path.isfile(file_path):
                # Write the file name in bold (Markdown-style)
                outfile.write(f"**{item}**\n\n")
                
                # Write the file's content
                with open(file_path, "r", encoding="utf-8", errors="ignore") as infile:
                    outfile.write(infile.read().rstrip("\n"))
                
                # Write a line breaker before the next file
                outfile.write("\n\n---\n\n")

if __name__ == "__main__":
    # Prompt for directory path
    folder_path_input = input("Enter the path of the folder containing the files to merge: ")
    
    # Optional: prompt for output filename if desired
    # output_file_input = input("Enter the desired output filename (default: merged_file.txt): ") or "merged_file.txt"
    
    merge_files_in_folder(folder_path_input)
    print("Files merged successfully into 'merged_file.txt'.")