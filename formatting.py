import openpyxl
import re
import tkinter as tk
from tkinter import filedialog
import os

def clean_text(original_text: str) -> str:
    """
    1. Remove all '...' (three consecutive periods).
    2. In any text enclosed in double or single quotes:
       - If the quoted text is exactly 'BOM' (case-insensitive), remove the quotes completely, replaced by BOM.
       - Else, if the quoted text consists only of single letters/digits separated by spaces,
         remove those spaces (e.g., "P R 3" -> "PR3").
       - Otherwise, leave the quoted text as is.
    3. Convert multiple spaces to single spaces.
    4. Capitalize the first letter.
    5. Ensure the text ends with a period.
    """

    # 1. Strip leading/trailing whitespace
    text = original_text.strip()

    # 2. Remove all occurrences of '...'
    text = text.replace("...", "")

    # 3. Process quoted text (both single and double quotes).
    #    We'll find all substrings in quotes and selectively remove or transform the content.
    def process_quoted(match):
        quote_char = match.group(1)  # The quote symbol (single or double)
        content    = match.group(2)  # The text inside the quotes

        # Trim leading/trailing spaces inside the quotes
        content_stripped = content.strip()

        # If the content is "BOM" (any case), remove quotes entirely => BOM
        if content_stripped.upper() == "BOM":
            return "BOM"

        # Else, check if all tokens are exactly one character (letter or digit)
        tokens = content_stripped.split()
        if all(len(t) == 1 for t in tokens):
            # Remove spaces by joining tokens, e.g. "P R 3" -> "PR3"
            content_stripped = "".join(tokens)

        # Return with the original quote characters preserved, unless it's BOM
        return f"{quote_char}{content_stripped}{quote_char}"

    # Regex explanation:
    # (["'])       -> capture a single or double quote in group(1)
    # (.*?)        -> capture anything (non-greedy) until ...
    # (\1)         -> the same quote that started it
    text = re.sub(r'(["\'])(.*?)(\1)', process_quoted, text)

    # 4. Convert multiple spaces to single space
    #    This will catch double, triple, etc. spaces and make them single
    text = re.sub(r'\s{2,}', ' ', text)

    # 5. Capitalize the first letter, if there's any text
    if text:
        text = text[0].upper() + text[1:]

    # 6. Ensure it ends with a period
    if not text.endswith("."):
        text += "."

    return text

def remove_blank_rows(input_filename: str, sheet_name: str = None):
    """
    Removes all blank rows from the specified (or active) sheet in an Excel file.
    A row is considered blank if all cells in that row are empty.
    """
    # Load the workbook
    wb = openpyxl.load_workbook(input_filename)
    
    # If sheet_name is given, use that sheet; otherwise use active
    if sheet_name:
        ws = wb[sheet_name]
    else:
        ws = wb.active
    
    # Find all blank rows (in reverse order to avoid index issues when deleting)
    blank_rows = []
    for row_idx in range(ws.max_row, 0, -1):
        is_blank = True
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if cell.value is not None and str(cell.value).strip() != "":
                is_blank = False
                break
        if is_blank:
            blank_rows.append(row_idx)
    
    # Delete the blank rows
    for row_idx in blank_rows:
        ws.delete_rows(row_idx)
    
    # Save the workbook
    wb.save(input_filename)
    
    print(f"Removed {len(blank_rows)} blank rows from {input_filename}.")
    return len(blank_rows)

def process_excel(input_filename: str, sheet_name: str = None):
    """
    Iterates down column A in the specified (or active) sheet.
    For each non-empty cell:
      - Clean the text using clean_text().
      - Write it back to the cell.
    Stops when two consecutive blank cells are found.
    """
    # Load the workbook
    wb = openpyxl.load_workbook(input_filename)
    
    # If sheet_name is given, use that sheet; otherwise use active
    if sheet_name:
        ws = wb[sheet_name]
    else:
        ws = wb.active

    consecutive_blank_count = 0
    row = 1  # Start from the first row
    
    while True:
        cell = ws.cell(row=row, column=1)
        cell_value = cell.value
        
        # Check for blank (or None) cell
        if cell_value is None or str(cell_value).strip() == "":
            consecutive_blank_count += 1
            # If we've reached two consecutive blanks, stop
            if consecutive_blank_count == 2:
                break
        else:
            # Reset blank count
            consecutive_blank_count = 0

            # Convert cell_value to string and clean it
            text = str(cell_value)
            new_text = clean_text(text)

            # Update the cell if changed
            if new_text != text:
                cell.value = new_text
        
        row += 1  # Move to the next row
    
    # Option A: Overwrite the same file (make sure it's not open in Excel)
    wb.save(input_filename)

    # Option B: Save to a new file to avoid overwriting
    # new_filename = os.path.join(
    #     os.path.dirname(input_filename),
    #     "cleaned_" + os.path.basename(input_filename)
    # )
    # wb.save(new_filename)

    print(f"Finished processing {input_filename}.")

def main():
    """
    Main function that provides a GUI for selecting an Excel file and processing it.
    """
    # Create a hidden root window for the file dialog
    root = tk.Tk()
    root.withdraw()

    print("Please select an Excel file to process...")
    
    # Ask the user to select the Excel file
    input_filename = filedialog.askopenfilename(
        title="Select an Excel file",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xltx *.xltm")]
    )
    
    # If user cancels, just exit
    if not input_filename:
        print("No file selected. Exiting.")
        return
    
    # Ask user which operation to perform
    print("\nChoose an operation:")
    print("1. Clean text in cells")
    print("2. Remove blank rows")
    print("3. Clean text AND remove blank rows")
    choice = input("Enter your choice (1, 2, or 3): ")
    
    if choice == "1":
        # Process the chosen file - clean text
        process_excel(input_filename)
    elif choice == "2":
        # Remove blank rows from the chosen file
        removed = remove_blank_rows(input_filename)
        print(f"Removed {removed} blank rows from the file.")
    elif choice == "3":
        # Perform both operations
        print("Performing complete formatting...")
        process_excel(input_filename)  # First clean the text
        removed = remove_blank_rows(input_filename)  # Then remove blank rows
        print(f"Cleaned text and removed {removed} blank rows from the file.")
    else:
        print("Invalid choice. Exiting.")

if __name__ == "__main__":
    main()
