import pandas as pd
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Combobox

def create_gui_and_run_script():
    """
    Creates a graphical user interface (GUI) to interact with the BOM processing script.
    It prompts the user for a CSV file path and a boolean value for checking part numbers.
    """
    
    # --- Code from the original script (functions) ---
    def remove_extension(file_path: str) -> str:
        """Removes the file extension from a given file path string."""
        base_name, _ = os.path.splitext(file_path)
        return base_name

    def remove_column(df: pd.DataFrame, columns_to_remove: list) -> None:
        """Removes specified columns from a Pandas DataFrame."""
        df.drop(columns=columns_to_remove, errors='ignore', inplace=True)
        return None

    # Robust Level tokenization (used where needed; does not change core flow)
    def _level_tokens(val) -> list:
        """
        Return a robust token list for a Level value.
        Handles floats like 10.0, strings like '10.1', blanks/NaN, spaces, trailing dots.
        """
        s = "" if pd.isna(val) else str(val).strip()
        if s == "" or s.lower() == "nan":
            return ["0"]
        s = s.replace(" ", "").strip(".")
        return s.split(".")

    def remove_welded_components(df, substrings):
        """Removes rows related to welded assemblies and their sub-components.

        FINAL precise rule:
        - Keep the welded parent row (tokens P, e.g., ['10','0']).
        - While in a welded region, drop any subsequent row whose first token equals P[0]
          AND tokens != P (this catches same-depth children like 10.1, 10.2, ... and deeper ones).
        - Stop dropping when the first token changes OR when tokens == P (a new 10.0 root).
        """
        to_drop = []
        welded_active = False
        parent_tokens = None
        parent_token0 = None

        for idx, row in df.iterrows():
            number_str = str(row.get("Number", ""))
            tokens = _level_tokens(row.get("Level"))
            token0 = tokens[0]

            if welded_active:
                # New root at same spot (exact same tokens as parent) → end previous region
                if tokens == parent_tokens:
                    welded_active = False
                    parent_tokens = None
                    parent_token0 = None
                    # fall through to possibly start a new welded region on this same row

                # Still under same top bucket → drop if not the parent itself
                elif token0 == parent_token0:
                    to_drop.append(idx)
                    continue
                else:
                    # First token changed → end welded region
                    welded_active = False
                    parent_tokens = None
                    parent_token0 = None
                    # fall through to possibly start a new one

            # Start welded region if this row is a welded assembly number
            if any(sub in number_str for sub in substrings):
                welded_active = True
                parent_tokens = tokens
                parent_token0 = token0

        return df.drop(index=to_drop)

    def remove_asm_rows(df, substrings):
        """Removes rows where the 'Number' column contains specified substrings."""
        to_drop = []
        for index, row in df.iterrows():
            number_str = str(row['Number'])
            if any(sub in number_str for sub in substrings):
                to_drop.append(index)
        new_df = df.drop(index=to_drop)
        return new_df

    def remove_OEM_sub_components(df: pd.DataFrame) -> None:
        """Removes any item with notation AA-A####-# from the BOM."""
        to_drop = []
        rows = []
        nums = []
        for index, row in df.iterrows():
            number_str = str(row['Number'])
            if len(number_str) > 8 and number_str[8] == "-":
                to_drop.append(index)
                rows.append(row)
        new_df = df.drop(index=to_drop)
        return new_df

    def keep_welded_components(df: pd.DataFrame, substrings) -> pd.DataFrame:
        """Keeps rows related to welded assemblies and their sub-components."""
        to_keep = []
        welded_part = False
        for index, row in df.iterrows():
            number_str = str(row['Number'])
            if welded_part:
                level_value = str(row['Level'])
                depth = len(level_value.split('.')) - 1
                if depth > welded_depth:
                    to_keep.append(index)
                else:
                    welded_part = False
            elif any(sub in number_str for sub in substrings):
                level_value = str(row['Level'])
                welded_part = True
                welded_depth = len(level_value.split('.')) - 1
                to_keep.append(index)
        df_new = df.loc[to_keep]
        return df_new

    def keep_rows_with(df: pd.DataFrame, substrings) -> pd.DataFrame:
        """Keeps only rows where the 'Number' column contains specified substrings."""
        to_drop = []
        for index, row in df.iterrows():
            number_str = str(row['Number'])
            if not any(sub in number_str for sub in substrings):
                to_drop.append(index)
        new_df = df.drop(index=to_drop, inplace=False)
        return new_df

    # --- Non-invasive Material/Finish helpers (adds fields, no logic change) ---
    def _normalize_col(col: str) -> str:
        return re.sub(r'[^a-z0-9]+', '', str(col).strip().lower())

    def find_best_source_column(df: pd.DataFrame, candidates: list) -> str:
        """Find the first df column that matches/contains a candidate name (robust to case/spacing)."""
        norm_map = {_normalize_col(c): c for c in df.columns}
        norm_cols = list(norm_map.keys())
        norm_candidates = [_normalize_col(c) for c in candidates]
        for nc in norm_candidates:
            if nc in norm_map:
                return norm_map[nc]
        for nc in norm_candidates:
            for ncol in norm_cols:
                if ncol.startswith(nc) or nc in ncol:
                    return norm_map[ncol]
        return None

    def coalesce_string_cols(df: pd.DataFrame, target_col: str, source_cols: list) -> pd.Series:
        """Create target_col as first non-empty among source_cols, preserving existing non-empty values."""
        out = pd.Series("", index=df.index, dtype="object")
        if target_col in df.columns:
            out = df[target_col].astype("string").fillna("")
        def clean(x):
            s = str(x).strip() if pd.notna(x) else ""
            return "" if s.lower() in ("nan", "none", "null") else s
        for col in source_cols:
            if col and col in df.columns:
                src = df[col].astype("string").map(clean)
                out = out.where(out != "", src)
        return out.fillna("")

    def combine_repeats(df: pd.DataFrame) -> pd.DataFrame:
        """Combines rows with duplicate parts and sums their quantities.
        Grouping logic is unchanged. After grouping, append Material/Finish
        via a first-nonempty lookup by Number (non-invasive).
        """
        combined_df = df.groupby(['Number', 'Revision', 'Description', 'State'])['Qty'].sum().reset_index()
        combined_df = combined_df[['Number', 'Revision', 'Description', 'Qty', 'State']]

        def first_nonempty(df_src: pd.DataFrame, col: str):
            if col not in df_src.columns:
                return pd.DataFrame(columns=['Number', col])
            tmp = df_src[['Number', col]].astype({col: 'string'})
            tmp[col] = tmp[col].fillna('').str.strip()
            tmp = tmp[tmp[col] != '']
            return tmp.drop_duplicates(subset=['Number'])

        if 'Material' in df.columns:
            combined_df = combined_df.merge(first_nonempty(df, 'Material'), on='Number', how='left')
        else:
            combined_df['Material'] = ""
        if 'Finish' in df.columns:
            combined_df = combined_df.merge(first_nonempty(df, 'Finish'), on='Number', how='left')
        else:
            combined_df['Finish'] = ""

        cols = ['Number', 'Revision', 'Description', 'Qty', 'State', 'Material', 'Finish']
        combined_df = combined_df[cols]
        return combined_df

    # --- End of original code functions ---

    def browse_file():
        """Opens a file dialog for the user to select a CSV file."""
        filepath = filedialog.askopenfilename(
            filetypes=[("CSV Files", "*.csv")],
            title="Select BOM CSV file"
        )
        if filepath:
            csv_path_entry.delete(0, tk.END)
            csv_path_entry.insert(0, filepath)

    def on_encoding_select(event):
        """Enable or disable the custom encoding entry based on dropdown selection."""
        if encoding_var.get() == "Custom...":
            custom_encoding_entry.config(state="normal")
        else:
            custom_encoding_entry.config(state="disabled")

    def run_script():
        """
        Executes the main logic of the original script using the values from the GUI.
        """
        csv_filepath = csv_path_entry.get()
        if not csv_filepath:
            messagebox.showerror("Error", "Please select a CSV file.")
            return

        # Get values from GUI widgets
        check_for_pn_str = check_for_pn_var.get()
        check_for_PN = check_for_pn_str == "True"
        
        # Get the lists and encoding from the GUI, splitting by newline for lists
        columns_to_remove = columns_entry.get("1.0", tk.END).strip().split('\n')
        welded_asm_names = welded_names_entry.get("1.0", tk.END).strip().split('\n')
        asm_names = asm_names_entry.get("1.0", tk.END).strip().split('\n')
        
        # Get encoding from the dropdown or custom entry
        selected_encoding = encoding_var.get()
        if selected_encoding == "Custom...":
            encoding = custom_encoding_entry.get().strip()
            if not encoding:
                messagebox.showerror("Error", "Please enter a custom encoding.")
                return
        else:
            encoding = selected_encoding

        # Handle empty lists from the GUI
        if columns_to_remove == ['']:
            columns_to_remove = []
        if welded_asm_names == ['']:
            welded_asm_names = []
        if asm_names == ['']:
            asm_names = []
            
        try:
            # Create the new "BOMs" directory
            new_dir_path = os.path.join(remove_extension(os.path.basename(csv_filepath)) + " BOMs")
            output_dir = os.path.join(os.path.dirname(csv_filepath), new_dir_path)
            os.makedirs(output_dir, exist_ok=True)
            
            # Read the CSV file
            df = pd.read_csv(csv_filepath, encoding=encoding)

            # --- Non-invasive: populate Material / Finish from PDM fields ---
            material_candidates = [
                "Material", "SW-Material", "SW Material", "Material Name",
                "Material@Part", "Material@Model", "Raw Material", "Mat"
            ]
            finish_candidates = [
                "Finish", "Surface Finish", "Surface-Finish", "Finish@Part",
                "Finish@Model", "Coating", "Treatment", "SurfaceTreatment"
            ]
            best_mat = find_best_source_column(df, material_candidates)
            best_fin = find_best_source_column(df, finish_candidates)

            def _dedupe(seq):
                seen, out = set(), []
                for x in seq:
                    if x and x not in seen:
                        seen.add(x)
                        out.append(x)
                return out
            mat_sources = _dedupe([best_mat] + material_candidates)
            fin_sources = _dedupe([best_fin] + finish_candidates)

            df["Material"] = coalesce_string_cols(df, "Material", mat_sources)
            df["Finish"]   = coalesce_string_cols(df, "Finish",   fin_sources)

            # --- Main script logic from original code ---
            if check_for_PN:
                isWrong = False
                print("The following files do not have matching File Names and PartNo in Solidworks custom properties. Files to be fixed:")
                error_log = []
                for index, row in df.iterrows():
                    PN = remove_extension(row["File Name"]).strip()
                    if PN != str(row["Number"]):
                        isWrong = True
                        error_log.append(row.to_string())
                if error_log:
                    messagebox.showerror("Validation Error", "File name and 'PartNo' custom properties MUST match for code to work. Please fix the files listed below. Then re-extract the BOM and re-run the code.\n\n" + "\n".join(error_log))
                    return
                else:
                    print("\tNone. All File Names match PartNo in custom properties.")
            else:
                for index, row in df.iterrows():
                    df.loc[index, "Number"] = remove_extension(row["File Name"]).strip()

            # Create EBOM file (Material/Finish now present)
            EBOM_filepath = os.path.join(output_dir, remove_extension(os.path.basename(csv_filepath)) + "_EBOM")
            EBOM = df.copy()  # Make a copy to avoid modifying the original df for subsequent steps
            remove_column(EBOM, columns_to_remove)
            EBOM.to_excel(EBOM_filepath + ".xlsx", index=False)

            # Create WBOM file
            WBOM_filepath = os.path.join(output_dir, remove_extension(os.path.basename(csv_filepath)) + "_WBOM")
            WBOM = keep_welded_components(df.copy(), welded_asm_names)
            remove_column(WBOM, columns_to_remove)
            first = True
            current_level_number = [['0'], 0]
            for index, row in WBOM.iterrows():
                level_value = str(row['Level'])
                depth_values = level_value.split('.')
                if first:
                    current_level_number = [depth_values, len(depth_values) - 1]
                    WBOM.loc[index, "Level"] = 0
                    first = False
                elif depth_values[0:current_level_number[1] + 1] == current_level_number[0]:
                    WBOM.loc[index, "Level"] = len(depth_values) - (current_level_number[1] + 1)
                else:
                    current_level_number = [depth_values, len(depth_values) - 1]
                    WBOM.loc[index, "Level"] = 0
            
            WBOM.to_excel(WBOM_filepath + ".xlsx", index=False)

            # Ensure Level is int before writing WBOM.txt (prevents float error)
            WBOM["Level"] = pd.to_numeric(WBOM["Level"], errors="coerce").fillna(0).astype(int)

            # Write WBOM.txt (append Material & Finish non-invasively)
            def _safe_str(x):
                s = "" if pd.isna(x) else str(x).strip()
                return "" if s.lower() in ("nan", "none", "null") else s

            with open(WBOM_filepath + ".txt", mode='w', encoding='utf-8') as f:
                for index, row in WBOM.iterrows():
                    indentation = "    " * row["Level"]
                    mat = _safe_str(row.get("Material", ""))
                    fin = _safe_str(row.get("Finish", ""))
                    f.write(indentation + str(row["Number"]).strip() + "    " +
                            str(row["Description"]) + "    " + mat + "    " + fin + "\n")

            # Create assembly tree file (append Material & Finish non-invasively)
            ASMtree_filepath = os.path.join(output_dir, remove_extension(os.path.basename(csv_filepath)) + "_ASMtree")
            ASMtree = keep_rows_with(df.copy(), asm_names)
            ASMtree_list = []
            for index, row in ASMtree.iterrows():
                level_value = str(row['Level'])
                depth = len(level_value.split('.')) - 1
                mat = _safe_str(row.get("Material", ""))
                fin = _safe_str(row.get("Finish", ""))
                ASMtree_list.append((str(row["Number"]) + "    " + str(row["Description"]) + "    " + mat + "    " + fin, depth))
            
            with open(ASMtree_filepath + ".txt", mode='w', encoding='utf-8') as f:
                for item in ASMtree_list:
                    indentation = "    " * item[1]
                    f.write(indentation + item[0].strip() + "\n")

            # Create MBOM file (grouping unchanged; remove welded children; then append Material/Finish)
            MBOM_filepath = os.path.join(output_dir, remove_extension(os.path.basename(csv_filepath)) + "_MBOM")
            MBOM = remove_OEM_sub_components(df.copy())
            MBOM = remove_welded_components(MBOM, welded_asm_names)  # now drops same-depth children like 10.1, 10.2, ...
            MBOM = remove_asm_rows(MBOM, asm_names)
            MBOM = combine_repeats(MBOM)
            MBOM.to_excel(MBOM_filepath + ".xlsx", index=False)

            messagebox.showinfo("Success", "BOM processing complete!\n\nFiles saved in the 'BOMs' directory.")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    # --- GUI Setup ---
    root = tk.Tk()
    root.title("BOM Processor")
    root.geometry("450x650")
    root.resizable(False, False)

    # Main frame for padding
    main_frame = tk.Frame(root, padx=10, pady=10)
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # CSV File Path entry and button
    csv_path_label = tk.Label(main_frame, text="Select CSV File:")
    csv_path_label.pack(anchor="w", pady=(5, 0))
    csv_path_entry = tk.Entry(main_frame, width=50)
    csv_path_entry.pack(fill=tk.X, pady=(0, 5))
    browse_button = tk.Button(main_frame, text="Browse...", command=browse_file)
    browse_button.pack(anchor="e", pady=(0, 10))

    # Check for PN dropdown
    check_for_pn_label = tk.Label(main_frame, text="Check that Part Numbers match File Name\nIf set to false, assumes file name is part number")
    check_for_pn_label.pack(anchor="w", pady=(5, 0))
    check_for_pn_var = tk.StringVar(value="True")  # Default to True
    check_for_pn_dropdown = Combobox(main_frame, textvariable=check_for_pn_var, state="readonly", width=10)
    check_for_pn_dropdown['values'] = ("True", "False")
    check_for_pn_dropdown.pack(anchor="w", pady=(0, 10))

    # Columns to Remove
    columns_label = tk.Label(main_frame, text="Columns to Remove from original .CSV file (one per line, case sensitive):")
    columns_label.pack(anchor="w", pady=(5, 0))
    columns_entry = tk.Text(main_frame, height=2, width=40)
    columns_entry.insert(tk.END, "File Name\nConfiguration") # Default value
    columns_entry.pack(pady=(0, 10), fill=tk.X)

    # Welded Assembly Names
    welded_names_label = tk.Label(main_frame, text="Welded Assembly Names (one per line):")
    welded_names_label.pack(anchor="w", pady=(5, 0))
    welded_names_entry = tk.Text(main_frame, height=4, width=40)
    welded_names_entry.insert(tk.END, "BS-W\nSS-W\nTP-W\nTS-W\nSC-W\nJB-W\nPB-W\nWJ-W\nSU-W") # Default value
    welded_names_entry.pack(pady=(0, 10), fill=tk.X)

    # Assembly Names
    asm_names_label = tk.Label(main_frame, text="Assembly Names (one per line):")
    asm_names_label.pack(anchor="w", pady=(5, 0))
    asm_names_entry = tk.Text(main_frame, height=4, width=40)
    asm_names_entry.insert(tk.END, "BS-A\nSS-A\nTP-A\nTS-A\nSC-A\nJB-A\nPB-A\nWJ-A\nSU-A\nEC-A") # Default value
    asm_names_entry.pack(pady=(0, 10), fill=tk.X)

    # Encoding
    encodings = [
        "utf-8", "utf-16", "utf-32", "latin-1", "ascii", "cp1252",
        "iso-8859-1", "iso-8859-2", "gbk", "shift_jis", "euc-jp",
        "big5", "koi8-r", "utf-8-sig", "windows-1250", "windows-1251",
        "windows-1253", "windows-1254", "windows-1255", "windows-1256",
        "Custom..."
    ]
    encoding_label = tk.Label(main_frame, text="CSV Encoding: (if getting error reading .csv file error, try utf-8)")
    encoding_label.pack(anchor="w", pady=(5, 0))

    encoding_var = tk.StringVar(value="utf-16")  # Default to utf-16
    encoding_dropdown = Combobox(main_frame, textvariable=encoding_var, state="readonly", width=15)
    encoding_dropdown['values'] = encodings
    encoding_dropdown.pack(anchor="w", pady=(0, 5))
    encoding_dropdown.bind("<<ComboboxSelected>>", on_encoding_select)
    
    custom_encoding_entry = tk.Entry(main_frame, width=20, state="disabled")
    custom_encoding_entry.pack(anchor="w", pady=(0, 10))
    
    # Run button
    run_button = tk.Button(main_frame, text="Run Script", command=run_script)
    run_button.pack(pady=20)

    root.mainloop()

# Call the function to start the GUI
if __name__ == '__main__':
    create_gui_and_run_script()
