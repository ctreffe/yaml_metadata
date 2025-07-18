import os
import re
import yaml
import argparse
from openpyxl import Workbook
from openpyxl.styles import Font
from collections import OrderedDict

# List of file name or file path patterns to ignore (partial match or full name)
IGNORED_FILES = [
    "_quarto.yml",  # Ignores Quarto project configuration files
    "\\renv"  # Ignores renv environment paths
]

def ordered_yaml_loader():
    """
    Creates a custom YAML loader that preserves the order of keys
    as they appear in the original YAML files.

    By default, Python dictionaries do not guarantee key order in all contexts.
    This function ensures that the order in which fields are written in the YAML files
    is maintained when they are loaded into Python dictionaries.

    This is especially useful when exporting the data to a CSV file,
    where consistent and meaningful column order improves readability.
    """
    class OrderedLoader(yaml.SafeLoader):
        pass

    def construct_mapping(loader, node):
        loader.flatten_mapping(node)
        return OrderedDict(loader.construct_pairs(node))

    OrderedLoader.add_constructor(
        yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG,
        construct_mapping,
    )

    return OrderedLoader

def load_yaml_files_recursively(root_path):
    """
    Searches through the specified folder and all of its subfolders to find YAML files,
    loads them as structured metadata (dictionaries), and returns them in a list.

    This function is designed to be robust and user-friendly, even when the YAML files
    have inconsistencies in formatting or encoding. It handles common issues that may
    prevent YAML files from being read correctly.

    Key features and error handling:

    1. File search:
       - The script looks for all files ending with .yaml or .yml
         in the specified folder and all subfolders.

    2. Character encoding fallback:
       - The function first tries to read each file using UTF-8 (the recommended encoding).
       - If this fails (due to special characters like ä, ö, ü, etc.), it automatically falls
         back to Windows-1252 (common in files created on Windows systems).
       - A warning is printed if the fallback is used.

    3. Tab character cleanup:
       - YAML does not allow tab characters for indentation, only spaces.
       - If any tab characters are found in a file, they are automatically replaced
         with four spaces to prevent parsing errors.
       - A note is printed when this happens, so users are aware of the correction.

    4. YAML parsing:
       - After reading and cleaning, each file is parsed using a YAML loader
         that preserves the original order of fields as written in the file.
       - If a file is valid, its contents are added to a list as a dictionary.
       - If a file is invalid and cannot be parsed, a warning is printed and
         the file is skipped.

    Parameters:
        root_path (str):
            The folder where the search should begin.
            All subfolders will also be included.

    Returns:
        List[dict]:
            A list of metadata dictionaries, each representing the contents
            of one successfully parsed YAML file.
    """
    yaml_dicts = []
    loader = ordered_yaml_loader()

    for dirpath, dirnames, filenames in os.walk(root_path):
        for filename in filenames:
            if filename.endswith((".yaml", ".yml")):
                if any(pattern in filename for pattern in IGNORED_FILES):
                    print(f"Skipping ignored file: {os.path.join(dirpath, filename)}")
                    continue

                if any(pattern in dirpath for pattern in IGNORED_FILES):
                    print(f"Skipping ignored filepath: {dirpath}")
                    continue

                full_path = os.path.join(dirpath, filename)
                file_content = None

                try:
                    with open(full_path, "r", encoding="utf-8") as file:
                        file_content = file.read()
                except UnicodeDecodeError:
                    try:
                        with open(full_path, "r", encoding="cp1252") as file:
                            file_content = file.read()
                        print(
                            f"Warning: '{full_path}' is not UTF-8. Read using cp1252 fallback."
                        )
                    except Exception as e:
                        print(f"Failed to read {full_path}: {e}")
                        continue

                if file_content is None:
                    continue

                if "\t" in file_content:
                    print(f"Note: Tabs found in '{full_path}', replacing with spaces.")
                    file_content = file_content.replace("\t", "    ")

                try:
                    data = yaml.load(file_content, Loader=loader)
                    if isinstance(data, dict):
                        # Add relative path as string to each data set
                        rel_path = os.path.relpath(dirpath, root_path)
                        print(rel_path)
                        data["filepath"] = rel_path.replace("\\", "/")  # Same format for Windows/Linux
                        yaml_dicts.append(data)
                except yaml.YAMLError as e:
                    print(f"Failed to parse YAML in {full_path}: {e}")

    return yaml_dicts

def write_dicts_to_excel(dicts, output_path):
    """
    Writes a list of metadata dictionaries to an Excel (.xlsx) file.

    This function takes structured metadata extracted from YAML files and
    writes them to a well-formatted Excel spreadsheet, where each row
    represents one metadata record and each column represents a specific field.

    The following features are included:

    1. Column ordering:
       - Fields that are present in all metadata dictionaries and appear first in
         the first file will appear first in the Excel sheet.
       - Additional fields found only in some files are appended afterward in
         alphabetical order.

    2. List handling:
       - List-type fields are expanded into multiple columns with indexed suffixes
         (e.g., "keywords_01", "keywords_02", ...).
       - The original list is preserved in its own column, while individual elements
         are accessible in their own columns for easier filtering or analysis.

    3. Text cleanup:
       - Line breaks (\r, \n) and excessive spacing in string values are removed to
         ensure clean, single-line output.
       - All values are converted to strings where necessary to prevent type issues.

    4. Excel formatting:
       - The first row is bolded to clearly separate headers from data.
       - Auto-filters are added to the header row to allow for easy sorting and filtering.
       - Column widths are automatically adjusted based on content length, with a
         reasonable maximum to avoid excessive width.

    Parameters:
        dicts (List[dict]):
            The list of metadata dictionaries loaded from YAML files.
        output_path (str):
            The file path where the Excel spreadsheet should be saved.
    """
    if not dicts:
        print("No data to write.")
        return

    first_keys = list(dicts[0].keys())
    common_keys = [key for key in first_keys if all(key in d for d in dicts[1:])]

    all_keys_seen = set()
    list_lengths = {}

    for d in dicts:
        for key, value in d.items():
            all_keys_seen.add(key)
            if isinstance(value, list):
                current_len = len(value)
                if key not in list_lengths or current_len > list_lengths[key]:
                    list_lengths[key] = current_len

    extra_keys = [key for key in sorted(all_keys_seen) if key not in common_keys]

    list_expanded_keys = []
    for key, length in list_lengths.items():
        for i in range(1, length + 1):
            list_expanded_keys.append(f"{key}_{i:02d}")

    all_keys = common_keys + extra_keys + list_expanded_keys

    wb = Workbook()
    ws = wb.active
    ws.title = "Metadata"

    ws.append(all_keys)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for d in dicts:
        cleaned_row = []

        for key in common_keys + extra_keys:
            value = d.get(key, "")
            if isinstance(value, list):
                value = ", ".join(str(v) for v in value)
            elif isinstance(value, dict):
                value = "; ".join(f"{k}: {v}" for k, v in value.items())
            elif not isinstance(value, str):
                value = str(value)

            value = value.replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
            value = re.sub(r"\s{2,}", " ", value).strip()
            cleaned_row.append(value)

        for key, max_len in list_lengths.items():
            values = d.get(key, [])
            if isinstance(values, list):
                for i in range(1, max_len + 1):
                    val = values[i - 1] if i - 1 < len(values) else ""
                    if isinstance(val, str):
                        val = val.replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
                        val = re.sub(r"\s{2,}", " ", val).strip()
                    cleaned_row.append(val)
            else:
                cleaned_row.extend([""] * max_len)

        ws.append(cleaned_row)

    ws.auto_filter.ref = ws.dimensions

    for col in ws.columns:
        max_length = max((len(str(cell.value or "")) for cell in col), default=0)
        adjusted_width = min(max_length + 2, 80)
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = adjusted_width

    try:
        wb.save(output_path)
        print(f"Excel file written to: {output_path}")
    except Exception as e:
        print(f"Failed to write Excel file: {e}")

if __name__ == "__main__":
    # Use argparse to allow optional command-line path input
    parser = argparse.ArgumentParser(description="Load YAML files and export them to XLSX.")
    parser.add_argument(
        "path",
        nargs="?",
        default=os.getcwd(),
        help="Path to the root folder containing YAML files (default: current working directory)."
    )
    args = parser.parse_args()

    # Resolve the absolute path to the target folder
    root_folder = os.path.abspath(args.path)

    # Load all YAML metadata files from the specified folder
    all_metadata = load_yaml_files_recursively(root_folder)

    print(f"Loaded {len(all_metadata)} YAML files.")

    if all_metadata:
        output_xlsx = os.path.join(root_folder, "metadata_summary.xlsx")
        write_dicts_to_excel(all_metadata, output_xlsx)
