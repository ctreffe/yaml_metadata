import os
import re
import yaml
import argparse
import csv
from collections import OrderedDict

# List of file name or file path patterns to ignore (partial match or full name)
IGNORED_FILES = [
    "_quarto.yml",  # Ignores Quarto project configuration files
    "\\renv",  # Ignores renv environment paths
    "_metadata.yaml"  # Ignore a specific template file
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
                # Check if file should be ignored
                if any(pattern in filename for pattern in IGNORED_FILES):
                    print(f"Skipping ignored file: {os.path.join(dirpath, filename)}")
                    continue

                if any(pattern in dirpath for pattern in IGNORED_FILES):
                    print(f"Skipping ignored filepath: {dirpath}")
                    continue

                full_path = os.path.join(dirpath, filename)

                # Try UTF-8 first, then fallback to cp1252
                file_content = None
                encoding_used = "utf-8"
                try:
                    with open(full_path, "r", encoding="utf-8") as file:
                        file_content = file.read()
                except UnicodeDecodeError:
                    try:
                        with open(full_path, "r", encoding="cp1252") as file:
                            file_content = file.read()
                        encoding_used = "cp1252"
                        print(
                            f"Warning: '{full_path}' is not UTF-8. Read using cp1252 fallback."
                        )
                    except Exception as e:
                        print(f"Failed to read {full_path}: {e}")
                        continue

                if file_content is None:
                    continue

                # Replace tabs with 4 spaces
                if "\t" in file_content:
                    print(f"Note: Tabs found in '{full_path}', replacing with spaces.")
                    file_content = file_content.replace("\t", "    ")

                # Try parsing YAML
                try:
                    data = yaml.load(file_content, Loader=loader)
                    if isinstance(data, dict):
                        # Add relative path as string to each data set
                        rel_path = os.path.relpath(dirpath, root_path)
                        data["filepath"] = rel_path.replace("\\", "/")  # Same format for Windows/Linux
                        yaml_dicts.append(data)
                except yaml.YAMLError as e:
                    print(f"Failed to parse YAML in {full_path}: {e}")

    return yaml_dicts



def write_dicts_to_csv(dicts, output_path):
    """
    Converts a list of metadata records (loaded from YAML files) into a clean, structured
    CSV file that is easy to read and analyze.

    Each metadata record is expected to be a dictionary containing field names and values.
    The script writes all records into a single CSV file, where each row corresponds to one
    YAML file, and each column represents one metadata field.

    Key features of this function:

    1. Column ordering for better readability:
       - Fields that appear in *all* metadata files and are listed first in the first file
         will be placed at the beginning of the table (preserving their original order).
       - Any additional or optional fields (those that are only present in some files)
         will be added at the end of the table in alphabetical order.

    2. Handling of list values:
       - If a metadata field contains a list (e.g., tags, keywords, authors),
         the original list remains visible in the table.
       - In addition, each list item is extracted and written into its own separate column.
         For example, a field called "keywords" with 3 items becomes:
         "keywords_01", "keywords_02", "keywords_03".
       - If some files have longer lists than others, more columns will be added accordingly.
         Missing values in shorter lists remain empty.

    3. Cleanup of formatting issues:
       - Line breaks in text fields (e.g., in descriptions) are removed and replaced with spaces,
         so that each metadata record fits neatly into one row in the table.
       - Unnecessary extra spaces are collapsed into single spaces.
       - Leading and trailing spaces are removed to keep the output clean.

    Parameters:
        dicts (List[dict]):
            A list of dictionaries containing metadata information from YAML files.

        output_path (str):
            The full path (including filename) where the resulting CSV file should be saved.

    Output:
        A well-formatted CSV file is written to the specified location.
        The file can be opened with Excel or any other spreadsheet program.
        It provides a clean overview of all metadata records in tabular form,
        with additional columns for list items.
    """
    if not dicts:
        print("No data to write.")
        return

    first_keys = list(dicts[0].keys())
    common_keys = [key for key in first_keys if all(key in d for d in dicts[1:])]

    # Collect all unique keys across all dicts
    all_keys_seen = set()
    list_lengths = {}  # Track maximum length of list per key

    for d in dicts:
        for key, value in d.items():
            all_keys_seen.add(key)
            if isinstance(value, list):
                current_len = len(value)
                if key not in list_lengths or current_len > list_lengths[key]:
                    list_lengths[key] = current_len

    # Extra keys not common across all
    extra_keys = [key for key in sorted(all_keys_seen) if key not in common_keys]

    # Prepare list-based expanded keys
    list_expanded_keys = []
    for key, length in list_lengths.items():
        for i in range(1, length + 1):
            list_expanded_keys.append(f"{key}_{i:02d}")

    # Final column order: common keys + extra keys + expanded list keys (no duplicates)
    all_keys = common_keys + extra_keys + list_expanded_keys

    try:
        with open(output_path, "w", encoding="utf-8", newline="") as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=all_keys)
            writer.writeheader()

            for d in dicts:
                cleaned_row = {}

                # First: process all standard fields
                for key in common_keys + extra_keys:
                    value = d.get(key, "")
                    if isinstance(value, str):
                        value = value.replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
                        value = re.sub(r"\s{2,}", " ", value).strip()
                    cleaned_row[key] = value

                # Then: process list fields into multiple columns
                for key, max_len in list_lengths.items():
                    values = d.get(key, [])
                    if isinstance(values, list):
                        for i in range(1, max_len + 1):
                            col_name = f"{key}_{i:02d}"
                            cleaned_value = values[i - 1] if i - 1 < len(values) else ""
                            if isinstance(cleaned_value, str):
                                cleaned_value = (
                                    cleaned_value.replace("\r\n", " ")
                                    .replace("\n", " ")
                                    .replace("\r", " ")
                                )
                                cleaned_value = re.sub(r"\s{2,}", " ", cleaned_value).strip()
                            cleaned_row[col_name] = cleaned_value
                    else:
                        # If not a list, leave list columns empty
                        for i in range(1, max_len + 1):
                            cleaned_row[f"{key}_{i:02d}"] = ""

                writer.writerow(cleaned_row)

        print(f"CSV file written to: {output_path}")
    except Exception as e:
        print(f"Failed to write CSV file: {e}")


if __name__ == "__main__":
    # Use argparse to allow optional command-line path input
    parser = argparse.ArgumentParser(description="Load YAML files and export them to CSV.")
    parser.add_argument(
        "path",
        nargs="?",
        default=os.getcwd(),
        help="Path to the root folder containing YAML files (default: current working directory)."
    )
    args = parser.parse_args()

    # Resolve the absolute path to the target folder
    root_folder = os.path.abspath(args.path)

    # Load all YAML metadata files into a list of dictionaries
    all_metadata = load_yaml_files_recursively(root_folder)

    print(f"Loaded {len(all_metadata)} YAML files.")

    # Write the collected metadata to a CSV file in the current folder
    if all_metadata:
        output_csv = os.path.join(root_folder, "metadata_summary.csv")
        write_dicts_to_csv(all_metadata, output_csv)
