import os
import re
import yaml
import argparse
import json
import datetime
from collections import OrderedDict

# List of file name patterns to ignore (partial match or full name)
IGNORED_FILES = [
    "_quarto.yml"  # Ignores Quarto project configuration files
]

def ordered_yaml_loader():
    """
    Creates a custom YAML loader that preserves the order of keys
    as they appear in the original YAML files.

    By default, Python dictionaries in earlier versions of Python
    do not guarantee the order of keys. Using OrderedDict ensures
    that the key order from the YAML source is preserved.

    This is especially useful when the field order has semantic meaning
    or should be preserved for human readability and consistent output.

    Returns:
        A customized PyYAML loader class that returns OrderedDict instances.
    """
    class OrderedLoader(yaml.SafeLoader):
        pass

    def construct_mapping(loader, node):
        loader.flatten_mapping(node)
        return OrderedDict(loader.construct_pairs(node))

    OrderedLoader.add_constructor(
        yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG, construct_mapping
    )

    return OrderedLoader

def load_yaml_files_recursively(root_path):
    """
    Recursively scans the specified folder and all its subfolders for YAML files,
    loads their contents into dictionaries while preserving key order,
    and returns a list of these dictionaries.

    This function includes robust error handling and preprocessing steps
    to deal with common issues such as encoding mismatches and tab characters
    (which are not allowed in YAML indentation).

    Features:
    - Recognizes .yaml and .yml files
    - Skips files based on predefined patterns
    - Falls back to cp1252 encoding if UTF-8 fails
    - Replaces tabs with four spaces to conform to YAML syntax
    - Uses an OrderedDict loader to preserve field order

    Parameters:
        root_path (str):
            The base directory to start scanning for YAML files.

    Returns:
        List[OrderedDict]:
            A list of dictionaries loaded from valid YAML files.
    """
    yaml_dicts = []
    loader = ordered_yaml_loader()

    for dirpath, dirnames, filenames in os.walk(root_path):
        for filename in filenames:
            if filename.endswith((".yaml", ".yml")):
                if any(pattern in filename for pattern in IGNORED_FILES):
                    print(f"Skipping ignored file: {os.path.join(dirpath, filename)}")
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
                        yaml_dicts.append(data)
                except yaml.YAMLError as e:
                    print(f"Failed to parse YAML in {full_path}: {e}")

    return yaml_dicts

class CustomJSONEncoder(json.JSONEncoder):
    """
    Custom JSON encoder that converts non-serializable types like
    datetime.date and datetime.datetime into ISO 8601 formatted strings.
    """
    def default(self, obj):
        if isinstance(obj, (datetime.date, datetime.datetime)):
            return obj.isoformat()
        return super().default(obj)

def write_dicts_to_json(dicts, output_path):
    """
    Writes a list of dictionaries to a JSON file in a human-readable format.

    Each dictionary in the list represents metadata from a single YAML file.
    This function serializes all dictionaries into one JSON array and writes
    it to the given file path. The output is UTF-8 encoded and pretty-printed
    with indentation.

    Parameters:
        dicts (List[dict]):
            A list of metadata dictionaries to write.

        output_path (str):
            Path (including filename) where the JSON file will be saved.

    Returns:
        None. The function prints a success message or an error if writing fails.
    """
    try:
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(dicts, f, indent=2, ensure_ascii=False, cls=CustomJSONEncoder)
        print(f"JSON file written to: {output_path}")
    except Exception as e:
        print(f"Failed to write JSON file: {e}")

if __name__ == "__main__":
    # Use argparse to allow optional command-line path input
    parser = argparse.ArgumentParser(description="Load YAML files and export them to JSON.")
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
        output_json = os.path.join(root_folder, "metadata_summary.json")
        write_dicts_to_json(all_metadata, output_json)