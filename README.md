# yaml_metadata

A YAML-based metadata template and Python scripts that build on it to support research data management.

For more information on the YAML syntax, visit https://yaml.org/

## Description

This metadata file template supports a minimalistic approach to research data management, based on two key assumptions:

1. Research data will only be managed professionally — and on a voluntary basis — if the process is as simple and accessible as possible, and can be carried out with minimal time investment.

2. As a common denominator across different researchers and technical environments, research data is typically stored in a folder structure — for example, on a personal computer, a network drive, or in the cloud.

Under these conditions, metadata from different studies can be stored in a distributed and decentralized manner using the provided template.

A folder structure with subdirectories for each dataset — following a suggested naming convention — is recommended.

With the help of Python scripts, the metadata can be aggregated in various formats and made usable — for example, as Excel spreadsheets or JSON files for import into a document-based database (e.g., MongoDB).

## Running the Scripts

The scripts require a few third-party Python packages to run (mainly a package to read YAML files). Follow the steps below to set up your environment and execute the code.

### Prerequisites

* [Git](https://git-scm.com/) – to clone the repository
* Python 3.10 or higher recommended
* (Optional) Virtual environment tool: `venv` or `virtualenv`

### Installation From Command Line

1. **Clone the repository:**

   ```bash
   git clone https://github.com/ctreffe/yaml_metadata.git
   cd yaml_metadata
   ```

2. **(Optional but recommended) Create and activate a virtual environment:**

   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies:**

   ```bash
   pip install -r requirements.txt
   ```

4. **Execute one of these commands:**

   ```bash
   python metadata_to_csv.py
   python metadata_to_json.py
   python metadata_to_xlsx.py

   # With Custom path (absolute or relative)
   python metadata_to_csv.py "/path/to/your/research data/folder"
   python metadata_to_json.py "/path/to/your/research data/folder"
   python metadata_to_xlsx.py "/path/to/your/research data/folder"
   ```