# SSIS_Parser

## 📘 Overview

`SSIS_Parser` is a Python tool that parses `.dtsx` files from SQL Server Integration Services (SSIS) projects.
It extracts executable structures, components, SQL statements, and connections,
then exports the data to `.xlsx` Excel files for further inspection or documentation.

This parser is helpful for:
- Documenting SSIS packages
- Auditing or migrating data pipelines
- Reverse engineering data flows

---

## 🚀 Features

- Parses `.dtsx` files recursively in a folder
- Extracts executables, pipelines, and SQL tasks
- Flattens nested structures into tabular format
- Saves output as `.xlsx` files (one per SSIS package)
- Supports:
  - `Execute SQL Task`
  - `Pipeline Task` with components and properties
  - `Execute Package Task`

---

## 🛠️ Requirements

- Python 3.8+
- [lxml](https://lxml.de/)
- [pandas](https://pandas.pydata.org/)
- [openpyxl](https://openpyxl.readthedocs.io/)

You can install dependencies via pip:

```bash
pip install lxml pandas openpyxl

You can test the project using SSIS packages from the following third-party GitHub repositories:
git clone https://github.com/AmitPNK/MSBI-Project.git
git clone https://github.com/RanaGaballah/DataWareHouse_SSIS.git
git clone https://github.com/omarkhalled/DataWarehouse-GalaxySchema-SSIS.git
