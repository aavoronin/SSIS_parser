import os
import json
from lxml import etree
import pandas as pd

class SSIS_Parser:
    def __init__(self):
        # List of .dtsx files to parse
        self.files_to_parse = []

    def flatten_executables(self, executables, parent=None):
        """
        Recursively flattens nested SSIS executables into a flat list of dictionaries.
        Adds component-level information for pipeline tasks.
        """
        rows = []

        for exe in executables:
            # Basic executable information
            base = {
                "ExecutableID": exe.get("ID"),
                "ExecutableType": exe.get("Type"),
                "ExecutableTag": exe.get("Tag"),
                "ExecutedPackageName": exe.get("ExecutedPackageName"),
                "SqlStatementSource": exe.get("SqlStatementSource"),
                "ConnectionID": exe.get("ConnectionID"),
            }

            if parent:
                # Add parent ID for hierarchy tracing if desired
                base["ParentExecutableID"] = parent.get("ID")

            if exe.get("Type") == "Microsoft.Pipeline" and "Components" in exe:
                # For pipelines: one row per component
                for comp in exe["Components"]:
                    row = base.copy()
                    row.update({
                        "ComponentName": comp.get("Name"),
                        "ComponentType": comp.get("Type"),
                    })
                    # Flatten component properties into individual columns
                    for prop_name, prop_val in comp.get("Properties", {}).items():
                        row[f"Property_{prop_name}"] = prop_val
                    rows.append(row)
            else:
                # For non-pipeline tasks: just the executable data
                rows.append(base)

            # Recursively process nested children executables
            children = exe.get("Children", [])
            if children:
                rows.extend(self.flatten_executables(children, parent=exe))

        return rows

    def parse(self, folder_path):
        """
        Entry point to parse all .dtsx files in a folder and save the extracted data as Excel files.
        """
        if not os.path.isdir(folder_path):
            raise ValueError(f"The provided path '{folder_path}' is not a valid directory.")

        # Walk through the folder tree and find .dtsx files
        for root, _, files in os.walk(folder_path):
            for file in files:
                if file.endswith(".dtsx"):
                    full_path = os.path.join(root, file)
                    self.files_to_parse.append(full_path)

        print(f'{len(self.files_to_parse)} files to parse')

        for dtsx_file in self.files_to_parse:
            parsed_data = {}
            self.parse_single_file(dtsx_file, parsed_data)
            print(json.dumps(parsed_data, indent=2))  # Optional debug print

            # Flatten hierarchical executable structure
            rows = self.flatten_executables(parsed_data.get("Executables", []))

            # Create DataFrame and export to Excel
            df = pd.DataFrame(rows)
            excel_file = os.path.splitext(dtsx_file)[0] + ".xlsx"
            df.to_excel(excel_file, index=False)
            print(f"Saved parsed data to {excel_file}")

    def parse_single_file(self, dtsx_file, parsed_data):
        """
        Parses a single .dtsx file and populates the parsed_data dictionary.
        """
        try:
            tree = etree.parse(dtsx_file)
            root = tree.getroot()
            print(f"\nParsing file: {dtsx_file}")

            # Define XML namespaces used in SSIS packages
            ns = {
                'DTS': "www.microsoft.com/SqlServer/Dts",
                'SQLTask': 'www.microsoft.com/sqlserver/dts/tasks/sqltask'
            }

            # Extract all top-level executables recursively
            parsed_data['Executables'] = self.parse_executables(root, level=0, ns=ns)

        except Exception as e:
            print(f"Error parsing {dtsx_file}: {e}")

    def parse_executables(self, xml_element, level=0, ns=None):
        """
        Recursively parses all <DTS:Executable> nodes from XML.
        """
        if ns is None:
            raise ValueError("Namespace dictionary 'ns' must be provided")

        # XPath to get child executables
        executables = xml_element.xpath("./DTS:Executables/DTS:Executable", namespaces=ns)
        result = []

        for exe in executables:
            executable_id = exe.get(f"{{{ns['DTS']}}}refId")
            executable_type = exe.get(f"{{{ns['DTS']}}}ExecutableType")
            tag_name = etree.QName(exe).localname

            # Base executable dictionary
            exe_dict = {
                "ID": executable_id,
                "Type": executable_type,
                "Tag": tag_name
            }

            # Specialized parsing for known executable types
            if executable_type == "Microsoft.ExecutePackageTask":
                package_name_nodes = exe.xpath(
                    "./DTS:ObjectData/ExecutePackageTask/PackageName/text()",
                    namespaces=ns
                )
                exe_dict["ExecutedPackageName"] = package_name_nodes[0] if package_name_nodes else None

            elif executable_type == "Microsoft.Pipeline":
                self.parse_pipeline(exe, exe_dict, ns)

            elif executable_type == "Microsoft.ExecuteSQLTask":
                self.parse_execute_sql_task(exe, exe_dict, ns)

            # Recursively parse nested executables
            children = self.parse_executables(exe, level=level + 1, ns=ns)
            exe_dict["Children"] = children

            result.append(exe_dict)

        return result

    def parse_execute_sql_task(self, exe, exe_dict, ns):
        """
        Extracts SQL command and connection for SQL Task executables.
        """
        sql_node = exe.xpath(
            "./DTS:ObjectData/SQLTask:SqlTaskData",
            namespaces=ns
        )

        if sql_node:
            sql_task = sql_node[0]
            connection_id = sql_task.get(f"{{{ns['SQLTask']}}}Connection")
            sql_statement = sql_task.get(f"{{{ns['SQLTask']}}}SqlStatementSource")
            exe_dict["ConnectionID"] = connection_id
            exe_dict["SqlStatementSource"] = sql_statement

    def parse_pipeline(self, exe, exe_dict, ns):
        """
        Parses the Data Flow (Pipeline) task, extracting components and their properties.
        """
        components = []
        comp_nodes = exe.xpath(
            "./DTS:ObjectData/pipeline/components/component",
            namespaces=ns
        )

        for comp in comp_nodes:
            comp_name = comp.get("name")
            comp_type = comp.get("componentClassID")

            comp_dict = {
                "Name": comp_name,
                "Type": comp_type,
                "Properties": {}
            }

            # Parse all <property> children under this component
            properties_nodes = comp.xpath("./properties/property", namespaces=ns)
            for prop in properties_nodes:
                prop_name = prop.get("name")
                prop_value = prop.text if prop.text is not None else ""
                comp_dict["Properties"][prop_name] = prop_value

            components.append(comp_dict)

        exe_dict["Components"] = components
