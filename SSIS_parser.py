import os
import json
from lxml import etree
import os
import json
import pandas as pd


class SSIS_Parser:
    def __init__(self):
        self.parsed_structure = {}
        self.files_to_parse = []

    def flatten_executables(self, executables, parent=None):
        """
        Recursively flatten executables list into a list of dicts.
        Each executable becomes a dict.
        For pipeline executables, each component generates a separate row,
        merging executable, component, and component properties.
        """
        rows = []

        for exe in executables:
            base = {
                "ExecutableID": exe.get("ID"),
                "ExecutableType": exe.get("Type"),
                "ExecutableTag": exe.get("Tag"),
                "ExecutedPackageName": exe.get("ExecutedPackageName"),
            }
            if parent:
                # Optionally add parent ID if you want hierarchy trace
                base["ParentExecutableID"] = parent.get("ID")

            if exe.get("Type") == "Microsoft.Pipeline" and "Components" in exe:
                # One row per component
                for comp in exe["Components"]:
                    row = base.copy()
                    row.update({
                        "ComponentName": comp.get("Name"),
                        "ComponentType": comp.get("Type"),
                    })
                    # Flatten properties dictionary
                    for prop_name, prop_val in comp.get("Properties", {}).items():
                        row[f"Property_{prop_name}"] = prop_val
                    rows.append(row)
            else:
                # Executable row only
                rows.append(base)

            # Recursively flatten children
            children = exe.get("Children", [])
            if children:
                rows.extend(self.flatten_executables(children, parent=exe))

        return rows

    def parse(self, folder_path):
        if not os.path.isdir(folder_path):
            raise ValueError(f"The provided path '{folder_path}' is not a valid directory.")

        for root, _, files in os.walk(folder_path):
            for file in files:
                if file.endswith(".dtsx"):
                    full_path = os.path.join(root, file)
                    self.files_to_parse.append(full_path)

        print(f'{len(self.files_to_parse)} files to parse')

        for dtsx_file in self.files_to_parse:
            parsed_data = {}
            self.parse_single_file(dtsx_file, parsed_data)
            print(json.dumps(parsed_data, indent=2))

            # Flatten parsed_data["Executables"] into rows
            rows = self.flatten_executables(parsed_data.get("Executables", []))

            # Create DataFrame
            df = pd.DataFrame(rows)

            # Excel file name
            excel_file = os.path.splitext(dtsx_file)[0] + ".xlsx"
            df.to_excel(excel_file, index=False)
            print(f"Saved parsed data to {excel_file}")

    def parse_single_file(self, dtsx_file, parsed_data):
        try:
            tree = etree.parse(dtsx_file)
            root = tree.getroot()
            print(f"\nParsing file: {dtsx_file}")

            nsmap = root.nsmap
            dts_ns = nsmap.get('DTS')
            if dts_ns is None:
                raise Exception("DTS namespace not found in the XML.")
            ns = {'DTS': dts_ns}

            parsed_data['Executables'] = self.parse_executables(root, level=0, ns=ns)

        except Exception as e:
            print(f"Error parsing {dtsx_file}: {e}")

    def parse_executables(self, xml_element, level=0, ns=None):
        if ns is None:
            raise ValueError("Namespace dictionary 'ns' must be provided")

        executables = xml_element.xpath("./DTS:Executables/DTS:Executable", namespaces=ns)
        result = []

        for exe in executables:
            executable_id = exe.get(f"{{{ns['DTS']}}}refId")
            executable_type = exe.get(f"{{{ns['DTS']}}}ExecutableType")
            tag_name = etree.QName(exe).localname

            exe_dict = {
                "ID": executable_id,
                "Type": executable_type,
                "Tag": tag_name
            }

            if executable_type == "Microsoft.ExecutePackageTask":
                package_name_nodes = exe.xpath(
                    "./DTS:ObjectData/ExecutePackageTask/PackageName/text()",
                    namespaces=ns
                )
                exe_dict["ExecutedPackageName"] = package_name_nodes[0] if package_name_nodes else None

            elif executable_type == "Microsoft.Pipeline":
                self.parse_pipeline(exe, exe_dict, ns)

            children = self.parse_executables(exe, level=level + 1, ns=ns)
            exe_dict["Children"] = children

            result.append(exe_dict)

        return result

    def parse_pipeline(self, exe, exe_dict, ns):
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

            # Parse properties if present
            properties_nodes = comp.xpath("./properties/property", namespaces=ns)
            for prop in properties_nodes:
                prop_name = prop.get("name")
                # Text content can be empty or None, fallback to empty string
                prop_value = prop.text if prop.text is not None else ""
                comp_dict["Properties"][prop_name] = prop_value

            components.append(comp_dict)

        exe_dict["Components"] = components



