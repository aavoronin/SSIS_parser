import os
from lxml import etree


class SSIS_Parser:
    def __init__(self):
        # Stores parsed SSIS package structures
        self.parsed_structure = {}
        # List of .dtsx files to be parsed
        self.files_to_parse = []

    def parse(self, folder_path):
        """
        Scans the given folder for .dtsx files and stores their full paths
        in the files_to_parse list.
        """
        if not os.path.isdir(folder_path):
            raise ValueError(f"The provided path '{folder_path}' is not a valid directory.")

        for root, _, files in os.walk(folder_path):
            for file in files:
                if file.endswith(".dtsx"):
                    full_path = os.path.join(root, file)
                    self.files_to_parse.append(full_path)

        print(f'{len(self.files_to_parse)} files to parse')
        
        for dtsx_file in self.files_to_parse:
            self.parse_single_file(dtsx_file)

    def parse_single_file(self, dtsx_file):
        try:
            tree = etree.parse(dtsx_file)
            root = tree.getroot()
            print(f"\nParsing file: {dtsx_file}")
            self.parse_executables(root, level=0)
        except Exception as e:
            print(f"Error parsing {dtsx_file}: {e}")

    def parse_executables(self, xml_element, level=0):
        nsmap = xml_element.nsmap
        ns = {'DTS': nsmap.get(None)}  # Default namespace

        for exe in xml_element.xpath("./DTS:Executables/DTS:Executable", namespaces=ns):
            executable_id = exe.get("DTS:refId")
            executable_name = exe.get("DTS:ExecutableType")
            executable_type = exe.tag.split('}')[-1]  # remove namespace

            indent = "  " * level
            print(f"{indent}- ID: {executable_id}, Name: {executable_name}, Type: {executable_type}")

            # Recursively parse any nested Executables
            self.parse_executables(exe, level=level+1)

