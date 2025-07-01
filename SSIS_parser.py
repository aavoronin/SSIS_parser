import os


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
