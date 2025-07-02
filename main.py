from SSIS_parser import SSIS_Parser

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    parser = SSIS_Parser()
    #git clone https://github.com/AmitPNK/MSBI-Project.git
    parser.parse(r'c:\Py\SSIS_parsing\MSBI-Project\SSIS_Project')
    #git clone https://github.com/RanaGaballah/DataWareHouse_SSIS.git
    parser.parse(r'c:\Py\SSIS_parsing\DataWareHouse_SSIS')
    #git clone https://github.com/omarkhalled/DataWarehouse-GalaxySchema-SSIS.git
    parser.parse(r'c:\Py\SSIS_parsing\DataWarehouse-GalaxySchema-SSIS')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
