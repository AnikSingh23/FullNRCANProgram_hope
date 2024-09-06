import os

# This file is used to create the exe's from the py files in the dist folder of the pycharm folder

# log_file = os.path.dirname(os.path.abspath(__file__)) + '/logfile.txt'

# Creates the exe files with all the hidden imports required
# os.system("pyinstaller --onefile AGG_Imputer.py --i=canadian-maple-leaf.ico --collect-submodules sklearn --collect-all xgboost ")
# os.system("pyinstaller --onefile AGR_Imputer.py --i=canadian-maple-leaf.ico --collect-submodules sklearn --collect-all xgboost ")
# os.system("pyinstaller --onefile COM_Imputer.py --i=canadian-maple-leaf.ico --collect-submodules sklearn --collect-all xgboost ")
# os.system("pyinstaller --onefile Combined_imputer.py --i=canadian-maple-leaf.ico --collect-submodules sklearn --collect-all xgboost ")
# os.system("pyinstaller --onefile IND_Imputer.py --i=canadian-maple-leaf.ico --collect-submodules sklearn --collect-all xgboost ")
# os.system("pyinstaller --onefile RES_Imputer.py --i=canadian-maple-leaf.ico --collect-submodules sklearn --collect-all xgboost ")
# os.system("pyinstaller --onefile TRA_Imputer.py --i=canadian-maple-leaf.ico --collect-submodules sklearn --collect-all xgboost ")


# os.system("pyinstaller --onefile  Complete_Program.py --i=canadian-maple-leaf.ico --collect-submodules sklearn --collect-all xgboost ")
os.system("pyinstaller --onefile  Just_Imputer.py --i=canadian-maple-leaf.ico --collect-submodules sklearn --collect-all xgboost ")

# os.system(
#     "pyinstaller --onefile  DelAndMoveFiles.py --i=canadian-maple-leaf.ico --collect-submodules sklearn --collect-all xgboost ")

