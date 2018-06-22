#!/Library/Frameworks/Python.framework/Versions/3.6/bin/python3
import sys

# Check Python version client is using at runtime (should be 3.x as specified by the shebang)
PYTHON_VERSION = "{0}.{1}.{2}".format(sys.version_info.major, sys.version_info.minor, sys.version_info.micro)
try:
    if int(sys.version_info.major) < 3:
        raise Exception("\n Need Python 3 to import 'xlrd' module necessary to parse excel files.\nExiting now...")
except Exception as err:
    print(err)
    sys.exit(0)  # Exit script cleanly
finally:
    print("\n Currently using Python version {}\n".format(PYTHON_VERSION))

from textwrap import dedent
import xlrd  # Only available in Python 3.X
import os

EXCEL_FILE_FORMATS = ['xls', 'xlsx', 'xlt', 'xltx']  # Different standard Excel formats
COLORS_ENABLED = True

# Defined outside of the '_colorText' function for run time efficiency purposes
ANSI_COLORS_DICT = {
    'red'     : '\033[1;31;40m',
    'green'   : '\033[1;32;40m',
    'yellow'  : '\033[1;33;40m',
    'blue'    : '\033[1;34;40m',
    'magenta' : '\033[1;35;40m',
    'cyan'    : '\033[1;36;40m',
    'white'   : '\033[1;37;40m',
}

def _colorText(msg, color):
    try:
        if COLORS_ENABLED:
            msg = "{0}{1}{2}".format(ANSI_COLORS_DICT[color], msg, ANSI_COLORS_DICT[colors])
    except KeyError:
        # Code to log error  (TO_DO)
        pass
    finally:
        return msg

def _commandLineArgsIn():
    """
    Detects if command line arguments were specified
    """
    return True if len(sys.argv) >= 2 else False

def _fileExists(file_path):
    return os.path.exists(file_path)

def _getNumExcelFilesInCWD():
    """
    Count the number of excel files in the current working directory
    """
    _excelFileCounter = 0
    for item in os.listdir(os.getcwd()):
        if item[-item.index('.'):] in EXCEL_FILE_FORMATS:
            _excelFileCounter += 1
    return _excelFileCounter

def _getNumExcelFilesInCWD():
    _excel_file_list = []
    for item in os.listdir(os.getcwd()):
        if item[-item.index('.'):] in EXCEL_FILE_FORMATS:
            _excel_file_list.append(item)
    return _excel_file_list


def getExcelFile():
    if _commandLineArgsIn() and _fileExists(sys.argv[2]):
        return sys.argv[2]  # User provided an excel file through command line argument
    else:
        if _getNumExcelFilesInCWD() == 0:
            _errMsg = """
                Error:  No excel file found in {}"
                Recommended:  Please move excel file to the folder you are running this parser in
                OR specify the full path to the excel file as a command line argument.
                """
            print(dedent(_errMsg.format(os.getcwd())))
            sys.exit(0)
        elif _getNumExcelFilesInCWD() == 1:  # (TO_DO)
            pass  
        else:  # number of excel files in CWD > 1  (TO_DO)
            pass


