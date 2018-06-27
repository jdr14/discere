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
import time
import os

EXCEL_FILE_FORMATS = ['xls', 'xlsx', 'xlt', 'xltx']  # Different standard Excel formats
COLORS_ENABLED = True

# Defined outside of the '_colorText' function for run time efficiency purposes
ANSI_COLORS_DICT = {
    'red'     : '1;31;40m',
    'green'   : '1;32;40m',
    'yellow'  : '1;33;40m',
    'blue'    : '1;34;40m',
    'magenta' : '1;35;40m',
    'cyan'    : '1;36;40m',
    'white'   : '1;37;40m',
    'escape'  : '0;0m',
}

def _colorText(msg, color):
    try:
        if COLORS_ENABLED:
            msg = "{0}{1}{2}".format('\033[{}'.format(ANSI_COLORS_DICT[color]), 
                                     msg, 
                                     '\033[{}'.format(ANSI_COLORS_DICT['reset']))
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
        if item[item.index('.') + 1:] in EXCEL_FILE_FORMATS:
            _excelFileCounter += 1
    return _excelFileCounter

def _getListExcelFilesInCWD():
    """
    Gets the number of excel files in 
    """
    _excel_file_list = []
    for item in os.listdir(os.getcwd()):
        if item[item.index('.') + 1:] in EXCEL_FILE_FORMATS:
            _excel_file_list.append(item)
    return _excel_file_list

def _checkQuit(input):
    """
    Helper function to cleanly exit the program
    """
    if input == 'q':
        print(_colorText(" Exiting now... Goodbye.", 'magenta'))
        time.sleep(2)  # Pause 
        sys.exit(0)

def getExcelFile():
    """
    """
    # User provided an excel file through command line argument
    if _commandLineArgsIn() and _fileExists(sys.argv[1]):
        return sys.argv[1]  
    
    # User did not provide a custom path to excel file.  Check CWD for valid file to parse
    if _getNumExcelFilesInCWD() == 0:
        _errMsg = """
            Error:  No excel file found in {}"
            Recommended:  Please move excel file to the folder you are running this parser in
            OR specify the full path to the excel file as a command line argument.
            """
        print(_colorText(dedent(_errMsg.format(os.getcwd())), 'red'))
        sys.exit(0)
    elif _getNumExcelFilesInCWD() == 1:  # (TO_DO)
        return _getListExcelFilesInCWD()[0]
    else:  # number of excel files in CWD > 1  (TO_DO)
        # Store the excel files in a list var to avoid reiterating (thereby improving performance)
        file_list = _getListExcelFilesInCWD()
        
        # Neatly print out all of the file options for the user to choose from with corresponding numbers
        for excel_file in file_list:
            print(_colorText(" ( {0} )  {1}".format(file_list.index(excel_file), excel_file), 'yellow'))
        
        # Get user input
        user_in = input("\n Please enter a number corresponding to the file you want to parse.\n File #> ")
        
        # Ensure the user input is valid
        while not user_in.isdigit() or not int(user_in) in range(0, len(file_list)):
            user_in = input("\n Please enter a number between 0 and {} or enter 'q' to quit.\n File #> ".format(len(file_list)))
            
            # Leave the user the option to cleanly exit the program
            _checkQuit(user_in)
        
        # User input should be valid now once at this stage.  Use it to select the excel file
        return file_list[int(user_in)]    

print(getExcelFile())



