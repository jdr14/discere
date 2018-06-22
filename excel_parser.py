#!/Library/Frameworks/Python.framework/Versions/3.6/bin/python3
import sys

# Check Python version client is using at runtime (should be 3.x as specified by the shebang)
PYTHON_VERSION = "{0}.{1}.{2}".format(sys.version_info.major, sys.version_info.minor, sys.version_info.micro)
try:
    if int(sys.version_info.major) < 3:
        raise Exception("\nNeed Python 3 to import 'xlrd' module necessary to parse excel files.\nExiting now...")
except Exception as err:
    print(err)
    sys.exit(0)  # Exit script cleanly
finally:
    print("\nCurrently using Python version {}\n".format(PYTHON_VERSION))

import xlrd
import os

def _commandLineArgsIn():
	return True if len(sys.argv) >= 2 else False

