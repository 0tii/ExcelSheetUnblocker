xlshelp = '''\033[93m
You have one or more .xls files in the target directory. That file format is not supported!
To use these files with this script, please follow these simple steps:
1. Open the xls file in Excel
2. Go to file -> Save as...
3. Save the file as .xlsx\n\033[0m
'''

oneslx = '''\033[93m
Your file is of type xls. That file format is not supported!
To use this file with this script, please follow these simple steps:
1. Open the xls file in Excel
2. Go to file -> Save as...
3. Save the file as .xlsx\n\033[0m
'''

nofiles = '''\033[91mThere are no files of xlsx format in the directory provided\033[0m'''

wrongfile = '''\033[91mThe path provided does not have the correct format (.xlsx) or is not a file\033[0m'''

nofile = '''\033[91mThe path provided does not exist\033[0m'''

nopath = '''\033[91mThe path provided does not exist\033[0m'''

pathgood = '''\033[92mPath accepted\033[0m'''

filegood = '''\033[92mFile accepted\033[0m'''

invalidarg = '''\033[91mInvalid launch parameters provided\033[0m'''

isfile = '''\033[91mPath is file, not directory\033[0m'''

ispath = '''\033[91mPath is directory, not file\033[0m'''