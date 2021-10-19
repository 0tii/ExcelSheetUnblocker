import zipfile, re, os, sys, resources.checks as checks
from resources.updatezip import UpdateableZipFile as uzip

file = not sys.argv.__contains__('--folder')
filepath, folderpath = '',''

# input
if file:
    filepath = input('Enter the path to the xlsx file. Both explicit and relative are valid.\n')
else:
    folderpath = input('Enter the path to the target folder. Both explicit and relative are valid.\n')

#regex pattern to exclude sheetProtection
strclean = re.compile('<sheetProtection.*?>')

def processfile(excelfile):
    sheetfiles = {}

    # get all target sheet paths
    with zipfile.ZipFile(excelfile) as zip:
        for filename in zip.namelist():
            if filename.startswith('xl/worksheets/sheet') and filename.endswith('.xml'):
                ofile = zip.open(filename)
                sheetfiles[filename] = re.sub(strclean, '', ofile.read().decode('utf-8'))

    # alter files
    with uzip(excelfile) as file:
        for fname, fcontent in sheetfiles.items():
            file.writestr(fname, fcontent.encode('utf-8'))

# execute
if file:
    if checks.checkfile(filepath):
        processfile(filepath)
        print('\033[92mFinished.\033[0m')
else:
    if checks.checkpath(folderpath):
        for item in os.listdir(folderpath):
            if item.endswith('.xlsx'):
                processfile(item)
        print('\033[92mFinished.\033[0m')
