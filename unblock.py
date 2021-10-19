import zipfile, re, os, sys, resources.checks as checks
from resources.updatezip import UpdateableZipFile as uzip

file = False
filepath = './test2.xlsx'
folderpath = './'

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
    if checks.checkFile(filepath):
        processfile(filepath)
        print('\033[92mFinished.\033[0m')
else:
    if checks.checkPath(folderpath):
        for item in os.listdir(folderpath):
            if item.endswith('.xlsx'):
                processfile(item)
        print('\033[92mFinished.\033[0m')
