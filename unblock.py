import zipfile, re, io
from resources.updatezip import UpdateableZipFile as uzip

excelfile = './test.xlsx'
sheetfiles = {}

strclean = re.compile('<sheetProtection.*?>') #regex exclude sheetProtection tag

# get all sheet paths
zip = zipfile.ZipFile(excelfile)
for filename in zip.namelist():
    if filename.startswith('xl/worksheets/sheet') and filename.endswith('.xml'):
        ofile = zip.open(filename)
        sheetfiles[filename] = re.sub(strclean, '', ofile.read().decode('utf-8'))
zip.close()

# alter files
with uzip(excelfile) as file:
    for fname, fcontent in sheetfiles.items():
        file.writestr(fname, fcontent.encode('utf-8'))


