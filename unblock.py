import zipfile, os, tempfile, fnmatch
from resources.updatezip import UpdateableZipFile as uzip

fpaths = []

# get all sheet paths
zip = zipfile.ZipFile('./test.xlsx')
for filename in zip.namelist():
    if filename.startswith('xl/worksheets/sheet') and filename.endswith('.xml'):
        fpaths.append(filename)


