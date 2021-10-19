# 🔓 Excel Sheet Unlocker

Remove sheet protection from .xlsx files.

## How to use

Run the script/packaged executable from the command line.

Universal unpackaged
```
py unblock.py [--cli] [--dir]
```

Powershell:
```PowerShell
& "C:\Users\...\unblock.exe" [--cli] [--dir]
```

Cmd
```
unblock.exe [--cli] [--dir]
```

### Options

## `--dir`
Select a directory instead of individual files. All .xlsx files in the directory will be processed.

## `--cli`
Replaces the GUI file/directory picker with cli-input paths. Paths can be specified both explicitly and relative.

## Requirements

Written in Python 3.9.7, should work on 3.x. Depends solely on std libraries.

## Output

Input-files will be overwritten with unblocked versions.

## Limits

This only works on .xlsx files. Older excel formats should be converted using Excel 'Save As' -> '.xlsx'