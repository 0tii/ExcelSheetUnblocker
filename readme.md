# 🔓 Excel Sheet Unlocker

Remove sheet protection from .xlsx files.

## How to use


### Run

Run the script/packaged executable from the command line.

Universal unpackaged
```PowerShell
py unblock.py [--cli] [--dir]
```

Powershell:
```PowerShell
& "C:\Users\...\unblock.exe" [--cli] [--dir]
```

Cmd
```PowerShell
unblock.exe [--cli] [--dir]
```

### Selecting files

- If starting without options, you will be presented a multi-file picker. Choose any amount of xlsx files to unblock, select 'Open' and you are done. ![Standard Picker](https://i.imgur.com/PHy4rrC.png)

- If starting with `--dir` option, you will be presented a directory picker. Select your directory and click 'Select Folder'. Unblocker will then check for and process all viable files in the directory. ![Directory Picker](https://i.imgur.com/aDC2FVj.png)

- Should you want to run this as `--cli` application, you will be prompted to enter paths to the target files. In standard mode, you will only be able to point to a single file, in `--dir` mode, you will be able to supply a directory path.

### Options

## `--dir`
Select a directory instead of individual files. All .xlsx files in the directory will be processed.

## `--cli`
Replaces the GUI file/directory picker with cli-input paths. Paths can be specified both explicitly and relative.

## Requirements

Written in Python 3.9.7, should work on 3.x. Depends solely on std libraries.

## Output

Input-files will be overwritten with unlocked versions.

## Limits

Only works on .xlsx files. Older excel formats should be converted using Excel, via 'Save As' -> '.xlsx'

## Todo

- More transparent logging
- Auto-convert xls to xlsx

