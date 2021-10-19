import os, resources.messages as messages

def checkpath(path):
    if not os.path.exists(path): 
        print(messages.nopath)
        return False
    if os.path.isfile(path):
        print(messages.isfile)
        return False
    count, hcount = 0, 0
    for file in os.listdir(path):
        if file.endswith('.xlsx'): count += 1
        split = file.rsplit('.', 1)[-1]
        if  split.startswith('xl') and not split.endswith('x'): hcount +=1
    if hcount > 0: 
        print(messages.xlshelp)
    if count == 0:
        print(messages.nofiles)
        return False
    print(messages.pathgood)
    return True

def checkfile(path):
    if not os.path.exists(path):
        print(messages.nofile)
        return False
    if not os.path.isfile(path):
        print(messages.ispath)
        return False
    split = path.rsplit('.', 1)[-1]
    if  split.startswith('xl') and not split.endswith('x'):
        print(messages.oneslx)
        return False
    if not path.endswith('.xlsx'): 
        print(messages.wrongfile)
        return False
    print(messages.filegood)
    return True

def checkfiles(paths):
    for path in paths:
        if not checkfile(path):
            return False
    if len(paths) == 0:
        print(messages.nopath)
        return False
    return True