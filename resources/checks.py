import os, resources.messages as messages

def checkpath(path):
    if not os.path.exists(path): 
        print(messages.nopath)
        return False
    if os.path.isfile(path):
        print(messages.isfile)
        return False
    count, scount = 0, 0
    for file in os.listdir(path):
        if file.endswith('.xlsx'): count += 1
        if file.endswith('xls'): scount +=1
        #other formats ...
    if scount > 0: 
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
    if path.endswith('.xls'):
        print(messages.oneslx)
        return False

    if not path.endswith('.xlsx'): 
        print(messages.wrongfile)
        return False
    print(messages.filegood)
    return True
