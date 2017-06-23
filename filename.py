table = dict( (ord(c), ord('_')) for c in "|\\?*<\":>+[]/'" )

def getWindowsName(name):
    return name.translate(table)