from pathlib import Path

def openW2wFile(path, mode):
    """
    Helper function to read/write all files with same encoding and line endings

    :param str|Path path: full path to file
    :param str mode: open mode: 'r' - read, 'w' - write, 'a' - append
    :return TextIO:
    """
    if isinstance(path, Path):
        path = str(path)

    # Windows line endings so that less advanced people can edit files, created on Unix in Windows Notepad
    return open(path, mode, encoding='utf-8', newline='\r\n')

def intToHtmlHex(val):
    return '#%X' % val

def iterUnoCollection(unoCollection):
    enum = unoCollection.createEnumeration()
    while enum.hasMoreElements():
        yield enum.nextElement()
