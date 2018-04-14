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
    return '#{:0>6X}'.format(val)

def iterUnoCollection(unoCollection):
    enum = unoCollection.createEnumeration()
    while enum.hasMoreElements():
        yield enum.nextElement()

def flushLogger():
    import logging as log

    for weakRef in log._handlerList:
        handler = weakRef()
        if not handler:
            continue

        print('flushing log handler:', handler)
        handler.acquire()
        try:
            handler.flush()
        finally:
            handler.release()


def surroundWithTag(content, tag, tagAttributes=''):
    if tagAttributes:
        tagAttributes = ' ' + tagAttributes
    return '<{0}{1}>{2}</{0}>'.format(tag, tagAttributes, content)