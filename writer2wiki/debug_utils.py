
def inspectUno(obj):
    import unohelper, sys
    unohelper.inspect(obj, sys.stdout)

def printCharCodes(text, tabsNum=1):
    [print('\t' * tabsNum + '`{}` => 0x{:X}'.format(c, ord(c))) for c in text]

def printTextPortionProperties(portion):
    print('>', '`%s`' % portion.getString(), ':: ', hex(portion.CharColor), portion.CharWeight, portion.CharPosture)
    print('---     :', portion.CharUnderline, portion.CharUnderlineHasColor, portion.CharUnderlineColor)
    print('---     :', portion.CharStrikeout, portion.CharCaseMap, portion.CharEscapement)
