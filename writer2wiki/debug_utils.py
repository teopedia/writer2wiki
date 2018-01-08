
def inspectUno(obj):
    """ Print all fields and methods `obj` supports """
    import unohelper, sys
    unohelper.inspect(obj, sys.stdout)

def printCentered(msg):
    print(" {} ".format(msg).center(80, '-'))

def inspectUnoProperties(obj, properties):
    """
    Call `inspectUno` for all `properties` of the obj.
    Example:
        inspectUnoProperties(textPortion, ['CharStyleNames', 'TextUserDefinedAttributes', 'CharWeight'])

        output:
            --------------- CharStyleNames is None ---------------
            --------------- TextUserDefinedAttributes ---------------
            Supported services:
              com.sun.star.xml.AttributeContainer
            Interfaces:
              com.sun.star.lang.XServiceInfo
            ...
            --------------- CharWeight = 100.0 ---------------
    """

    for p in properties:
        if not hasattr(obj, p):
            printCentered("{} doesn't exist".format(p))
            continue

        propValue = getattr(obj, p)
        if propValue is None:
            printCentered("{} is None".format(p))
            continue
        if isinstance(propValue, (bool, int, float, str, dict, list, set)):
            printCentered("{} = {}".format(p, propValue))
            continue

        printCentered(p)
        inspectUno(propValue)
        printCentered('end of {}'.format(p))

def printCharCodes(text, tabsNum=1):
    [print('\t' * tabsNum + '`{}` => 0x{:X}'.format(c, ord(c))) for c in text]

def printTextPortionProperties(portion):
    print('>', '`%s`' % portion.getString(), ':: ', hex(portion.CharColor), portion.CharWeight, portion.CharPosture)
    print('---     :', portion.CharUnderline, portion.CharUnderlineHasColor, portion.CharUnderlineColor)
    print('---     :', portion.CharStrikeout, portion.CharCaseMap, portion.CharEscapement)
