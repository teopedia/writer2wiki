#           Copyright Alexander Malahov 2017.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE_1_0.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


# TODO before release:
#       print() -> logging
#       delete unused files (lo_import ?)
#       inline simple functions used once (Service.objectSupports ?)
#       put licence into all files, see http://www.boost.org/users/license.html
#       localize dialogs


""" Office Writer To Wiki. Convert Libre Office Writer documents to MediaWiki markup

See README.md for debugging tips and a full list of supported styles
"""

import sys

try:
    # if variable exists, we are running as macro from inside Office and need to add parent dir to sys.path
    # (alternatively we could put all files except `main.py` inside folder named `pythonpath`, see
    #  https://wiki.openoffice.org/wiki/Python/Transfer_from_Basic_to_Python#Importing_Modules)
    XSCRIPTCONTEXT

    from inspect import getsourcefile
    from os.path import dirname, join, abspath, pardir

    # a hack to get this file's location, because `__file__` and `sys.argv` are not defined inside macro
    thisFilePath = getsourcefile(lambda: 0)

    # relative path to parent dir like `<path to py macros>\writer2wiki-ext\writer2wiki\..`
    parentDir = join(dirname(thisFilePath), pardir)

    parentDirAbs = abspath(parentDir)
    if parentDirAbs not in sys.path:
        sys.path.append(parentDirAbs)
except NameError:
    # if XSCRIPTCONTEXT is not defined, we are running from IDE or command line and all paths should be OK
    pass


import uno
import unohelper
from com.sun.star.connection import NoConnectException

from writer2wiki.OfficeUi import OfficeUi
from writer2wiki.WikiConverter import WikiConverter
from writer2wiki.w2w_office.lo_enums import \
    CaseMap, FontSlant, TextPortionType, FontStrikeout, FontWeight, FontUnderline
from writer2wiki.w2w_office.service import Service
import writer2wiki.debug_utils as dbg


def getTargetFilename(docLocation, targetExtension):
    import os
    fileWithoutExt, ext = os.path.splitext(docLocation)
    return fileWithoutExt + '.' + targetExtension

def saveStringToFile(appContext, filename, content):
    fileAccessService = Service.create(Service.SIMPLE_FILE_ACCESS, appContext)

    # in case file already exists and contains longer string then we are writing here, the remaining of original string
    # will be left in file
    # Example: file contains 'abc' => saveStringToFile(..., 'AB') => now file contains 'ABc'
    if fileAccessService.exists(filename):
        # TODO rename original file and kill it only after save has succeeded
        # TODO handle exception when we can't delete file (e.g. because of permissions, com.sun.star.uno.Exception)
        fileAccessService.kill(filename)

    outputFile = fileAccessService.openFileWrite(filename)
    outputStream = Service.create(Service.TEXT_OUTPUT_STREAM)
    outputStream.setEncoding('UTF-8')

    outputStream.setOutputStream(outputFile)
    outputStream.writeString(content)
    outputStream.closeOutput()

def getOfficeAppContext(haveTriedToStartOffice=False):
    resolver = Service.create(Service.UNO_URL_RESOLVER)
    try:
        context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    except NoConnectException:
        if haveTriedToStartOffice:
            # we've already tried to start office, some other problem occurred
            raise

        import os
        portNumber = 2002
        os.system('soffice --writer --accept="socket,port=%d;urp;StarOffice.ServiceManager"' % portNumber)

        print("It seems no Libre Office instance is running, let's try to start one")
        context = getOfficeAppContext(True)

    return context


def convert(context, converter):
    ui = OfficeUi(context)
    desktop = Service.create(Service.DESKTOP, context)
    document = desktop.getCurrentComponent()

    if not Service.objectSupports(document, Service.TEXT_DOCUMENT):
        # TODO more specific message: either no document is opened at all or we can't convert, for example, Calc
        ui.messageBox('Please, open LibreOffice Writer document to convert')
        return

    if not document.hasLocation():
        ui.messageBox('Save you document - converted file will be saved to the same folder')
        return

    textModel = document.getText()

    # TODO write generator to iterate UNO enumerations
    textIterator = textModel.createEnumeration()
    while textIterator.hasMoreElements():
        paragraph = textIterator.nextElement()
        if Service.objectSupports(paragraph, Service.TEXT_TABLE):
            print('skip text table')
            continue
        print('para iter')
        paragraphResult = ''

        textPortionsEnum = paragraph.createEnumeration()
        while textPortionsEnum.hasMoreElements():
            portion = textPortionsEnum.nextElement()

            portionType = portion.TextPortionType
            if portionType != TextPortionType.TEXT:
                print('skip non-text portion: ' + portionType)
                continue

            text = converter.replaceNonBreakingChars(portion.getString())  # type: str
            if not text:  # blank line
                continue

            portionDecorator = converter.makeTextPortionDecorator(text)

            link = portion.HyperLinkURL
            if link:  # link should go first for proper wiki markup
                portionDecorator.addHyperLink(link)

            if not text.isspace():
                if portion.CharPosture != FontSlant.NONE:                   # italic
                    portionDecorator.addPosture(portion.CharPosture)

                if portion.CharWeight != FontWeight.NORMAL and not link:    # bold
                    # FIX CONVERT: handle non-bold links (possible in Office)
                    portionDecorator.addWeight(portion.CharWeight)

                if portion.CharCaseMap != CaseMap.NONE:
                    portionDecorator.addCaseMap(portion.CharCaseMap)

                if portion.CharColor != -1 and not link:
                    # FIX CONVERT: handle custom-colored links (possible in Office)
                    portionDecorator.addFontColor(portion.CharColor)

                # TODO check if sub/sup is compatible with styles above (it works in Office, but how is it displayed in wiki?)
                if portion.CharEscapement < 0:
                    portionDecorator.addSubScript()
                if portion.CharEscapement > 0:
                    portionDecorator.addSuperScript()

            if portion.CharStrikeout != FontStrikeout.NONE:
                portionDecorator.addStrikeout(portion.CharStrikeout)

            if portion.CharUnderline not in [FontUnderline.NONE, FontUnderline.DONTKNOW] and not link:
                # FIX CONVERT: handle links without underlines (possible in Office)
                underlineColor = portion.CharUnderlineColor if portion.CharUnderlineHasColor else None
                portionDecorator.addUnderLine(portion.CharUnderline, underlineColor)

            paragraphResult += portionDecorator.getResult()

        converter.addParagraph(paragraphResult)

    print('---')
    print('--- result: ' + converter.getResult())

    targetFilename = getTargetFilename(document.getLocation(), converter.getFileExtension())
    saveStringToFile(context, targetFilename, converter.getResult())

    ui.messageBox('Saved converted file to ' + uno.fileUrlToSystemPath(targetFilename))

def convertToWiki():
    try:
        # if variable exists, we are running as macro from inside Office
        appContext = XSCRIPTCONTEXT.getComponentContext()   # UNO type: XScriptContext
    except NameError:
        # when not running as a macro, try to connect to Office through socket
        appContext = getOfficeAppContext()

    convert(appContext, WikiConverter())


if __name__ == '__main__':
    convertToWiki()

# comma (,) at the end is significant, don't remove
g_exportedScripts = convertToWiki,
