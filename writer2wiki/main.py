#           Copyright Alexander Malahov 2017-2018.
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
from com.sun.star.connection import NoConnectException

from pathlib import Path

from writer2wiki.OfficeUi import OfficeUi
from writer2wiki.WikiConverter import WikiConverter
from writer2wiki.w2w_office.lo_enums import \
    CaseMap, FontSlant, TextPortionType, FontStrikeout, FontWeight, FontUnderline
from writer2wiki.w2w_office.service import Service
from writer2wiki.UserStylesMapper import UserStylesMapper
from writer2wiki.util import *
from writer2wiki import ui_text
import writer2wiki.debug_utils as dbg


def getOfficeAppContext(haveTriedToStartOffice=False):
    resolver = Service.create(Service.UNO_URL_RESOLVER)
    try:
        context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    except NoConnectException:
        if haveTriedToStartOffice:
            # we've already tried to start office, some other problem occurred
            raise

        print("It seems no Libre Office instance is running, let's try to start one")

        import os
        portNumber = 2002
        os.system('soffice --writer --accept="socket,port=%d;urp;StarOffice.ServiceManager"' % portNumber)

        context = getOfficeAppContext(True)

    return context


def convert(context, converter):
    ui = OfficeUi(context)
    desktop = Service.create(Service.DESKTOP, context)
    document = desktop.getCurrentComponent()

    if not Service.objectSupports(document, Service.TEXT_DOCUMENT):
        # TODO more specific message: either no document is opened at all or we can't convert, for example, Calc
        ui.messageBox(ui_text.noWriterDocumentOpened())
        return

    if not document.hasLocation():
        ui.messageBox(ui_text.docHasNoFile())
        return

    docPath = Path(uno.fileUrlToSystemPath(document.getLocation()))
    userStylesMapper = UserStylesMapper(docPath.parent / 'wiki-styles.txt')
    textModel = document.getText()

    # TODO write generator to iterate UNO enumerations
    textIterator = textModel.createEnumeration()
    while textIterator.hasMoreElements():
        paragraph = textIterator.nextElement()
        if Service.objectSupports(paragraph, Service.TEXT_TABLE):
            print('skip text table')
            continue
        dbg.printCentered('para iter')
        paragraphDecorator = converter.makeParagraphDecorator(paragraph, userStylesMapper)

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

            paragraphDecorator.addPortion(portionDecorator)

        converter.addParagraph(paragraphDecorator)

    dbg.printCentered('done')
    print('result:\n' + converter.getResult())

    targetFile = docPath.with_suffix(converter.getFileExtension())
    with openW2wFile(targetFile, 'w') as f:
        f.write(converter.getResult())

    if not userStylesMapper.saveStyles():
        ui.messageBox(ui_text.failedToSaveMappingsFile(userStylesMapper.getFilePath()))

    ui.messageBox(ui_text.conversionDone(targetFile, userStylesMapper))

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
