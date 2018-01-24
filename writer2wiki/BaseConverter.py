#           Copyright Alexander Malahov 2017-2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


import uno
from abc import ABCMeta, abstractclassmethod, abstractmethod
from pathlib import Path

from writer2wiki.OfficeUi import OfficeUi
from writer2wiki.w2w_office.lo_enums import \
    CaseMap, FontSlant, TextPortionType, FontStrikeout, FontWeight, FontUnderline
from writer2wiki.w2w_office.service import Service
from writer2wiki.UserStylesMapper import UserStylesMapper
from writer2wiki.util import *
from writer2wiki import ui_text
import writer2wiki.debug_utils as dbg

class BaseConverter(metaclass=ABCMeta):

    @classmethod
    @abstractclassmethod
    def makeTextPortionDecorator(cls, text): pass

    @classmethod
    @abstractclassmethod
    def makeParagraphDecorator(cls, paragraphUNO, userStylesMap): pass

    @classmethod
    @abstractclassmethod
    def getFileExtension(cls):
        # type: () -> str
        pass

    @classmethod
    @abstractclassmethod
    def replaceNonBreakingChars(cls, text):
        # type: (str) -> str
        """
        Replace non-breaking space and dash with html entities for better readability of wiki-pages sources and safe
        copy-pasting to editors without proper Unicode support
        """

        # full list of non-breaking (glue) chars: http://unicode.org/reports/tr14/#GL
        pass

    @abstractmethod
    def addParagraph(self, paragraphDecorator):
        # type: () -> None
        pass

    @abstractmethod
    def getResult(self):
        # type: (None) -> str
        pass

    def convert(self, context):
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
            paragraphDecorator = self.makeParagraphDecorator(paragraph, userStylesMapper)

            textPortionsEnum = paragraph.createEnumeration()
            while textPortionsEnum.hasMoreElements():
                portion = textPortionsEnum.nextElement()

                portionType = portion.TextPortionType
                if portionType != TextPortionType.TEXT:
                    print('skip non-text portion: ' + portionType)
                    continue

                text = self.replaceNonBreakingChars(portion.getString())
                if not text:  # blank line
                    continue

                portionDecorator = self.makeTextPortionDecorator(text)

                link = portion.HyperLinkURL
                if link:  # link should go first for proper wiki markup
                    portionDecorator.addHyperLink(link)

                if not text.isspace():
                    if portion.CharPosture != FontSlant.NONE:  # italic
                        portionDecorator.addPosture(portion.CharPosture)

                    if portion.CharWeight != FontWeight.NORMAL and not link:  # bold
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

            self.addParagraph(paragraphDecorator)

        dbg.printCentered('done')
        print('result:\n' + self.getResult())

        targetFile = docPath.with_suffix(self.getFileExtension())
        with openW2wFile(targetFile, 'w') as f:
            f.write(self.getResult())

        if not userStylesMapper.saveStyles():
            ui.messageBox(ui_text.failedToSaveMappingsFile(userStylesMapper.getFilePath()))

        ui.messageBox(ui_text.conversionDone(targetFile, userStylesMapper))

