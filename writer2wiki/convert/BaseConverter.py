#           Copyright Alexander Malahov 2017-2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


import uno
from abc import ABCMeta, abstractclassmethod, abstractmethod

from writer2wiki.OfficeUi import OfficeUi
from writer2wiki.w2w_office.lo_enums import \
    CaseMap, FontSlant, TextPortionType, FontStrikeout, FontWeight, FontUnderline
from writer2wiki.w2w_office.service import Service
from writer2wiki.convert.UserStylesMapper import UserStylesMapper
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
        # type: () -> str
        pass

    def __init__(self, context):
        desktop = Service.create(Service.DESKTOP, context)

        self._context = context
        self._document = desktop.getCurrentComponent()
        self._ui = OfficeUi(context)
        self._hasFootnotes = False

    def checkCanConvert(self):
        if not Service.objectSupports(self._document, Service.TEXT_DOCUMENT):
            # TODO more specific message: either no document is opened at all or we can't convert, for example, Calc
            self._ui.messageBox(ui_text.noWriterDocumentOpened())
            return False

        if not self._document.hasLocation():
            self._ui.messageBox(ui_text.docHasNoFile())
            return False

        return True

    def convertCurrentDocument(self):
        docPath = Path(uno.fileUrlToSystemPath(self._document.getLocation()))
        userStylesMapper = UserStylesMapper(docPath.parent / 'wiki-styles.txt')
        textModel = self._document.getText()

        self._convertXTextObject(textModel, userStylesMapper)

        dbg.printCentered('done')
        print('result:\n' + self.getResult())

        targetFile = docPath.with_suffix(self.getFileExtension())
        if targetFile.exists():
            from writer2wiki.w2w_office.lo_enums import MbType, MbButtons, MbResult
            answer = self._ui.messageBox(
                ui_text.conversionDoneAndTargetFileExists(targetFile, userStylesMapper),
                boxType=MbType.QUERYBOX,
                buttons=MbButtons.BUTTONS_OK_CANCEL)
            if answer == MbResult.CANCEL:
                return
        else:
            self._ui.messageBox(ui_text.conversionDoneAndTargetFileDoesNotExist(targetFile, userStylesMapper))

        with openW2wFile(targetFile, 'w') as f:
            f.write(self.getResult())

        if not userStylesMapper.saveStyles():
            self._ui.messageBox(ui_text.failedToSaveMappingsFile(userStylesMapper.getFilePath()))

    def _convertXTextObject(self, textUno, userStylesMapper):
        from writer2wiki.util import iterUnoCollection
        for paragraph in iterUnoCollection(textUno):
            if Service.objectSupports(paragraph, Service.TEXT_TABLE):
                print('skip text table')
                continue

            dbg.printCentered('para iter')
            paragraphDecorator = self.makeParagraphDecorator(paragraph, userStylesMapper)

            for portion in iterUnoCollection(paragraph):
                portionType = portion.TextPortionType
                if portionType == TextPortionType.TEXT:
                    portionDecorator = self._buildTextPortionTypeText(portion)
                    if portionDecorator is not None:
                        paragraphDecorator.addPortion(portionDecorator)

                elif portionType == TextPortionType.FOOTNOTE:
                    # TODO convert: recognize endnotes - it has same portion type
                    self._hasFootnotes = True
                    caption = portion.getString()

                    # FIXME design: we don't need `context` here, this means the method should be in separate class -
                    #               XTextObjectConverter or something like that
                    footConverter = self.__class__(self._context)
                    footConverter._convertXTextObject(portion.Footnote, userStylesMapper)
                    paragraphDecorator.addFootnote(caption, footConverter.getResult())

                else:
                    print('skip portion with not supported type: ' + portionType)
                    continue

            self.addParagraph(paragraphDecorator)

    def _buildTextPortionTypeText(self, portionUno):
        text = self.replaceNonBreakingChars(portionUno.getString())
        if not text:  # blank line
            return None

        portionDecorator = self.makeTextPortionDecorator(text)

        link = portionUno.HyperLinkURL
        if link:  # link should go first for proper wiki markup
            portionDecorator.addHyperLink(link)

        if not text.isspace():
            if portionUno.CharPosture != FontSlant.NONE:  # italic
                portionDecorator.addPosture(portionUno.CharPosture)

            if portionUno.CharWeight != FontWeight.NORMAL and not link:  # bold
                # FIX CONVERT: handle non-bold links (possible in Office)
                portionDecorator.addWeight(portionUno.CharWeight)

            if portionUno.CharCaseMap != CaseMap.NONE:
                portionDecorator.addCaseMap(portionUno.CharCaseMap)

            if portionUno.CharColor != -1 and not link:
                # FIX CONVERT: handle custom-colored links (possible in Office)
                portionDecorator.addFontColor(portionUno.CharColor)

            if portionUno.CharEscapement < 0:
                portionDecorator.addSubScript()
            if portionUno.CharEscapement > 0:
                portionDecorator.addSuperScript()

        if portionUno.CharStrikeout != FontStrikeout.NONE:
            portionDecorator.addStrikeout(portionUno.CharStrikeout)

        if portionUno.CharUnderline not in [FontUnderline.NONE, FontUnderline.DONTKNOW] and not link:
            # FIX CONVERT: handle links without underlines (possible in Office)
            underlineColor = portionUno.CharUnderlineColor if portionUno.CharUnderlineHasColor else None
            portionDecorator.addUnderLine(portionUno.CharUnderline, underlineColor)

        return portionDecorator
