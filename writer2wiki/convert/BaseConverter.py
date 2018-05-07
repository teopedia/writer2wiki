#           Copyright Alexander Malahov 2017-2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


import uno
from abc import ABCMeta, abstractmethod

from convert.Paragraph import Paragraph
from convert.TextPortion import TextPortion
from convert.WikiParagraphDecorator import WikiParagraphDecorator
from writer2wiki.OfficeUi import OfficeUi
from writer2wiki.w2w_office.lo_enums import TextPortionType
from writer2wiki.w2w_office.service import Service
from writer2wiki.convert.UserStylesMapper import UserStylesMapper
from writer2wiki.util import *
from writer2wiki import ui_text
import writer2wiki.debug_utils as dbg

class BaseConverter(metaclass=ABCMeta):

    @classmethod
    @abstractmethod
    def makeParagraphDecorator(cls, paragraphUNO, userStylesMap) -> WikiParagraphDecorator: pass

    @classmethod
    @abstractmethod
    def getFileExtension(cls) -> str: pass

    @classmethod
    @abstractmethod
    def replaceNonBreakingChars(cls, text: str) -> str:
        """
        Replace non-breaking space and dash with html entities for better readability of wiki-pages sources and safe
        copy-pasting to editors without proper Unicode support
        """

        # full list of non-breaking (glue) chars: http://unicode.org/reports/tr14/#GL
        pass

    @abstractmethod
    def addParagraph(self, paragraphDecorator: WikiParagraphDecorator) -> None:
        pass

    @abstractmethod
    def getResult(self) -> str:
        pass

    def __init__(self, context):
        desktop = Service.create(Service.DESKTOP, context)

        self._context = context
        self._document = desktop.getCurrentComponent()
        self._ui = OfficeUi(context)
        self._hasFootnotes = False

    def checkCanConvert(self) -> bool:
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
        print('result:\n', self.getResult())

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

        for index, paragraphUno in enumerate(iterUnoCollection(textUno)):
            if (index + 1) % 5 == 0:
                print('iter #', index + 1, 'out of', self._document.ParagraphCount)
            if Service.objectSupports(paragraphUno, Service.TEXT_TABLE):
                print('skip text table')
                continue

            paragraphDecorator = self.makeParagraphDecorator(paragraphUno, userStylesMapper)
            portionsList = Paragraph(paragraphUno, userStylesMapper)
            supportedProperties = paragraphDecorator.makeTextPortionDecorator().getSupportedUnoProperties()

            for portionUno in iterUnoCollection(paragraphUno):
                portionType = portionUno.TextPortionType
                if portionType == TextPortionType.TEXT:

                    portion = TextPortion(portionUno,
                                          self._document.getStyleFamilies(),
                                          userStylesMapper,
                                          supportedProperties
                                          )
                    if not portion.isEmpty():
                        paragraphDecorator.addPortion(portion)
                        # portionsList.appendPortion(portion)

                elif portionType == TextPortionType.FOOTNOTE:
                    # TODO convert: recognize endnotes - it has same portion type
                    self._hasFootnotes = True
                    caption = portionUno.getString()

                    # FIXME design: we don't need `context` here, this means the method should be in separate class -
                    #               XTextObjectConverter or something like that
                    footConverter = self.__class__(self._context)
                    footConverter._convertXTextObject(portionUno.Footnote, userStylesMapper)
                    paragraphDecorator.addFootnote(caption, footConverter.getResult())
                    # portionsList.appendFootnote(caption, footConverter.getResult())

                else:
                    print('skip portion with not supported type: ' + portionType)
                    continue

            self.addParagraph(paragraphDecorator)
