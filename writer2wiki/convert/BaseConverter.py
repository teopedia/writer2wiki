#           Copyright Alexander Malahov 2017-2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


from typing import List
from abc import ABCMeta, abstractmethod

import uno

from writer2wiki.convert.Paragraph import Paragraph
from writer2wiki.convert.TextPortion import TextPortion
from writer2wiki.convert.WikiParagraphDecorator import WikiParagraphDecorator
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
    def makeParagraphDecorator(cls) -> WikiParagraphDecorator: pass

    @classmethod
    @abstractmethod
    def getFileExtension(cls) -> str: pass

    @abstractmethod
    def getResult(self) -> str: pass

    def __init__(self, context):
        desktop = Service.create(Service.DESKTOP, context)

        self._context = context
        self._document = desktop.getCurrentComponent()
        self._ui = OfficeUi(context)
        self._hasFootnotes = False
        self._paragraphs = []  # type: List[Paragraph]

    def addParagraph(self, p: Paragraph) -> None:
        if p.isEmpty():
            print('>> skip empty paragraph')
            return

        self._paragraphs.append(p)

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

            paragraphDecorator = self.makeParagraphDecorator()
            paragraph = Paragraph(paragraphUno, userStylesMapper)
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
                        paragraph.appendPortion(portion)

                elif portionType == TextPortionType.FOOTNOTE:
                    # TODO convert: recognize endnotes - it has same portion type
                    self._hasFootnotes = True
                    caption = portionUno.getString()

                    # TODO design: we don't need `context` here, this means the method should be in separate class -
                    #              XTextObjectConverter or something like that
                    footConverter = self.__class__(self._context)
                    footConverter._convertXTextObject(portionUno.Footnote, userStylesMapper)
                    paragraph.appendFootnote(caption, footConverter.getResult())

                else:
                    print('skip portion with not supported type: ' + portionType)
                    continue

            self.addParagraph(paragraph)
