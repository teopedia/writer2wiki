#           Copyright Alexander Malahov 2017-2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


from typing import List

from writer2wiki.convert.BaseConverter import BaseConverter
from writer2wiki.convert.WikiTextPortionDecorator import WikiTextPortionDecorator
from writer2wiki.convert.WikiParagraphDecorator import WikiParagraphDecorator


class WikiConverter(BaseConverter):

    def __init__(self):
        self._paragraphs = []  # type: List[WikiParagraphDecorator]

    @classmethod
    def makeTextPortionDecorator(cls, text):
        return WikiTextPortionDecorator(text)

    @classmethod
    def makeParagraphDecorator(cls, paragraphUNO, userStylesMap):
        """
        :param paragraphUNO:
        :param userStylesMap:
        :return WikiParagraphDecorator:
        """
        return WikiParagraphDecorator(paragraphUNO, userStylesMap)

    @classmethod
    def getFileExtension(cls):
        return '.wiki.txt'

    @classmethod
    def replaceNonBreakingChars(cls, text):
        """
        Replace non-breaking space and dash with html entities for better readability of wiki-pages sources and safe
        copy-pasting to editors without proper Unicode support

        :param str text:
        :return: modified text
        """

        # full list of non-breaking (glue) chars: http://unicode.org/reports/tr14/#GL

        return text.translate({
            0x00A0: '&nbsp;',   # non-breaking space
            0x2011: '&#x2011;'  # non-breaking dash
        })

    def addParagraph(self, paragraphDecorator):
        """
        :param WikiParagraphDecorator paragraphDecorator:
        :return void:
        """

        # TODO handle ParagraphAdjust {LEFT, RIGHT, ...}
        if paragraphDecorator.isEmpty():
            print('>> skip empty paragraph')
            return

        self._paragraphs.append(paragraphDecorator)

        print('>> para add:', paragraphDecorator)

    def getResult(self):
        if len(self._paragraphs) == 0:
            return ''

        result = ''
        currentStyle = self._paragraphs[0].getStyle()
        sameStyleBuffer = ''

        def flushBuffer():
            nonlocal sameStyleBuffer, result
            sameStyleBuffer = sameStyleBuffer[: -2]  # remove last para separator before wrapping in style
            result += WikiParagraphDecorator.getStyledContent(currentStyle, sameStyleBuffer) + '\n\n'
            sameStyleBuffer = ''

        for p in self._paragraphs:
            if p.getStyle() != currentStyle:
                flushBuffer()
                currentStyle = p.getStyle()

            if p.isListItem():
                sameStyleBuffer = sameStyleBuffer[:-1]  # remove 1 line feed
                listChar = '#' if p.isNumberedList() else '*'
                sameStyleBuffer += listChar * p.getListLevel() + ' '

            sameStyleBuffer += p.getContent() + '\n\n'

        flushBuffer()  # the last style in text will not be flushed inside loop

        return result
