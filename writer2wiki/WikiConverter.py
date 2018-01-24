#           Copyright Alexander Malahov 2017-2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


from writer2wiki.BaseConverter import BaseConverter
from writer2wiki.WikiTextPortionDecorator import WikiTextPortionDecorator
from writer2wiki.WikiParagraphDecorator import WikiParagraphDecorator


class WikiConverter(BaseConverter):

    def __init__(self):
        self._result = ''

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

        print('>> para add:', paragraphDecorator)
        self._result += paragraphDecorator.getResult() + '\n\n'

    def getResult(self):
        return self._result
