#           Copyright Alexander Malahov 2017-2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


from typing import List

from writer2wiki.convert.BaseConverter import BaseConverter
from writer2wiki.convert.WikiTextPortionDecorator import WikiTextPortionDecorator
from writer2wiki.convert.WikiParagraphDecorator import WikiParagraphDecorator


class WikiConverter(BaseConverter):

    def __init__(self, context):
        super(WikiConverter, self).__init__(context)

    @classmethod
    def makeParagraphDecorator(cls):
        return WikiParagraphDecorator()

    @classmethod
    def getFileExtension(cls):
        return '.wiki.txt'

    def getResult(self):
        # TODO handle ParagraphAdjust {LEFT, RIGHT, ...}

        if len(self._paragraphs) == 0:
            return ''

        paraDecorator = self.makeParagraphDecorator()
        result = ''
        currentStyle = self._paragraphs[0].getStyleName()
        sameStyleBuffer = ''

        def flushBuffer():
            from writer2wiki.convert.wiki_util import getStyledContent
            nonlocal sameStyleBuffer, result

            sameStyleBuffer = sameStyleBuffer[: -2]  # remove last para separator before wrapping in style
            result += getStyledContent(currentStyle, sameStyleBuffer) + '\n\n'
            sameStyleBuffer = ''

        for para in self._paragraphs:
            if para.getStyleName() != currentStyle:
                flushBuffer()
                currentStyle = para.getStyleName()

            if para.isListItem():
                sameStyleBuffer = sameStyleBuffer[:-1]  # remove 1 line feed
                listChar = '#' if para.isNumberedList() else '*'
                sameStyleBuffer += listChar * para.getListLevel() + ' '

            sameStyleBuffer += paraDecorator.getDecorated(para) + '\n\n'

        flushBuffer()  # the last style in text will not be flushed inside loop

        if self._hasFootnotes:
            result += '<references/>\n'

        return result
