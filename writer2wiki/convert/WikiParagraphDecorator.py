#           Copyright Alexander Malahov 2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


from typing import List

from convert.TextPortion import TextPortion
from convert.UserStylesMapper import UserStylesMapper
from convert.wiki_util import getStyledContent
from convert.WikiTextPortionDecorator import WikiTextPortionDecorator


class WikiParagraphDecorator:

    def __init__(self, paragraphUNO, userStylesMapper: UserStylesMapper):
        self._paragraphUNO = paragraphUNO
        self._userStylesMapper = userStylesMapper
        self._wikiStyle = self._userStylesMapper.getMappedStyle(self._paragraphUNO.ParaStyleName)
        self._portions = []  # type: List[TextPortion]

    def __str__(self):
        return getStyledContent(self.getStyle(), self.getContent())

    @classmethod
    def makeTextPortionDecorator(cls):
        return WikiTextPortionDecorator()

    def _noPortions(self):
        return len(self._portions) == 0

    def addPortion(self, portion: TextPortion) -> None:
        print('adding portion:', portion)
        if not self._noPortions() and self._portions[-1].hasSameProperties(portion):
            print('merge `{}` with `{}`'.format(self._portions[-1].getRawText(), portion.getRawText()))
            self._portions[-1].appendRawText(portion.getRawText())
        else:
            self._portions.append(portion)

    def addFootnote(self, caption, content: str):
        if self._noPortions():
            # FIXME convert: handle this case; caption should be TextPortion to account for styles
            print("NOT IMPLEMENTED: can't handle footnote at the start of paragraph")
            return

        self._portions[-1].appendRawText("<ref>{}</ref>".format(content.strip()))

    def isEmpty(self):
        return len(self.getContent()) == 0

    def isListItem(self):
        return self._paragraphUNO.ListId != ''

    def getListLevel(self):
        # In ideal world `getListLevel` and `isNumberedList` methods should be removed from this class and defined in
        # separate class like WikiListItemParagraphDecorator, which inherits from this one.
        # But since I hasn't learned yet how to init base sub-object (which holds `_paragraphUNO` etc), I think that
        # would be an overkill for these 2 one-liners.
        #
        # Maybe something like
        #   from copy import copy
        #   ...
        #       self = copy(arg)
        return self._paragraphUNO.NumberingLevel + 1

    def isNumberedList(self):
        # TODO convert. ListLabelString is empty for unordered list items and contains strings like "1.",  "I.", "(a)"
        #      for numbered lists. Most likely, code below is not going to work in case of numbered list with custom
        #      icons instead of text labels. But this should be OK in most cases.
        #
        #      I hadn't found out what proper solution should be, look at Paragraph::NumberingStyleName and
        #      Document::getStyleFamilies() and then getByName("NumberingStyles") or something like that

        return len(self._paragraphUNO.ListLabelString) > 0

    def getStyle(self):
        return self._wikiStyle

    def getContent(self):
        if len(self._portions) == 0:
            return ''

        portionDecorator = self.makeTextPortionDecorator()
        result = ''
        currentStyle = self._portions[0].getStyleName()
        sameStyleBuffer = ''

        for p in self._portions:
            if p.getStyleName() != currentStyle:
                result += getStyledContent(currentStyle, sameStyleBuffer)
                sameStyleBuffer = ''
                currentStyle = p.getStyleName()

            sameStyleBuffer += portionDecorator.getDecoratedText(p)

        # the last style in paragraph will not be flushed inside loop
        result += getStyledContent(currentStyle, sameStyleBuffer)

        return result
