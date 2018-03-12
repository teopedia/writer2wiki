#           Copyright Alexander Malahov 2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


from writer2wiki.convert.WikiTextPortionDecorator import WikiTextPortionDecorator


class WikiParagraphDecorator:

    def __init__(self, paragraphUNO, userStylesMapper):
        """
        :param paragraphUNO:
        :param UserStylesMapper userStylesMapper:
        """
        self._paragraphUNO = paragraphUNO
        self._userStylesMapper = userStylesMapper
        self._content = ''
        self._wikiStyle = self._userStylesMapper.getParagraphMappedStyle(self._paragraphUNO)

    def __str__(self):
        return self.getStyledContent(self.getStyle(), self.getContent())

    def addPortion(self, portion):
        """
        :param WikiTextPortionDecorator portion:
        :return void:
        """
        self._content += portion.getResult()

    def addFootnote(self, caption, content: str):
        self._content += "<ref>{}</ref>".format(content.strip())

    def isEmpty(self):
        return len(self.getContent()) == 0

    def isListItem(self):
        return self._paragraphUNO.ListId != ''

    def getListLevel(self):
        # In ideal world `getListLevel` and `isNumberedList` methods should be removed from this class and defined in
        # separate class like WikiListItemParagraphDecorator, which inherits from this one.
        # But since I didn't learned yet how to init base sub-object (which holds `_paragraphUNO` etc), I think that
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
        return self._content

    @staticmethod
    def getStyledContent(style, content):
        if style is None or not content:
            return content

        return '{{' + style + '|' + content + '}}'
