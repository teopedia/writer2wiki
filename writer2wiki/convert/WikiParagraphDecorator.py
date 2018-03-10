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

    def addFootnote(self, caption, content):
        self._content += "<ref>{}</ref>".format(content)

    def isEmpty(self):
        return len(self.getContent()) == 0

    def getStyle(self):
        return self._wikiStyle

    def getContent(self):
        return self._content

    @staticmethod
    def getStyledContent(style, content):
        if style is None or not content:
            return content

        return '{{' + style + '|' + content + '}}'
