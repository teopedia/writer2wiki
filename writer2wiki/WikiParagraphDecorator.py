#           Copyright Alexander Malahov 2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


from writer2wiki.WikiTextPortionDecorator import WikiTextPortionDecorator


class WikiParagraphDecorator:

    def __init__(self, paragraphUNO, userStylesMapper):
        """
        :param paragraphUNO:
        :param UserStylesMapper userStylesMapper:
        """
        self._paragraphUNO = paragraphUNO
        self._userStylesMapper = userStylesMapper
        self._result = ''

    def __str__(self):
        return self.getResult()

    def addPortion(self, portion):
        """
        :param WikiTextPortionDecorator portion:
        :return void:
        """
        self._result += portion.getResult()

    def getResult(self):
        wikiStyle = self._userStylesMapper.getParagraphMappedStyle(self._paragraphUNO)
        if wikiStyle is None or not self._result:
            return self._result

        return '{{' + wikiStyle + '|' + self._result + '}}'

    def isEmpty(self):
        return len(self.getResult()) == 0
