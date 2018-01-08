#           Copyright Alexander Malahov 2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE_1_0.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


from writer2wiki.WikiTextPortionDecorator import WikiTextPortionDecorator

class WikiParagraphDecorator:

    _missingCustomStyles = {}

    def __init__(self, paragraphUNO):
        self._paragraphUNO = paragraphUNO
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
        return self._result

    def isEmpty(self):
        return len(self.getResult()) == 0
