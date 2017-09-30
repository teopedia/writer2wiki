#           Copyright Alexander Malahov 2017.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE_1_0.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


from writer2wiki.WikiTextPortionDecorator import WikiTextPortionDecorator


class WikiConverter:

    def __init__(self):
        self._result = ''

    @classmethod
    def makeTextPortionDecorator(cls, text):
        return WikiTextPortionDecorator(text)

    @classmethod
    def getFileExtension(cls):
        return 'txt'

    def addParagraph(self, text):
        # TODO handle ParagraphAdjust {LEFT, RIGHT, ...}
        print('>> para add:', text)
        if not text:
            return
        self._result += text + '\r\n\r\n'

    def getResult(self):
        return self._result
