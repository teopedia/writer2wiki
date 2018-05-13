#           Copyright Alexander Malahov 2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


from writer2wiki.convert.Paragraph import Paragraph
from writer2wiki.convert.wiki_util import getStyledContent
from writer2wiki.convert.WikiTextPortionDecorator import WikiTextPortionDecorator


class WikiParagraphDecorator:

    @classmethod
    def makeTextPortionDecorator(cls):
        return WikiTextPortionDecorator()

    def getDecorated(self, para: Paragraph):
        if para.isEmpty():
            print('BUG: empty paragraph')
            return ''

        portions = para.getPortions()
        portionDecorator = self.makeTextPortionDecorator()
        result = ''
        currentStyle = portions[0].getStyleName()
        sameStyleBuffer = ''

        for p in portions:
            if p.getStyleName() != currentStyle:
                result += getStyledContent(currentStyle, sameStyleBuffer)
                sameStyleBuffer = ''
                currentStyle = p.getStyleName()

            sameStyleBuffer += portionDecorator.getDecoratedText(p)

        # the last style in paragraph will not be flushed inside loop
        result += getStyledContent(currentStyle, sameStyleBuffer)

        return result
