#           Copyright Alexander Malahov 2017.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE_1_0.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


from writer2wiki.css_enums import CssTextDecorationStyle
from writer2wiki.w2w_office.lo_enums import *


class WikiTextPortionDecorator:

    def __init__(self, text):
        self._originalText = text
        self._cssStyles = {}
        self._result = text

    def _addCssStyle(self, name, value):
        if name in self._cssStyles:
            print('WARN: change `%s` value from %s to %s' % (name, self._cssStyles[name], value))
        self._cssStyles[name] = value

    def _surroundWithTag(self, tag, tagAttributes=''):
        if tagAttributes:
            tagAttributes = ' ' + tagAttributes
        self._result = '<{0}{1}>{2}</{0}>'.format(tag, tagAttributes, self._result)

    def _surround(self, with_string):
        self._result = with_string + self._result + with_string

    def addHyperLink(self, targetUrl):
        targetUrl = targetUrl + ' ' if targetUrl != self._originalText else ''
        self._result = '[' + targetUrl + self._originalText + ']'

    def addPosture(self, posture):
        """italic etc"""
        if posture != FontSlant.ITALIC:
            print('unexpected posture:', posture)
        self._surround("''")

    def addWeight(self, weight):
        if weight < FontWeight.NORMAL:
            print('thin weights are not supported, got:', weight)
            return
        if weight != FontWeight.BOLD:
            print('unexpected boldness:', weight)

        self._surround("'''")

    def _addTextDecorationStyle(self, decorationType, officeStyleKind, styleMap):
        if officeStyleKind not in styleMap:
            print('ignore unexpected %s kind: %s' % (decorationType, officeStyleKind))
            return

        self._addCssStyle('text-decoration', decorationType)

        mappedStyle = styleMap[officeStyleKind]
        if mappedStyle != CssTextDecorationStyle.SOLID:  # solid is default
            self._addCssStyle('text-decoration-style', mappedStyle)

    def addStrikeout(self, strikeoutKind):
        STYLES = {FontStrikeout.SINGLE: CssTextDecorationStyle.SOLID,
                  FontStrikeout.DOUBLE: CssTextDecorationStyle.DOUBLE,
                  FontStrikeout.BOLD  : CssTextDecorationStyle.SOLID,
                  FontStrikeout.SLASH : CssTextDecorationStyle.DOUBLE,
                  FontStrikeout.X     : CssTextDecorationStyle.DOUBLE
                  }
        self._addTextDecorationStyle('line-through', strikeoutKind, STYLES)

    def addUnderLine(self, underlineKind, color):
        STYLES = {FontUnderline.SINGLE:         CssTextDecorationStyle.SOLID,
                  FontUnderline.DOUBLE:         CssTextDecorationStyle.DOUBLE,
                  FontUnderline.DOTTED:         CssTextDecorationStyle.DOTTED,
                  FontUnderline.DASH:           CssTextDecorationStyle.DASHED,
                  FontUnderline.LONGDASH:       CssTextDecorationStyle.DASHED,
                  FontUnderline.DASHDOT:        CssTextDecorationStyle.DASHED,
                  FontUnderline.DASHDOTDOT:     CssTextDecorationStyle.DOTTED,
                  FontUnderline.SMALLWAVE:      CssTextDecorationStyle.WAVY,
                  FontUnderline.WAVE:           CssTextDecorationStyle.WAVY,
                  FontUnderline.DOUBLEWAVE:     CssTextDecorationStyle.WAVY,
                  FontUnderline.BOLD:           CssTextDecorationStyle.SOLID,
                  FontUnderline.BOLDDOTTED:     CssTextDecorationStyle.DOTTED,
                  FontUnderline.BOLDDASH:       CssTextDecorationStyle.DASHED,
                  FontUnderline.BOLDLONGDASH:   CssTextDecorationStyle.DASHED,
                  FontUnderline.BOLDDASHDOT:    CssTextDecorationStyle.DASHED,
                  FontUnderline.BOLDDASHDOTDOT: CssTextDecorationStyle.DASHED,
                  FontUnderline.BOLDWAVE:       CssTextDecorationStyle.WAVY
                  }
        self._addTextDecorationStyle('underline', underlineKind, STYLES)
        if color:
            self._addCssStyle('text-decoration-color', intToHtmlHex(color))

    def addCaseMap(self, caseMapKind):
        STYLES = {CaseMap.UPPERCASE: ['text-transform', 'uppercase'],
                  CaseMap.LOWERCASE: ['text-transform', 'lowercase'],
                  CaseMap.TITLE:     ['text-transform', 'capitalize'],
                  CaseMap.SMALLCAPS: ['font-variant',   'small-caps'],
                  }
        if caseMapKind not in STYLES:
            print('unexpected CaseMap: ', caseMapKind)
            return
        style = STYLES[caseMapKind]
        self._addCssStyle(style[0], style[1])

    def addFontColor(self, color):
        self._addCssStyle('color', intToHtmlHex(color))

    def addSubScript(self):
        self._surroundWithTag('sub')

    def addSuperScript(self):
        self._surroundWithTag('sup')

    def getResult(self):
        if len(self._cssStyles):
            style = ''
            for name, value in sorted(self._cssStyles.items()):
                style += name + ':' + value + ';'
            style = style[:-1]  # remove last semicolon
            self._surroundWithTag('span', 'style="%s"' % style)

        return self._result


def intToHtmlHex(val):
    return '#%X' % val
