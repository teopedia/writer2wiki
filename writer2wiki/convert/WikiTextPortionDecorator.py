#           Copyright Alexander Malahov 2017-2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


from convert.BaseTextPortionDecorator import BaseTextPortionDecorator
from convert.TextPortion import TextPortion
from convert.css_enums import CssTextDecorationStyle
from w2w_office.lo_enums import *
from util import *


class WikiTextPortionDecorator(BaseTextPortionDecorator):

    def __init__(self):
        super().__init__()
        self._cssStyles = {}

    def __str__(self):
        return self.getDecoratedText()

    @classmethod
    def getSupportedUnoProperties(cls):
        return [
            'CharPosture',      # italic
            'CharWeight',       # bold
            'CharCaseMap',      # small-caps, capitalize etc
            'CharColor',        # font color
            'CharEscapement',   # subscript / superscript
            'CharStrikeout',
            'CharUnderline',
            'CharUnderlineColor'
        ]

    @classmethod
    def _replaceNonBreakingChars(cls, text: str) -> str:
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

    def _addCssStyle(self, name, value, appendIfExist=False):
        if name in self._cssStyles:
            if appendIfExist:
                self._cssStyles[name] = self._cssStyles[name] + ' ' + value
            # elif self._cssStyles[name] != value:
            else:
                print('WARN: change `{}` value from {} to {}'.format(name, self._cssStyles[name], value))
        self._cssStyles[name] = value

    def _surround(self, with_string):
        self._result = with_string + self._result + with_string

    def applyHyperLinkURL(self, targetUrl):
        targetUrl = targetUrl + ' ' if targetUrl != self._originalText else ''
        self._result = '[' + targetUrl + self._result + ']'

    def applyCharPosture(self, posture):
        """italic etc"""
        if posture != FontSlant.ITALIC:
            print('unexpected posture:', posture)
        self._surround("''")

    def applyCharWeight(self, weight):
        if weight < FontWeight.NORMAL:
            print('thin weights are not supported, got:', weight)
            return
        if weight != FontWeight.BOLD:
            print('unexpected boldness:', weight)

        self._surround("'''")

    def _addTextDecorationStyle(self, decorationType, officeStyleKind, styleMap):
        if officeStyleKind not in styleMap:
            print('ignore unexpected {} kind: {}'.format(decorationType, officeStyleKind))
            return

        mappedStyle = styleMap[officeStyleKind]
        if mappedStyle is None:
            # We add properties only if they differ from portion or paragraph style.
            # If mapped style is None (no underline or strike-through), it means we override style.
            # For example, style has underline and we want portion without underline.
            # LIMITATION: 'display: inline-block' will remove all text-decoration's (underline, strike-through)
            #             inherited from style, but this case should be pretty rare, so we will not complicate
            #             code to handle that
            self._addCssStyle('display', 'inline-block')
            return

        self._addCssStyle('text-decoration', decorationType, True)
        if mappedStyle != CssTextDecorationStyle.SOLID:  # solid is default
            self._addCssStyle('text-decoration-style', mappedStyle)

    def applyCharStrikeout(self, strikeoutKind):
        STYLES = {FontStrikeout.NONE:   None,
                  FontStrikeout.SINGLE: CssTextDecorationStyle.SOLID,
                  FontStrikeout.DOUBLE: CssTextDecorationStyle.DOUBLE,
                  FontStrikeout.BOLD  : CssTextDecorationStyle.SOLID,
                  FontStrikeout.SLASH : CssTextDecorationStyle.DOUBLE,
                  FontStrikeout.X     : CssTextDecorationStyle.DOUBLE
                  }
        self._addTextDecorationStyle('line-through', strikeoutKind, STYLES)

    def applyCharUnderline(self, underlineKind):
        STYLES = {FontUnderline.NONE:           None,
                  FontUnderline.SINGLE:         CssTextDecorationStyle.SOLID,
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

    def applyCharUnderlineColor(self, color):
        if color == -1:
            print('WARN: tried to handle underline color == -1')
            return
        self._addCssStyle('text-decoration-color', intToHtmlHex(color))

    def applyCharCaseMap(self, caseMapKind):
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

    def applyCharColor(self, color):
        self._addCssStyle('color', intToHtmlHex(color))

    def applyCharEscapement(self, escapement):
        if escapement == 0:
            print('BUG: invoked handleCharEscapement with 0 escapement')
            return

        if escapement > 0:
            self._result = surroundWithTag(self._result, 'sup')
        else:
            self._result = surroundWithTag(self._result, 'sub')

    def _afterPropertiesApplied(self):
        if len(self._cssStyles) == 0:
            return

        style = ''

        # TODO PY: replace with something like `str.join()`
        for name, value in sorted(self._cssStyles.items()):
            style += name + ':' + value + ';'
        style = style[:-1]  # remove last semicolon
        self._result = surroundWithTag(self._result, 'span', 'style="%s"' % style)

        # workaround for wikitext limitation: <span> tag not rendered inside wiki {{templates}}
        # (we get templates from paragraph's named styles)
        self._result = '{{#tag:span|' + self._result + '}}'
