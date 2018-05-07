#           Copyright Alexander Malahov 2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../../LICENSE.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


from typing import List

from convert.TextPortion import TextPortion
from convert.UserStylesMapper import UserStylesMapper


class Paragraph:
    """ Wrapper for Office's Paragraph UNO object.
        Merges TextPortions with identical CharProperties upon addition of new portions
    """

    # Improvement idea: it would be nice to merge text portions with nearly identical char properties.
    # Example:
    #   * we have 3 portions: {"aaa", italic}, {"bbb", bold, italic}, {"ccc", italic}
    #   * merge them to: {["aaa", {"bbb", bold}, "ccc"], italic}
    #
    # Implementation ideas:
    #   * Of course do it after all paragraph's portions are added to the list (initial compacting is
    #     done and portions with identical styles are already merged)
    #   * TextPortion._rawText should be able to hold either plain text or array
    #     of TextPortion's. See "Composite pattern" on Wikipedia.
    #   * Do merge only if portions differ by 1 property
    #   * Think if we should do merge for 2 portions or better to set a minimum bar of 3
    #   * Think what if we merge portions [{1,2,3}, {4,5}], but [{1,2},{3,4,5}] would be a better match
    #     (at a first glance, this shouldn't be a problem at least if we keep "differ by 1 property"
    #      rule, because if {3,4,5} have had same properties, they would be already merged on prev step)
    #
    # If that's implemented, we can do merge for named styles directly on Paragraph
    # instead of doing it in ParagraphDecorator (and hence make the merger format-agnostic)

    def __init__(self, paragraphUno, userStylesMapper: UserStylesMapper):
        self._uno = paragraphUno
        self._namedStyle = userStylesMapper.getMappedStyle(paragraphUno.ParaStyleName)
        self._portions = []  # type: List[TextPortion]

    def isEmpty(self):
        return len(self._portions) == 0

    def _last(self):  # FIXME apply after debug
        return self._portions[-1]

    def appendPortion(self, portion: TextPortion):
        if portion.isEmpty():
            # this shouldn't happen
            print('BUG: tried to add empty portion')
            return

        if not self.isEmpty() and self._portions[-1].hasSameProperties(portion):
            print('merge `{}` with `{}`'.format(self._portions[-1].getRawText(), portion.getRawText()))
            self._portions[-1].appendRawText(portion.getRawText())
        else:
            self._portions.append(portion)

    def appendFootnote(self, caption, content):
        # FIXME convert: this implementation is wiki-specific
        # To make it format agnostic, one solution would be to create subclass for footnote portions:
        #   1. char properties and raw text are taken from caption
        #   2. content is completely converted to text and assigned in BaseConverter (as is now already)
        #   3. this portion is appended to the `_portions` as usual portion
        #   4. if styles of caption and the last portion are the same, merge them by ??"promoting"
        #      the last portion to footnote portion??
        #
        # This will also allow to handle footnotes at the beginning of paragraph

        if self.isEmpty():
            print("NOT IMPLEMENTED: can't handle footnote at the start of paragraph")
            return

        self._portions[-1].appendRawText(content)

    def getPortions(self):
        return self._portions

    def __str__(self) -> str:
        return self.__class__.__name__ + "({})".format([str(p) for p in self._portions])

