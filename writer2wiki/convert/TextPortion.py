#           Copyright Alexander Malahov 2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../../LICENSE.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


from collections import OrderedDict
from typing import List

from writer2wiki.convert.UserStylesMapper import UserStylesMapper


class TextPortion:
    """
    Wrapper class for Office's TextPortion UNO. Used to merge adjacent portions with identical styles
    """

    def __init__(self, portionUno,
                 styleFamilies,
                 userStylesMapper: UserStylesMapper,
                 supportedStyles: List[str]):
        self._rawText = portionUno.getString()
        self._namedStyle = userStylesMapper.getMappedStyle(portionUno.CharStyleName)
        self._nonDefaultProperties = OrderedDict()

        if portionUno.HyperLinkURL:
            self._nonDefaultProperties['HyperLinkURL'] = portionUno.HyperLinkURL

        for unoPropName in supportedStyles:
            propValue = getattr(portionUno, unoPropName, None)
            if propValue is None:
                print('ERR: portion UNO has no property `{}`'.format(unoPropName))
                continue
            if __class__._propertyIsInStyleOrIsDefault(portionUno, unoPropName, styleFamilies):
                continue

            self._nonDefaultProperties[unoPropName] = propValue

    def __str__(self) -> str:
        return __class__.__name__ + "(text: {}, style: {}, properties: {})".format(
            self._rawText, self._namedStyle, self._nonDefaultProperties)

    @staticmethod
    def _propertyIsInStyleOrIsDefault(portionUno, unoPropName, styleFamilies):
        # styles docs: https://wiki.openoffice.org/wiki/Documentation/DevGuide/Text/Overall_Document_Features

        def propertyIsInStyle(styleFamilyName, styleName):
            nonlocal portionUno, unoPropName, styleFamilies

            if styleName == '':  # no para or char style for property
                # print('style `{:<18}` is not set'.format(styleFamilyName))
                return False

            familyStyles = styleFamilies.getByName(styleFamilyName)

            if not familyStyles.hasByName(styleName):
                print("ERR. Style family '{}' has no style '{}'".format(styleFamilyName, styleName))
                return False

            style = familyStyles.getByName(styleName)
            stylePropValue = getattr(style, unoPropName)
            portionPropValue = getattr(portionUno, unoPropName)

            # print('style `{:<18}`, prop {:<14} | portionVal: {}, styleVal: {} | equals: {}'.
            #       format(styleName, unoPropName, portionPropValue, stylePropValue, stylePropValue==portionPropValue))

            return stylePropValue == portionPropValue

        # FIXME: 'Default Style' has only localized name for every locale. Docs [1] say only DisplayName should be
        #        localized and that's true for all built-in styles (e.g. headings) except default one.
        #        Probably a bug. For now we will check localized names for Russian and English locales.
        #        [1] https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1CharacterStyle.html
        inDefaultStyle = propertyIsInStyle('CharacterStyles', 'Базовый') \
                      or propertyIsInStyle('CharacterStyles', 'Default Style')
        inPortionStyle = propertyIsInStyle('CharacterStyles', portionUno.CharStyleName)
        inParaStyle    = propertyIsInStyle('ParagraphStyles', portionUno.ParaStyleName)

        if not inDefaultStyle and inPortionStyle:
            return True

        if portionUno.ParaStyleName != '':
            return inParaStyle

        return inDefaultStyle

    def isEmpty(self) -> bool:
        return not bool(self._rawText)

    def appendRawText(self, text):
        self._rawText += text

    def getRawText(self):
        return self._rawText

    def getStyleName(self):
        return self._namedStyle

    def hasSameProperties(self, other):  # type: (TextPortion) -> bool
        return    self._namedStyle           == other._namedStyle           \
              and self._nonDefaultProperties == other._nonDefaultProperties

    def getProperties(self):
        return self._nonDefaultProperties
