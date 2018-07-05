#           Copyright Alexander Malahov 2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../../LICENSE.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


from collections import OrderedDict
from typing import List

from writer2wiki.convert.ConversionSettings import ConversionSettings


class TextPortion:
    """
    Wrapper class for Office's TextPortion UNO. Used to merge adjacent portions with identical styles
    """

    def __init__(self, portionUno,
                 styleFamilies,
                 conversionSettings: ConversionSettings,
                 supportedStyles: List[str]):
        self._rawText = portionUno.getString()
        self._namedStyle = conversionSettings.getMappedStyle(portionUno.CharStyleName)
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
            # stylePropValue = getattr(style, unoPropName)
            stylePropValue = style.getPropertyValue(unoPropName)
            # portionPropValue = getattr(portionUno, unoPropName)
            portionPropValue = portionUno.getPropertyValue(unoPropName)

            # print('style `{:<18}`, prop {:<14} | portionVal: {}, styleVal: {} | equals: {}'.
            #       format(styleName, unoPropName, portionPropValue, stylePropValue, stylePropValue==portionPropValue))

            return stylePropValue == portionPropValue

        # It would be nice to handle 'Default Style' uniformly with portion and paragraph styles,
        # but I've failed to figure out how to get its' XStyle object in locale-agnostic way.
        #
        # Name (string) 'Default Style' is localized string - e.g. in Russian locale it's 'Базовый'.
        # Docs [1] say XStyle object has localized name in `DisplayName` field and locale-agnostic
        # name in `Name` field. If that was true, we could iterate all styles in CharacterStyles
        # family and find default one upon script startup. However, default style has localized
        # string in both these fields (at least when we iterate family). Probably a bug in Office.
        #
        # [1] https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1CharacterStyle.html

        inDefaultStyle = portionUno.getPropertyDefault(unoPropName) == portionUno.getPropertyValue(unoPropName)
        inPortionStyle = propertyIsInStyle('CharacterStyles', portionUno.CharStyleName)
        inParaStyle    = propertyIsInStyle('ParagraphStyles', portionUno.ParaStyleName)

        # print("'{:<5}' prop: {:<18}, def: {:<1}, port: {:<1}, para: {:<1}"
        #       .format(portionUno.getString(), unoPropName, inDefaultStyle, inPortionStyle, inParaStyle))

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
