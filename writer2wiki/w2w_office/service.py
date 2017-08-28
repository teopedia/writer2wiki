#           Copyright Alexander Malahov 2017.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../../LICENSE_1_0.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


import uno

import functools

# Constants and function are inside class (as opposed to at file's top level) for consistency with other UNO enums
class Service:
    # paragraph/text services
    TEXT_CONTENT                 = 'com.sun.star.text.TextContent'
    TEXT_TABLE                   = 'com.sun.star.text.TextTable'
    PARAGRAPH                    = 'com.sun.star.text.Paragraph'
    CHARACTER_PROPERTIES         = 'com.sun.star.style.CharacterProperties'
    CHARACTER_PROPERTIES_ASIAN   = 'com.sun.star.style.CharacterPropertiesAsian'
    CHARACTER_PROPERTIES_COMPLEX = 'com.sun.star.style.CharacterPropertiesComplex'
    PARAGRAPH_PROPERTIES         = 'com.sun.star.style.ParagraphProperties'
    PARAGRAPH_PROPERTIES_ASIAN   = 'com.sun.star.style.ParagraphPropertiesAsian'
    PARAGRAPH_PROPERTIES_COMPLEX = 'com.sun.star.style.ParagraphPropertiesComplex'

    TEXT_OUTPUT_STREAM = 'com.sun.star.io.TextOutputStream'
    UNO_URL_RESOLVER   = 'com.sun.star.bridge.UnoUrlResolver'
    DESKTOP            = 'com.sun.star.frame.Desktop'
    SIMPLE_FILE_ACCESS = 'com.sun.star.ucb.SimpleFileAccess'
    TOOLKIT            = 'com.sun.star.awt.Toolkit'

    TEXT_DOCUMENT = 'com.sun.star.text.TextDocument'

    @staticmethod
    # @functools.lru_cache()  # may be useful for optimization, but it should be benchmarked first
    def create(serviceName, context=None):
        if context is None:
            context = uno.getComponentContext()
        manager = context.ServiceManager
        return manager.createInstanceWithContext(serviceName, context)

    @staticmethod
    def objectSupports(obj, serviceName):
        return obj.supportsService(serviceName)
