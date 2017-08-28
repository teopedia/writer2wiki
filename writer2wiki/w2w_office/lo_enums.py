#           Copyright Alexander Malahov 2017.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../../LICENSE_1_0.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


from writer2wiki.w2w_office import lo_import

"""https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1awt.html#a362a86d3ebca4a201d13bc3e7b94340e"""
FontSlant = lo_import.enum('awt.FontSlant',
                           'NONE', 'OBLIQUE', 'ITALIC', 'DONTKNOW', 'REVERSE_OBLIQUE', 'REVERSE_ITALIC')

class TextPortionType:
    """https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextPortion.html#a7ecd2de53df4ec8d3fffa94c2e80d651"""
    TEXT                = 'Text'
    TEXT_FIELD          = 'TextField'
    TEXT_CONTENT        = 'TextContent'
    FOOTNOTE            = 'Footnote'
    CONTROL_CHARACTER   = 'ControlCharacter'
    REFERENCE_MARK      = 'ReferenceMark'
    DOCUMENT_INDEX_MARK = 'DocumentIndexMark'
    BOOKMARK            = 'Bookmark'
    REDLINE             = 'Redline'
    RUBY                = 'Ruby'
    FRAME               = 'Frame'
    SOFT_PAGE_BREAK     = 'SoftPageBreak'
    IN_CONTENT_METADATA = 'InContentMetadata'

class FontWeight:
    DONTKNOW   = 0.0
    THIN       = 50.0
    ULTRALIGHT = 60.0
    LIGHT      = 75.0
    SEMILIGHT  = 90.0
    NORMAL     = 100.0
    SEMIBOLD   = 110.0
    BOLD       = 150.0
    ULTRABOLD  = 175.0
    BLACK      = 200.0

class CaseMap:
    """https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1style_1_1CaseMap.html"""
    NONE      = 0
    UPPERCASE = 1
    LOWERCASE = 2
    TITLE     = 3
    SMALLCAPS = 4

class FontStrikeout:
    """https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1awt_1_1FontStrikeout.html"""
    NONE     = 0
    SINGLE   = 1
    DOUBLE   = 2
    DONTKNOW = 3
    BOLD     = 4
    SLASH    = 5
    X        = 6

class FontUnderline:
    """https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1awt_1_1FontUnderline.html"""
    NONE            = 0
    SINGLE          = 1
    DOUBLE          = 2
    DOTTED          = 3
    DONTKNOW        = 4
    DASH            = 5
    LONGDASH        = 6
    DASHDOT         = 7
    DASHDOTDOT      = 8
    SMALLWAVE       = 9
    WAVE            = 10
    DOUBLEWAVE      = 11
    BOLD            = 12
    BOLDDOTTED      = 13
    BOLDDASH        = 14
    BOLDLONGDASH    = 15
    BOLDDASHDOT     = 16
    BOLDDASHDOTDOT  = 17
    BOLDWAVE        = 18


"""https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1awt.html#ad249d76933bdf54c35f4eaf51a5b7965"""
MbType = lo_import.enum('awt.MessageBoxType',
                    'MESSAGEBOX', 'INFOBOX', 'WARNINGBOX', 'ERRORBOX', 'QUERYBOX')

class MbButtons:
    """https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1awt_1_1MessageBoxButtons.html"""
    BUTTONS_OK                 = 1
    BUTTONS_OK_CANCEL          = 2
    BUTTONS_YES_NO             = 3
    BUTTONS_YES_NO_CANCEL      = 4
    BUTTONS_RETRY_CANCEL       = 5
    BUTTONS_ABORT_IGNORE_RETRY = 6
    DEFAULT_BUTTON_OK          = 0x10000
    DEFAULT_BUTTON_CANCEL      = 0x20000
    DEFAULT_BUTTON_RETRY       = 0x30000
    DEFAULT_BUTTON_YES         = 0x40000
    DEFAULT_BUTTON_NO          = 0x50000
    DEFAULT_BUTTON_IGNORE      = 0x60000
