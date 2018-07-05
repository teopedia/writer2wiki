#           Copyright Alexander Malahov 2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


import configparser
from collections import Counter
from pathlib import Path

from writer2wiki.util import openW2wFile


class ConversionSettings:
    """ Conversion settings, read from file `writer2wiki-folder-settings.txt` in document's folder.
        If this file doesn't exist and the document contains custom styles, it will be created.
    """

    _KEY_STYLES_SECTION = 'styles'
    _KEY_OPTIONS_SECTION = 'options'
    _OPTION_IGNORE_FONT_COLOR = 'ignore font color'

    def __init__(self, documentFilePath: Path):
        self._docPath = documentFilePath
        self._settingsFilePath = self._docPath.parent / 'writer2wiki-folder-settings.txt'

        # Save because we need to check this after file is created
        self._settingsFileExisted = self._settingsFilePath.exists()

        self._missingStyles = set()
        self._stylesUsageCount = Counter()

        # TODO delete
        self._errors = []

        # args to disallow defaults we don't need
        self._settings = configparser.ConfigParser(comment_prefixes=('#',), delimiters=('=',),
                                                   empty_lines_in_values=False)
        """ Don't use it directly to read/write individual settings, prefer aliases `_styleMap`
            and `_options` (for brevity and to avoid typos in map keys)
        """

        # disable conversion of keys to lower case
        self._settings.optionxform = lambda option: option

        if self.settingsFileExisted():
            self._settings.read(str(self._settingsFilePath), encoding='utf-8')
        else:
            self._settings[self._KEY_STYLES_SECTION] = {}
            self._settings[self._KEY_OPTIONS_SECTION] = {}
        self._styleMap = self._settings[self._KEY_STYLES_SECTION]  # type: configparser.SectionProxy
        self._options = self._settings[self._KEY_OPTIONS_SECTION]  # type: configparser.SectionProxy

        self._legacyMapFile = self._docPath.parent / 'wiki-styles.txt'

        if self.hadOnlyLegacyMapFile():
            self._initStyleMapFromFile()

    def ignoreFontColor(self) -> bool:
        return self._options.get(self._OPTION_IGNORE_FONT_COLOR, 'no').strip().lower() == 'yes'

    # TODO delete
    def hadOnlyLegacyMapFile(self):
        return not self.settingsFileExisted() and self._legacyMapFile.exists()

    def _addParseError(self, userMsg, styleLine):
        msg = "{}: {}".format(userMsg, styleLine)
        self._errors.append(msg)
        print('ERR: style-map parse error. ' + msg)

    # TODO: deprecated, delete
    def _handleLine(self, line):
        """
        :param str line: line of style-map file
        :return void:
        """
        line = line.strip()         # skip empty line
        if not line:
            return
        if line.startswith('#'):    # skip comments
            return

        parts = [x.strip() for x in line.split('=')]
        if len(parts) > 2:
            self._addParseError("this style-map line contains more than one symbol '='", line)
            return
        if len(parts) == 1:
            self._addParseError("this style-map line contains no symbol '='", line)
            return

        docName, wikiName = parts
        if docName in self._styleMap:
            self._addParseError('style `{}` is already in the style-map, will overwrite it'.format(docName), line)

        self._styleMap[docName] = wikiName

    # TODO: deprecated, delete
    def _initStyleMapFromFile(self):
        if not self._legacyMapFile.exists():
            print("style-map file does't exist:", self._legacyMapFile)
            return

        with openW2wFile(self._legacyMapFile, 'r') as f:
            for line in f.readlines():
                self._handleLine(line)

    def getMappedStyle(self, styleName):
        if styleName in ['Standard', '']:
            return None

        self._stylesUsageCount[styleName] += 1

        if styleName not in self._styleMap:
            self._missingStyles.add(styleName)
            self._styleMap[styleName] = styleName

        return self._styleMap[styleName]

    def mostCommonMissingStyles(self, num):
        missingStylesDict = {name: count for name, count in self._stylesUsageCount.items()
                             if name in self._missingStyles}
        return Counter(missingStylesDict).most_common(num)

    def getMissingStyles(self):
        return self._missingStyles

    def hasMissingStyles(self):
        return len(self._missingStyles) != 0

    def getFilePath(self):
        return self._settingsFilePath

    def _newFileIntroText(self):
        from textwrap import dedent
        return dedent("""\
            # Writer2Wiki conversion settings file.
            #
            # The settings are automatically applied when converting any Writer file in the same
            # folder as this settings file. 
            #
            # Words surrounded by square brackets (e.g. [{options}]) are so called "sections", used
            # to group settings together.
            # IMPORTANT: don't remove or change order of sections
            #
            # Spaces and tabs at the beginnings and endings of lines, or around `=` symbol are ignored
            # and has no effect on conversion process.
            # Lines beginning with `#` symbol are ignored as well (like this note), you can use that to
            # leave comments in the file

            
            #{section_sep}
            # This section contains various document-wide options
            [{options}] 
            
            # Apply or not font color property to resulting wiki file  
            # Values: yes/no
            # Default: no
            {opt_ignore_font_color} = no

            
            #{section_sep}
            # This section sets mappings of Office user-defined (custom) styles to wiki templates.
            # 
            # For example, if Office document contains paragraph with text `Introduction` and style
            # `my-title`, and there is following line in this file:
            #     my-title = my-wiki-namespace:big-title
            # 
            # Converted wiki-text will look like:
            #     {{my-wiki-namespace:big-title|Introduction}}
            # 
            # To ignore style, leave right-hand side of an assignment blank:
            #      my-ignored-style =
            [{styles}]  
            
            """.format(options=self._KEY_OPTIONS_SECTION, styles=self._KEY_STYLES_SECTION,
                       opt_ignore_font_color=self._OPTION_IGNORE_FONT_COLOR, section_sep='-' * 79))

    def saveStyles(self):

        # TODO delete
        if len(self._errors) > 0:
            # TODO show errors to user
            print('style-map parse errors:', self._errors)

        if not self.hasMissingStyles() and not self.hadOnlyLegacyMapFile():
            return True

        # We write file manually instead of configparser.write() because configparser can't
        # retain existing comments. We could overwrite file and print intro on every save,
        # but comments on newly appended styles (see code below) would be lost.
        # If in future we would need more flexibility (e.g. change [options] programmatically),
        # probably the easiest solution would be to use ConfigObj library.
        with openW2wFile(self._settingsFilePath, 'a') as f:
            if not f.writable():
                return False

            if self.settingsFileExisted():
                from datetime import datetime
                f.write("\n\n# below are styles automatically added by Writer2Wiki from '{}' "
                        "on {:%Y-%m-%d %H:%M:%S}, change them as you see fit"
                        .format(self._docPath.name, datetime.today()))
            else:
                f.write(self._newFileIntroText())

                # TODO delete
                if self.hadOnlyLegacyMapFile():
                    for styleName in sorted(self._styleMap.keys() - self._missingStyles):
                        f.write('{} = {}\n'.format(styleName, self._styleMap[styleName]))

            for styleName in sorted(self.getMissingStyles()):
                f.write('{0} = {0}\n'.format(styleName))

        return True

    def settingsFileExisted(self):
        return self._settingsFileExisted
