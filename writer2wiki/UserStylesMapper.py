#           Copyright Alexander Malahov 2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)

from collections import Counter
from pathlib import Path

from writer2wiki.util import *


class UserStylesMapper:

    def __init__(self, styleMapFileName):
        self._mapFilePath = styleMapFileName  # type: Path
        self._mapFileExisted = self._mapFilePath.exists()
        self._missingStyles = set()
        self._stylesUsageCount = Counter()
        self._errors = []
        self._styleMap = {}
        self._initStyleMapFromFile()

    def _addParseError(self, userMsg, styleLine):
        msg = "{}: {}".format(userMsg, styleLine)
        self._errors.append(msg)
        print('ERR: style-map parse error. ' + msg)

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
        print('---qqq adding {} = {}'.format(docName, wikiName))

        self._styleMap[docName] = wikiName

    def _initStyleMapFromFile(self):
        if not self._mapFilePath.exists():
            print("style-map file does't exist:", self._mapFilePath)
            return

        with openW2wFile(self._mapFilePath, 'r') as f:
            for line in f.readlines():
                self._handleLine(line)

    def getParagraphMappedStyle(self, paragraphUNO):
        styleName = paragraphUNO.ParaStyleName
        if styleName == 'Standard':
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
        return self._mapFilePath

    def saveStyles(self):
        if len(self._errors) > 0:
            # TODO show errors to user
            print('style-map parse errors:', self._errors)

        if not self.hasMissingStyles():
            return True

        with openW2wFile(self._mapFilePath, 'a') as f:
            if not f.writable():
                return False

            if self.mapFileExisted():
                f.write('\n\n# below are styles automatically added by Writer2Wiki, change them as you see fit\n')
            else:
                intro = \
                    '# Mappings of Office user-defined styles to wiki templates. \n' \
                    '#  \n' \
                    '# For example, if Office document contains paragraph with text `Introduction` and style \n' \
                    '# `my-title`, and there is following line in this file: \n' \
                    '#  \n' \
                    '#     my-title = my-wiki-namespace:big-title \n' \
                    '#  \n' \
                    '# Converted wiki-text will look like: \n' \
                    '#  \n' \
                    '#     {{my-wiki-namespace:big-title|Introduction}} \n' \
                    '#  \n' \
                    '# Spaces and tabs at the beginnings and endings of lines, or around `=` symbol are ignored. \n' \
                    '# Lines beginning with `#` symbol are ignored as well (like this note), you can use that to \n' \
                    '# leave comments in the file \n' \
                    '\n'
                f.write(intro)

            for styleName in sorted(self.getMissingStyles()):
                f.write('{0} = {0}\n'.format(styleName))

        return True

    def mapFileExisted(self):
        return self._mapFileExisted
