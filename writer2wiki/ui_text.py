#           Copyright Alexander Malahov 2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


# TODO add localization
from pathlib import Path

from writer2wiki.convert.ConversionSettings import ConversionSettings


def noWriterDocumentOpened():
    return 'Please, open LibreOffice Writer document to convert'

def docHasNoFile():
    return 'Save you document - converted file will be saved to the same folder'

def _missingStylesDescription(conversionSettings: ConversionSettings):
    """
    Get string with text description of missing styles for UI dialog
    """
    result = ''
    printedNamesCount = 2
    mostCommon = conversionSettings.mostCommonMissingStyles(printedNamesCount)

    # TODO PY: replace with `join()`
    for name, count in mostCommon:
        result += "'{}', ".format(name)
    result = result[:-2]  # remove last ', '

    moreElementsCount = len(conversionSettings.getMissingStyles()) - printedNamesCount
    if moreElementsCount > 0:
        result += ' and {} more styles'.format(moreElementsCount)

    return result

def conversionDoneAndTargetFileDoesNotExist(convertedFilename: Path,
                                            conversionSettings: ConversionSettings):
    msg = 'Saved converted file to {}'.format(convertedFilename)

    if not conversionSettings.hasMissingStyles():
        return msg

    msg += '.\n\n'
    stylesDesc = _missingStylesDescription(conversionSettings)
    if conversionSettings.settingsFileExisted():
        msg += 'Some user styles ({}) not found in style-map file, {}. We will map them to themselfs ' \
               'and append them at the end of style-map file.' \
               .format(stylesDesc, conversionSettings.getFilePath())
    else:
        msg += 'Document contains user-defined styles ({}), but no style-map file found (it should ' \
               'be {}), we will generate it and map all styles to themselfs. ' \
               .format(stylesDesc, conversionSettings.getFilePath())

    msg += '\n\nOpen style-map file in any editor to correct mappings (instructions on how to do ' \
           'that are in the beginning of the file).'

    return msg

def conversionDoneAndTargetFileExists(targetFile: Path, mapper: ConversionSettings):
    msg = "Target file '{}' already exists. Overwrite it?".format(targetFile)

    # if file exists, assume user is familiar with the extension and knows what style-map file is
    if mapper.hasMissingStyles():
        msg += '\n\n(added {} to style-map file, {})'.\
            format(_missingStylesDescription(mapper), mapper.getFilePath())

    return msg

def failedToSaveMappingsFile(filePath):
    return "Can't write conversion settings file, {}. Check that it's writable (for example, " \
           "you are not saving it to some system directory).".format(filePath)