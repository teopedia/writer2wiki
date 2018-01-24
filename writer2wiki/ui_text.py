#           Copyright Alexander Malahov 2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)

# TODO add localization

from writer2wiki.UserStylesMapper import UserStylesMapper


def noWriterDocumentOpened():
    return 'Please, open LibreOffice Writer document to convert'

def docHasNoFile():
    return 'Save you document - converted file will be saved to the same folder'

def _missingStylesDescription(userStylesMapper):
    """
    Get string with text description of missing styles for UI dialog

    :param UserStylesMapper userStylesMapper:
    :return:
    """
    result = ''
    printedNamesCount = 2
    mostCommon = userStylesMapper.mostCommonMissingStyles(printedNamesCount)

    # TODO PY: replace with `join()`
    for name, count in mostCommon:
        result += "'{}', ".format(name)
    result = result[:-2]  # remove last ', '

    moreElementsCount = len(userStylesMapper.getMissingStyles()) - printedNamesCount
    if moreElementsCount > 0:
        result += ' and {} more styles'.format(moreElementsCount)

    return result

def conversionDone(convertedFilename, mapper):
    """
    :param str|Path convertedFilename:
    :param UserStylesMapper mapper:
    :return:
    """

    msg = 'Saved converted file to {}'.format(convertedFilename)

    if not mapper.hasMissingStyles():
        return msg

    msg += '.\n\n'
    stylesDesc = _missingStylesDescription(mapper)
    if mapper.mapFileExisted():
        msg += 'Some user styles ({}) not found in style-map file, {}. We will map them to themselfs and append '\
               'them at the end of style-map file.' \
               .format(stylesDesc, mapper.getFilePath())
    else:
        msg += 'Document contains user-defined styles ({}), but no style-map file found (it should be {}), ' \
               'we will generate it and map all styles to themselfs. ' \
               .format(stylesDesc, mapper.getFilePath())

    msg += '\n\nOpen style-map file in any editor to correct mappings (instructions on how to do that are in the ' \
           'beginning of the file).' \

    return msg

def failedToSaveMappingsFile(filePath):
    return "Can't write styles-mapping file, {}. Check that it's writable (for example, you are not saving it to "\
           "some system directory).".format(filePath)