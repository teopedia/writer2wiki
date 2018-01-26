#           Copyright Alexander Malahov 2017-2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


# TODO before release:
#       print() -> logging
#       delete unused files (lo_import ?)
#       inline simple functions used once (Service.objectSupports ?)
#       put licence into all files, see http://www.boost.org/users/license.html
#       localize dialogs


""" Office Writer To Wiki. Convert Libre Office Writer documents to MediaWiki markup

See README.md for debugging tips and a full list of supported styles
"""

import sys

try:
    # if variable exists, we are running as macro from inside Office and need to add parent dir to sys.path
    # (alternatively we could put all files except `main.py` inside folder named `pythonpath`, see
    #  https://wiki.openoffice.org/wiki/Python/Transfer_from_Basic_to_Python#Importing_Modules)
    XSCRIPTCONTEXT

    from inspect import getsourcefile
    from os.path import dirname, join, abspath, pardir

    # a hack to get this file's location, because `__file__` and `sys.argv` are not defined inside macro
    thisFilePath = getsourcefile(lambda: 0)

    # relative path to parent dir like `<path to py macros>\writer2wiki-ext\writer2wiki\..`
    parentDir = join(dirname(thisFilePath), pardir)

    parentDirAbs = abspath(parentDir)
    if parentDirAbs not in sys.path:
        sys.path.append(parentDirAbs)
except NameError:
    # if XSCRIPTCONTEXT is not defined, we are running from IDE or command line and all paths should be OK
    pass

from writer2wiki.WikiConverter import WikiConverter
from writer2wiki.w2w_office.service import Service


def getOfficeAppContext(haveTriedToStartOffice=False):
    # need this for `from com.star... import ...` to work - it's re-defined with `uno._uno_import()`
    # noinspection PyUnresolvedReferences
    import uno

    from com.sun.star.connection import NoConnectException

    resolver = Service.create(Service.UNO_URL_RESOLVER)
    try:
        context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    except NoConnectException:
        if haveTriedToStartOffice:
            # we've already tried to start office, some other problem occurred
            raise

        print("It seems no Libre Office instance is running, let's try to start one")

        import os
        portNumber = 2002
        os.system('soffice --writer --accept="socket,port=%d;urp;StarOffice.ServiceManager"' % portNumber)

        context = getOfficeAppContext(True)

    return context

def convertToWiki():
    try:
        # if variable exists, we are running as macro from inside Office
        appContext = XSCRIPTCONTEXT.getComponentContext()   # UNO type: XScriptContext
    except NameError:
        # when not running as a macro, try to connect to Office through socket
        appContext = getOfficeAppContext()

    c = WikiConverter()
    c.convert(appContext)


if __name__ == '__main__':
    convertToWiki()

# comma (,) at the end is significant, don't remove
g_exportedScripts = convertToWiki,
