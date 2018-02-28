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


import unohelper
from com.sun.star.task import XJobExecutor

import logging as log
import os.path

# '~' will be expanded to: 'C:\Users\my-user-name\' on Windows, '/home/my-user-name/' on Linux
LOG_FILE_NAME = os.path.normpath(os.path.expanduser('~/writer2wiki.log'))
# LOG_FILE_NAME = 'C:/Alexander/Programming/Python/w2w.log'
log.basicConfig(
    filename=LOG_FILE_NAME,
    format='%(asctime)s [%(levelname)s] %(message)s',
    level=log.DEBUG)
log.debug(' logging has started '.center(80, '-'))


LOG_FIX_PATH = True

def logFixPath(msg):
    if LOG_FIX_PATH:
        log.debug(msg)

def fixPythonImportPath():
    """
    Add main.py's folder to Python's import paths.

    We need this because by default Macros and UNO components can only import files located in `pythonpath` folder,
    which must be in extension's root folder. This requires extra configuration in IDE and project structure becomes
    somewhat ugly
    """

    logFixPath('entered fix path')
    import sys
    from inspect import getsourcefile
    from os.path import dirname, join, abspath, pardir

    logFixPath('sys path: %s' % sys.path)

    # a hack to get this file's location, because `__file__` and `sys.argv` are not defined inside macro
    thisFilePath = getsourcefile(lambda: 0)

    logFixPath('this file name: %s' % thisFilePath)

    # relative path to parent dir like `<path to py macros or extension>\writer2wiki-ext\writer2wiki\..`
    parentDir = join(dirname(thisFilePath), pardir)

    parentDirAbs = abspath(parentDir)
    if parentDirAbs not in sys.path:
        logFixPath('appending dir: %s' % parentDirAbs)
        sys.path.append(parentDirAbs)
    else:
        logFixPath('NOT appending %s' % parentDirAbs)
    logFixPath('leaving fix path')


fixPythonImportPath()


def getOfficeAppContext(haveTriedToStartOffice=False):
    from writer2wiki.w2w_office.service import Service
    # need this import for `from com.star... import ...` to work - it's re-defined with `uno._uno_import()`
    # noinspection PyUnresolvedReferences
    import uno

    # noinspection PyUnresolvedReferences
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

def convertToWiki(appContext=None):
    log.info(' Conversion started '.center(80, '-'))
    try:
        from writer2wiki.WikiConverter import WikiConverter

        if appContext is None:  # this must be the case only when we run as a macro or from command line / IDE
            try:
                # this variable is implicitly defined for macros
                appContext = XSCRIPTCONTEXT.getComponentContext()  # UNO type: XScriptContext
            except NameError:
                # when not running as a macro, try to connect to Office through socket
                appContext = getOfficeAppContext()

        c = WikiConverter()
        c.convert(appContext)
        log.info(' Conversion done OK '.center(80, '-'))
    except Exception:
        log.critical("Unexpected exception", exc_info=True)
        log.critical(' Conversion failed '.center(80, '-'))
    log.shutdown()


class Writer2WikiComp(unohelper.Base, XJobExecutor):
    # IMPORTANT. This must be the same string as description.xml::<identifier value>
    EXTENSION_ID = 'com.github.teopedia.writer2wiki'

    def __init__(self, context, *args):
        log.debug('component init')
        self._context = context

    # method from XJobExecutor
    def trigger(self, argsString):
        log.debug("`trigger` start with args: '%s'", str(argsString))
        convertToWiki(self._context)


# For use as UNO component in extension
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    Writer2WikiComp, Writer2WikiComp.EXTENSION_ID, (), )


# For use from IDE or command line.
# Use Python interpreter from LibreOffice's distribution.
# On Windows it's in <Program Files>\LibreOffice 5\program\python.exe
if __name__ == '__main__':
    convertToWiki()


# For use as a macro.
# Comma (,) at the end is significant, don't remove
g_exportedScripts = convertToWiki,
