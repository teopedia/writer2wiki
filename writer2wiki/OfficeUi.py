#           Copyright Alexander Malahov 2017-2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


from writer2wiki.w2w_office.lo_enums import MbType, MbButtons
from writer2wiki.w2w_office.service import Service


class OfficeUi:
    def __init__(self, appContext):
        self._context = appContext
        desktop = Service.create(Service.DESKTOP, self._context)
        self._window = desktop.getComponentWindow()
        self._writeToLog = True  # TODO after we add proper logging, enable dialogs logging only for DEBUG log level

    def _log(self, *messages):
        if self._writeToLog:
            print(*messages)

    def messageBox(self, message, title='Writer to Wiki Converter',
                   boxType=MbType.MESSAGEBOX, buttons=MbButtons.BUTTONS_OK):
        toolkit = Service.create(Service.TOOLKIT, self._context)
        box = toolkit.createMessageBox(self._window, boxType, buttons, title, message)

        self._log('user dialog:', message)
        dialogResult = box.execute()
        self._log('user response:', dialogResult)

        return dialogResult
