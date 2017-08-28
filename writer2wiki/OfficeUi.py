#           Copyright Alexander Malahov 2017.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE_1_0.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


from writer2wiki.w2w_office.lo_enums import MbType, MbButtons
from writer2wiki.w2w_office.service import Service


class OfficeUi:
    def __init__(self, appContext):
        self._context = appContext
        desktop = Service.create(Service.DESKTOP, self._context)
        self._window = desktop.getComponentWindow()

    def messageBox(self, message, title='Writer to Wiki Converter',
                   boxType=MbType.MESSAGEBOX, buttons=MbButtons.BUTTONS_OK):

        toolkit = Service.create(Service.TOOLKIT, self._context)
        box = toolkit.createMessageBox(self._window, boxType, buttons, title, message)
        return box.execute()
