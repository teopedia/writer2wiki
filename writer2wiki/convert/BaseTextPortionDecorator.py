#           Copyright Alexander Malahov 2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../LICENSE.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


from abc import abstractmethod, ABCMeta

from writer2wiki.convert.TextPortion import TextPortion


class BaseTextPortionDecorator(metaclass=ABCMeta):

    @classmethod
    @abstractmethod
    def getSupportedUnoProperties(cls): pass

    @classmethod
    @abstractmethod
    def _replaceNonBreakingChars(cls, text: str) -> str: pass

    # @abstractmethod
    # def afterPropertyHandlersApplied(self) -> None: pass

    @abstractmethod
    def getDecoratedText(self, textPortion: TextPortion): pass

    def __init__(self):
        self._originalText = ''
        self._result = ''

    def _initFromPortion(self, textPortion: TextPortion):
        self._originalText = self._replaceNonBreakingChars(textPortion.getRawText())
        self._result = self._originalText

        # apply decorator's methods for char properties, they all must start with 'handle', e.g. handleCharWeight(...)
        for unoPropName, propValue in textPortion.getProperties().items():
            method = getattr(self, 'handle' + unoPropName, None)
            if method is None:
                print('ERR: `{}` has no handler method for property `{}`'.format(self.__class__, unoPropName))
                continue

            method(propValue)

        return self
