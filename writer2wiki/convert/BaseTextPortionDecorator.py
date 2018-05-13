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

    @abstractmethod
    def _afterPropertiesApplied(self) -> None: pass

    def __init__(self):
        self._originalText = ''
        self._result = ''
        self._isLink = False

    def getDecoratedText(self, textPortion: TextPortion):
        self._originalText = textPortion.getRawText()
        self._result = self._replaceNonBreakingChars(self._originalText)
        self._isLink = 'HyperLinkURL' in textPortion.getProperties()

        # call decorator's methods to apply char properties to raw text, all method names
        # must start with 'apply', e.g. applyCharWeight(...)
        for unoPropName, propValue in textPortion.getProperties().items():
            method = getattr(self, 'apply' + unoPropName, None)
            if method is None:
                print('ERR: `{}` has no handler method for property `{}`'.format(self.__class__, unoPropName))
                continue

            method(propValue)

        self._afterPropertiesApplied()

        return self._result
