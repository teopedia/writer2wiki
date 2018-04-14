class PseudoPortionUno:
    """
    Class to emulate UNO for text portion to reduce special cases for footnotes when building paragraph's result
    """

    def __init__(self, content):
        self._content = content

    def getStyle(self):
        return None

    def getContent(self):
        return self._content