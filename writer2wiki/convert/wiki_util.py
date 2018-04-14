def getStyledContent(style, content):
    if style is None or not content:
        return content

    return '{{' + style + '|' + content + '}}'
