def getStyledContent(style, content):
    if not style or not content:
        return content

    return '{{' + style + '|' + content + '}}'
