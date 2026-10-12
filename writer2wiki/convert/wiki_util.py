INLINE_STYLE_PLACEHOLDER = '$$$'

def getStyledContent(style: str, content: str) -> str:
    if not style or not content:
        return content

    contentPlaceholderIdx = style.find(INLINE_STYLE_PLACEHOLDER)
    if contentPlaceholderIdx == -1:  # Regular named style
        return '{{' + style + '|' + content + '}}'
    else:  # Inline style
        return style.replace(INLINE_STYLE_PLACEHOLDER, content)
