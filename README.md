# Convert Libre Office Writer documents to MediaWiki markup

This is still a work in progress, but it's usable.


Currently supported styles:
* bold, underline, italic, underline, strike through (with variations like wavy underline)
* SMALL CAPS, uppercase, lowercase, capitalize
* paragraphs
* sub- and superscript (as in 2 power 10)
* hyperlinks
* text color

See `examples/wiki-sample.odt` for an example of what works.

## Using as a LibreOffice macro
1. open folder *(your office install path)/share/Scripts/python*
2. copy *writer2wiki* folder from this repository there
3. start Libre Office, open your Writer document
4. Menu Tools --> Macros --> Organize Macros --> Python (here you can get an error message about damaged JRE - just click on OK and ignore it)
5. a new window should appear, choose LibreOffice macros --> writer2wiki --> main --> convertToWiki
6. you are done


## Packaging
TODO packaging into .oxt


## Dev tips

### Getting started

Hello world: [Scripting LibreOffice with Python](https://onesheep.org/scripting-libreoffice-python/)

Setting up [PyCharm](https://www.jetbrains.com/pycharm/download):
1. create new project
2. go to menu: File | Settings... | Project: (your project) | Project interpreter
3. set `Project interpreter` to `(your Office install path)/program/python` or `...program\python\python.exe` on Windows

Running from an IDE:
1. run *main.py* file (`SHIFT` + `F10` with default PyCharm key-mapping) - Office will start if it isn't already
2. run second time to convert currently open document to wiki


### Contributing

*1.* If in doubt just make a pull request! We will handle any details in the comments, or I'll just fix everything myself
*2.* Naming convention: `ClassName`, `methodName`, `argName`, `localVariable`. Yes, it's against [PEP-8](https://www.python.org/dev/peps/pep-0008/), but it's consistent with UNO

#### 3. Classes/code organization

`convert` method in *main.py* should be format-agnostic (one other backend I have in mind is StackOverflow markdown).
When other backend will be added, we will need to create base abstract classes like `Converter`, `TextPortionDecorator`.
Also when more document features will be supported we will likely need various `ParagraphDecorator`s (for text, tables, ...)
and maybe a separate abstract factory to create these classes.

But for now let's keep things simple. Just make sure *main.py* doesn't contain wiki-specific code. If you're not sure
how to do that, see rule #1 :)


### Useful links
Wikitext docs:
* [WikiMedia markup examples](https://meta.wikimedia.org/wiki/Help:Wikitext_examples)
* [WikiMedia: HTML in wikitext](https://meta.wikimedia.org/wiki/Help:HTML_in_wikitext) (seems to be better than in Wikipedia)
* [Wikipedia short markup examples](https://en.wikipedia.org/wiki/Help:Cheatsheet)
* [Wikipedia elaborate markup examples](https://en.wikipedia.org/wiki/Help:Wiki_markup) (don't know what's the difference with Wikipedia)
* [Wikipedia: HTML in wikitext](https://en.wikipedia.org/wiki/Help:HTML_in_wikitext)

I was unable to find LibreOffice dev guide as good as OpenOffice's one and the docs are still relevant
[OpenOffice.org Developers Guide](https://wiki.openoffice.org/wiki/Documentation/DevGuide/OpenOffice.org_Developers_Guide)

There is a lot of info on macro's programming in Star Basic, here are tips on translating to Python:
[Transfer from Basic to Python](https://wiki.openoffice.org/wiki/Python/Transfer_from_Basic_to_Python) (there are many useful code samples there, including UI)

UI controls reference: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1awt.html

Services (like FileAccess) reference with links to tutorials
https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1lang_1_1ServiceManager.html


### Inspect UNO variable
See component's available services, methods, fields etc: `unohelper.inspect(myUnoObj, sys.stdout)`


### Importing enums
You must import enum *keys*, not enum type itself. Example, [com.sun.star.awt.FontSlant](https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1awt.html#a362a86d3ebca4a201d13bc3e7b94340e)
is defined in IDL as
```java
enum FontSlant {
  NONE, OBLIQUE, ITALIC, DONTKNOW,
  REVERSE_OBLIQUE, REVERSE_ITALIC
}
```

in python code you import that as follows:
```python
from com.sun.star.awt.FontSlant import ITALIC, NONE
# or
FontSlant = lo_import.enum('awt.FontSlant', 'ITALIC', 'NONE')
```

### Generic import errors
    * only `from ... import ...` is supported (see uno.py:_uno_import(...))
    * TODO give an example of erroneous import, right import and wrong usage (reproduce with com.sun.star.connection.NoConnectException)

### `DeploymentException: null process service factory`
Something is wrong with your Context. For example you've passed local context where global one is expected:
```python
def open_file(context):
    smgr = context.ServiceManager
    file_access_service = smgr.createInstanceWithContext("com.sun.star.ucb.SimpleFileAccess", context)
    output_file = file_access_service.openFileWrite("example.txt")

### this fails
file = open_file(uno.getComponentContext())

### this works

local_context = uno.getComponentContext()
resolver = local_context.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local_context)
# assuming office was stated as: `soffice --writer --accept="socket,port=2002;urp;StarOffice.ServiceManager"`
context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
file = open_file(context)
```

### AttributeError: 'NoneType' object has no attribute 'ComponentWindow'
Problem with context
