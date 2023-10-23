# arabic-ocr-with-azure
An Azure Document Intelligence-based MacOS program, built on [a source code by Microsoft Software Engineering Manager Anatoly Ponomarev](https://techcommunity.microsoft.com/t5/azure-ai-services-blog/generate-searchable-pdfs-with-azure-form-recognizer/ba-p/3652024), using the Reportlab library to generate PDF files bearing an OCR layer from Document Intelligence's response.

The GUI was created on Tkinter, with [Tkinter-Designer](https://github.com/ParthJadhav/Tkinter-Designer) and [Ttk-Bootstrap](https://ttkbootstrap.readthedocs.io/en/latest/).

To install:

If you're using MacOS 14 (Sonoma), start by installing the latest version (8.6.13) of Tcl/Tk with Homebrew, then install a Python distribution with Pyenv, in order to link Python and Tcl/Tk:

```brew install tcl-tk```

Then:

```echo 'export PATH="/usr/local/opt/tcl-tk/bin:$PATH"' >> ~/.zshrc```

Next, restart the terminal:

```source ~/.zshrc```

Now enter the following three commands:

```
export LDFLAGS="-L/usr/local/opt/tcl-tk/lib"
export CPPFLAGS="-I/usr/local/opt/tcl-tk/include"
export PKG_CONFIG_PATH="/usr/local/opt/tcl-tk/lib/pkgconfig"
```

To configure the environment variable to be used by the Python build:

```
PYTHON_CONFIGURE_OPTS="--with-tcltk-includes='-I/usr/local/opt/tcl-tk/include'
--with-tcltk-libs='-L/usr/local/opt/tcl-tk/lib -ltcl8.6.13 -ltk8.6.13'"
```

Then install a Python distribution with Pyenv:

```pyenv install <version>```

To test, run:

```pyenv global <installed version>```

Then run:

```
import sys

try:
    import Tkinter as tk # Python 2
except ImportError:
    import tkinter as tk # Python 3

print("Tcl Version: {}".format(tk.Tcl().eval('info patchlevel')))
print("Tk Version: {}".format(tk.Tk().eval('info patchlevel')))
sys.exit()
```

The terminal should display:

```
Tcl Version: 8.6.13
Tk Version: 8.6.13
```

Clone the Github repo, launch a virtual environment, then:

```pip install -r requirements.txt```

Lastly, copy the theme to the appropriate folder:

```cp -r "user.py" <your-virtual-environment>/lib/python<version>/site-packages/ttkbootstrap/themes```
