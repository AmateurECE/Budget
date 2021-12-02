# LibreOffice Macros for Budget Calculations

This repo is set up to be cloned as a Homeshick castle.

# Running the Budgetizer in Development

First, start up an Office server instance:

```
$ soffice --calc \
    --accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager"
```

Next, run the script, which will load the package and execute the budgetizer:

```
$ python3 DevelopmentRunner.py
```

# Silently Reloading the Document

The Basic macro to do this can be found in `SilentlyReload.macro`. To install:

1. Go to _Tools_->_Macros_->_Organize Macros_->_Basic_.
2. In the window that opens, expand _My Macros_, select _Standard_,
   and click _New_.
3. In the window that opens, paste the contents of
   `SilentlyReload.macro`.
4. Press <kbd>Ctrl</kbd> + <kbd>S</kbd> to save the macro and close
   the window.

# To Assign a Macro to a Shortcut Key

1. Go to _Tools_->_Customize_
2. In the window that opens, select the shortcut you wish to use for
   the macro, for example, <kbd>F3</kbd>, in the _Shortcut Keys_
   section.
3. In the _Category_ section select _LibreOffice Macros_->_My Macros_,
   expand _Standard_, select _Module 1_.
4. In the _Function_ section select _SilentlyReload_.
5. Click _Modify_ and then _OK_.
