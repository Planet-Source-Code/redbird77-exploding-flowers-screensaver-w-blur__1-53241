--------------------------------------------------------------------------
CONTENTS: APPINFO, ABOUT, UN/INSTALLATION, CREDITS, TO DO, REVISON HISTORY
--------------------------------------------------------------------------

[APPINFO]
name   : Exploding Flowers Screensaver (w/Blur)
version: 2004.04.19
author : redbird77
email  : redbird77@earthlink.net
www    : http://home.earthlink.net/~redbird77

[ABOUT]
Exploding Flowers Screensaver is a modification of Paul Bahlawan's original Planet Source code post.  My version adds a few extra things, most notably a reasonably fast (for VB, that is) blur effect.

It also includes:
- the abililty to dynamically set the screen resoultion and buffer size (for varying degrees of speed/smoothness)

- transparent flowers

[UN/INSTALLATION]
To install:
Simply compile the executable, making sure the extension is SCR not EXE, and place in your system directory (usually something like C:\WINDOWS\SYSTEM)

To uninstall:
Delete the SCR executable and the same-ly named INI file.

[CREDITS]
Paul Bahlawan - original concept and flower geometry.
Carles PV - uber nifty cDIB32 class.
Carlos J. Quintero - author of the must-have VB addin - MZ-Tools.  The "TabIndex Assistant" feature is great for configuration dialogs with zillions of controls.

[TO DO]
- add password-protection (parse the /a command line switch)

- make the configuration dialog actually appear modally against the display properties dialog (any ideas?)

- try to go all API by using CreateWindow and subclassing the window, responding to various WM_ messages like WM_TIMER (then it could be created in C, ha ha ha!)

- tweak the load/unload/show/destroy, etc. lifecycle of a screensaver

- turn it into a screensaver template

- remove any memory leaks/resource unallocation/logic errors (I'm sure they exist!)

[REVISON HISTORY]
2004.04.19
Initial release.