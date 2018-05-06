# VBA.HtmlHelp #
![Help](https://raw.githubusercontent.com/GustavBrock/VBA.HtmlHelp/master/images/EE%20HTML%20Help.png)

### Viewing a compressed help file ###
An *HTML help file* itself is of a special compressed format - with the extension .*chm* - which requires a special viewer. That viewer is often the native viewer of Windows, *Microsoft HTML Help Viewer*, an OCX control.

From VBA, the viewer is opened via one of two API calls to the control, *hhctrl.ocx*, which normally is present on all newer Windows installs. This would be very simple, if not for these traps:

* If a context ID is passed for no purpose, the application may crash
* If the help viewer is not closed before exiting the application, the application may crash
* the help viewer will not by itself move to the Contents tab when a context is opened

The documentation on the control is sparse and rudimentary, almost nil, and these traps are to be found by trial and error or to be extracted bit by bit at various fora and blog posts. However, once taken care of, the usage is straight-forward. Only one function is needed.

### Code ###
Code has been tested with both 32-bit and 64-bit *Microsoft Access 2016* and *365*.

### Documentation ###
Full documentation can be found here:

![EE Logo](https://raw.githubusercontent.com/GustavBrock/VBA.HelpFile/master/images/EE%20Logo.png)

[Control the HTML Help Viewer OCX control](https://www.experts-exchange.com/articles/32054/Control-the-HTML-Help-Viewer-OCX-control.html)

Included is a Microsoft Access example application.