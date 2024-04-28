# MyAddin
LibreOffice Calc Addin, written in python.

This addin implements several utility functions 
for Libre Office Calc.

It's written in python and relies on the following packages
being available either at system or user level:
 - numpy
 - pandas
Before installing the addin, make sure that those packages are installed.

In addition, you need to install the libreoffice-sdk package,
in order to have the following Uno utilities:
 - unoidl-write
 - unoidl-read

After making sure the above dependencies are satisfied,
build the addin as follows:
 - clone the repo 
 - from the terminal invoke build.bash
 The addin will be packaged and installed for your current user
 
 Alternatively you can install the pre-built MyAddin.oxt and test it with MyAddin.ods.
 
 
 
