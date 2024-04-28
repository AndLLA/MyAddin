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
 
 The functions provided by the addin are:
 - myPython: returns a string with the indication of the python version
 - myDebug: turn on/off the console debugging infos
 - myInterp: linear interpolation with flat extrapolation
 - myTextSplit: split a string, given a separator
 - myFilter: returns only elements with a matching true condition
 - myUnique: return an unique list of the inputs (unsorted)
 - mySort: sort the inputs
 - myXLookup: a vlookup with individual parameters for lookup and return array
 - myWebCsv: downloads (and caches in memory) a csv file from a URL
 - myWebCsvList: list the keys (URLs) for the cached contents
 - myWebCsvClear: cleas the myWebCsv memory cache
 
 
 Tested on Fedora 39, 40.
 
 
