Reference Manual
================
The std211tobsync package provides a number of different ways to translate the
Standard 211 spreadsheet into BuildingSync-formatted XML. A script is provided
for command line use:

>>> python read211.py my_data.xlsx -o my_data.xml

For further information, see :doc:`usage/cli`.

Convenience Functions
---------------------
There are several functions in the package that take a spreadsheet file name
and generate Python strings of the BuildingSync XML:

.. autofunction:: read211.map_std211_xlsx_to_prettystring

.. autofunction:: read211.map_std211_xlsx_to_string

Other Translation Functions
---------------------------
Behind the scenes, there are two main functions that do most of the work. These
functions are used by the command line script and the convenience functions to
translate the spreadsheet:

.. autofunction:: read211.read_std211_xlsx

.. autofunction:: read211.map_to_buildingsync

loadxl Module
-------------
The Standard 211 spreadsheet uses a quite a few controls (primarily checkboxes),
and unfortunately (as of version 2.4.9) openpyxl does not support access to 
these controls. To bridge that gap, a small module was developed to dig through
the spreadsheet files (which are just a collection of zipped XML files) and get
the information that we need and then attaches that information to the openpyxl
data structures. The openpyxl workbook object is extended to include two
dictionaries:

* `controls` contains a dictionary (by the control name) of control objects
* `textboxes` contains a dictionary (by the textbox name) of the string contents

The control structure contains only data and defines no methods:

.. autoclass:: loadxl.Control
   :members:

The main function that does the loading is:

.. autofunction:: loadxl.load_workbook


