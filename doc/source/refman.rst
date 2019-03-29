Reference Manual
================

The std211tobsync package provides a number of different ways to translate the
Standard 211 spreadsheet into BuildingSync-formatted XML. A script is provided
for command line use:

>>> python read211.py my_data.xlsx -o my_data.xml

Behind the scenes, there are two main functions that do most of the work

.. autofunction:: read211.read_std211_xlsx

.. autofunction:: read211.map_to_buildingsync

