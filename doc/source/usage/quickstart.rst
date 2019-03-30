Quickstart
==========
The std211tobsync package provides several different methods to convert a
Standard 211 spreadsheet into a BuildingSync XML file. After installation,
the package can create the files at the command line or within a Python script.

Command Line
------------
Given a Standard 211 spreadsheet in :file:`std211.xlsx`, the command line script
can be used to generate an XML file:

>>> python --pretty std211.xlsx

This generates a file called :file:`std211.xml`. To control the name of the
output file, add the option :file:`--output compact.xml` and to forego pretty
printing, omit the :file:`--pretty`:

>>> python --output compact.xml std211.py

For further information, see :doc:`cli`.

Code Example
------------
Usage in code is easiest when the desired output is a string containing the XML
data and the input is an Excel :file:`xlsx` file. The following code prints out
pretty-printed BuildingSync XML data:

.. code-block:: python

   import read211
   ...
   string_xml = read211.map_std211_xlsx_to_prettystring('my_std211_file.xlsx')
   print(string_xml)

