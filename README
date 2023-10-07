# oodbpyges


## Purpose

The file ``Psmacros.py`` can be used as a LibreOffice (and probably OpenOffice)
macro file, provided python is enabled as a scripting language. The macros it
provides fetch lists of articles from a MS-SQL database and create a nicely
formatted spreadsheet that is ready to print.

The actual file will probably be of interest for only a handful of people (see
the next section for details), but this seemingly small project got quickly
out of bounds and I learned a **lot** about the inner workings of LibreOffice,
and how everything can be tied together and documented that to some
detail. This might help other people wanting to achieve similar things.
This documentation can be found in :doc:`notes`.

## Background

The project started in a small community run organic food store in south-west
Germany. Near to an electronic balance hangs a paper sheet listing the food
and vegetable items on sale together with price and the three digit code you
need to type in for each. To fit on a single sheet of paper, the list needs a
two column layout. The sheet was usually created in the morning, by opening
their inventory management system BioOFFICE, cut-and-paste the data into a
spreadsheet and then rearraning the columns manually until the result fit onto
a single sheet of paper.

The database backend for BioOFFICE is a standard MS-SQL database and a custom
readonly view had already been created that prevent interference with the
inventory system. LibreOffice allows to access databases using an odbc driver,
and can be scripted in python, so how hard can that be?

And as usual in IT, the answer is: Much harder than you think initially. To
e.g. automate such a seemingly simple task as using number format for a
column, you need to select the locale, know the magic constant representing
the standard numberformat and how right call to actually use this constant.
