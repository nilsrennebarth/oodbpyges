Database access and scripting for Plattsalat printouts
======================================================

Accessing the database
----------------------
We choose jdbc to access the database, which is a MS-SQL running on a virtual
Windows Machine at host bodb (172.16.0.129). No idea if odbc would work as
well. Install the package `libjtds-java`, supplying a JDBC driver for
microsoft and sybase databases, then create a new database (no, don't worry,
it will just be a reference to a database).

Use the following as the Datasource URL:
  jtds:sqlserver://<host>:<port>/<dbname>;password=<password>
where

  host
    bodb (meaning bio office database of course)
  port
    1433 (the standard port for MS-SQL)
  dbname
    Extras (the separate schema, containing our own tables and views)

As the JDBC driver class use `net.sourceforge.jtds.jdbc.Driver` Note that
libreoffice needs the standard java classpath, where libtds is installed
as well, to be set explicitly to `/usr/share/java` in the LibreOffice
configuration at LibreOffice - Advanced - Java Options - Class Path, and
obviously, you need a java runtime registered for LibreOffice. In our case,
this is 1.8.0_121, provided by the debian package `openjdk-8-jre`

The password may be omitted, in which case your are asked for it, every time
the database is (re)opened.

Python for scripting
--------------------
Install the package `libreoffice-script-provider-python` (btw., our
libreoffice version, standard Ubuntu 16.04 is 5.1.6) to allow python
macro programming.

Python scripts can be placed in the `Scripts/python` subdirectory of your
users profile (note the capital S!), in Files ending in .py The Filename
(without .py) and any directories below are show as branches of a tree. The
top level functions are shown as the leaves that can be run.

If the directory `pythonpath` exists as a sibling of your script, it will be
added to the script's search path for python modules.

Basic scripts go into the `basic` subdirectory, in Files ending in .xba, the
naming is the same, first all directories below basic, then the name of the
file (without the extension) and finally as leaves the top level functions
(are nested functions even possible in the LibreOffice basic dialect?)

Create shortcut
~~~~~~~~~~~~~~~
Calling a macro is rather cumbersome: Tools - Macros - Run Macro, then two
more clicks to open the tree, and finally select you macro and click 'Run',
so better create a shortcut early. The macro must exist for that though,
even though its content is not important. Note that no restart of a
LibreOffice application is necessary to make tha macro appear.

Now go to Tools - Customize - Toolbars as this is the binding used in the
office, it is easy reachable and won't conflict with keybindings that already
may have a predefined meaning (Although for testing that is probably ok)

Create a new Toolbar, save in the general location - i.e. not document
specific - then Add... and select your macro from Macros.

Actions can also be assigned to events using the "Events" tab in the Tools -
Customize Menu.

Interact with Libreoffice from the python shell
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
This is only for development (and maybe the coolness factor), so it is not
strictly needed for python development, but may be fairly useful. See
https://bitbucket.org/t2y/unotools for details.

The unotools package allows to talk to a running LibreOffice process over a
local socket, provided you started libreoffice as::

  libreoffice --calc \
    --accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager"

The starting point to interact with LibreOffice is the context. In a macro
that is called from within Libreoffice, the context is in the global variable
`XSCRIPTCONTEXT` In an interactive python session, it can be obtained by::

  import uno
  localContext = uno.getComponentContext()
  resolver = localContext.ServiceManager.createInstanceWithContext(
    "com.sun.star.bridge.UnoUrlResolver",
    localContext
  )
  ctx = resolver.resolve(
    "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"
  )

Next you probably want the service manager, a desktop the model.  The desktop
lets you open a new document, which is what we want when the macro is run, The
model gives you access to the currently loaded document. The service manager
is needed to instantiate various classes directly given their name.
Continuing the interactive session, you do::

  smgr = ctx.ServiceManager
  desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
  model = desktop.getCurrentComponent()

In a macro to get the same objects you do::

  ctx = XSCRIPTCONTEXT.getComponentContext()
  smgr = ctx.ServiceManager
  desktop = XSCRIPTCONTEXT.getDesktop()
  model = desktop.getCurrentComponent()

Scripts can be embedded in a document directly. Remember that any LibreOffice
document file is a zip file. Python scripts go into the Scripts/python
directory of the zip root, and must be mentioned in the manifest
`META-INF/manifest.xml`, but we won't need that for now.

Methods and properties of XSCRIPTCONTEXT

  getDocument
    get current document
  getDesktop
    get Desktop object

Relation to Basic
~~~~~~~~~~~~~~~~~
Most of the documentation that mentions basic as the programming language now
is useful as well, as the methods and objects described there are the same.

However, the CreateUnoService must be replaced with a call to::

  context.getServiceManager().createInstanceWithContext("service-string", context)

If the CreateUnoService takes arguments, use createUnstanceWithArgumentsAndContext

For more python <-> Basic information see
https://wiki.openoffice.org/wiki/Python/Transfer_from_Basic_to_Python

Using Uno
---------
Uno is the tree of classes and interfaces in OpenOffice.

See http://www.openoffice.org/api/docs/common/ref/ for API reference info

To obtain the path of the current document in the filesystem use
the property `URL`, a String containing the 'file:' URL, so if you leave out the
leading 7 characters `file://`, you get the actual path. The URL may be empty
if the Document is new and hasn't been saved yet.

THe general unit for length is 1/100 mm

Introspection
~~~~~~~~~~~~~
There is an
introspection facility MRI (written in Python btw.), at
http://extensions.services.openoffice.org/project/MRI
It must be installed as an extension (works for LibreOffice as well), and can
then be run via Tools - Add Ons - MRI

It shows properties, methods, interfaces and services, starting with the
current Document. It is supposed to show the reference IDL documentation in
the browser, but that button does not work. As the class or interface name is
shown, you can navigate directly by the browser though.

Using Spreadsheets
~~~~~~~~~~~~~~~~~~
A calc document contains one or more sheets. Create a new document by::

  calc = desktop.loadComponentFromURL(
    "private:factory/scalc", "_blank", 0, ()
  )
  sheet = calc.Sheets.getByIndex(0)

This will open a new window as well, which by default holds a single sheet, at
index 0. 

Create a new sheet in an existing Document::

  calc.Sheets.insertNewByName("<name>", position)
  sheet = calc.Sheets.getByIndex(position)

Use property `Name` to get/set the name. Remove a sheet with::

  calc.Sheets.removeByName("<name>")

Test if given name exists with::

  calc.Sheets.hasByName("<name>")

or obtain all names with::

  calc.Sheets.getElementNames()

Note however, that the order of the returned list need not correspond to the
sheet's indices.

You can obtain a cell using the getCellByPosition(x,y) method. x and y are
zero based. You can obtain a cell range (mostly equivalent to selecting some
cells using the mouse in the GUI) with getCellRangeByPosition(x0,y0,x1,y1),
where (x0,y0) is the top left and (x1,y1) the bottom right corner. There also
is getCellRangeByName("A1:C15")

A cell can be assigned a string using their `String` property, or a numeric
(or other type) value using their `Value` property.

Full rows or columns can be obtained via::

  sheet.getColumns().getByIndex(n)
  sheet.getRows().getByIndex

Columns have a Width, Rows have a Height, both setable and getable. By
assigning to the Column property `OptimalWidth`, the column can be made just
wide enough to never clip its content. This is a one-time action, the width
can later be set manually, and if the content changes it won't be adjusted
automatically.

Cells (either a single or a range) can be formatted by assigning to their
proprties:

  CharHeight
    Font height (in points)
  CharWeight
    Can be used to select bold face, see below
  CellBackColor
    Background color, as 0xrrggbb
  TopBorder
  BottomBorder
  LeftBorder
  RightBorder
    Borders. use a BorderLine2 object (note the 2), with Properties
    `Color`, `InnerLineWidth`, `OuterLineWidth`, `LineDistance`, `LineStyle`

For other properties see
`com.sun.star.style.CharacterProperties` and
`com.sun.star.style.ParagraphProperties` and for special formatting
`com.sun.star.table.CellProperties`

Bold face is a special constant, obtained in python via::

  bf = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")

Split/Merge cells by getting a cell range and call the .merge(True) method for
merging, merge(False) for splitting. The option in the dialog where content is
moved up is not available.


Page properties
~~~~~~~~~~~~~~~
Printing margins, paper size, etc. are set via
`com.sun.star.style.PageProperties`: LeftMargin, RightMargin, TopMargin,
BottomMargin (all in hundredths of a millimeter, or 10 µm::

  Doc = ThisComponent
  StyleFamilies = Doc.StyleFamilies
  PageStyles = StyleFamilies.getByName("PageStyles")
  DefPage = PageStyles.getByName("Default")
 
  DefPage.LeftMargin = 1000
  DefPage.RightMargin = 1000


Using a database
~~~~~~~~~~~~~~~~
Get a db query::

  DatabaseContext = createUnoService( "com.sun.star.sdb.DatabaseContext" )
  DataSource = DatabaseContext.GetByName("bodb")
  query = DataSource.getByName("WObst")
  cmd = query.QueryDefinition.Command

  DBConn = DBSource.GetConnection("", "")

  DBStmt = DBConn.createStatement()
  DBRes = DBStmt.executeQuery("SELECT ...")

  While DBRes.next
    DBRes.getString(1)
  Wend

The executeQuery method returns an object of type `ResultSet`,
having methods `getString`, `getInt`,... other types are Byte, Short,
Double, Boolean, Date, Time, Timestamp, To navigate, there are
methods `next()`,

The purpose of the query services available at a DataSource is to define and
edit queries. The query services by themselves do not offer methods to execute
queries.


Libre Office general notes
--------------------------
The user profile is the folder storing all user related data like extensions,
custom dictionaries, templates, etc. It is located in

  Windows
    %APPDATA%/libreoffice\4\user (where APPDATA usually is

      Windows XP
        C:\Documents and Settings\<username>\Application Data
      Vista+
        C:\Users\<username>\AppData\Roaming

  GNU/Linux
    $HOME/.config/libreoffice/4/user

  MacOS
    $HOME/Library/Application Support/LibreOffice/4/user


Other URLs
----------

- https://wiki.openoffice.org/wiki/Python_as_a_macro_language
- https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/Formatting_Spreadsheet_Documents

Snippets
========

	model = desktop.getCurrentComponent()
	asheet = model.CurrentController.ActiveSheet
	cell = asheet.getCellRangeByName("C4")
	cell.String = "Foobar"

import uno
localContext = uno.getComponentContext()
resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",localContext)
ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
smgr = ctx.ServiceManager
desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop",ctx)
model=desktop.getCurrentComponent()
asheet = model.CurrentController.ActiveSheet

Local BioOffice notes
---------------------
The most used table is V_Artikelinfo which is actually a view::

  CREATE VIEW V_ArtikelInfo AS
  SELECT Artikel.WG, Artikel.EAN, Artikel.Bezeichnung, Artikel.VKEinheit,
  Artikel.Wiegeartikel, Artikel.Land, Einkauf.LiefID, Einkauf.ArtNr,
  Einkauf.EK0, Einkauf.VKGH, Verkauf.VK1, Verkauf.VK0, Verkauf.MwSt,
  Verkauf.LadenID, Verkauf.Waage
  FROM (BOArt.dbo.Verkauf FULL JOIN BOArt.dbo.Einkauf
    ON Verkauf.EAN = Einkauf.EAN)
  LEFT JOIN BOArt.dbo.Artikel
    ON (Artikel.EAN = Verkauf.EAN) OR (Artikel.EAN = Einkauf.EAN)
  WHERE Verkauf.Sortiment = 1

