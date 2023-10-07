Notes
=====

Installation
------------
For every user, LibreOffice and a java runtime needs to be installed.

- Put jtds-1.3.1.jar either into the system's classpath (linux)
  or somehere in a user's application subdir (Windows)
- Add the java runtime to LibreOffice in Tools -> Options -> LibreOffice
  -> Advanced
- Create a LibreOffice datasource by New Database -> Select Database ->
  Connect to an existing database, JDBC, use the settings from the next
  section and create it with name bodb. Save it somwhere in the user's
  path
- Install the python scripts in
  - (Windows) C:\Users\<username>\AppData\Roaming
    \libreoffice\4\user\Scripts\python
  - (Linux) ~/.config/libreoffice/4/user/Scripts/python
- Create the Extra Toolbar and connect it with the macros.

Accessing the database
----------------------
We choose jdbc to access the database, which is a MS-SQL running on a virtual
Windows Machine at host bodb (172.16.0.129). No idea if odbc would work as
well. Install the package `libjtds-java`, supplying a JDBC driver for
microsoft and sybase databases, then create a new database (no, don't worry,
it will just be a reference to a database).

As Datasource URL use
`jdbc: jtds:sqlserver://bodb:1433/Extras;password=***REMOVED***`
As user name
`sa`
and as JDBC driver class use
`net.sourceforge.jtds.jdbc.Driver`, 

Explanation for the URL:

  bodb
    Host
  1433
    Port (1433 is standard and can probably be omitted
  Extras
    Name of database. This is the schema where ps specific views and
    tables are stored to not interfere with the normal BioOffice

Note that libreoffice needs the standard
java classpath, where libtds is installed as well, to be set explicitly to
`/usr/share/java` in the LibreOffice configuration at LibreOffice - Advanced -
Java Options - Class Path, and obviously, you need a java runtime registered
for LibreOffice. In our case, this is 1.8.0_121, provided by the debian
package `openjdk-8-jre`

The password may be omitted, in which case your are asked for it, every time
the database is (re)opened.

Commandline queries with jdbc
-----------------------------
There is a small java program that allows to use a jdbc driver via
commandline at: https://jdbcsql.sourceforge.net/

Download the file jdbcsql-1.0.zip from the project page, unpack the zip file
to a temporary directory, say ``jdbc/`` then perform the following
modifications:

- Add file jtds-1.3.1.jar to the directory
- Edit file META-INF/MANIFEST.MF and repair the damaged line starting with
  Rsrc-Class-Path: by putting everything into one line, space separated, and
  add jtds-1.3.1.jar to the line
- Add the following lines to the file DBCConfig.properties::

    # MSSQL settings
    mssql_driver = net.sourceforge.jtds.jdbc.Driver
    mssql_url = jdbc:jtds:sqlserver://host:port/dbname;password=******

  where the asterisks are replaced by the actual password. Alternatively, you
  can omit ";password" and need to put the password into your commandline
  invocation.

Now zip the directory with zip -r . ../jdbcsql.jar and then you can issue an
SQL query using::

  java -jar jdbcsql.jar -m mssql -h bodb -d Extras -U sa -Px <sql string>

Note that the last argument must be seen as a single commandline arg by the
java program, i.e. it must be put in quotes. Also note that strings in SQL
must be delimited by single quotes and table names in mssql might be put in
double quotes, so you need to surround `sql string` by double quotes, and use
backslash escaped double quotes inside to delimit column names.

Commandline query with sqlcmd
-----------------------------

You can obtain the debian package `mssql-tools` using the following
sources.list file::

  deb [arch=amd64 signed-by=/usr/share/keyrings/microsoft-products.gpg]
  https://packages.microsoft.com/ubuntu/22.04/prod jammy main

The package insists on the microsoft way of putting binaries in package
specific subdirectories, instead of following the unix convention of putting
these in /usr/bin, so either add /opt/mssql-tools/bin to your path or (what I
prefer)::

  ln -s /opt/mssql-tools/bin/sqlcmd /usr/bin

To use sqlcmd first put user name and password of our database user in the
environment::

  export SQLCMDUSER=waage
  export SQLCMDPASSWORD=***Removed***

Now call sqlcmd with::

  sqlcmd -S db.core.plattsalat.de -d Extras -Q "<sqlquery>"

Example Query::

  SELECT DISTINCT CAST(CAST(EAN AS DECIMAL(20)) AS VARCHAR(20)),
  Bezeichnung, Land, VK1, VK0, VKEinheit FROM V_Artikelinfo
  WHERE Waage = 'A' AND wg = '0001' ORDER BY 2

All on one line. Other environment variables, to allow shortening the
commandline::

  export SQLCMDSERVER=db.core.plattsalat.de
  export SQLCMDDBNAME=Extras

When running sqlcmd interactively note that sql commands are only run when you
enter the command::

  go

on a single line all by itself. Output in different formats (json, xml) is not
supported under linux. Note that sqlcmd apparently has::

  SET QUOTED_IDENTIFIER OFF

so identifiers (variables, table and column names) must comply to the
identfier rules: Start with letter or _ and not being a mssql keyword.
You can still delimit identifiers by enclosing them in square brackets.


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

Btw: To check the syntax of a python script do::

  python -m py_compile foo.py



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

Libreoffice stores its user specific configuration in a os specific
folder. For linux it is $HOME/.config/libreoffice/4/user. In the subdirectory
config/soffice.cfg/modules/scalc/toolbar, there is a file standardbar.xml
a normal xml file that is surprisingly readable. You can use the Libreoffice
GUI to add a macro call to see how an entry looks like and later edit the xml
file accordingly. Warning: You need to stop all libreoffice processes to make
it rearead the general config, even those that have no window open, like the
one started with the pywithcalc script.

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

See
http://www.openoffice.org/api/docs/common/ref/com/sun/star/module-ix.html
for API reference info

To obtain the path of the current document in the filesystem use
the property `URL`, a String containing the 'file:' URL, so if you leave out the
leading 7 characters `file://`, you get the actual path. The URL may be empty
if the Document is new and hasn't been saved yet.

The general unit for length is 1/100 mm

For an explanation on how to use various uno objects in python see
https://www.openoffice.org/de/doc/entwicklung/python_bruecke.html

or

http://www.openoffice.org/udk/python/python-bridge.html for the English
version.

There is a BASIC Programming guide as well, at
https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide
that is useful for python programming as well, because it mentions the proper
services, interface names and concepts, The documentation for Header and
Footers in spreadsheets e.g. was at
https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/Formatting_Spreadsheet_Documents

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

To get/set the active sheet of a calc, use the property
`calc.CurrentController.ActiveSheet`.

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
'
  bf = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")

Split/Merge cells by getting a cell range and call the .merge(True) method for
merging, merge(False) for splitting. The option in the dialog where content is
moved up is not available.

Number formats are of type Long, because they are specified by an index into
a table where rather complex format may be specified. All numbre formats of
a document are listed in its `NumberFormats` property. There are some standard
formats defined for a locale, `CURRENCY`, `DATE`, `TIME`. `PERCENT`, ...
To put it all together::

  from com.sun.star.lang import Locale
  loc = Locale('de','DE','')
  curr = uno.getConstantByName("com.sun.star.util.NumberFormat.CURRENCY")
  cf = doc.NumberFormats.getStandardFormat(curr, loc)

Indices for existing entries can be obtained by::

  numberformats.queryKey(numberformatstring, localformat, bool)

Types can be obtained by importing them from the right `com.sun.star`
module. E.g. a sheet has the property TitleRows, to set the rows that will be
printed on every page on the top, if the printout runs over several pages.
The property is of the type ``com.sun.star.table.CellRangeAddress``.
The constructor allows each of the components Sheet, StartColumn, StartRow,
EndColumn, EndRow to be set via keyword arguments. So the following code will
make Row 0 and 1 be repeated on top of every printed page (sheet is the
spreadsheet object)::

  from com.sun.star.table import CellRangeAddress
  sheet.setTitleRows(CellRangeAddress(StartRow=0,EndRow=1))

To set only part of a cell text in boldface, you need a cursor, move it (the
second argument is True to expand the selection) and then set the property::

  cell = sheet.getCellByPosition(1,0)
  cur  = cell.Text.createTextCuror()
  cur.goLeft(4, False)
  cur.goLeft(3. True)
  cur.setPropertyValue("CharWeight", bf)

Where bf is obtained as above.

To add or remove manual breaks use the IsStartOfNewPage property of a row.

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
  DefPage.IsLandscape = True

The Content of headers and footers can be set by the HeaderFooterContent
service. The service can be obtained by the page style (`DefPage` in the
preceding example) as {Right|Left}Page{Header|Footer}Content, and offers three
properties: `LeftText`, `CenterText` and `RightText`. Obviously, Header is for
the headers and footer for the footers. Right is for odd numbered pages and
Left for even numbered pages. If Left is not set, all pages are done the style
set by Right.

You need to obtain the service as an object, modify the Properties and then
assign it back to the service to make the changes take effect::

  hs = DefPage.RightPageHeaderContent
  hs.LeftText.String = "foobar"
  DefPage.RightPageHeaderContent = hs



Using a database
~~~~~~~~~~~~~~~~
In your general LibreOffice configuration, you need to register a database,
this is Extras / Options / LibreOffice Base / Databases / New
(The New means Register a new one, not to create a new one) You choose the .odb
file containing your connection data and register it withe name "bodb" because
that is what the python macro uses.

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

Plattsalat servera
------------------
The python program is stored on ``vserver2:/srv/samba/data/software/psmacros/``
as Psmacros.py User is nils-rennebarth To copy it, use::

  rsync -av /home/nils/src/ps/Psmacros.py \
    nils-rennebarth@vserver2:/srv/samba/data/software/psmacros


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

VK1 ist der Mitglieder-Verkaufspreis, VK0 der allgemeine Verkaufspreis

Kassenliste braucht Spalten EAN, Bezeichnung, VKEinheit, Land, VK1, VK0


Fleisch
-------
Ganze seite hochformat, Spalten: (VK1 = Mitgliederpreis)
EAN Bezeichnung VKEinheit VK1 VK0

Preise eher kleiner, EAN, Bez

LiefID= URIA FAUSER UNTERWEGER

eine Seite pro LiefID

Lose Produkte Lebensmittel
WG=0585, Unique Bezeichnung, da selbes Produkt von mehreren Lieferanten
eine Seite

Lose Prdukte Wasch
WG=0590

Eine Seite
1. Saft
WG=0400, iWG="HH"

2. 5Elemente
Lieferid

WG='0070' 0200 0280 0340

Tennental

Lieferid=Tennnental
WG 0020 0025 0060

Am besten als Menü
