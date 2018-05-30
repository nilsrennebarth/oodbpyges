# Plattsalat specific python macros
import collections
import uno
import types
from com.sun.star.lang import Locale
from com.sun.star.table.CellVertJustify import CENTER as vertCenter
from com.sun.star.table.CellHoriJustify import CENTER as horCenter
from com.sun.star.table.CellHoriJustify import RIGHT as horRight
from com.sun.star.table.CellHoriJustify import LEFT as horLeft
from com.sun.star.table import CellRangeAddress

class BioOfficeConn:
	"""Connection to our Bio-Office database"""

	def __init__(self):
		# Obtain connection to our database.
		# Needs the registered data source "bodb"
		ctx = XSCRIPTCONTEXT.getComponentContext()
		self.dbconn = ctx.ServiceManager.createInstanceWithContext(
			"com.sun.star.sdb.DatabaseContext", ctx
		).getByName("bodb").getConnection('', '')

	def queryResult(self, sql, types):
		"""Get the results of an SQL query as a list

		sql is the query as a string, types is a string specifying
		the types in each row. I is for Int, S for String, D for Double.
		"""
		meths = []
		result = []
		dbres = self.dbconn.createStatement().executeQuery(sql)

		# create a list of methods from the type string
		for c in types:
			if c == 'I':
				meths.append(getattr(dbres, 'getLong'))
			elif c == 'S':
				meths.append(getattr(dbres, 'getString'))
			elif c == 'D':
				meths.append(getattr(dbres, 'getDouble'))

		while dbres.next():
			result.append([meths[i](i+1) for i in range(len(meths))])
		return result

Pos = collections.namedtuple('Pos', 'x y')

class ColumnDef(types.SimpleNamespace):
	"""Options for a single column in a table

	This is mostly a container for various options. The following
	options are currently recognized:
	- width (int) width in mm
	- height (int) char height
	- tryOptWidth (boolean) First try to set the width to its optimum
	  value. Only if that is too big, set it to the given width
	- bold (boolean) set typeface to bold
	- greyUnit (boolean) set background to grey if the text appears
	  to represent discrete units

	"""

	colDefaults = dict(
		bold=False,
		greyUnit=False,
		tryOptWidth=False,
		width=10,
		height=12,
		hcenter=False,
		hright=False,
		hleft=False
	)

	def __init__(self, **opts):
		self.__dict__.update(ColumnDef.colDefaults)
		super().__init__(**opts)


class Sheet:
	"""A single sheet to be filled with tables"""

	def __init__(self, name, cols, titlerows=0):
		desktop = XSCRIPTCONTEXT.getDesktop()
		# Create a new calc and use its first sheet
		self.calc = desktop.loadComponentFromURL(
			"private:factory/scalc", "_blank", 0, ()
		)
		self.sheet = self.calc.Sheets.getByIndex(0)
		self.sheet.Name = name
		self.cols = cols
		self.titlerows = titlerows
		self.currencyformat = self.calc.NumberFormats.getStandardFormat(
			uno.getConstantByName("com.sun.star.util.NumberFormat.CURRENCY"),
			Locale('de','DE','')
		)
		self.Linestyle = uno.createUnoStruct("com.sun.star.table.BorderLine2")
		self.Linestyle.OuterLineWidth = 5
		self.ColDefs = []
		self.Boldface = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")
		# Get the default cell style
		# and use it to set use a 12pt Font Size by default
		cs = self.calc.StyleFamilies.CellStyles.getByName('Default')
		cs.CharHeight=12


	def addColumns(self, cols):
		self.ColDefs += cols

	def getCell(self, x, y):
		return self.sheet.getCellByPosition(x, y)

	def getCol(self, col):
		return self.sheet.getColumns().getByIndex(col)

	def getRow(self, row):
		return self.sheet.getRows().getByIndex(row)

	def styleBlock(self, x, y, n):
		"""Style a row, Blocks with lines everywhere.
		"""
		cells = self.sheet.getCellRangeByPosition(x,y,x+n-1,y)
		cells.LeftBorder   = self.Linestyle
		cells.RightBorder  = self.Linestyle
		cells.TopBorder    = self.Linestyle
		cells.BottomBorder = self.Linestyle
		cells.ParaRightMargin = 100
		cells.ParaLeftMargin  = 100

	def styleAltGrey(self, x, y, n):
		"""Style a row, Alternating grey background
		"""
		self.getCell(x,y).LeftBorder = self.Linestyle
		self.getCell(x+n-1,y).RightBorder = self.Linestyle
		if (y & 1) == 1:
			cells = self.sheet.getCellRangeByPosition(x,y,x+n-1,y)
			cells.CellBackColor = 0xdddddd

	def addData(self, *lists, style = 'Block'):
		mysheet=self

		class Cellpos:
			def __init__(self, cols, rows):
				self.x = 0
				self.y = mysheet.titlerows
				self.cols = cols
				self.rows = rows

			def advance(self):
				# go one down
				self.y += 1
				# if at bottom row, go to top and left
				if self.y == self.rows + mysheet.titlerows:
					self.x = self.x + self.cols + 1
					self.y = mysheet.titlerows

		# N is sum of list members
		N = 0
		for list in lists: N += len(list)
		# colCols is the number of columns in each list. All lists
		# are supposed to have the same number of columns
		self.colCols = len(lists[0][0])
		self.HeaderPositions = []
		self.totalCols = self.cols * (self.colCols + 1) - 1

		c = self.getCell(10,1)
		# Each list starts with a Label, using a single row
		# then one row for each member and another row to separate
		# the list from the next one. The total numer of rows
		# is TR = <number of lists> * 2 - 1 + <sum of list lengths>
		needed = len(lists) * 2 - 1 + N
		# We want to divide these equally over all columns,
		# so we round up to the next multiple of cols and
		# get the actual number of sheet rows
		self.totalRows = (needed + self.cols-1) // self.cols
		rest = self.totalRows * self.cols - needed

		pos = Cellpos(self.colCols, self.totalRows)
		styler = getattr(self, 'style'+style)
		for list in lists:
			self.HeaderPositions.append(Pos(pos.x, pos.y))
			# advance once, to get room for the label
			pos.advance()
			for row in list:
				for i, val in enumerate(row):
					cell = self.getCell(pos.x + i, pos.y)
					if isinstance(val, str):
						cell.String = val
					else:
						cell.Value  = val
					if isinstance(val, float):
						cell.NumberFormat = self.currencyformat
				styler(pos.x, pos.y, self.colCols)
				pos.advance()
			# advance once at the end of a list
			pos.advance()
			if rest > 0:
				pos.advance()
				rest -= 1

	def getOptimalScale(self):
		"""Calculate the optimal scale factor in percent
		"""
		w=0
		for i in range(self.totalCols):
			w += self.getCol(i).Width
		h=0
		for i in range(self.totalRows):
			h += self.getRow(i).Height
		if h==0 or w==0: return 100 # should not happen
		ws = 19500 / w # factor to scale to 195mm width
		hs = 28200 / h # factor to scale to 270mm height
		# We must use the smaller of the two for scaling.
		# If hs is smaller, the resulting height is at the maximum,
		# and we only might make the Columns a bit wider, but we don't
		if hs < ws: return int(hs * 100)
		# If ws is smaller, the resulting width is at the maximum.
		# In that case we can still make each row a bit higher to increase
		# readability
		hstretch = 28200 / (h * ws)
		if hstretch > 1.5: hstretch = 1.5
		for i in range(self.totalRows):
			self.getRow(i).Height = self.getRow(i).Height * hstretch
		return int(ws * 100)

	def getOptimalScaleExt(self, landscape, pages):
		nrows = (self.totalRows + pages-1) // pages
		w = 0
		for i in range(self.totalCols):
			w += self.getCol(i).Width
		h=0
		for i in range(nrows):
			h += self.getRow(i+self.titlerows).Height
		for i in range(self.titlerows):
			h += self.getRow(i).Height
		if h==0 or w==0: return 100 # should not happen
		if landscape:
			towidth = 28400
			toheight = 19800
		else:
			towidth = 19800
			toheight = 28400
		ws = towidth / w
		hs = toheight / h
		if hs < ws: return int(hs * 100)
		hstretch = toheight / (h * ws)
		if hstretch > 1.8: hstretch = 1.8
		for i in range(self.titlerows, self.totalRows):
			self.getRow(i).Height = self.getRow(i).Height * hstretch
		return int(ws * 100)

	def addGrey(self, col):
		for i in range(self.totalRows):
			cell = self.getCell(col, i)
			if len(cell.String) == 2 and cell.String != 'Kg':
				# cell.CellBackColor = 0xdddddd
				cell.CharWeight = self.Boldface

	def formatCol(self, i, cdef):
		col = self.getCol(i)
		if cdef.tryOptWidth:
			col.OptimalWidth = True
			if col.Width > cdef.width * 100:
				col.Width = cdef.width * 100
		else:
			col.Width = cdef.width * 100
		if cdef.bold:
			col.CharWeight = self.Boldface
		if cdef.greyUnit:
			self.addGrey(i)
		if cdef.height != 12:
			col.CharHeight = cdef.height
		if cdef.hright:
			col.HoriJustify = horRight
		if cdef.hleft:
			col.HoriJustify = horLeft
		col.VertJustify = vertCenter

	def formatColumns(self):
		for t in range(self.cols):
			for i,cdef in enumerate(self.ColDefs):
				self.formatCol(t * (self.colCols + 1) + i, cdef)
			if t < self.cols-1:
				self.getCol((t+1) * (self.colCols + 1) - 1).Width = 800

	def setListLabels(self, *labels, cheight=14):
		for i,l in enumerate(labels):
			p = self.HeaderPositions[i]
			cell = self.getCell(p.x + 1, p.y)
			cell.String = l
			cell.CharHeight = cheight
			cell.CharWeight = self.Boldface

	def setPageStyle(self, landscape=False, pages=1):
		defp = self.calc.StyleFamilies.PageStyles.getByName("Default")
		defp.LeftMargin   = 500
		defp.TopMargin    = 500
		defp.BottomMargin = 500
		defp.RightMargin  = 500
		defp.HeaderIsOn=False
		defp.FooterIsOn=False
		defp.CenterHorizontally=True
		defp.CenterVertically=False
		if landscape or pages > 1:
			if landscape:
				defp.Width = 29700
				defp.Height = 21000
				defp.IsLandscape = True
			defp.PageScale = self.getOptimalScaleExt(landscape, pages)
		else:
			defp.PageScale = self.getOptimalScale()

	def setHeaderRow(self, titles):
		self.sheet.setTitleRows(CellRangeAddress(StartRow=0, EndRow=0))
		for i in range(self.cols):
			for title in titles:
				pos = title[0]
				cdef = title[2]
				cell = self.getCell(i * (self.colCols + 1) + pos, 0)
				cell.String  = title[1]
				if cdef.bold:
					cell.CharWeight = self.Boldface
				if cdef.height != 12:
					cell.CharHeight = cdef.height
				if cdef.hcenter:
					cell.HoriJustify = horCenter

def Waagenliste(*args):
	"""Lists for the electronic balances
	
	Create a ready to print spreadsheet for the
	electronic balances, containing the EAN numbers,
	the names and the unit
	"""

	db = BioOfficeConn()

	sql = 'SELECT DISTINCT "EAN", "Bezeichnung", "Land", "VK1", ' \
	 + '                   "VK0", "VKEinheit" '\
	 + 'FROM "V_Artikelinfo" ' \
	 + 'WHERE "Waage" = \'A\' AND "LadenID" = \'PLATTSALAT\' AND "WG" = %i ' \
	 + 'ORDER BY "Bezeichnung"'

	# Obtain lists from DB via sql query
	listGemuese = db.queryResult(sql % 1, 'ISSDDS')
	listObst    = db.queryResult(sql % 3, 'ISSDDS')

	# Use a consistant capitalization for the unit
	for r in listGemuese: r[5] = r[5].capitalize()
	for r in listObst:    r[5] = r[5].capitalize()

	sheet = Sheet('Waagenliste', 2, titlerows=1)
	sheet.addData(listGemuese, listObst, style='AltGrey')
	sheet.addColumns([
		ColumnDef(height=13, width=10, bold=True, hleft=True),
		ColumnDef(height=13, width=57, bold=True, tryOptWidth=True),
		ColumnDef(width=7),
		ColumnDef(height=14, width=21),
		ColumnDef(height=14, width=21),
		ColumnDef(width=8, greyUnit=True, hright=True)
	]);
	sheet.formatColumns()
	sheet.setListLabels("Gemüse", "Obst", cheight=15)
	sheet.setHeaderRow([
		[2,'Land',            ColumnDef(hcenter=True, height=9)],
		[3,'Mitglieder',      ColumnDef(hcenter=True, height=10, bold=True)],
		[4,'Nicht-\nmitglieder', ColumnDef(hcenter=True, height=10, bold=True)]
	])
	sheet.setPageStyle(landscape=True, pages=2)
	
	return None


def KassenlisteGemuese(*args):
	db = BioOfficeConn()

	sql = 'SELECT DISTINCT "EAN", "Bezeichnung", "Land", "VKEinheit", ' \
	 + '      "VK1", "VK0" ' \
	 + 'FROM "V_Artikelinfo" ' \
	 + 'WHERE "Waage" = \'A\' AND "LadenID" = \'PLATTSALAT\' AND "WG" = %i ' \
	 + 'ORDER BY "Bezeichnung"'

	# Obtain lists from DB via sql query
	listGemuese = db.queryResult(sql % 1, 'ISSSDD')
	listObst    = db.queryResult(sql % 3, 'ISSSDD')

	# Use a consistant capitalization for the unit
	for r in listGemuese: r[3] = r[3].capitalize()
	for r in listObst:    r[3] = r[3].capitalize()

	sheet = Sheet('Kassenliste', 2)
	sheet.addData(listGemuese, listObst)
	sheet.addColumns([
		ColumnDef(width=10, bold=True),        # EAN
		ColumnDef(width=50, tryOptWidth=True), # Bezeichnung
		ColumnDef(width=8),                    # Land
		ColumnDef(width=8, greyUnit=True),     # VKEinheit
		ColumnDef(width=17), # Preis Mitglieder
		ColumnDef(width=17)  # Preis Andere
	])
	sheet.formatColumns()
	sheet.setListLabels("Gemüse", "Obst")
	sheet.setPageStyle()
	return None


sql_brot = 'SELECT DISTINCT "EAN", "Bezeichnung", "VKEinheit", ' \
	 + '      "VK1", "VK0" ' \
	 + 'FROM "V_Artikelinfo" ' \
	 + 'WHERE "LadenID" = \'PLATTSALAT\' AND "WG" = \'%s\' ' \
	 + '  AND "LiefID" = \'%s\' ' \
	 + '  AND "EAN" <= 9999 AND "EAN" >= 1000 ' \
	 + 'ORDER BY "Bezeichnung"'

def KassenlisteBrot(name, id):
	db = BioOfficeConn()

	# Obtain lists from DB via sql query
	lst1 = db.queryResult(sql_brot % ('0020', id), 'ISSDD')
	lst2 = db.queryResult(sql_brot % ('0025', id), 'ISSDD')

	# Use a consistant capitalization for the unit
	for r in lst1: r[2] = r[2].capitalize()
	for r in lst2: r[2] = r[2].capitalize()

	sheet = Sheet('KassenlisteBrot'+id, 2)
	sheet.addData(lst1, lst2)
	sheet.addColumns([
		ColumnDef(width=15, bold=True),        # EAN
		ColumnDef(width=50, tryOptWidth=True), # Bezeichnung
		ColumnDef(width=12, greyUnit=True),    # VKEinheit
		ColumnDef(width=14, height=10), # Preis Mitglieder
		ColumnDef(width=14, height=10)  # Preis Andere
	])
	sheet.formatColumns()
	sheet.setListLabels(name + ' Brot', name + ' Kleingebäck')
	sheet.setPageStyle()
	return None

def KassenlisteBrotS(*args):
	return KassenlisteBrot('Schäfer', 'SCHÄFERBROT')

def KassenlisteBrotW(*args):
	return KassenlisteBrot('Weber', 'WEBER')

sql_fleisch = """SELECT
  "EAN", "Bezeichnung", "VKEinheit", "VK1", "VK0"
FROM "V_Artikelinfo"
WHERE "LadenID" = 'PLATTSALAT'
  AND "WG" = '0090'
  AND "LiefID" = '%s'
ORDER BY "Bezeichnung" """

def KassenlisteFleisch(name, id):
	db = BioOfficeConn()

	lst = db.queryResult(sql_fleisch % id, 'ISSDD')
	for r in lst: r[2] = r[2].capitalize()

	sheet = Sheet('KassenlisteFleisch'+name, 2)
	sheet.addData(lst)
	sheet.addColumns([
		ColumnDef(width=15, bold=True),        # EAN
		ColumnDef(width=50, tryOptWidth=True), # Bezeichnung
		ColumnDef(width=12, greyUnit=True),    # VKEinheit
		ColumnDef(width=14, height=10), # Preis Mitglieder
		ColumnDef(width=14, height=10)  # Preis Andere
	])
	sheet.formatColumns()
	sheet.setListLabels('Fleisch ' + name)
	sheet.setPageStyle()
	return None

def KassenlisteFleischFau(*args):
	return KassenlisteFleisch('Fauser', 'FAUSER')

def KassenlisteFleischUnt(*args):
	return KassenlisteFleisch('Unterweger', 'UNTERWEGER')

def KassenlisteFleischUri(*args):
	return KassenlisteFleisch('Uria', 'URIA')

def wglist(*args):
	return "'" + "', '".join(args) + "'"

sql_loses1 = """SELECT DISTINCT
  "EAN", "Bezeichnung", "VKEinheit", "VK1", "VK0"
FROM "V_Artikelinfo"
WHERE "LadenID" = 'PLATTSALAT'
  AND "WG" = '%s'
ORDER BY "Bezeichnung" """

sql_loses2 = """SELECT DISTINCT
  "EAN", "Bezeichnung", "VKEinheit", "VK1", "VK0"
FROM "V_Artikelinfo"
WHERE "LadenID" = 'PLATTSALAT'
  AND "WG" in (%s)
  AND "iWG" = 'HH'
ORDER BY "Bezeichnung" """

sql_loses3 = """SELECT DISTINCT
  "EAN", "Bezeichnung", "VKEinheit", "VK1", "VK0"
FROM "V_Artikelinfo"
WHERE "LadenID" = 'PLATTSALAT'
  AND "LiefID" = 'TENNENTAL'
  AND "WG" in (%s)
  AND "iWG" = 'HH'
ORDER BY "Bezeichnung" """

def KassenlisteLoseWare(*args):
	db = BioOfficeConn()

	lst1 = db.queryResult(sql_loses1 % '0585', 'ISSDD')
	lst2 = db.queryResult(sql_loses1 % '0590', 'ISSDD')
	lst3 = db.queryResult(sql_loses2 % wglist('0400'), 'ISSDD')
	lst4 = db.queryResult(sql_loses2 %
						  wglist('0070', '0200', '0280', '0340'), 'ISSDD')
	lst5 = db.queryResult(sql_loses2 %
						  wglist('0020', '0025', '0060'), 'ISSDD')
	for r in lst1: r[2] = r[2].capitalize()
	for r in lst2: r[2] = r[2].capitalize()
	for r in lst3: r[2] = r[2].capitalize()
	for r in lst4: r[2] = r[2].capitalize()
	for r in lst5: r[2] = r[2].capitalize()

	sheet = Sheet('KassenlisteLoseWare', 2)
	sheet.addData(lst1, lst2, lst3, lst4, lst5)
	sheet.addColumns([
		ColumnDef(width=17, bold=True),        # EAN
		ColumnDef(width=50, tryOptWidth=True), # Bezeichnung
		ColumnDef(width=12, greyUnit=True),    # VKEinheit
		ColumnDef(width=16, height=10), # Preis Mitglieder
		ColumnDef(width=16, height=10)  # Preis Andere
	])
	sheet.formatColumns()
	sheet.setListLabels('Lose Lebensmittel', 'Lose Waschmittel',
						'Säfte', '5 Elemente', 'Tennental')
	sheet.setPageStyle()
	return None

# Only export the public functions as macros
g_exportedScripts = [
	Waagenliste,
	KassenlisteGemuese,
	KassenlisteBrotS,
	KassenlisteBrotW,
	KassenlisteFleischFau,
	KassenlisteFleischUnt,
	KassenlisteFleischUri,
	KassenlisteLoseWare
]
