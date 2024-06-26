# Plattsalat specific python macros
import collections
import datetime
import numbers
import types
import logging
import uno
from bodb import Query, mkcctx
from com.sun.star.lang import Locale
from com.sun.star.table.CellVertJustify import CENTER as vertCenter
from com.sun.star.table.CellHoriJustify import CENTER as horCenter
from com.sun.star.table.CellHoriJustify import RIGHT as horRight
from com.sun.star.table.CellHoriJustify import LEFT as horLeft
from com.sun.star.table import CellRangeAddress


def do_log(fname='/home/nils/tmp/oodebug.log'):
	global log

	logging.basicConfig(filename=fname)
	log = logging.getLogger('libreoffice')
	log.setLevel(logging.DEBUG)


Pos = collections.namedtuple('Pos', 'x y')


class ColumnDef(types.SimpleNamespace):
	"""Options for a single column in a table

	This is mostly a container for various options. The following
	options are currently recognized:
	- width (int) width in mm
	- height (int) char height
	- tryOptWidth (boolean) First try to set the width to its optimum
	. value. Only if that is too big, set it to the given width
	- bold (boolean) set typeface to bold
	- greyUnit (boolean) set background to grey if the text appears
	. to represent discrete units
	- hcenter (boolean) Center horizontally
	- hright (boolean) Align on the right

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
			Locale('de', 'DE', '')
		)
		self.Linestyle = uno.createUnoStruct("com.sun.star.table.BorderLine2")
		self.Linestyle.OuterLineWidth = 5
		self.ColDefs = []
		self.Boldface = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")
		# Get the default cell style
		# and use it to set use a 12pt Font Size by default
		cs = self.calc.StyleFamilies.CellStyles.getByName('Default')
		cs.CharHeight = 12

	def addColumns(self, cols):
		self.ColDefs += cols

	def getCell(self, x, y):
		return self.sheet.getCellByPosition(x, y)

	def getMergeCell(self, x, y):
		r = self.sheet.getCellRangeByPosition(x, y, x, y + 1)
		r.merge(True)
		return self.sheet.getCellByPosition(x, y)

	def getCol(self, col):
		return self.sheet.getColumns().getByIndex(col)

	def getRow(self, row):
		return self.sheet.getRows().getByIndex(row)

	def styleBlock(self, x, y, n):
		"""Style a row, Blocks with lines everywhere.
		"""
		cells = self.sheet.getCellRangeByPosition(x, y, x + n - 1, y)
		cells.LeftBorder = self.Linestyle
		cells.RightBorder = self.Linestyle
		cells.TopBorder = self.Linestyle
		cells.BottomBorder = self.Linestyle
		cells.ParaRightMargin = 100
		cells.ParaLeftMargin = 100

	def styleAltGrey(self, x, y, n):
		"""Style a row, Alternating grey background
		"""
		self.getCell(x, y).LeftBorder = self.Linestyle
		self.getCell(x + n - 1, y).RightBorder = self.Linestyle
		if (y & 1) == 1:
			cells = self.sheet.getCellRangeByPosition(
				x, y, x + n - 1, y
			)
			cells.CellBackColor = 0xdddddd

	def addData(self, *lists, style='Block'):
		mysheet = self

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
		# are supposed to have the same number of columns.
		self.colCols = max(len(ll[0]) if len(ll) > 0 else 0 for ll in lists)
		if self.colCols == 0:
			raise ValueError('All lists are empty')
		self.HeaderPositions = []
		self.totalCols = self.cols * (self.colCols + 1) - 1
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
			if len(lists) > 1:
				pos.advance()
			for row in list:
				for i, val in enumerate(row):
					cell = self.getCell(pos.x + i, pos.y)
					if isinstance(val, numbers.Number) and val < 2000000000:
						cell.Value = val
					else:
						cell.String = val
					if isinstance(val, float):
						cell.NumberFormat = self.currencyformat
				styler(pos.x, pos.y, self.colCols)
				pos.advance()
			# advance once at the end of a list
			pos.advance()
			if rest > 0:
				pos.advance()
				rest -= 1

	def addPagelistrow(self, row):
		cell = self.getMergeCell(0, self.crow)
		cell.String = row[0]
		cell = self.getMergeCell(1, self.crow)
		cell.String = row[1]
		cell = self.getCell(2, self.crow)
		cell.String = row[2]
		cell = self.getCell(2, self.crow+1)
		cell.String = row[3]
		cell = self.getMergeCell(3, self.crow)
		cell.Value = row[4]
		cell.NumberFormat = self.currencyformat
		cell = self.getMergeCell(4, self.crow)
		cell.Value = row[5]
		cell.NumberFormat = self.currencyformat
		self.crow += 2

	def addPagelist(self, *lists, style='Block', hstretch=1.2):
		"""Add a single page list in fixed layout

		Solely used by Wagenlisten, which produces several pages,
		one for each location.
		"""
		self.crow = self.titlerows
		self.colCols = len(lists[0][0])
		self.HeaderPositions = []
		self.totalCols = self.colCols
		styler = getattr(self, 'style'+style)
		for list in lists:
			if self.crow > self.titlerows:
				self.getRow(self.crow).IsStartOfNewPage = True
			for row in list:
				self.addPagelistrow(row)
				styler(0, self.crow-2, self.colCols-1)
				styler(0, self.crow-1, self.colCols-1)

	def getOptimalScale(self, header=False):
		"""Calculate the optimal scale factor in percent
		"""
		w = 0
		for i in range(self.totalCols):
			w += self.getCol(i).Width
		h = 0
		for i in range(self.totalRows):
			h += self.getRow(i).Height
		if h == 0 or w == 0: return 100  # should not happen
		ws = 19500 / w  # factor to scale to 195mm width
		hs = 28200 / h  # factor to scale to 270mm height
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

	def getOptimalScaleExt(self, landscape, pages, header=False):
		nrows = (self.totalRows + pages-1) // pages
		w = 0
		for i in range(self.totalCols):
			w += self.getCol(i).Width
		h = 0
		for i in range(nrows):
			h += self.getRow(i+self.titlerows).Height
		for i in range(self.titlerows):
			h += self.getRow(i).Height
		if h == 0 or w == 0: return 100  # should not happen
		if landscape:
			towidth = 28400
			toheight = 19800
		else:
			towidth = 19800
			toheight = 28400
		if header:
			toheight -= 900
		ws = towidth / w
		hs = toheight / h
		if hs < ws: return int(hs * 100)
		hstretch = toheight / (h * ws)
		if hstretch > 1.8: hstretch = 1.8
		for i in range(self.titlerows, self.totalRows):
			self.getRow(i).Height = self.getRow(i).Height * hstretch
		return int(ws * 100)

	def pieceMarker(self, x, y):
		cell = self.getCell(x, y)
		if len(cell.String) == 2 and cell.String != 'Kg':
			# cell.CellBackColor = 0xdddddd
			cell.CharWeight = self.Boldface

	def pieceMarkCol(self, col):
		for i in range(self.totalRows):
			self.pieceMarker(col, i)

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
			self.pieceMarkCol(i)
		if cdef.height != 12:
			col.CharHeight = cdef.height
		if cdef.hright:
			col.HoriJustify = horRight
		if cdef.hleft:
			col.HoriJustify = horLeft
		col.VertJustify = vertCenter

	def formatColumns(self):
		for t in range(self.cols):
			for i, cdef in enumerate(self.ColDefs):
				self.formatCol(t * (self.colCols + 1) + i, cdef)
			if t < self.cols-1:
				self.getCol((t+1) * (self.colCols + 1) - 1).Width = 800

	def setListLabels(self, *labels, cheight=14):
		for i, l in enumerate(labels):
			p = self.HeaderPositions[i]
			cell = self.getCell(p.x + 1, p.y)
			cell.String = l
			cell.CharHeight = cheight
			cell.CharWeight = self.Boldface

	def setPageStyle(self, landscape=False, maxscale=True, pages=1, date=False):
		defp = self.calc.StyleFamilies.PageStyles.getByName("Default")
		defp.LeftMargin = 500
		defp.TopMargin = 500
		defp.BottomMargin = 500
		defp.RightMargin = 500
		defp.HeaderIsOn = False
		defp.FooterIsOn = False
		defp.CenterHorizontally = True
		defp.CenterVertically = False
		if landscape:
			defp.Width = 29700
			defp.Height = 21000
			defp.IsLandscape = True
		if date:
			defp.HeaderIsOn = True
			hs = defp.RightPageHeaderContent
			hs.LeftText.String = datetime.date.today().strftime('%d.%m.%Y')
			hs.CenterText.String = ''
			defp.RightPageHeaderContent = hs
		if maxscale:
			if landscape or pages > 1:
				defp.PageScale = self.getOptimalScaleExt(landscape, pages, header=date)
			else:
				defp.PageScale = self.getOptimalScale(header=date)

	def setHeaderRow(self, titles):
		self.sheet.setTitleRows(CellRangeAddress(StartRow=0, EndRow=0))
		for i in range(self.cols):
			for title in titles:
				pos = title[0]
				cdef = title[2]
				cell = self.getCell(i * (self.colCols + 1) + pos, 0)
				cell.String = title[1]
				if cdef.bold:
					cell.CharWeight = self.Boldface
				if cdef.height != 12:
					cell.CharHeight = cdef.height
				if cdef.hcenter:
					cell.HoriJustify = horCenter


class WaagenlistenQuery(Query):
	Cols = ["EAN", "Bezeichnung", "Land", "VKEinheit", "VK1", "VK0"]
	SCols = "SSSSDD"
	CONDS = [
		"Waage = 'A'",
		"WG IN ('0001', '0003')"
	]


def Waagenlisten(*args):
	"""
	Location based lists

	For each of the 7 locations create a landscape formatted
	page with large items, all on one sheet with page breaks
	ready to print.

	These lists will be placed at the various places where
	fruits and vegetables can be found.
	"""

	locs = [
		'Apfel', 'Kartoffel', 'Knoblauch', 'kühl links', 'kühl rechts',
		'Pilze', 'Zitrone', 'Zwiebel'
	]

	lists = []

	for loc in locs:
		# Obtain list for location
		L = WaagenlistenQuery(iwg=loc).run()
		# Use consistent capitalization for the unit
		for r in L: r[3] = r[3].capitalize()
		lists.append(L)

	sheet = Sheet('Waagenliste', 1, titlerows=1)
	sheet.addPagelist(*lists)

	sheet.addColumns([
		ColumnDef(height=24, width=18, bold=True, hleft=True),
		ColumnDef(height=29, width=100, bold=True),
		ColumnDef(width=8),
		ColumnDef(height=22, width=35),
		ColumnDef(height=22, width=35),
	])
	sheet.formatColumns()
	sheet.setHeaderRow([
		[2, '', ColumnDef(hcenter=True, height=9)],
		[3, 'Mitglieder', ColumnDef(hcenter=True, height=10, bold=True)],
		[4, 'Nichtmitglieder', ColumnDef(hcenter=True, height=10, bold=True)]
	])
	sheet.setPageStyle(maxscale=False, date=True)


class WaageQuery(Query):
	Cols = ["EAN", "Bezeichnung", "Land", "VK1", "VK0", "VKEinheit"]
	SCols = "SSSDDS"
	CONDS = ["Waage = 'A'"]


def Waagenliste(*args):
	"""Lists for the electronic balances

	Create a ready to print spreadsheet for the
	electronic balances, containing the EAN numbers,
	the names and the unit

	The list is in landscape format and fitted to two pages.
	"""
	# Obtain lists from DB via sql query
	listGemuese = WaageQuery(wg='0001').run()
	listObst = WaageQuery(wg='0003').run()

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
	])
	sheet.formatColumns()
	sheet.setListLabels("Gemüse", "Obst", cheight=15)
	sheet.setHeaderRow([
		[2, 'Land', ColumnDef(hcenter=True, height=9)],
		[3, 'Mitglieder', ColumnDef(hcenter=True, height=10, bold=True)],
		[4, 'Nicht-\nmitglieder', ColumnDef(hcenter=True, height=10, bold=True)]
	])
	sheet.setPageStyle(landscape=True, pages=2, date=True)

	return None


class WaagenupQuery(Query):
	Cols = ["EAN", "Bezeichnung", "VK1", "VKEinheit"]
	SCols = "SSDS"
	CONDS = ["Waage = 'A'"]


def WaagenlisteUp(*args):
	"""Lists for the electronic balances

	Create a ready to print spreadsheet for the
	electronic balances, containing the EAN numbers,
	the names and the unit.

	The list is in portrait format and fitted onto a single page.
	"""
	# Obtain lists from DB via sql query
	listGemuese = WaagenupQuery(wg='0001').run()
	listObst = WaagenupQuery(wg='0003').run()

	# Use a consistant capitalization for the unit
	for r in listGemuese: r[3] = r[3].capitalize()
	for r in listObst:    r[3] = r[3].capitalize()

	sheet = Sheet('Waagenliste', 2)
	sheet.addData(listGemuese, listObst)
	sheet.addColumns([
		ColumnDef(width=10, bold=True),
		ColumnDef(width=50, tryOptWidth=True),
		ColumnDef(width=17),
		ColumnDef(width=10, greyUnit=True)
	])
	sheet.formatColumns()
	sheet.setListLabels("Gemüse", "Obst")
	sheet.setPageStyle()

	return None


class SchrankQuery(Query):
	Cols = ["EAN", "Bezeichnung", "Land", "VK1", "VK0", "Hersteller"]
	SCols = 'SSSDDS'


def _schrankliste(title, iwg, **pageopts):
	"""Lists for the Refridgerators"""
	data = SchrankQuery(iwg=iwg).run()

	sheet = Sheet(title, 1, titlerows=1)
	sheet.addData(data)
	sheet.setHeaderRow([
		[0, 'EAN', ColumnDef(bold=True, hcenter=True)],
		[1, 'Bezeichnung', ColumnDef()],
		[2, 'Land', ColumnDef()],
		[3, 'Mitglieder', ColumnDef(hcenter=True, height=10, bold=True)],
		[4, 'Nicht-\nmitglieder', ColumnDef(hcenter=True, height=10, bold=True)],
		[5, 'Hersteller', ColumnDef(hcenter=True)]
	])
	sheet.addColumns([
		ColumnDef(width=35, bold=True),
		ColumnDef(width=90),
		ColumnDef(width=10),
		ColumnDef(width=25),
		ColumnDef(width=25),
		ColumnDef(width=30)
	])
	sheet.formatColumns()
	sheet.setPageStyle(**pageopts)

	return None


def KuehlschrankLinks(*args):
	_schrankliste("Kühlschrank links", "kühl links")


def KuehlschrankRechts(*args):
	_schrankliste("Kühlschrank rechts", "kühl rechts")


def KuehlschrankMopro1(*args):
	_schrankliste("Kühlschrank Molkereiprodukte 1", "1Mopro")


def KuehlschrankMopro2(*args):
	_schrankliste("Kühlschrank Molkereiprodukte 2", "2Mopro")


def KuehlschrankMix(*args):
	_schrankliste("Kühlschrank Mix", "3Mix")


def KuehlschrankVegan(*args):
	_schrankliste("Kühlschrank Vegan", "4Vegan")


def KuehlschrankFleisch(*args):
	_schrankliste("Kühlschrank Fleisch", "5Fleisch", pages=2)


class KassenlandQuery(Query):
	Cols = ["EAN", "Bezeichnung", "Land", "VKEinheit", "VK1", "VK0"]
	SCols = "SSSSDD"
	CONDS = ["Waage = 'A'"]


def KassenlisteGemuese(*args):
	# Obtain lists from DB via sql query
	listGemuese = KassenlandQuery(wg='0001').run()
	listObst = KassenlandQuery(wg='0003').run()

	# Use a consistant capitalization for the unit
	for r in listGemuese: r[3] = r[3].capitalize()
	for r in listObst:    r[3] = r[3].capitalize()

	sheet = Sheet('Kassenliste', 2)
	sheet.addData(listGemuese, listObst)
	sheet.addColumns([
		ColumnDef(width=10, bold=True),         # EAN
		ColumnDef(width=50, tryOptWidth=True),  # Bezeichnung
		ColumnDef(width=8),                     # Land
		ColumnDef(width=8, greyUnit=True),      # VKEinheit
		ColumnDef(width=17),  # Preis Mitglieder
		ColumnDef(width=17)   # Preis Andere
	])
	sheet.formatColumns()
	sheet.setListLabels("Gemüse", "Obst")
	sheet.setPageStyle()
	return None


class KassenQuery(Query):
	Cols = ["EAN", "Bezeichnung", "VKEinheit", "VK1", "VK0"]
	SCols = "SSSDD"


def KassenlisteBrot(name, id):
	# Obtain lists from DB via sql query
	lst1 = KassenQuery(wg='0020', liefer=id).run()
	lst2 = KassenQuery(wg='0025', liefer=id).run()

	# Use a consistant capitalization for the unit
	for r in lst1: r[2] = r[2].capitalize()
	for r in lst2: r[2] = r[2].capitalize()

	sheet = Sheet('KassenlisteBrot'+id, 2)
	sheet.addData(lst1, lst2)
	sheet.addColumns([
		ColumnDef(width=15, bold=True),         # EAN
		ColumnDef(width=50, tryOptWidth=True),  # Bezeichnung
		ColumnDef(width=12, greyUnit=True),     # VKEinheit
		ColumnDef(width=14, height=10),  # Preis Mitglieder
		ColumnDef(width=14, height=10)   # Preis Andere
	])
	sheet.formatColumns()
	sheet.setListLabels(name + ' Brot', name + ' Kleingebäck')
	sheet.setPageStyle()
	return None


def KassenlisteBrotS(*args):
	return KassenlisteBrot('Schäfer', 'SCHÄFERBROT')


def KassenlisteBrotW(*args):
	return KassenlisteBrot('Weber', 'WEBER')


def KassenlisteFleisch(name, id):
	lst = KassenQuery(wg='0090', liefer=id).run()
	for r in lst: r[2] = r[2].capitalize()

	sheet = Sheet('KassenlisteFleisch'+name, 2)
	sheet.addData(lst)
	sheet.addColumns([
		ColumnDef(width=15, bold=True),         # EAN
		ColumnDef(width=50, tryOptWidth=True),  # Bezeichnung
		ColumnDef(width=12, greyUnit=True),     # VKEinheit
		ColumnDef(width=14, height=10),  # Preis Mitglieder
		ColumnDef(width=14, height=10)   # Preis Andere
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


def KassenlisteLoseWare(*args):
	lst1 = KassenQuery(wg='0585').run()
	lst2 = KassenQuery(wg='0590').run()
	lst3 = KassenQuery(iwg='HH', wg='0400').run()
	lst4 = KassenQuery(iwg='HH', wg=['0070', '0200', '0280', '0340']).run()
	lst5 = KassenQuery(iwg='HH', wg=['0020', '0025', '0060']).run()

	for r in lst1: r[2] = r[2].capitalize()
	for r in lst2: r[2] = r[2].capitalize()
	for r in lst3: r[2] = r[2].capitalize()
	for r in lst4: r[2] = r[2].capitalize()
	for r in lst5: r[2] = r[2].capitalize()

	sheet = Sheet('KassenlisteLoseWare', 2)
	sheet.addData(lst1, lst2, lst3, lst4, lst5)
	sheet.addColumns([
		ColumnDef(width=32, bold=True),         # EAN
		ColumnDef(width=50, tryOptWidth=True),  # Bezeichnung
		ColumnDef(width=12, greyUnit=True),     # VKEinheit
		ColumnDef(width=16, height=10),  # Preis Mitglieder
		ColumnDef(width=16, height=10)   # Preis Andere
	])
	sheet.formatColumns()
	sheet.setListLabels(
		'Lose Lebensmittel', 'Lose Waschmittel',
		'Säfte', '5 Elemente', 'Tennental'
	)
	sheet.setPageStyle()
	return None


# Only export the public functions as macros
mkcctx(XSCRIPTCONTEXT.getComponentContext())
g_exportedScripts = [
	KassenlisteBrotS,
	KassenlisteBrotW,
	KassenlisteFleischFau,
	KassenlisteFleischUnt,
	KassenlisteFleischUri,
	KassenlisteGemuese,
	KassenlisteLoseWare,
	KuehlschrankMopro1,
	KuehlschrankMopro2,
	KuehlschrankMix,
	KuehlschrankVegan,
	KuehlschrankFleisch,
	Waagenliste,
	WaagenlisteUp,
	Waagenlisten
]
