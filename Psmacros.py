# Plattsalat specific python macros
import collections
import uno
from com.sun.star.lang import Locale

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
				meths.append(getattr(dbres, 'getInt'))
			elif c == 'S':
				meths.append(getattr(dbres, 'getString'))
			elif c == 'D':
				meths.append(getattr(dbres, 'getDouble'))

		while dbres.next():
			result.append([meths[i](i+1) for i in range(len(meths))])
		return result

Pos = collections.namedtuple('Pos', 'x y')

class Sheet:
	"""A single sheet to be filled with tables"""

	def __init__(self, name, cols):
		desktop = XSCRIPTCONTEXT.getDesktop()
		# Create a new calc and use its first sheet
		self.calc = desktop.loadComponentFromURL(
			"private:factory/scalc", "_blank", 0, ()
		)
		self.sheet = self.calc.Sheets.getByIndex(0)
		self.sheet.Name = name
		self.cols = cols
		self.currencyformat = self.calc.NumberFormats.getStandardFormat(
			uno.getConstantByName("com.sun.star.util.NumberFormat.CURRENCY"),
			Locale('de','DE','')
		)
		self.Linestyle = uno.createUnoStruct("com.sun.star.table.BorderLine2")
		self.Linestyle.OuterLineWidth = 4

	def getCell(self, x, y):
		return self.sheet.getCellByPosition(x, y)

	def getCol(self, col):
		return self.sheet.getColumns().getByIndex(col)

	def getRow(self, row):
		return self.sheet.getRows().getByIndex(row)

	def setColWidth(self, col, width):
		"""Set column of a sheet to a certain width

		col: Column index (0 based)
		width: width in mm
		"""
		self.sheet.getColumns().getByIndex(col).Width = width * 100

	def setColOptimal(self, col, maxwidth):
		"""Set column to optimal width

		col: Column index (0 based)
		maxwidth: maximal width in mm
		"""
		col = self.sheet.getColumns().getByIndex(col)
		col.OptimalWidth = True
		if col.Width > maxwidth * 100:
			col.Width = maxwidth * 100

	def addLines(self, x, y, n):
		"""Add lines around all cells in part of a row"""
		cells = self.sheet.getCellRangeByPosition(x,y,x+n-1,y)
		cells.LeftBorder   = self.Linestyle
		cells.RightBorder  = self.Linestyle
		cells.TopBorder    = self.Linestyle
		cells.BottomBorder = self.Linestyle

	def addData(self, *lists):

		class Cellpos:
			def __init__(self, cols, rows):
				self.x = 0
				self.y = 0
				self.cols = cols
				self.rows = rows

			def advance(self):
				# go one down
				self.y += 1
				# if at bottom row, go to top and left
				if self.y == self.rows:
					self.x = self.x + self.cols + 1
					self.y = 0

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
				self.addLines(pos.x, pos.y, self.colCols)
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
		ws = 19900 / w # factor to scale to 199mm width
		hs = 28600 / h # factor to scale to 286mm height
		# return the smaller one
		return ws * 100 if ws < hs else hs * 100


	def addGrey(self, col):
		for i in range(self.totalRows):
			cell = self.getCell(col, i)
			if len(cell.String) == 2 and cell.String != 'Kg':
				cell.CellBackColor = 0xdddddd


def Waagenliste(*args):
	"""Lists for the electronic balances
	
	Create a ready to print spreadsheet for the
	electronic balances, containing the EAN numbers,
	the names and the unit
	"""

	db = BioOfficeConn()

	sql = 'SELECT DISTINCT "EAN", "Bezeichnung", "VK0", "VKEinheit" ' \
	 + 'FROM "V_Artikelinfo" ' \
	 + 'WHERE "Waage" = \'A\' AND "LadenID" = \'PLATTSALAT\' AND "WG" = %i ' \
	 + 'ORDER BY "Bezeichnung"'

	# Obtain lists from DB via sql query
	listGemuese = db.queryResult(sql % 1, 'ISDS')
	listObst    = db.queryResult(sql % 3, 'ISDS')

	# Use a consistant capitalization for the unit
	for r in listGemuese: r[3] = r[3].capitalize()
	for r in listObst:    r[3] = r[3].capitalize()

	sheet = Sheet('Waagenliste', 2)
	doc = sheet.calc

	# Get the default cell style
	# and use it to set use a 12pt Font Size by default
	cs = doc.getStyleFamilies().getByName('CellStyles').getByName('Default')
	cs.CharHeight=12

	sheet.addData(listGemuese, listObst)

	bf = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")

	p = sheet.HeaderPositions[0]
	cell = sheet.getCell(p.x + 1, p.y)
	cell.String = "Gemüse"
	cell.CharHeight = 14
	cell.CharWeight = bf

	p = sheet.HeaderPositions[1]
	cell = sheet.getCell(p.x + 1, p.y)
	cell.String = "Obst"
	cell.CharHeight = 14
	cell.CharWeight = bf

	# set column width (mm)
	# EAN Columns fixed to 10mm
	sheet.setColWidth(0, 10)
	sheet.setColWidth(5, 10)
	# Name Columns dynamically
	sheet.setColOptimal(1, 55)
	sheet.setColOptimal(6, 55)
	# Price Coumns to 1.7mm
	sheet.setColWidth(2, 17)
	sheet.setColWidth(7, 17)
	# Unit Columns fixed to 10mm
	sheet.setColWidth(3, 10)
	sheet.setColWidth(8, 10)
	# Empty colum D fixed to 8mm
	sheet.setColWidth(4, 8)

	# a little bit of margin
	sheet.getCol(0).ParaRightMargin = 100
	sheet.getCol(5).ParaRightMargin = 100
	sheet.getCol(3).ParaLeftMargin = 100
	sheet.getCol(8).ParaLeftMargin = 100

	# EAN numbers in bold
	sheet.getCol(0).CharWeight = bf
	sheet.getCol(5).CharWeight = bf

	sheet.addGrey(3)
	sheet.addGrey(8)

	# Set Page style
	defp = doc.StyleFamilies.getByName("PageStyles").getByName("Default")
	defp.LeftMargin   = 500
	defp.TopMargin    = 500
	defp.BottomMargin = 500
	defp.RightMargin  = 500

	defp.HeaderIsOn=False
	defp.FooterIsOn=False
	defp.CenterHorizontally=True
	defp.CenterVertically=True
	defp.PageScale = sheet.getOptimalScale()
	
	return None

# Only export the public functions as macros
g_exportedScripts = Waagenliste,
