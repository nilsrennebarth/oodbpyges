# Plattsalat specific python macros
import uno

def _getQuery(sql, types, conn):
	"""Get the results of an SQL query as a list

	sql is the query as a string, types is a string specifying
	the types in each row. I is for Int, S for String, conn
	is the database connection
	"""
	meths = []
	result = []
	dbres = conn.createStatement().executeQuery(sql)
	
	# create a list of methods from the type string
	for c in types:
		if c == 'I':
			meths.append(getattr(dbres, 'getInt'))
		elif c == 'S':
			meths.append(getattr(dbres, 'getString'))

	while dbres.next():
		result.append([meths[i](i+1) for i in range(len(meths))])
	return result

def _setColWidth(sheet, col, width):
	"""Set column of a sheet to a certain width

	col: Column index (0 based)
	width: width in mm

	"""
	sheet.getColumns().getByIndex(col).Width = width * 100

def _setAllLines(range, style):
	"""Set left, rigth top, bottom borders"""
	range.LeftBorder = style
	range.RightBorder = style
	range.TopBorder = style
	range.BottomBorder = style

def _greyBg(cell):
	if cell.String == 'Kg': return
	cell.CellBackColor = 0xeeeeee


def optimalScale(sheet, nc, nr):
	"""Calculate the optimal scale factor in percent
	"""
	w=0
	for i in range(nc):
		w += sheet.getColumns().getByIndex(i).Width
	h=0
	for i in range(nr):
		h += sheet.getRows().getByIndex(i).Height
	ws = 19900 / w # factor to scale to 199mm width
	hs = 28600 / h # factor to scale to 286mm height
	# return the smaller one
	return ws * 100 if ws < hs else hs * 100

def Waagenliste(*args):
	"""Lists for the electronic balances
	
	Create a ready to print spreadsheet for the
	electronic balances, containing the EAN numbers,
	the names and the unit
	"""
	ctx = XSCRIPTCONTEXT.getComponentContext()
	desktop = XSCRIPTCONTEXT.getDesktop()
	smgr = ctx.ServiceManager
	dbctx = smgr.createInstanceWithContext(
		"com.sun.star.sdb.DatabaseContext", ctx
	)
	dbconn = dbctx.getByName("bodb").getConnection('', '')

	sql = 'SELECT DISTINCT "EAN", "Bezeichnung", "VKEinheit" ' \
	 + 'FROM "V_Artikelinfo" ' \
	 + 'WHERE "Waage" = \'A\' AND "LadenID" = \'PLATTSALAT\' AND "WG" = %i ' \
	 + 'ORDER BY "Bezeichnung"'

	# Obtain lists from DB via sql query
	listGemuese = _getQuery(sql % 1, 'ISS', dbconn)
	listObst = _getQuery(sql % 3, 'ISS', dbconn)

	# Create a new calc and use its first sheet
	calc = desktop.loadComponentFromURL(
		"private:factory/scalc", "_blank", 0, ()
	)
	sheet = calc.Sheets.getByIndex(0)

	# Get the default cell style
	cs = calc.getStyleFamilies().getByName('CellStyles').getByName('Default')
	# and use it to set use a 12pt Font Size by default
	cs.CharHeight=12

	ng = len(listGemuese)
	no = len(listObst)

	for r in listGemuese: r[2] = r[2].capitalize()
	for r in listObst: r[2] = r[2].capitalize()
	if (no + ng) % 2 == 0:
		nr = (no + ng + 2) // 2
	else:
		nr = (no + ng + 1) // 2

	# Now insert list items into sheet at proper positions
	for i in range(1, nr+1):
		r = listGemuese[i-1]
		sheet.getCellByPosition(0, i).Value  = r[0]
		sheet.getCellByPosition(1, i).String = r[1]
		sheet.getCellByPosition(2, i).String = r[2]
		_greyBg(sheet.getCellByPosition(2, i))

		if i <= ng - nr:
			r = listGemuese[i+nr-1]
			sheet.getCellByPosition(4, i-1).Value  = r[0]
			sheet.getCellByPosition(5, i-1).String = r[1]
			sheet.getCellByPosition(6, i-1).String = r[2]
			_greyBg(sheet.getCellByPosition(6, i-1))
		elif i >= nr - no + 1:
			r = listObst[i - nr + no - 1]
			sheet.getCellByPosition(4, i).Value  = r[0]
			sheet.getCellByPosition(5, i).String = r[1]
			sheet.getCellByPosition(6, i).String = r[2]
			_greyBg(sheet.getCellByPosition(6, i))

	bf = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")
	
	cell = sheet.getCellByPosition(1,0)
	cell.String = "Gemüse"
	cell.CharHeight = 14
	cell.CharWeight = bf

	cell = sheet.getCellByPosition(5,nr-no)
	cell.String = "Obst"
	cell.CharHeight = 14
	cell.CharWeight = bf

	# set column width (mm)
	# EAN Columns A and E fixed to 10mm
	_setColWidth(sheet, 0, 10)
	_setColWidth(sheet, 4, 10)
	# Unit Columns C and G fixed to 10mm
	_setColWidth(sheet, 2, 10)
	_setColWidth(sheet, 6, 10)
	# Empty colum D fixed to 8mm
	_setColWidth(sheet, 3, 8)	
	# Name Columns B and F dynamically
	sheet.getColumns().getByIndex(1).OptimalWidth=True
	sheet.getColumns().getByIndex(5).OptimalWidth=True

	# set right margin for EAN numbers
	sheet.getColumns().getByIndex(0).ParaRightMargin = 100
	sheet.getColumns().getByIndex(4).ParaRightMargin = 100
	sheet.getColumns().getByIndex(2).ParaLeftMargin = 100
	sheet.getColumns().getByIndex(6).ParaLeftMargin = 100

	sheet.getColumns().getByIndex(0).CharWeight = bf
	sheet.getColumns().getByIndex(4).CharWeight = bf

	# Set lines
	ls = uno.createUnoStruct("com.sun.star.table.BorderLine2")
	ls.OuterLineWidth = 2
	ls.LineWidth = 2
	area = sheet.getCellRangeByPosition(0,1,2,nr)
	_setAllLines(area, ls)
	area = sheet.getCellRangeByPosition(4,0,6,ng-nr-1)
	_setAllLines(area, ls)
	area = sheet.getCellRangeByPosition(4,nr-no+1,6,nr)
	_setAllLines(area, ls)

	# Set Page style
	defp = calc.StyleFamilies.getByName("PageStyles").getByName("Default")
	defp.LeftMargin   = 500
	defp.TopMargin    = 500
	defp.BottomMargin = 500
	defp.RightMargin  = 500

	defp.HeaderIsOn=False
	defp.FooterIsOn=False
	defp.CenterHorizontally=True
	defp.CenterVertically=True
	defp.PageScale = optimalScale(sheet, 7, nr+1)
	
	return None

# Only export the public functions as macros
g_exportedScripts = Waagenliste,
