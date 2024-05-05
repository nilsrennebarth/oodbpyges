import types


def mkcctx(cc):
	global CC
	CC = cc


class BioOfficeConn:
	"""Connection to our Bio-Office database"""

	def __init__(self):
		# Obtain connection to our database.
		# Needs the registered data source "bodb"
		self.dbconn = CC.ServiceManager.createInstanceWithContext(
			"com.sun.star.sdb.DatabaseContext", CC
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


def mkincond(name, value):
	lst = ','.join(f"'{v}'" for v in value)
	return f'{name} IN ({lst})'


def mkeqcond(name, value):
	return f"{name} = '{value}'"


class Query(types.SimpleNamespace):
	SQL = 'SELECT DISTINCT {cols} FROM V_Artikelinfo '\
		"WHERE LadenID = 'PLATTSALAT' AND {cons} " \
		'ORDER BY Bezeichnung'

	EAN = 'CAST(CAST(EAN AS DECIMAL(20)) AS VARCHAR(20))'

	Cols = ["EAN", "Bezeichnung", "Land", "VK1", "VK0", "VKEinheit"]
	SCols = "SSSDDS"

	CONDS = []

	def __init__(self, wg=None, iwg=None, liefer=None) -> None:
		self.wg, self.iwg, self.liefer = wg, iwg, liefer

	def run(self):
		self.cols = ','.join(self.EAN if c == 'EAN' else f'{c}' for c in self.Cols)
		conditions = self.CONDS.copy()
		for n, name in dict(iwg='iWG', liefer='LiefID', wg='WG').items():
			value = self.__dict__[n]
			if value is None: continue
			if isinstance(value, list):
				conditions.append(mkincond(name, value))
			else:
				conditions.append(mkeqcond(name, value))
		self.cons = ' AND '.join(conditions)
		self.sql = self.SQL.format_map(self.__dict__)
		# log.debug(f'Query: {self.sql}')
		return BioOfficeConn().queryResult(self.sql, self.SCols)
