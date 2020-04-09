=========
xlOil SQL
=========

The SQL plugin uses sqlite3 to provide functions which query Excel 
arrays (or ranges) as if they were tables in a database. Multiple tables can be queried and joined.

.. contents::
	:local:

.. function:: xloSql(Query, Meta, Table1, Table2, Table3, ...)

	Excecutes the SQL query on the provided tables, returning an
	array. The tables will be named Table1, Table2, etc for the 
	purposes of the query unless overrided by the meta.

		query:
			a string or array of string (which will be concatenated) 
			describing a query in SQL (sqlite3). 

		meta: 
			optional array of string. The first column contains the 
			names of the tables. Subsequent columns are interpreted
			as column headings for the table. Providing a blank table
			name or few names than tables results in the un-named
			tables retaining their default name of tableN
		
		tableN:
			each table argument should point to a array of data with
			columns as fields and records as rows. Unless column
			names are specified in the meta, the first row is interpreted as column names

