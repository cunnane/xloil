=========
xlOil SQL
=========

The SQL plugin uses sqlite3 to provide functions which query Excel arrays (or
ranges) as if they were tables in a database. Multiple tables can be queried 
and joined.

.. contents::
    :local:

.. _sql-getting-started:

Getting Started
---------------

xlOil_SQL does not require any settings and is automatically loaded in a default 
xlOil installation.  It should appear in the plugin list in
`%APPDATA%/xlOil/xlOil.ini`:

::

    Plugins=["xlOil_SQL.dll"]

You can open the example spreadsheet at :ref:`core-example-sheets` to see it in action.

For a quick demo, create a 3 column table of data in an new Excel workbook. 
Make the headings 'Foo', 'Bar' and 'Baz'.  The contents of the data can be 
anything you like.

Suppose the table is in cells A1:C5, then in another cell type 

::

    =xloSql("SELECT Bar, Baz FROM Table1", A1:C5)

Make the output an array formula with Ctrl-Shift-Enter and size it 
appropriately.


xloSql: executes a query on multiples data arrays
-------------------------------------------------

.. function:: xloSql(Query, [Meta], [Table1], [Table2], [Table3], ...)

    Excecutes the SQL query on the provided tables, returning the 
    result in an array. The tables will be named Table1, Table2, etc in the 
    query but this can overrided by the `meta` parameter

        Query:
            a string or array of string (which will be concatenated) 
            describing a query in SQL (sqlite3). 

        Meta: 
            optional array of string. The first column contains the 
            names of the tables. Subsequent columns are interpreted
            as column headings for the table. Providing a blank table
            name or few names than tables results in the un-named
            tables retaining their default name of tableN
        
        TableN:
            each table argument should point to an array of data with
            columns as fields and records as rows. Unless column
            names are specified in the meta, the first row is interpreted as column names


    **Examples**
    
    (Arguments pointing to array data are surrouned by `{}`)

    ::

        =xloSql("SELECT table1.A, B, C FROM table1 ",  { A    B  } , { A     C } )
                "INNER JOIN table2                 "   { Foo  1  }   { Bar   2 } )
                "ON table1.A == table2.A           "   { Baz  7  }   { Foo   3 } )

        --> Foo 1 3

Stateful Database Functions
---------------------------

This family of functions can be used to build up and repeatedly query an 
in-memory database for cases where building the database on the fly using 
:any:`xloSql` is not performant.

xloSqlDB
~~~~~~~~

.. function:: xloSqlDB()

    Returns a reference to a new database object. The functions :any:`xloSqlDB`, :any:`xloSqlTable`
    and :any:`xloSqlQuery` can be used to build up an in-memory database for the cases where
    building these objects on the fly using :any:`xloSql` is not performant.

xloSqlTable
~~~~~~~~~~~

.. function:: xloSqlTable(Database, Data, Name, [Headings], [Query])

    Creates a table in a database created with :any:`xloSqlDB`.  The function returns a reference 
    to the database: it is recommended to chain xloSqlTable calls to force execution order
    in Excel. This ensures tables are added to the database before any queries are run

        Database:
            a reference to a database created with `xloSqlDB`. 

        Data: 
            an array of data with columns as fields and records as rows. Unless column
            headings are specified, the first row is interpreted as column names

        Name:
            The name of the table in the database. This must be unique.
        
        Headings:
            optional column headings for the data. If these are specified, data is read
            from the first input row
        
        Query:
            An optional query to process the data as it is copied into the database.
            If ommitted, "SELECT * FROM name" is used.

xloSqlQuery
~~~~~~~~~~~
.. function:: xloSqlQuery(Database, Query)

        Database:
            A reference to a database originally created with :any:`xloSqlDB` but which has
            passed through calls to :any:`xloSqlTable`.

        Query:
            A SQL query to execute. Tables referenced in the query must have been added 
            to the database by :any:`xloSqlTable` before this function is called.


   **Examples**

    ::

        .              A                               B       C       D   
        1 =xloSqlDB()                                  MyTab   Foo     Bar
        2                                                      7       2
        3 =xloSqlTable(A1, C1:D4, B1)                          4       1
        4                                                      8       4
        5
        6 =xloSqlQuery(A3, "SELECT Bar FROM MyTab")

        Cell A6 will contain the array [2, 1, 4]

xloSqlTables
~~~~~~~~~~~~

.. function:: xloSqlTables(Database)

    Returns an array of all table names in the database
