# Python Pivot Table function for MySQL
#
#For example, given data as shown below:
#
# ID          year        type        amt         
#--------- ----------- ----------- ----------- 
# 7           1999        Vinyl       23
# 8           1999        Tape        44
# 9           1999        CD          55
# 10          2000        Vinyl       66
# 11          2000        Tape        77
# 12          2000        CD          88
# 13          1999        Vinyl       11
#
#
# ... you can pivot the data to show the years down the side 
# and the types across the top...
#
# year        Vinyl      Tape         CD          Total    
#--------- ----------- ----------- ----------- ----------- 
# 1999        34          44          55          133
# 2000        66          77          88          231
#
#
#... and get the SQL which would produce this table
#	SELECT Pivot_Data.*
#	FROM (SELECT [year],  
#	SUM(CASE [type] WHEN '1' THEN [amt] END) AS [Vinyl],  
#	SUM(CASE [type] WHEN '2' THEN [amt] END) AS [Tape],  
#	SUM(CASE [type] WHEN '3' THEN [amt] END) AS [CD]
#	 FROM mediaSales 
#	GROUP BY [year]) AS Pivot_Data ORDER BY [year]
#.. by calling the procedure as shown below:
# transform( 'SUM', '[amt]', 'mediaSales', '[year]', '[type]')

from re import *

def transform(curs,
              Operation,	# SUM, PRODUCT, count, etc
              Op_Argument,	# Column to use as argument in operation
              Base_Data_SQL,	# Table to use as recordsource to build final crosstab qry
              Row_Headings,	# Comma-separated list of rows to use as groupings of data
              Column_Heading	# Column to use as heading, ie. pivot column
) :
    if Operation not in ["SUM", "COUNT", "MAX", "MIN", "AVG", "STDEV", "VAR", "VARP", "STDEVP"] : raise "illegal operation error"

    CaseElse = ''
    # One doesn't want to count nulls, but one wants to sum them as 0. No longer necessary.
    #if Operation == 'SUM' : CaseElse = ' else 0'

    SQL = "SELECT " + Row_Headings + ", "

    Column_Head_SQL = "SELECT DISTINCT " + Column_Heading + " from " + Base_Data_SQL + " order by " + Column_Heading
    # print Column_Head_SQL # debug check

    # Get list of distinct column headings
    rowsAffected = curs.execute(Column_Head_SQL)
    rows = curs.fetchall()
    # print rows # debug check

    for row in rows :
        if row[0] == '' :
            pivot_sql = Operation + '(CASE WHEN ' + Column_Heading + ' is null THEN ' + Op_Argument + CaseElse	+ ' END) AS [NULL]'
        else :
            pivot_sql = Operation + '(CASE ' + Column_Heading + ' WHEN \'' + row[0] + '\' THEN ' + Op_Argument + CaseElse + ' END) as \'' + row[0] + '\''
#        print pivot_sql
        SQL = SQL + pivot_sql + ', '
    groupBy = sub(r'as [^,\n\r]+(,|$)', r'\1', Row_Headings)
    SQL = SQL + Operation + '(' + Op_Argument + ') as Total FROM ' + Base_Data_SQL + ' GROUP BY ' + groupBy + ' ORDER BY ' + groupBy
    print SQL
    return curs.execute(SQL)

