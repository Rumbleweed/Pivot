-- Pivot Table stored procedure for MS SQL
-- Get pivot table results in a SQL command line interface. For example
-- p_Transform 'SUM', '[amt]', 'mediaSales', '[year]', '[type]'

if exists (select * from dbo.sysobjects 
	where id = object_id('p_Transform') and
	OBJECTPROPERTY(id, 'IsProcedure') = 1)
drop procedure p_Transform
go
-- Pivot Table stored procedure for MS SQL
-- Description:
-- Generates simple pivot tables on data, and outputs the final SQL statement via Print.
-- The syntax is as close to the Access transform...pivot model as possible.

-- For example, given data as shown below:
--
-- ID          year        type        amt         
----------- ----------- ----------- ----------- 
-- 7           1999        Vinyl       23
-- 8           1999        Tape        44
-- 9           1999        CD          55
-- 10          2000        Vinyl       66
-- 11          2000        Tape        77
-- 12          2000        CD          88
-- 13          1999        Vinyl       11
--
--
-- ... you can pivot the data to show the years down the side 
-- and the types across the top...
--
-- year        Vinyl      Tape         CD          Total    
----------- ----------- ----------- ----------- ----------- 
-- 1999        34          44          55          133
-- 2000        66          77          88          231
--
-- ... and get the SQL which would produce this table
-- 	SELECT Pivot_Data.*
-- 	FROM (SELECT [year],  
-- 	SUM(CASE [type] WHEN 'Vinyl' THEN [amt] END) AS [Vinyl],  
-- 	SUM(CASE [type] WHEN 'Tape' THEN [amt] END) AS [Tape],  
-- 	SUM(CASE [type] WHEN 'CD' THEN [amt] END) AS [CD]
-- 	 FROM mediaSales 
-- 	GROUP BY [year]) AS Pivot_Data ORDER BY [year]
-- 
-- 
-- .. by calling the procedure as shown below:
--
-- exec p_Transform 'SUM', '[amt]', 'mediaSales', '[year]', '[type]', 'Y', 'Y'
--
-- SUM column [amt]
-- from Base_Data_SQL mediaSales
-- group by year
-- use the type column data values as headings (this is the pivot column)
-- order by expression (this is optional, default no ordering)
-- include a summary for the row (this is optional, default = 'N')

-- Procedure 
-- Arguments:
-- 	Operation		SUM, PRODUCT, COUNT, etc
-- 	Op_Argument		Column to use as argument to operation
-- 	Base_Data_SQL		SQL that returns data to be summarized
-- 	Row_Headings		Comma-separated list of rows to use as groupings of data
-- 	Column_Heading		Column to use as heading (ie. pivot data)
-- 	Order_By_Grouping	Optional expression to order the results
-- 	Row_Total		Optional indication of whether to include row total in results
-- 
-- Steps in Routine: 
-- 	1. Get list of distinct column headings
-- 	2. Looping through column headings, alter SQL string for pivot
-- 	3. Add summary, ordering SQL if required
-- 	4. Execute
-- 
-- 
-- History:
-- Jeff Zohrab		Aug 13, 2001		Initial release
-- S. Anderson		Feb 13, 2002		Rearranged arguments, removed
-- 						column_head_sql, made row summary
-- 						optional, removed else in case
-- 						to make count work properly, allowed
-- 						simple table name for base sql.
-- S. Anderson		May 5, 2003		Made case else depend on operation,
-- 						0 for sum, product, no else otherwise.
-- 						Use cursor for loop instead of lame loop.
-- Anderson   Feb 2, 2016   Put into github.
CREATE   Procedure p_Transform
		@Operation		varchar(10),  	-- SUM, PRODUCT, count, etc
		@Op_Argument		varchar(255),  	-- Column to use as argument in operation
		@Base_Data_SQL		varchar(2000),  -- Table to use as recordsource to build final crosstab qry
		@Row_Headings		varchar(255),  	-- Comma-separated list of rows to use as groupings of data
		@Column_Heading		varchar(255),  	-- Column to use as heading, ie. pivot column
		@Order_By_Row		varchar(255) = NULL,	-- to order by the row groupings, put in arg 4
							-- you can also order by RowTotal, and ASC, DESC also work
		@Add_Row_Summary	char(1)='N'	-- 'Y' to include summary, 'N' to omit
AS 
   Declare @Column_Head_SQL	varchar(8000)
   Declare @SQL 		varchar(8000)
--   Declare @Summary_SQL 	varchar(8000)		-- to summarize each row

   set nocount on

   IF @Operation NOT IN ('SUM', 'COUNT', 'MAX', 'MIN', 'AVG', 'STDEV', 'VAR', 'VARP', 'STDEVP')    
   BEGIN RAISERROR ('Invalid aggregate function: %s', 10, 1, @Operation) END  
   ELSE  
   BEGIN  
   SET @SQL = 'SELECT ' + @Row_Headings + ', '

   -- if base table is a select statement, make into subquery
   if SUBSTRING(@Base_Data_SQL, 1, 6) in ('select')
		Set @Base_Data_SQL = '(' + @Base_Data_SQL + ') AS Base_Data '

   Set @Column_Head_SQL = 'SELECT DISTINCT ' + @Column_Heading + ' from ' + @Base_Data_SQL
   --print @Column_Head_SQL -- debug check

   -- Get list of distinct column headings
   CREATE TABLE #Col_Heads (
		Col_ID int identity(1,1),
		Col_Head varchar(200) NULL
		)
   Exec ('INSERT INTO #Col_Heads(Col_Head) ' + @Column_Head_SQL)

   --select * from #Col_Heads -- debug check

   DECLARE @Col_ID_Curr int,				-- column being checked
		@Curr_Col_Head	varchar(200),
		@Pivot_SQL varchar(200),		-- pivot SQL for current column
		@CaseElse varchar(10)

   declare csrColHead cursor for
			SELECT Col_ID, Col_Head 
			FROM #Col_Heads
			ORDER BY Col_ID ASC

   -- One doesn't want to count nulls, but one want to sum them as 0.
   if(@Operation like 'sum') set @CaseElse = ' else 0'
   else set @CaseElse = ''

   -- loop through all columns, build pivot strings
   open csrColHead
   fetch next from csrColHead into @Col_ID_Curr, @Curr_Col_Head
   while (@@fetch_status = 0) 
   begin
		--print 'Adding pivot line for heading ' + @Curr_Col_Head -- debug check

		if (@Curr_Col_Head is null)
			Set @Pivot_SQL = char(13) + @Operation
			+ '(CASE WHEN ' + @Column_Heading 
			+ ' is null THEN ' + @Op_Argument 
			+ @CaseElse
			+ ' END) AS [NULL]'
		else
			Set @Pivot_SQL = char(13) + @Operation
			+ '(CASE ' + @Column_Heading 
			+ ' WHEN ''' + @Curr_Col_Head + ''' THEN ' + @Op_Argument 
			+ @CaseElse
			+ ' END) AS [' + @Curr_Col_Head +']'
		
		Set @SQL = @SQL + ' ' + @Pivot_SQL

		-- Get the next column head
		fetch next from csrColHead into @Col_ID_Curr, @Curr_Col_Head
		if(@@fetch_status = 0)
		begin
			Set @SQL = @SQL  + ', '
		end
   end 
   close csrColHead
   deallocate csrColHead

   -- release objects
   DROP TABLE #Col_Heads

   -- Finish SQL
   If (@Add_Row_Summary='Y')
		Set @SQL = @SQL + char(13) + ', ' + @Operation + '(' + @Op_Argument + ') as Total FROM ' +  @Base_Data_SQL 
			+ char(13) + 'GROUP BY ' + @Row_Headings
   else
		Set @SQL = @SQL + char(13) + ' FROM ' +  @Base_Data_SQL 
			+ char(13) + 'GROUP BY ' + @Row_Headings

   -- Order by at the end of it all
   If (@Order_By_Row is not null)
		Set @SQL = @SQL + ' ORDER BY ' + @Order_By_Row

   -- Done
   Print @SQL
   Exec (@SQL)
   END
go
