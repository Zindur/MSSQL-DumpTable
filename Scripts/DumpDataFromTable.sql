
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[DumpDataFromTable]') AND type in (N'P', N'PC'))
    DROP PROCEDURE dbo.[DumpDataFromTable]
GO
 
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
 
-- =============================================
-- Author:    Oleg Ciobanu
-- Create date: 20171214
-- Version 1.01
-- Description:
-- dump data in 2 formats
-- @BuildMethod = 1 INSERT INTO format
-- @BuildMethod = 2 SELECT * FROM format
--
-- SQL must have permission to create files, if is not set-up then exec follow line once
-- EXEC sp_configure 'Ole Automation Procedures', 1; RECONFIGURE WITH OVERRIDE;
--
-- =============================================
CREATE PROCEDURE [dbo].[DumpDataFromTable]
(
     @SchemaName nvarchar(128) --= 'dbo'
    ,@TableName nvarchar(128) --= 'testTable'
    ,@WhereClause nvarchar (1000) = '' -- must start with AND
    ,@BuildMethod int = 1 -- taking values 1 for INSERT INTO forrmat or 2 for SELECT from value Table
    ,@PathOut nvarchar(250) = N'c:\tmp\scripts\' -- folder must exist !!!'
    ,@AsFileNAme nvarchar(250) = NULL -- if is passed then will use this value as FileName
    ,@DebugMode int = 0
)
AS
BEGIN  
    SET NOCOUNT ON;
 
        -- run follow next line if you get permission deny  for sp_OACreate,sp_OAMethod
        -- EXEC sp_configure 'Ole Automation Procedures', 1; RECONFIGURE WITH OVERRIDE;
 
    DECLARE @Sql nvarchar (max)
    DECLARE @SqlInsert nvarchar (max) = ''
    DECLARE @Columns nvarchar(max)
    DECLARE @ColumnsCast nvarchar(max)
 
    -- cleanUp/prepraring data
    SET @SchemaName = REPLACE(REPLACE(@SchemaName,'[',''),']','')
    SET @TableName = REPLACE(REPLACE(@TableName,'[',''),']','')
    SET @AsFileNAme = NULLIF(@AsFileNAme,'')
    SET @AsFileNAme = REPLACE(@AsFileNAme,'.','_')
    SET @AsFileNAme = COALESCE(@PathOut + @AsFileNAme + '.sql', @PathOut + @SchemaName + ISNULL('_' + @TableName,N'') + '.sql')
 
 
    --debug
    IF @DebugMode = 1
        print @AsFileNAme
 
        -- Create temp SP what will be responsable for generating script files
    DECLARE @PRC_WritereadFile VARCHAR(max) =
        'IF EXISTS (SELECT * FROM sys.objects WHERE type = ''P'' AND name = ''PRC_WritereadFile'')
       BEGIN
          DROP  Procedure  PRC_WritereadFile
       END;'
    EXEC  (@PRC_WritereadFile)
       -- '  
    SET @PRC_WritereadFile =
    'CREATE Procedure PRC_WritereadFile (
        @FileMode INT -- Recreate = 0 or Append Mode 1
       ,@Path NVARCHAR(1000)
       ,@AsFileNAme NVARCHAR(500)
       ,@FileBody NVARCHAR(MAX)   
       )
    AS
        DECLARE @OLEResult INT
        DECLARE @FS INT
        DECLARE @FileID INT
        DECLARE @hr INT
        DECLARE @FullFileName NVARCHAR(1500) = @Path + @AsFileNAme
     
        -- Create Object
        EXECUTE @OLEResult = sp_OACreate ''Scripting.FileSystemObject'', @FS OUTPUT
        IF @OLEResult <> 0 BEGIN
            PRINT ''Scripting.FileSystemObject''
            GOTO Error_Handler
        END    
 
        IF @FileMode = 0 BEGIN  -- Create
            EXECUTE @OLEResult = sp_OAMethod @FS,''CreateTextFile'',@FileID OUTPUT, @FullFileName
            IF @OLEResult <> 0 BEGIN
                PRINT ''CreateTextFile''
                GOTO Error_Handler
            END
        END ELSE BEGIN          -- Append
            EXECUTE @OLEResult = sp_OAMethod @FS,''OpenTextFile'',@FileID OUTPUT, @FullFileName, 8, 0 -- 8- forappending
            IF @OLEResult <> 0 BEGIN
                PRINT ''OpenTextFile''
                GOTO Error_Handler
            END            
        END
     
        EXECUTE @OLEResult = sp_OAMethod @FileID, ''WriteLine'', NULL, @FileBody
        IF @OLEResult <> 0 BEGIN
            PRINT ''WriteLine''
            GOTO Error_Handler
        END     
 
        EXECUTE @OLEResult = sp_OAMethod @FileID,''Close''
        IF @OLEResult <> 0 BEGIN
            PRINT ''Close''
            GOTO Error_Handler
        END
     
        EXECUTE sp_OADestroy @FS
        EXECUTE sp_OADestroy @FileID
     
        GOTO Done
 
        Error_Handler:
            DECLARE @source varchar(30), @desc varchar (200)       
            EXEC @hr = sp_OAGetErrorInfo null, @source OUT, @desc OUT
            PRINT ''*** ERROR ***''
            SELECT OLEResult = @OLEResult, hr = CONVERT (binary(4), @hr), source = @source, description = @desc
 
       Done:
    ';
        -- '
    EXEC  (@PRC_WritereadFile) 
    EXEC PRC_WritereadFile 0 /*Create*/, '', @AsFileNAme, ''
     
 
    ;WITH steColumns AS (
        SELECT
            1 as rn,
            c.ORDINAL_POSITION
            ,c.COLUMN_NAME as ColumnName
            ,c.DATA_TYPE as ColumnType
        FROM INFORMATION_SCHEMA.COLUMNS c
        WHERE 1 = 1
        AND c.TABLE_SCHEMA = @SchemaName
        AND c.TABLE_NAME = @TableName
    )
 
    --SELECT *
 
       SELECT
            @ColumnsCast = ( SELECT
                                    CASE WHEN ColumnType IN ('date','time','datetime2','datetimeoffset','smalldatetime','datetime','timestamp')
                                        THEN
                                            'convert(nvarchar(1001), s.[' + ColumnName + ']' + ' , 121) AS [' + ColumnName + '],'
                                            --,convert(nvarchar, [DateTimeScriptApplied], 121) as [DateTimeScriptApplied]
                                        ELSE
                                            'CAST(s.[' + ColumnName + ']' + ' AS NVARCHAR(1001)) AS [' + ColumnName + '],'
                                    END
                                     as 'data()'                                  
                                    FROM
                                      steColumns t2
                                    WHERE 1 =1
                                      AND t1.rn = t2.rn
                                    FOR xml PATH('')
                                   )
            ,@Columns = ( SELECT
                                    '[' + ColumnName + '],' as 'data()'                                  
                                    FROM
                                      steColumns t2
                                    WHERE 1 =1
                                      AND t1.rn = t2.rn
                                    FOR xml PATH('')
                                   )
 
    FROM steColumns t1
 
    -- remove last char
    IF lEN(@Columns) > 0 BEGIN
        SET @Columns = SUBSTRING(@Columns, 1, LEN(@Columns)-1);
        SET @ColumnsCast = SUBSTRING(@ColumnsCast, 1, LEN(@ColumnsCast)-1);
    END
 
    -- debug
    IF @DebugMode = 1 BEGIN
        print @ColumnsCast
        print @Columns
        select @ColumnsCast ,  @Columns
    END
 
    -- build unpivoted Data
    SET @SQL = '
    SELECT
        u.rn
        , c.ORDINAL_POSITION as ColumnPosition
        , c.DATA_TYPE as ColumnType
        , u.ColumnName
        , u.ColumnValue
    FROM
    (SELECT
        ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS rn,
    '
    + CHAR(13) + @ColumnsCast
    + CHAR(13) + 'FROM [' + @SchemaName + '].[' + @TableName + '] s'
    + CHAR(13) + 'WHERE 1 = 1'
    + CHAR(13) + COALESCE(@WhereClause,'')
    + CHAR(13) + ') tt
    UNPIVOT
    (
      ColumnValue
      FOR ColumnName in (
    ' + CHAR(13) + @Columns
    + CHAR(13)
    + '
     )
    ) u
 
    LEFT JOIN INFORMATION_SCHEMA.COLUMNS c ON c.COLUMN_NAME = u.ColumnName
        AND c.TABLE_SCHEMA = '''+ @SchemaName + '''
        AND c.TABLE_NAME = ''' + @TableName +'''
    ORDER BY u.rn
            , c.ORDINAL_POSITION
    '
 
    -- debug
    IF @DebugMode = 1 BEGIN
        print @Sql     
    END
 
    EXEC (@Sql)
 
    -- prepare data for cursor
 
    IF OBJECT_ID('tempdb..#tmp') IS NOT NULL
        DROP TABLE #tmp
    CREATE TABLE #tmp
    (
        rn bigint
        ,ColumnPosition int
        ,ColumnType varchar (128)
        ,ColumnName varchar (128)
        ,ColumnValue nvarchar (2000) -- I hope this size will be enough for storring values
    )
    SET @Sql = 'INSERT INTO  #tmp ' + CHAR(13)  + @Sql
 
    -- debug
    IF @DebugMode = 1 BEGIN
        print @Sql
    END
 
    EXEC (@Sql)
 
    IF @DebugMode = 1 BEGIN
        SELECT * FROM #tmp
    END
 
    DECLARE @rn bigint
        ,@ColumnPosition int
        ,@ColumnType varchar (128)
        ,@ColumnName varchar (128)
        ,@ColumnValue nvarchar (2000)
        ,@i int = -1 -- counter/flag
        ,@ColumnsInsert varchar(max) = NULL
        ,@ValuesInsert nvarchar(max) = NULL
 
    DECLARE cur CURSOR FOR
    SELECT rn, ColumnPosition, ColumnType, ColumnName, ColumnValue
    FROM #tmp
    ORDER BY rn, ColumnPosition -- note order is really important !!!
    OPEN cur
 
    FETCH NEXT FROM cur
    INTO @rn, @ColumnPosition, @ColumnType, @ColumnName, @ColumnValue
 
    IF @BuildMethod = 1
    BEGIN
    	SET @SqlInsert = 'SET NOCOUNT ON;' + CHAR(13);
		EXEC PRC_WritereadFile 1 /*Add*/, '', @FileName, @SqlInsert
        SET @SqlInsert = ''
    END
    ELSE BEGIN
	    SET @SqlInsert = 'SET NOCOUNT ON;' + CHAR(13);
		SET @SqlInsert = @SqlInsert
						+ 'SELECT *'
                        + CHAR(13) + 'FROM ('
                        + CHAR(13) + 'VALUES'
        EXEC PRC_WritereadFile 1 /*Add*/, '', @AsFileNAme, @SqlInsert
        SET @SqlInsert = NULL
    END
 
    SET @i = @rn
 
    WHILE @@FETCH_STATUS = 0
    BEGIN
     
        IF (@i <> @rn) -- is a new row
        BEGIN
            IF @BuildMethod = 1
            -- build as INSERT INTO -- as Default
            BEGIN
                SET @SqlInsert = 'INSERT INTO [' + @SchemaName + '].[' + @TableName + '] ('
                                + CHAR(13) + @ColumnsInsert + ')'
                                + CHAR(13) + 'VALUES ('
                                + @ValuesInsert
                                + CHAR(13) + ');'
            END
            ELSE
            BEGIN
                -- build as Table select
                IF (@i <> @rn) -- is a new row
                BEGIN
                    SET @SqlInsert = COALESCE(@SqlInsert + ',','') +  '(' + @ValuesInsert+ ')'
                    EXEC PRC_WritereadFile 1 /*Add*/, '', @AsFileNAme, @SqlInsert
                    SET @SqlInsert = '' -- in method 2 we should clear script
                END            
            END
            -- debug
            IF @DebugMode = 1
                PRINT @SqlInsert
            EXEC PRC_WritereadFile 1 /*Add*/, '', @AsFileNAme, @SqlInsert
 
            -- we have new row
            -- initialise variables
            SET @i = @rn
            SET @ColumnsInsert = NULL
            SET @ValuesInsert = NULL
        END
 
        -- build insert values
        IF (@i = @rn) -- is same row
        BEGIN
            SET @ColumnsInsert = COALESCE(@ColumnsInsert + ',','') + '[' + @ColumnName + ']'
            SET @ValuesInsert =  CASE                              
                                    -- date
                                    --WHEN
                                    --  @ColumnType IN ('date','time','datetime2','datetimeoffset','smalldatetime','datetime','timestamp')
                                    --THEN
                                    --  COALESCE(@ValuesInsert + ',','') + '''''' + ISNULL(RTRIM(@ColumnValue),'NULL') + ''''''
                                    -- numeric
                                    WHEN
                                        @ColumnType IN ('bit','tinyint','smallint','int','bigint'
                                                        ,'money','real','','float','decimal','numeric','smallmoney')
                                    THEN
                                        COALESCE(@ValuesInsert + ',','') + '' + ISNULL(RTRIM(@ColumnValue),'NULL') + ''
                                    -- other types treat as string
                                    ELSE
										COALESCE(@ValuesInsert + ',','') + '''' + ISNULL(RTRIM( 
																							-- escape single quote
																							REPLACE(@ColumnValue, '''', '''''') 
																							  ),'NULL') + ''''		   
                                END
        END
 
 
        FETCH NEXT FROM cur
        INTO @rn, @ColumnPosition, @ColumnType, @ColumnName, @ColumnValue
 
        -- debug
        IF @DebugMode = 1
        BEGIN
            print CAST(@rn AS VARCHAR) + '-' + CAST(@ColumnPosition AS VARCHAR)
        END
    END
    CLOSE cur
    DEALLOCATE cur
 
    IF @BuildMethod = 1
    BEGIN
        PRINT 'ignore'
    END
    ELSE BEGIN
        SET @SqlInsert = CHAR(13) + ') AS vtable '
                        + CHAR(13) + ' (' + @Columns
                        + CHAR(13) + ')'
        EXEC PRC_WritereadFile 1 /*Add*/, '', @AsFileNAme, @SqlInsert
        SET @SqlInsert = NULL
    END
END
