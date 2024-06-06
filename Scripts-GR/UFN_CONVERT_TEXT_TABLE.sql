IF EXISTS ( SELECT TOP 1
                    S.SPECIFIC_NAME
            FROM    information_schema.routines s
            WHERE   s.ROUTINE_TYPE = 'FUNCTION'
                    AND S.ROUTINE_NAME = 'UFN_CONVERT_TEXT_TABLE' )
    BEGIN
        DROP FUNCTION [dbo].[UFN_CONVERT_TEXT_TABLE]
    END
GO
/*
select * from UFN_CONVERT_TEXT_TABLE('001-1,001-2,001-3,001-4',',')
*/

CREATE FUNCTION dbo.UFN_CONVERT_TEXT_TABLE
(
    @Input NVARCHAR(MAX),
    @Delimiter CHAR(1)
)
RETURNS @OutputTable TABLE (Part1 CHAR(1), Part2 CHAR(3), Part3 INT)
WITH ENCRYPTION
AS
BEGIN
    DECLARE @Start INT, @End INT, @Segment NVARCHAR(50)
    DECLARE @FirstHyphenPos INT, @SecondHyphenPos INT
    
    SELECT @Start = 1, @End = CHARINDEX(@Delimiter, @Input)
    
    WHILE @Start <= LEN(@Input)
    BEGIN
        IF @End = 0 
            SET @End = LEN(@Input) + 1
        
        SET @Segment = SUBSTRING(@Input, @Start, @End - @Start)
        SET @FirstHyphenPos = CHARINDEX('-', @Segment)
        SET @SecondHyphenPos = CHARINDEX('-', @Segment, @FirstHyphenPos + 1)
        
        IF @FirstHyphenPos > 0 AND @SecondHyphenPos > 0
        BEGIN
            INSERT INTO @OutputTable (Part1, Part2, Part3)
            VALUES(
                SUBSTRING(@Segment, 1, @FirstHyphenPos - 1), 
                SUBSTRING(@Segment, @FirstHyphenPos + 1, @SecondHyphenPos - @FirstHyphenPos - 1), 
                CAST(SUBSTRING(@Segment, @SecondHyphenPos + 1, LEN(@Segment) - @SecondHyphenPos) AS INT)
            )
        END
        
        SET @Start = @End + 1
        SET @End = CHARINDEX(@Delimiter, @Input, @Start)
    END
    
    RETURN
END