IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'USP_VENTAS_FILL_CAB'
)
BEGIN
    DROP PROC [dbo].[USP_VENTAS_FILL_CAB];
END;
GO
/*
USP_VENTAS_FILL 57339,'20240101','20241201','01'
SELECT * FROM dbo.CLIENTES c WHERE c.CLI_CODCLIE=57339
*/

CREATE PROCEDURE [dbo].[USP_VENTAS_FILL_CAB]
    @IDCLIENTE BIGINT,
    @INI CHAR(8),
    @FIN CHAR(8),
    @CODCIA CHAR(2)
WITH ENCRYPTION
AS
SET NOCOUNT ON;

SELECT RIGHT('000' + RTRIM(LTRIM(a.ALL_NUMSER)), 3) AS 'SERIE',
       a.ALL_NUMFAC AS 'NUMERO',
       a.ALL_FBG AS 'TIPO',
       CONVERT(CHAR(10), a.ALL_FECHA_DIA, 103) AS 'FECHA',
       a.ALL_NETO AS 'TOTAL'
FROM dbo.ALLOG a
WHERE a.ALL_CODCIA = @CODCIA
      AND CONVERT(CHAR(8), a.ALL_FECHA_DIA, 112)
      BETWEEN @INI AND @FIN
      AND a.ALL_CODCLIE = @IDCLIENTE
      AND a.ALL_CODTRA = 2401
      AND a.ALL_FLAG_EXT = 'N'
ORDER BY a.ALL_NUMFAC;

GO