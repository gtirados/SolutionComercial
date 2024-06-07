IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'USP_GUIA_REMISION_FILL_CAB'
)
BEGIN
    DROP PROC [dbo].[USP_GUIA_REMISION_FILL_CAB];
END;
GO
/*
USP_GUIA_REMISION_FILL_CAB'20240101','20241201','01'

*/

CREATE PROCEDURE [dbo].[USP_GUIA_REMISION_FILL_CAB]
    @INI CHAR(8),
    @FIN CHAR(8),
    @CODCIA CHAR(2)
WITH ENCRYPTION
AS
SET NOCOUNT ON;

SELECT grc.SERIE,
       grc.NUMERO,
       CONVERT(VARCHAR(10), grc.FECHAEMISION, 103) AS 'FECHA',
       RTRIM(LTRIM(c.CLI_NOMBRE)) AS 'CLIENTE',
       grc.PESO_TOTAL AS 'PESOTOTAL',
	   COALESCE(grc.respuesta_sunat_descripcion,'') AS 'RPTASUNAT'
FROM dbo.GUIA_REMISION_CAB grc
    INNER JOIN dbo.CLIENTES c
        ON grc.IDCLIENTE = c.CLI_CODCLIE
           AND grc.CODCIA = c.CLI_CODCIA
WHERE grc.FECHAEMISION
BETWEEN @INI AND @FIN
ORDER BY grc.SERIE,
         grc.NUMERO;

GO