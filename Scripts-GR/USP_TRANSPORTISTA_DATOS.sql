IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'USP_TRANSPORTISTA_DATOS'
)
BEGIN
    DROP PROC [dbo].[USP_TRANSPORTISTA_DATOS];
END;
GO
/*
USP_TRANSPORTISTA_DATOS '01',13
*/
CREATE PROCEDURE [dbo].[USP_TRANSPORTISTA_DATOS]
    @CODCIA CHAR(2),
    @IDTRANSPORTE INT
WITH ENCRYPTION
AS
SET NOCOUNT ON;

SELECT RTRIM(LTRIM(t.TRN_DNI)) AS 'DNI',
       t.TRN_NOMBRE AS 'NOMBRES',
       t.TRN_BREVETE AS 'LICENCIA',
       t.TRN_PLACA AS 'PLACA'
FROM dbo.TRANSPORTE t
WHERE t.TRN_CODCIA = @CODCIA
      AND t.TRN_KEY = @IDTRANSPORTE;

GO