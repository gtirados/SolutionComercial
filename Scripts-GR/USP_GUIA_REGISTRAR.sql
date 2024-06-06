IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'USP_GUIA_REGISTRAR'
)
BEGIN
    DROP PROC [dbo].[USP_GUIA_REGISTRAR];
END;
GO
/*
USP_GUIA_REGISTRAR '01','T001',1,'<r><d idp="123" cant="3" /></r>'
*/
CREATE PROCEDURE [dbo].[USP_GUIA_REGISTRAR]
    @CODCIA CHAR(2),
    @IDCLIENTE BIGINT,
    @SERIE CHAR(4),
    @NUMERO BIGINT,
    @FECHAEMISION CHAR(8),
    @FECHATRASLADO CHAR(8),
    @CODIGOMOTIVOTRASLADO CHAR(2),
    @CODIGOMODALIDADTRASLADO CHAR(2),
    @IDTRASPORTISTA INT,
    @PESO DECIMAL(6, 2),
    @BULTOS INT,
    @PRODUCTOS NVARCHAR(MAX),
    @DESTINATARIOTIPO CHAR(1) = NULL,
    @DESTINATARIONUMERO VARCHAR(11) = NULL,
    @DESTINATARIORS VARCHAR(100) = NULL
WITH ENCRYPTION
AS
SET NOCOUNT ON;
DECLARE @EXITO VARCHAR(300) = '0=Guia registrada correctamente.';

BEGIN TRAN;
BEGIN TRY

    INSERT INTO dbo.GUIA_REMISION_CAB
    (
        CODCIA,
        IDCLIENTE,
        SERIE,
        NUMERO,
        FECHAEMISION,
        HORAEMISION,
        FECHATRASLADO,
        CODIGO_MOTIVO_TRASLADO,
        CODIGO_MODALIDAD_TRASLADO,
        IDTRANSPORTISTA,
        DESTINATARIO_TIPO,
        DESTINATARIO_NUMERO,
        DESTINATARIO_RAZONSOCIAL,
        PESO_TOTAL,
        BULTOS
    )
    VALUES
    (   @CODCIA,                  -- CODCIA - char(2)
        @IDCLIENTE, @SERIE,       -- SERIE - char(4)
        @NUMERO,                  -- NUMERO - bigint
        @FECHAEMISION,            -- FECHAEMISION - date
        CONVERT(TIME, GETDATE()), -- HORAEMISION - time(7)
        @FECHATRASLADO,           -- FECHATRASLADO - date
        @CODIGOMOTIVOTRASLADO,    -- CODIGO_MOTIVO_TRASLADO - char(2)
        @CODIGOMODALIDADTRASLADO, -- CODIGO_MODALIDAD_TRASLADO - char(2)
        @IDTRASPORTISTA,          -- IDTRANSPORTISTA - int
        @DESTINATARIOTIPO,        -- DESTINATARIO_TIPO - int
        @DESTINATARIONUMERO,      -- DESTINATARIO_NUMERO - varchar(11)
        @DESTINATARIORS,          -- DESTINATARIO_RAZONSOCIAL - varchar(100)
        @PESO,                    -- PESO_TOTAL - decimal(6, 2)
        @BULTOS                   -- BULTOS - int
        );

    DECLARE @idoc INT;


    EXEC sp_xml_preparedocument @idoc OUTPUT, @PRODUCTOS;

    INSERT INTO dbo.GUIA_REMISION_DET
    (
        CODCIA,
        SERIE,
        NUMERO,
        IDPRODUCTO,
        CANTIDAD
    )
    SELECT @CODCIA,
           @SERIE,
           @NUMERO,
           idp,
           cant
    FROM
        OPENXML(@idoc, '/r/d', 1)WITH (idp BIGINT, cant NUMERIC(9, 2));

    EXEC sp_xml_removedocument @idoc;


    UPDATE dbo.NUMERACION_DOCUMENTOS
    SET NUMERO_ACTUAL = NUMERO_ACTUAL + 1
    WHERE CODCIA = @CODCIA
          AND TIPO_DOCUMENTO = 'G';



--INSERT INTO dbo.GUIA_REMISION_DET
--(
--    CODCIA,
--    SERIE,
--    NUMERO,
--    IDPRODUCTO,
--    CANTIDAD
--)
--SELECT @CODCIA,@SERIE,@NUMERO,p.idproducto,p.cantidad FROM @PRODUCTOS p
END TRY
BEGIN CATCH
    SET @EXITO = RTRIM(LTRIM(STR(ERROR_NUMBER()))) + '=' + ERROR_MESSAGE();
    ROLLBACK TRAN;
    GOTO Terminar;
END CATCH;


IF @@TRANCOUNT > 0
    COMMIT;

Terminar:
SELECT @EXITO;
GO