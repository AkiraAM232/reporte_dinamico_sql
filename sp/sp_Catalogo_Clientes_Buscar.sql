USE [VISTA]
GO
/****** Object:  StoredProcedure [dbo].[sp_Catalogo_Clientes_Buscar]    Script Date: 10/07/2026 02:22:40 p. m. ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




-- sp_Catalogo_Clientes_Buscar
-- Catálogo de RFCs para el selector del front (Perfiles..Tb_Clientes).
-- Recibe un texto de búsqueda libre (RFC o Razón Social) y regresa
-- máximo @Top filas para que el combo no cargue todos los clientes de
-- un jalón. Solo clientes Activo = 1.
-- =====================================================================
ALTER PROCEDURE [dbo].[sp_Catalogo_Clientes_Buscar]
    @Busqueda NVARCHAR(200) = NULL,
    @Top      INT           = 50
AS
BEGIN
    SET NOCOUNT ON;

    IF @Top IS NULL OR @Top < 1 OR @Top > 100
        SET @Top = 50;

    -- Normaliza vacío a NULL para que el filtro de abajo se salte limpio
    IF @Busqueda IS NOT NULL AND LTRIM(RTRIM(@Busqueda)) = N''
        SET @Busqueda = NULL;

    SELECT TOP (@Top)
        c.RFC,
        c.Razon_Social,
        c.Industria,
        Identificador = c.RFC + N' - ' + c.Razon_Social
    FROM PERFILES.dbo.Tb_Clientes c
    WHERE c.Activo = 1
      AND (
            @Busqueda IS NULL
            OR c.RFC          LIKE N'%' + @Busqueda + N'%'
            OR c.Razon_Social LIKE N'%' + @Busqueda + N'%'
          )
    ORDER BY c.Razon_Social;
END
