USE [VISTA]
GO
/****** Object:  StoredProcedure [dbo].[sp_Catalogo_ObtenerCompleto]    Script Date: 10/07/2026 02:16:23 p. m. ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =====================================================================
-- sp_Catalogo_ObtenerCompleto
-- Regresa 3 result sets: Tablas visibles (con EsHub), sus Campos (con
-- EsCalculado/ExpresionCalculada), y las Relaciones (con
-- TipoJoin/CondicionExtra).
-- =====================================================================
ALTER   PROCEDURE [dbo].[sp_Catalogo_ObtenerCompleto]
AS
BEGIN
    SET NOCOUNT ON;

    SELECT Id, NombreTabla, Alias, Visible, EsHub
    FROM dbo.Rep_CatTablas
    WHERE Visible = 1
    ORDER BY Alias;

    SELECT cc.Id, cc.TablaId, cc.NombreColumna, cc.Alias, cc.TipoDatoSql,
           cc.TipoFiltroUI, cc.EsLlaveJoin, cc.EsLlavePrimaria,
           cc.EsCalculado, cc.ExpresionCalculada
    FROM dbo.Rep_CatCampos cc
    JOIN dbo.Rep_CatTablas ct ON ct.Id = cc.TablaId
    WHERE ct.Visible = 1
    ORDER BY cc.TablaId, cc.Alias;

    SELECT Id, TablaOrigenId, CampoOrigen, TablaDestinoId, CampoDestino,
           TipoJoin, CondicionExtra
    FROM dbo.Rep_CatRelaciones;
END
