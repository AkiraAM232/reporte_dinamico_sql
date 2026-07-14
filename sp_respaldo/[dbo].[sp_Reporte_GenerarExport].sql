USE [VISTA]
GO
/****** Object:  StoredProcedure [dbo].[sp_Reporte_GenerarExport]    Script Date: 14/07/2026 02:12:31 p. m. ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =====================================================================
-- sp_Reporte_GenerarExport
-- Regresa como máximo los primeros 100 registros (TOP 100) — así lo
-- pidieron: no se necesita el listado completo, solo un vistazo
-- rápido de los primeros resultados, tanto en pantalla como en el
-- Excel exportado. El API lo sigue leyendo con streaming
-- (SqlDataReader) por diseño, aunque con 100 filas como máximo eso ya
-- no es estrictamente necesario — se deja así por si el tope cambia
-- más adelante.
-- =====================================================================
ALTER   PROCEDURE [dbo].[sp_Reporte_GenerarExport]
    @CampoIds     NVARCHAR(MAX),
    @FiltrosJson  NVARCHAR(MAX) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    ------------------------------------------------------------------
    -- 1) Resolver columnas pedidas contra el catálogo (whitelist)
    ------------------------------------------------------------------
    DECLARE @CampoIdsTabla TABLE (CampoId INT, Orden INT IDENTITY(1,1));
    INSERT INTO @CampoIdsTabla (CampoId)
    SELECT TRY_CAST(value AS INT) FROM STRING_SPLIT(@CampoIds, ',');

    IF EXISTS (SELECT 1 FROM @CampoIdsTabla WHERE CampoId IS NULL)
    BEGIN
        RAISERROR('@CampoIds contiene un valor que no es un entero válido.', 16, 1);
        RETURN;
    END

    SELECT
        p.Orden,
        cc.Id            AS CampoId,
        cc.NombreColumna,
        cc.Alias,
        ct.Id            AS TablaId,
        ct.NombreTabla,
        cc.EsCalculado,
        cc.ExpresionCalculada,
        ColumnaSql = CASE
            WHEN cc.EsCalculado = 1 THEN cc.ExpresionCalculada
            ELSE N'dbo.[' + ct.NombreTabla + N'].[' + cc.NombreColumna + N']'
        END
    INTO #Seleccion
    FROM @CampoIdsTabla p
    JOIN dbo.Rep_CatCampos cc ON cc.Id = p.CampoId
    JOIN dbo.Rep_CatTablas ct ON ct.Id = cc.TablaId AND ct.Visible = 1;

    IF (SELECT COUNT(*) FROM #Seleccion) <> (SELECT COUNT(*) FROM @CampoIdsTabla)
    BEGIN
        RAISERROR('Una o más columnas no existen en el catálogo o pertenecen a una tabla no visible.', 16, 1);
        RETURN;
    END

    ------------------------------------------------------------------
    -- 2) Tablas involucradas + expansión por grupo obligatorio
    ------------------------------------------------------------------
    SELECT DISTINCT s.TablaId, ct.NombreTabla, ct.EsHub
    INTO #TablasUsadas
    FROM #Seleccion s
    JOIN dbo.Rep_CatTablas ct ON ct.Id = s.TablaId;

    -- Un grupo solo se expande si su tabla HUB (ej. SO130120) ya está
    -- en juego — nunca por la sola presencia de un miembro como
    -- VU150501, que también es el hub del star original.
    ;WITH GruposActivos AS (
        SELECT DISTINCT gh.NombreGrupo
        FROM dbo.Rep_CatGrupoObligatorio gh
        JOIN #TablasUsadas t ON t.TablaId = gh.TablaId
        WHERE gh.EsHubDelGrupo = 1
    )
    INSERT INTO #TablasUsadas (TablaId, NombreTabla, EsHub)
    SELECT ct.Id, ct.NombreTabla, ct.EsHub
    FROM dbo.Rep_CatGrupoObligatorio g
    JOIN GruposActivos ga ON ga.NombreGrupo = g.NombreGrupo
    JOIN dbo.Rep_CatTablas ct ON ct.Id = g.TablaId
    WHERE NOT EXISTS (SELECT 1 FROM #TablasUsadas t2 WHERE t2.TablaId = g.TablaId);

    DECLARE @TotalTablas INT = (SELECT COUNT(*) FROM #TablasUsadas);

    ------------------------------------------------------------------
    -- 2b) Requisitos de filtro obligatorio por grupo (ej. RFC, rango
    -- de fechas). Aquí NO se aplica el tope de 100 filas — el export
    -- está pensado para traer todo el resultado, vía streaming.
    --
    -- Un requisito se da por cumplido con Valor (Igual/Contiene) O con
    -- ValorDesde+ValorHasta (Rango, ej. fechas) — mismo criterio que
    -- sp_Reporte_Generar.
    ------------------------------------------------------------------
    DECLARE @CamposFaltantes NVARCHAR(MAX);

    SELECT @CamposFaltantes = STRING_AGG(
        CAST(req.CampoCatalogoId AS NVARCHAR(20)) + N' (' + cc.Alias + N')', N', '
    )
    FROM dbo.Rep_CatGrupoObligatorio g
    JOIN #TablasUsadas t ON t.TablaId = g.TablaId
    JOIN dbo.Rep_CatGrupoRequisitos req ON req.NombreGrupo = g.NombreGrupo
    JOIN dbo.Rep_CatCampos cc ON cc.Id = req.CampoCatalogoId
    WHERE NOT EXISTS (
        SELECT 1 FROM OPENJSON(@FiltrosJson) WITH (
            CampoId     INT           '$.CampoId',
            Valor       NVARCHAR(500) '$.Valor',
            ValorDesde  NVARCHAR(500) '$.ValorDesde',
            ValorHasta  NVARCHAR(500) '$.ValorHasta'
        ) f
        WHERE f.CampoId = req.CampoCatalogoId
          AND (
                (f.Valor IS NOT NULL AND f.Valor <> '')
                OR (f.ValorDesde IS NOT NULL AND f.ValorDesde <> '' AND f.ValorHasta IS NOT NULL AND f.ValorHasta <> '')
              )
    );

    IF @CamposFaltantes IS NOT NULL
    BEGIN
        RAISERROR('Este tipo de reporte requiere que captures como filtro (con valor) el/los CampoCatalogoId: %s.', 16, 1, @CamposFaltantes);
        RETURN;
    END

    ------------------------------------------------------------------
    -- 3) Armar FROM + JOINs contra el hub del grupo
    ------------------------------------------------------------------
    DECLARE @From NVARCHAR(MAX);
    DECLARE @HubTablaId INT, @HubNombreTabla NVARCHAR(128);

    IF @TotalTablas = 1
    BEGIN
        SELECT @From = N'dbo.[' + NombreTabla + N']' FROM #TablasUsadas;
    END
    ELSE
    BEGIN
        -- Primero: ¿alguna tabla activa pertenece a un grupo con hub
        -- propio marcado?
        SELECT TOP 1 @HubTablaId = ct.Id, @HubNombreTabla = ct.NombreTabla
        FROM #TablasUsadas t
        JOIN dbo.Rep_CatGrupoObligatorio g ON g.TablaId = t.TablaId
        JOIN dbo.Rep_CatGrupoObligatorio gh
            ON gh.NombreGrupo = g.NombreGrupo AND gh.EsHubDelGrupo = 1
        JOIN dbo.Rep_CatTablas ct ON ct.Id = gh.TablaId;

        IF @HubTablaId IS NULL
        SELECT @HubTablaId = TablaId, @HubNombreTabla = NombreTabla
        FROM #TablasUsadas WHERE EsHub = 1;

        IF @HubTablaId IS NULL
        BEGIN
            RAISERROR('Las tablas seleccionadas no comparten un grupo válido (falta la tabla central / hub).', 16, 1);
            RETURN;
        END

        SET @From = N'dbo.[' + @HubNombreTabla + N']';

        SELECT @From = @From + N'
' + r.TipoJoin + N' JOIN dbo.[' + t.NombreTabla + N'] ON dbo.[' + @HubNombreTabla + N'].[' + r.CampoDestino + N'] = dbo.[' + t.NombreTabla + N'].[' + r.CampoOrigen + N']' + ISNULL(N' ' + r.CondicionExtra, N'')
        FROM #TablasUsadas t
        JOIN dbo.Rep_CatRelaciones r
            ON r.TablaOrigenId = t.TablaId AND r.TablaDestinoId = @HubTablaId
        WHERE t.TablaId <> @HubTablaId;

        IF EXISTS (
            SELECT 1 FROM #TablasUsadas t
            WHERE t.TablaId <> @HubTablaId
              AND NOT EXISTS (
                  SELECT 1 FROM dbo.Rep_CatRelaciones r
                  WHERE r.TablaOrigenId = t.TablaId AND r.TablaDestinoId = @HubTablaId
              )
        )
        BEGIN
            RAISERROR('Hay una tabla seleccionada sin relación definida hacia la tabla central de su grupo.', 16, 1);
            RETURN;
        END
    END

    ------------------------------------------------------------------
    -- 4) SELECT list
    ------------------------------------------------------------------
    DECLARE @Select NVARCHAR(MAX);
    SELECT @Select = STRING_AGG(
        CAST(ColumnaSql + N' AS [' + REPLACE(Alias, ']', ']]') + N']' AS NVARCHAR(MAX)),
        N',
    '
    ) WITHIN GROUP (ORDER BY Orden)
    FROM #Seleccion;

    ------------------------------------------------------------------
    -- 5) WHERE a partir de @FiltrosJson (SIN concatenar valores)
    ------------------------------------------------------------------
    IF OBJECT_ID('tempdb..#Filtros') IS NOT NULL DROP TABLE #Filtros;

    SELECT
        f.CampoId,
        f.Operador
    INTO #Filtros
    FROM OPENJSON(@FiltrosJson) WITH (
        CampoId   INT           '$.CampoId',
        Operador  NVARCHAR(20)  '$.Operador'
    ) f
    WHERE @FiltrosJson IS NOT NULL;

    IF EXISTS (SELECT 1 FROM #Filtros f WHERE NOT EXISTS (SELECT 1 FROM #Seleccion s WHERE s.CampoId = f.CampoId))
    BEGIN
        RAISERROR('Solo se puede filtrar por columnas que ya están seleccionadas en el reporte.', 16, 1);
        RETURN;
    END

    DECLARE @Where NVARCHAR(MAX) = N'1 = 1';

    SELECT @Where = @Where + N'
  AND ' +
        CASE f.Operador
            WHEN 'Igual' THEN
                s.ColumnaSql + N' = ' +
                N'(SELECT JSON_VALUE(value, ''$.Valor'') FROM OPENJSON(@FiltrosJson) WHERE JSON_VALUE(value, ''$.CampoId'') = ''' + CAST(f.CampoId AS NVARCHAR(20)) + N''')'
            WHEN 'Contiene' THEN
                s.ColumnaSql + N' LIKE ' +
                N'''%'' + (SELECT JSON_VALUE(value, ''$.Valor'') FROM OPENJSON(@FiltrosJson) WHERE JSON_VALUE(value, ''$.CampoId'') = ''' + CAST(f.CampoId AS NVARCHAR(20)) + N''') + ''%'''
            WHEN 'Rango' THEN
                N'((SELECT JSON_VALUE(value, ''$.ValorDesde'') FROM OPENJSON(@FiltrosJson) WHERE JSON_VALUE(value, ''$.CampoId'') = ''' + CAST(f.CampoId AS NVARCHAR(20)) + N''') IS NULL OR ' + s.ColumnaSql + N' >= (SELECT JSON_VALUE(value, ''$.ValorDesde'') FROM OPENJSON(@FiltrosJson) WHERE JSON_VALUE(value, ''$.CampoId'') = ''' + CAST(f.CampoId AS NVARCHAR(20)) + N'''))
  AND ((SELECT JSON_VALUE(value, ''$.ValorHasta'') FROM OPENJSON(@FiltrosJson) WHERE JSON_VALUE(value, ''$.CampoId'') = ''' + CAST(f.CampoId AS NVARCHAR(20)) + N''') IS NULL OR ' + s.ColumnaSql + N' <= (SELECT JSON_VALUE(value, ''$.ValorHasta'') FROM OPENJSON(@FiltrosJson) WHERE JSON_VALUE(value, ''$.CampoId'') = ''' + CAST(f.CampoId AS NVARCHAR(20)) + N'''))'
            WHEN 'EnLista' THEN
                s.ColumnaSql + N' IN (SELECT TRIM(value) FROM OPENJSON((SELECT JSON_VALUE(value, ''$.Valor'') FROM OPENJSON(@FiltrosJson) WHERE JSON_VALUE(value, ''$.CampoId'') = ''' + CAST(f.CampoId AS NVARCHAR(20)) + N'''), ''$'') WITH (value NVARCHAR(500) ''$''))'
            ELSE N'1 = 1'
        END
    FROM #Filtros f
    JOIN #Seleccion s ON s.CampoId = f.CampoId;

    ------------------------------------------------------------------
    -- 6) SELECT único con TOP 100 — ya no se exporta el listado
    -- completo, solo los primeros 100 registros (mismo criterio que
    -- la vista previa). DISTINCT por el mismo motivo que en
    -- sp_Reporte_Generar: el grupo obligatorio une tablas 1-a-muchos
    -- (partidas, fechas, archivos) y sin una columna que las
    -- distinga, esas filas salen repetidas en el Excel.
    ------------------------------------------------------------------
    DECLARE @OrderBy NVARCHAR(MAX);
    SELECT TOP 1 @OrderBy = ColumnaSql FROM #Seleccion ORDER BY Orden;

    DECLARE @Sql NVARCHAR(MAX) = N'
		SELECT DISTINCT TOP 100
			' + @Select + N'
		FROM ' + @From + N'
		WHERE ' + @Where + N'
		ORDER BY ' + @OrderBy + N';';

    EXEC sp_executesql
        @Sql,
        N'@FiltrosJson NVARCHAR(MAX)',
        @FiltrosJson = @FiltrosJson;
END
