USE [VISTA]
GO
/****** Object:  StoredProcedure [dbo].[sp_Reporte_Generar]    Script Date: 10/07/2026 02:15:19 p. m. ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =====================================================================
-- sp_Reporte_Generar (vista previa paginada)
--
-- NUEVO en esta versión respecto a la anterior:
--   1. Ya no asume que el hub siempre es VU150501 — usa la columna
--      Rep_CatTablas.EsHub, así que puede trabajar con distintos
--      grupos de tablas (ej. el grupo VU150501 original, y el grupo
--      nuevo de SO130120 ).
--   2. Respeta Rep_CatGrupoObligatorio: si el usuario selecciona
--      cualquier columna de una tabla que pertenece a un grupo
--      obligatorio, TODAS las tablas de ese grupo se unen — aunque no
--      se haya pedido ninguna columna de ellas.
--   3. Respeta TipoJoin (INNER/LEFT) y CondicionExtra por relación.
--   4. Soporta campos calculados (Rep_CatCampos.EsCalculado = 1):
--      usa ExpresionCalculada en vez de "tabla.columna".
--
-- Las reglas de seguridad de siempre se mantienen intactas: toda
-- tabla/columna sale del catálogo (whitelist), los valores de los
-- filtros nunca se concatenan (viajan en @FiltrosJson y se leen con
-- OPENJSON en tiempo de ejecución).
-- =====================================================================
ALTER   PROCEDURE [dbo].[sp_Reporte_Generar]
    @CampoIds     NVARCHAR(MAX),
    @FiltrosJson  NVARCHAR(MAX) = NULL,
    @Pagina       INT = 1,
    @TamanioPagina INT = 50
AS
BEGIN
    SET NOCOUNT ON;

    IF @Pagina < 1 SET @Pagina = 1;
    IF @TamanioPagina < 1 OR @TamanioPagina > 100
    BEGIN
        RAISERROR('El tamaño de página debe estar entre 1 y 100.', 16, 1);
        RETURN;
    END

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
    -- VU150501, que también es el hub del star original y participa
    -- en montones de reportes que no tienen nada que ver con el
    -- grupo nuevo.
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
    -- 2b) Requisitos de filtro obligatorio por grupo (ej. RFC)
    ------------------------------------------------------------------
    IF EXISTS (
        SELECT 1
        FROM dbo.Rep_CatGrupoObligatorio g
        JOIN #TablasUsadas t ON t.TablaId = g.TablaId
        JOIN dbo.Rep_CatGrupoRequisitos req ON req.NombreGrupo = g.NombreGrupo
        WHERE NOT EXISTS (
            SELECT 1 FROM OPENJSON(@FiltrosJson) WITH (
                CampoId INT           '$.CampoId',
                Valor   NVARCHAR(500) '$.Valor'
            ) f
            WHERE f.CampoId = req.CampoCatalogoId AND f.Valor IS NOT NULL AND f.Valor <> ''
        )
    )
    BEGIN
        RAISERROR('Este tipo de reporte requiere que captures el RFC del cliente como filtro.', 16, 1);
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
        -- propio marcado? (ej. SO130120 dentro de PedimentosPorRfc).
        SELECT TOP 1 @HubTablaId = ct.Id, @HubNombreTabla = ct.NombreTabla
        FROM #TablasUsadas t
        JOIN dbo.Rep_CatGrupoObligatorio g ON g.TablaId = t.TablaId
        JOIN dbo.Rep_CatGrupoObligatorio gh
            ON gh.NombreGrupo = g.NombreGrupo AND gh.EsHubDelGrupo = 1
        JOIN dbo.Rep_CatTablas ct ON ct.Id = gh.TablaId;

        -- Si no hay grupo con hub propio activo, cae al EsHub general
        -- (comportamiento del star original, sin cambios).
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

        -- Si alguna tabla no-hub no tiene relación definida hacia el hub, se quedaría fuera del FROM sin avisar — se valida aquí.
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
    -- 4) SELECT list (usa ColumnaSql: columna real o fórmula calculada)
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
    -- 6) Armar y ejecutar (paginado, 100% parametrizado en valores)
    --
    -- IMPORTANTE: la vista previa NUNCA debe contar ni paginar sobre
    -- el resultado filtrado completo (puede ser de cientos de miles
    -- de filas) — así fue como terminamos con "359181 registros" y
    -- 17960 páginas en pantalla. El universo de la vista previa se
    -- topa en @TopUniverso (100, igual que sp_Reporte_GenerarExport):
    -- primero se materializa ESE TOP 100 en una tabla temporal
    -- (#PreviewSet), y tanto el conteo como el OFFSET/FETCH de la
    -- página se hacen sobre esa tabla (100 filas, prácticamente
    -- gratis), nunca sobre la tabla/join completo. El total que ve el
    -- front como consecuencia queda entre 0 y 100 — nunca el total
    -- real de la consulta.
    --
    -- DISTINCT: el grupo obligatorio une tablas 1-a-muchos contra el
    -- pedimento (partidas, fechas por TipoFecha, archivos/acuses). Si
    -- el usuario no selecciona ninguna columna que distinga esas
    -- filas (ej. número de partida), el JOIN las multiplica pero se
    -- ven idénticas en pantalla — de ahí filas "duplicadas" que en
    -- realidad corresponden a combinaciones distintas por debajo.
    -- DISTINCT las colapsa cuando de verdad son iguales en las
    -- columnas elegidas, y dejar de colapsarlas automáticamente en
    -- cuanto el usuario agregue una columna que sí las distinga.
    ------------------------------------------------------------------
    DECLARE @TopUniverso INT = 100;

    DECLARE @OrderBy NVARCHAR(MAX);
    SELECT TOP 1 @OrderBy = ColumnaSql FROM #Seleccion ORDER BY Orden;

    DECLARE @OrderByAlias NVARCHAR(400);
    SELECT TOP 1 @OrderByAlias = N'[' + REPLACE(Alias, ']', ']]') + N']' FROM #Seleccion ORDER BY Orden;

    DECLARE @Offset INT = (@Pagina - 1) * @TamanioPagina;

    DECLARE @Sql NVARCHAR(MAX) = N'
		SELECT DISTINCT TOP (@pTopUniverso)
			' + @Select + N'
		INTO #PreviewSet
		FROM ' + @From + N'
		WHERE ' + @Where + N'
		ORDER BY ' + @OrderBy + N';

		SELECT COUNT_BIG(*) FROM #PreviewSet;

		SELECT * FROM #PreviewSet
		ORDER BY ' + @OrderByAlias + N'
		OFFSET @pOffset ROWS FETCH NEXT @pTamanioPagina ROWS ONLY;';

    EXEC sp_executesql
        @Sql,
        N'@FiltrosJson NVARCHAR(MAX), @pTopUniverso INT, @pOffset INT, @pTamanioPagina INT',
        @FiltrosJson = @FiltrosJson,
        @pTopUniverso = @TopUniverso,
        @pOffset = @Offset,
        @pTamanioPagina = @TamanioPagina;
END
