import marimo

__generated_with = "0.14.12"
app = marimo.App(
    width="full",
    app_title="Compulab",
    layout_file="layouts/compulab.grid.json",
    sql_output="native",
)

with app.setup:
    # Initialization code that runs before all other cells
    import os
    import marimo as mo
    import ibis
    import pandas as pd
    from decimal import Decimal, ROUND_HALF_UP
    from dotenv import load_dotenv
    from bs4 import BeautifulSoup as bs
    from collections import defaultdict


@app.cell
def config():
    # Ruta archivos parametros
    load_dotenv()
    parametros_path = os.environ["PARAMETROS_PATH"]
    f572_path = os.environ["F572_PATH"]

    # archivos de parametros
    fuente_items_liquidacion = rf"{parametros_path}/items_liquidacion.xlsx"
    fuente_bases_imponibles = rf"{parametros_path}/bases_imponibles.xlsx"
    fuente_ganancias_art_94 = rf"{parametros_path}/ganancias_art_94.xlsx"
    fuente_ganancias_deducciones_personales = (
        rf"{parametros_path}/ganancias_deducciones_personales.xlsx"
    )
    fuente_ganancias_deducciones_generales = (
        rf"{parametros_path}/ganancias_deducciones_generales.xlsx"
    )
    fuente_ganancias_art_94 = rf"{parametros_path}/ganancias_art_94.xlsx"
    fuente_f931 = rf"{parametros_path}/f931.xlsx"
    fuente_ffep = rf"{parametros_path}/ffep.xlsx"
    fuente_scvo = rf"{parametros_path}/scvo.xlsx"
    fuente_smvm = rf"{parametros_path}/smvm.xlsx"

    # Ibis modo interactivo
    ibis.options.interactive = True

    # Conexión a DuckDB (in-memory)
    con = ibis.duckdb.connect()
    # ibis.set_backend(con)

    # Conexión a MSSQL
    try:
        con_mssql = ibis.mssql.connect(
            user=os.environ["MSSQL_USER"],
            password=os.environ["MSSQL_PASS"],
            host=os.environ["MSSQL_HOST"],
            database=os.environ["MSSQL_DB"],
            driver=os.environ["MSSQL_DRIVER"],
            port=os.environ["MSSQL_PORT"],
        )
    except Exception as e:
        print(f"No se pudo conectar a {e}")

    # table_conceptos_liquidacion = con.read_csv(
    #     conceptos_liquidacion,
    #     auto_detect=False,
    #     header=True,
    #     decimal_separator=",",
    #     delim=";",
    #     encoding="utf-8",
    #     nullstr="NULL",
    #     strict_mode=True,
    #     columns={
    #         "idLiquidacionItem": "INT",
    #         "nombre": "VARCHAR",
    #         "categoria": "VARCHAR",
    #         "subcategoria": "VARCHAR",
    #         "ganancias": "VARCHAR",
    #         "r4r8": "VARCHAR",
    #         "r9": "BOOLEAN",
    #     },
    # ).execute()
    return (
        con,
        con_mssql,
        f572_path,
        fuente_bases_imponibles,
        fuente_ganancias_art_94,
        fuente_ganancias_deducciones_generales,
        fuente_ganancias_deducciones_personales,
        fuente_items_liquidacion,
    )


@app.cell
def ui():
    dropdown_tipo_liquidacion = mo.ui.dropdown(
        options={
            "Embarcados": "Emb",
            "Jornalizados": "Jor",
            "Mensualizados": "Men",
        },
        value="Embarcados",
        allow_select_none=False,
        label="Liquidación tipo: ",
    )
    return (dropdown_tipo_liquidacion,)


@app.cell
def _(con_mssql, dropdown_tipo_liquidacion):
    tipo_liquidacion = dropdown_tipo_liquidacion.value

    list_año = mo.sql(
        f"""
        SELECT TOP 5 Año
        FROM (
            SELECT
                DISTINCT "PeriodoAnio" AS Año
            FROM 
                "Liquidaciones{tipo_liquidacion}"
                ) AS periodos
        ORDER BY Año DESC
        """,
        engine=con_mssql,
    ).to_pandas()

    list_año = list_año.sort_values(by="Año", ascending=False)

    dropdown_periodo_año = mo.ui.dropdown.from_series(
        list_año["Año"],
        label="Año",
        searchable=True,
        allow_select_none=False,
        value="2025",
    )
    return dropdown_periodo_año, tipo_liquidacion


@app.cell
def _(Liquidacionesnull, con_mssql, dropdown_periodo_año, tipo_liquidacion):
    list_mes = mo.sql(
        f"""
        SELECT
            DISTINCT "PeriodoMes" AS Mes
        FROM
            "Liquidaciones{tipo_liquidacion}"
        WHERE
            "PeriodoAnio" LIKE {dropdown_periodo_año.value}
        """,
        engine=con_mssql,
    ).to_pandas()

    list_mes = list_mes.sort_values(by="Mes", ascending=True)

    dropdown_periodo_mes = mo.ui.dropdown.from_series(
        list_mes["Mes"],
        searchable=True,
        label="Mes",
        allow_select_none=False,
        value="1",
    )

    button_generar_control = mo.ui.run_button(label="Generar hoja de control")

    dropdown_metodo_ganancias = mo.ui.dropdown(
        label="Remuneraciones: ",
        options=["Percibidas", "Devengadas"],
        allow_select_none=False,
        value="Percibidas",
    )

    dropdown_no_habituales = mo.ui.dropdown(
        label="Prorratear no habituales: ",
        options=["Sí", "No"],
        allow_select_none=False,
        value="Sí",
    )

    dropdown_tratamiento_sac = mo.ui.dropdown(
        label="Prorrateo SAC: ",
        options=["Semestral", "Anual"],
        allow_select_none=False,
        value="Semestral",
    )

    dropdown_reportes = mo.ui.multiselect(
        label="Informes a generar: ",
        options=[
            "Resúmen de importes",
            "Control Descuentos",
            "Control Imp. Ganancias",
            "Control F.931 / LSD",
        ],
        value=[
            "Resúmen de importes",
            "Control Descuentos",
            "Control Imp. Ganancias",
        ],
    )

    button_refrescar = mo.ui.run_button(label="Actualizar")

    button_exportar_items_todos = mo.ui.run_button(label="Exportar ítems (todos)")
    button_exportar_items_faltantes = mo.ui.run_button(
        label="Exportar ítems (sin parametrizar)"
    )

    button_exportar_liquidaciones = mo.ui.run_button(
        label="Exportar liquidaciones"
    )
    return (
        button_exportar_items_faltantes,
        button_exportar_items_todos,
        button_exportar_liquidaciones,
        button_generar_control,
        dropdown_metodo_ganancias,
        dropdown_no_habituales,
        dropdown_periodo_mes,
        dropdown_reportes,
        dropdown_tratamiento_sac,
    )


@app.cell
def _(
    dropdown_metodo_ganancias,
    dropdown_no_habituales,
    dropdown_tratamiento_sac,
):
    mo.hstack(
        [
            "Computo Imp. Ganancias - ",
            dropdown_metodo_ganancias,
            dropdown_no_habituales,
            dropdown_tratamiento_sac,
        ],
        justify="start",
        align="center",
    )
    return


@app.cell
def _(
    button_generar_control,
    dropdown_periodo_año,
    dropdown_periodo_mes,
    dropdown_reportes,
    dropdown_tipo_liquidacion,
):
    mo.hstack(
        [
            dropdown_tipo_liquidacion,
            dropdown_periodo_año,
            dropdown_periodo_mes,
            dropdown_reportes,
            button_generar_control,
        ],
        justify="start",
    )
    return


@app.cell
def _(con, con_mssql):
    con.raw_sql(
        """
        DROP TABLE IF EXISTS "items_liquidacion"
        """
    )

    con.create_table(
        "items_liquidacion",
        con_mssql.sql(
            """
            SELECT
                idLiquidacionItem,
                Nombre
            FROM
                LiquidacionItems
            """
        ).execute(),
    )
    return


@app.cell
def _(
    con,
    dropdown_periodo_año,
    dropdown_periodo_mes,
    fuente_bases_imponibles,
):
    con.raw_sql(
        """
        DROP TABLE IF EXISTS "bases_imponibles"
        """
    )

    bases_imponibles = con.create_table(
        "bases_imponibles",
        mo.sql(
            f"""
            SELECT 
                Periodo AS periodo,
                CAST (Maxima AS DECIMAL(14, 2)) AS tope_maximo,
                CAST (Minima AS DECIMAL(10, 2)) AS tope_minimo
            FROM
                read_xlsx(
                    '{fuente_bases_imponibles}',
                    sheet='bases_imponibles'                
                )
            WHERE 
                CAST(YEAR(Periodo) AS VARCHAR) LIKE '{dropdown_periodo_año.value}'
            AND
                CAST(MONTH(Periodo) AS VARCHAR) LIKE '{dropdown_periodo_mes.value}';
            """
        ),
    )

    tope_periodo_maximo = str(
        con.execute(bases_imponibles["tope_maximo"].as_scalar())
    )
    tope_periodo_minimo = str(
        con.execute(bases_imponibles["tope_minimo"].as_scalar())
    )
    return tope_periodo_maximo, tope_periodo_minimo


@app.cell
def _(tope_periodo_maximo):
    mo.md(f"""Tope maximo: {tope_periodo_maximo}""")
    return


@app.cell
def _(tope_periodo_minimo):
    mo.md(f"""Tope minimo: {tope_periodo_minimo}""")
    return


@app.cell
def _(con, fuente_items_liquidacion):
    con.raw_sql(
        """
        DROP TABLE IF EXISTS "items_parametrizados"
        """
    )

    items_parametrizados = con.create_table(
        "items_parametrizados",
        mo.sql(
            f""" 
            SELECT
                CAST(idLiquidacionItem AS INTEGER) AS idLiquidacionItem,
                CAST(nombre AS STRING) AS nombre_item,
                CAST(categoria AS STRING) AS categoria_item,
                COALESCE(CAST(subcategoria AS STRING), 'n/a') AS subcategoria_item,
                CAST(ganancias AS STRING) AS ganancias_item,
                CAST(r4r8 AS STRING) AS r4r8,
                CAST(r9 AS BOOLEAN) AS r9
            FROM 
                read_xlsx(
                    '{fuente_items_liquidacion}',
                    sheet='items_liquidacion',
                    stop_at_empty=false,
                    all_varchar=true
                    );
        """,
            engine=con,
        ),
    )
    return (items_parametrizados,)


@app.cell
def _(con, items_liquidacion, items_parametrizados):
    items_sin_parametrizar = mo.sql(
        f"""
        SELECT    
            "t0"."idLiquidacionItem",    
            "t0"."Nombre" 
        FROM "items_liquidacion" AS "t0"
        LEFT JOIN "items_parametrizados" AS "t1"
            ON "t0"."idLiquidacionItem" = "t1"."idLiquidacionItem"
        WHERE "t1"."idLiquidacionItem" IS NULL
        """,
        output=False,
        engine=con
    )
    return (items_sin_parametrizar,)


@app.cell
def _(con, items_sin_parametrizar):
    con.raw_sql(
        """
        DROP TABLE IF EXISTS "items_sin_parametrizar";
        """
    )

    con.create_table("items_sin_parametrizar", items_sin_parametrizar)

    tabla_items_sin_parametrizar = con.table("items_sin_parametrizar")
    return


@app.cell
def _(
    DatosPersonales,
    DatosPersonalesDocumentos,
    Liquidacionesnull,
    LiquidacionesnullDetalle,
    con_mssql,
    dropdown_periodo_año,
    dropdown_periodo_mes,
    tipo_liquidacion,
):
    items_liquidados = mo.sql(
        f"""
        SELECT
            "t1"."idDatoPersonal" AS 'id_dato_personal',
            "t2"."Apellido" AS 'apellido',
            "t2"."Nombres" AS 'nombre',
            "t0"."idLiquidacion" AS 'id_liquidacion',
            "t1"."idLiquidacionTipo" AS 'id_liquidacion_tipo',
            "t0"."idLiquidacionItem" AS 'id_liquidacion_item',
            "t0"."Cantidad" AS 'cantidad',
            COALESCE("t0"."Haber", 0) AS 'haber',
            COALESCE("t0"."Dscto", 0) AS 'dscto',
            "t0"."NombreItem" AS 'nombre_item',    
            "t1"."PeriodoMes" AS 'periodo_mes',
            "t3"."NumeroDocumento" AS 'numero_documento'
            --"t0"."Unidades"
            --"t0"."idEstado"
        FROM
            "Liquidaciones{tipo_liquidacion}Detalle" AS "t0"
            JOIN "Liquidaciones{tipo_liquidacion}" AS "t1" ON "t0"."idLiquidacion" = "t1"."idLiquidacion"
            JOIN "DatosPersonales" AS "t2" ON "t1"."idDatoPersonal" = "t2"."idDatoPersonal"
            JOIN "DatosPersonalesDocumentos" AS "t3" ON "t2"."idDatoPersonal" = "t3"."idDatoPersonal"
                AND "t3"."idTipoDocumento" = '33'
        WHERE
            --Filtra de principio de año hasta el mes elegido así ganancias computa ese tramo
            --el resto de los controles toman el mes puntual que necesitan.
            PeriodoAnio = '{dropdown_periodo_año.value}'
             AND PeriodoMes <= '{dropdown_periodo_mes.value}'

        ORDER BY
            "t1"."PeriodoMes" ASC;
        """,
        output=False,
        engine=con_mssql
    )
    return (items_liquidados,)


@app.cell
def _(con, items_liquidados):
    # Consulta para extraer liquidaciones confeccionadas (Embarcados)
    con.raw_sql(
        """
        DROP TABLE IF EXISTS "items_liquidados"
        """
    )

    # Convertimos a pandas porque no permite crear tabla desde otro backend
    tabla_items_liquidados = items_liquidados.to_pandas()

    con.create_table("items_liquidados", tabla_items_liquidados)

    tabla_items_liquidados = con.table("items_liquidados")

    con.raw_sql(
        """
        DROP TABLE IF EXISTS "temp_items_liquidados";
        """
    )

    con.raw_sql(
        """
        CREATE TABLE temp_items_liquidados AS
        SELECT 
            *
        FROM
            "items_liquidados" AS "t0"
        LEFT JOIN "items_parametrizados" AS "t1" ON "t0"."id_liquidacion_item" = "t1"."idLiquidacionItem"         
        ORDER BY
            "t0"."periodo_mes" ASC,
            "t0"."apellido" ASC,
            "t0"."nombre" ASC,
            CASE
                WHEN "t1"."categoria_item" = 'remunerativo' THEN 1
                WHEN "t1"."categoria_item" = 'no remunerativo' THEN 2
                WHEN "t1"."categoria_item" = 'descuento' THEN 3
                WHEN "t1"."categoria_item" = 'ganancias' THEN 4
                WHEN "t1"."categoria_item" = 'plantilla' THEN 5
                ELSE 6
            END;
        """
    )

    con.raw_sql(
        """
        DROP TABLE "items_liquidados";
        """
    )

    con.raw_sql(
        """
        ALTER TABLE "temp_items_liquidados" RENAME TO "items_liquidados";
        """
    )

    tabla_items_liquidados = con.table("items_liquidados")
    return (tabla_items_liquidados,)


@app.cell
def _(
    DatosPersonales,
    DatosPersonalesLaborales,
    LiquidacionItems,
    Liquidacionesnull,
    LiquidacionesnullDetalle,
    MAE_ConveniosTipos,
    con_mssql,
    dropdown_periodo_año,
    tipo_liquidacion,
):
    datos_adicionales = mo.sql(
        f"""
        SELECT
            DISTINCT "t1"."idDatoPersonal" AS "id_dato_personal",
            CAST(CASE WHEN "idCodigoCondicion" = 2 THEN 'TRUE' ELSE 'FALSE' END AS BIT) AS 'jubilado',
            "t_conv"."ConvenioTipo" AS "convenio",
            "t_os"."Nombre" AS "obra_social",
            "t_sind"."Nombre" AS "sindicato",
            CAST("SindEsAfiliado" AS BIT) AS "sind_afiliado",
            "t3"."FechaBaja" AS 'fecha_baja'
        FROM
            "Liquidaciones{tipo_liquidacion}Detalle" AS "t0"
        JOIN "Liquidaciones{tipo_liquidacion}" AS "t1" ON "t0"."idLiquidacion" = "t1"."idLiquidacion"
        JOIN "DatosPersonalesLaborales" AS "t2" ON "t1"."idDatoPersonal" = "t2"."idDatoPersonal"
        JOIN "DatosPersonales" AS "t3" ON "t1"."idDatoPersonal" = "t3"."idDatoPersonal"
        LEFT JOIN "LiquidacionItems" AS "t_os" ON "t2"."idObraSocial" = "t_os"."idLiquidacionItem"
        LEFT JOIN "LiquidacionItems" AS "t_sind" ON "t2"."idSindicato" = "t_sind"."idLiquidacionItem"
        LEFT JOIN "MAE_ConveniosTipos" AS "t_conv" ON "t_conv"."idConvenioTipo" = "t2"."idConvenioTipo"
        WHERE
            PeriodoAnio LIKE '{dropdown_periodo_año.value}'
        """,
        output=False,
        engine=con_mssql
    )
    return (datos_adicionales,)


@app.cell
def _(con, datos_adicionales):
    # Tabla de datos adicionales con obra social, sindicato, condición de jubilado de empleados
    con.raw_sql(
        """
        DROP TABLE IF EXISTS "datos_adicionales"
        """
    )

    tabla_datos_adicionales = datos_adicionales.to_pandas()

    con.create_table("datos_adicionales", tabla_datos_adicionales)

    tabla_datos_adicionales = con.table("datos_adicionales")
    return


@app.cell
def _(
    con,
    dropdown_periodo_año,
    dropdown_periodo_mes,
    fuente_ganancias_art_94,
    fuente_ganancias_deducciones_generales,
    fuente_ganancias_deducciones_personales,
):
    con.raw_sql(
        """
        DROP TABLE IF EXISTS "ganancias_art_94";
        DROP TABLE IF EXISTS "ganancias_deducciones_personales";
        DROP TABLE IF EXISTS "ganancias_deducciones_generales";
        """
    )

    ganancias_art_94 = con.create_table(
        "ganancias_art_94",
        mo.sql(
            f"""
            SELECT
                CAST(periodo AS DATE) AS "periodo",
                CAST(desde AS DECIMAL(14, 2)) AS "desde",
                CAST(hasta AS DECIMAL(14, 2)) AS "hasta",
                CAST(suma_fija AS DECIMAL(10, 2)) AS "suma_fija",
                CAST(coeficiente AS TINYINT) AS "coeficiente",
                CAST(excedente AS DECIMAL(10, 2)) AS "excedente",
                CAST(codigo_porcentaje AS TINYINT) AS "cod_%",
                CAST(tipo AS TINYINT) AS "tipo"
            FROM
                read_xlsx(
                    '{fuente_ganancias_art_94}',
                    sheet='gan_art_94'
                )
            WHERE 
                YEAR(Periodo) = '{dropdown_periodo_año.value}'
            AND (
                MONTH(Periodo) = '{dropdown_periodo_mes.value}'
                OR MONTH(Periodo) = '12'
                    AND Tipo = 1
                    OR Tipo = 2
                    );
                """,
            engine=con,
        ),
    )

    ganancias_deducciones_personales = con.create_table(
        "ganancias_deducciones_personales",
        mo.sql(
            f"""
            SELECT
                CAST(Periodo AS DATE) AS "periodo",
                CAST(GNI AS DECIMAL(10, 2)) AS "ded_gni",
                CAST(Especial AS DECIMAL(10, 2)) AS "ded_especial",
                CAST(Conyuge AS DECIMAL(10, 2)) AS "ded_conyuge",
                CAST(Hijo AS DECIMAL(10, 2)) AS "ded_hijo",
                CAST(Hijo_Incapacitado AS DECIMAL(10, 2)) AS "ded_hijo_inc",
                CAST(Tipo AS TINYINT) AS "tipo"
            FROM
                read_xlsx(
                    '{fuente_ganancias_deducciones_personales}',
                    sheet='gan_deducciones_personales'
                )
            WHERE 
                YEAR(Periodo) = '{dropdown_periodo_año.value}'
            AND(
                MONTH(Periodo) = '{dropdown_periodo_mes.value}'
                OR MONTH(Periodo) = '12'
                    AND Tipo = 1
                    OR Tipo = 2
                    );
            """,
            engine=con,
        ),
    )

    ganancias_deducciones_generales = con.create_table(
        "ganancias_deducciones_generales",
        mo.sql(
            f"""
            SELECT
                CAST(Periodo AS DATE) AS "periodo",
                CAST(Deduccion AS VARCHAR) AS "deduccion",
                CAST(Codigo AS SMALLINT) AS "codigo",
                Tipo AS "tipo",
                CAST(Tope AS DECIMAL(12, 2)) AS "tope",
                CAST(Porcentaje AS TINYINT) AS "porcentaje",
            FROM
                read_xlsx(
                    '{fuente_ganancias_deducciones_generales}',
                    sheet='gan_deducciones_generales'
                )
            WHERE 
                YEAR(Periodo) = '{dropdown_periodo_año.value}';
            """,
            engine=con,
        ),
    )
    return


@app.cell
def _(con, dropdown_periodo_año, dropdown_periodo_mes):
    monto_gni = (
        con.sql(f"""
        SELECT
            CAST("ded_gni" AS DECIMAL(12, 2))
        FROM
            "ganancias_deducciones_personales"
        WHERE
            YEAR(periodo) = {dropdown_periodo_año.value} AND
            MONTH(periodo) = {dropdown_periodo_mes.value}
        """)
        .execute()
        .iat[0, 0]
    )

    monto_ded_conyuge = (
        con.sql(f"""
        SELECT
            CAST("ded_conyuge" AS DECIMAL(12, 2))
        FROM
            "ganancias_deducciones_personales"
        WHERE
            YEAR(periodo) = {dropdown_periodo_año.value} AND
            MONTH(periodo) = {dropdown_periodo_mes.value}
        """)
        .execute()
        .iat[0, 0]
    )

    monto_ded_hijo = (
        con.sql(f"""
        SELECT
            CAST("ded_hijo" AS DECIMAL(12, 2))
        FROM
            "ganancias_deducciones_personales"
        WHERE
            YEAR(periodo) = {dropdown_periodo_año.value} AND
            MONTH(periodo) = {dropdown_periodo_mes.value}
        """)
        .execute()
        .iat[0, 0]
    )

    monto_ded_hijo_inc = (
        con.sql(f"""
        SELECT
            CAST("ded_hijo_inc" AS DECIMAL(12, 2))
        FROM
            "ganancias_deducciones_personales"
        WHERE
            YEAR(periodo) = {dropdown_periodo_año.value} AND
            MONTH(periodo) = {dropdown_periodo_mes.value}
        """)
        .execute()
        .iat[0, 0]
    )

    tope_seguro_muerte = (
        con.sql(f"""
        SELECT
            CAST("tope" AS DECIMAL(12, 2))
        FROM
            "ganancias_deducciones_generales"
        WHERE
            YEAR(periodo) = {dropdown_periodo_año.value}
        AND
            "codigo" = 2
        """)
        .execute()
        .iat[0, 0]
    )

    tope_intereses_hipotecarios = (
        con.sql(f"""
        SELECT
            CAST("tope" AS DECIMAL(12, 2))
        FROM
            "ganancias_deducciones_generales"
        WHERE
            YEAR(periodo) = {dropdown_periodo_año.value}
        AND
            "codigo" = 4
        """)
        .execute()
        .iat[0, 0]
    )

    tope_gastos_sepelio = (
        con.sql(f"""
        SELECT
            CAST("tope" AS DECIMAL(12, 2))
        FROM
            "ganancias_deducciones_generales"
        WHERE
            YEAR(periodo) = {dropdown_periodo_año.value}
        AND
            "codigo" = 5
        """)
        .execute()
        .iat[0, 0]
    )

    tope_casas_particulares = monto_gni * 40 / 100

    tope_movilidad_viaticos = (
        con.sql(f"""
        SELECT
            CAST("tope" AS DECIMAL(12, 2))
        FROM
            "ganancias_deducciones_generales"
        WHERE
            YEAR(periodo) = {dropdown_periodo_año.value}
        AND
            "codigo" = 11
        """)
        .execute()
        .iat[0, 0]
    )

    tope_alquileres_40 = monto_gni * 40 / 100

    tope_seguro_mixto = (
        con.sql(f"""
        SELECT
            CAST("tope" AS DECIMAL(12, 2))
        FROM
            "ganancias_deducciones_generales"
        WHERE
            YEAR(periodo) = {dropdown_periodo_año.value}
        AND
            "codigo" = 23
        """)
        .execute()
        .iat[0, 0]
    )

    tope_seguro_retiro = (
        con.sql(f"""
        SELECT
            CAST("tope" AS DECIMAL(12, 2))
        FROM
            "ganancias_deducciones_generales"
        WHERE
            YEAR(periodo) = {dropdown_periodo_año.value}
        AND
            "codigo" = 24
        """)
        .execute()
        .iat[0, 0]
    )

    tope_fines_educativos = monto_gni * 40 / 100
    return (
        monto_ded_conyuge,
        monto_ded_hijo,
        monto_ded_hijo_inc,
        tope_alquileres_40,
        tope_casas_particulares,
        tope_fines_educativos,
        tope_gastos_sepelio,
        tope_intereses_hipotecarios,
        tope_movilidad_viaticos,
        tope_seguro_mixto,
        tope_seguro_muerte,
        tope_seguro_retiro,
    )


@app.cell
def _(
    con,
    dropdown_periodo_mes,
    items_conyuge,
    items_hijos,
    items_hijos_incapacitados,
    monto_ded_conyuge,
    monto_ded_hijo,
    monto_ded_hijo_inc,
):
    con.raw_sql(f"""
    DROP TABLE IF EXISTS "deducciones_familiares_procesadas";
    DROP TABLE IF EXISTS "deducciones_familiares_agrupadas";
    DROP TABLE IF EXISTS "deducciones_generales_procesadas";
    DROP TABLE IF EXISTS "deducciones_generales_agrupadas";

    CREATE TABLE "deducciones_familiares_procesadas" AS
    SELECT
        "cuil",
        "parentesco",
        CASE
            WHEN "parentesco" IN {items_conyuge} THEN 'conyuge'
            WHEN "parentesco" IN {items_hijos} THEN 'hijo'
            WHEN "parentesco" IN {items_hijos_incapacitados} THEN 'hijo incapacitado'
        END AS "tipo",
        CAST(
            CASE
                WHEN {dropdown_periodo_mes.value} NOT BETWEEN "mes_desde" AND "mes_hasta" THEN 0
                ELSE 1
            END AS TINYINT) AS "aplica",
        "porcentaje",
        "mes_desde",
        "mes_hasta",
        CASE
            WHEN "tipo" = 'conyuge' THEN {monto_ded_conyuge} * "porcentaje" / 100 * "aplica"
            WHEN "tipo" = 'hijo' THEN {monto_ded_hijo} * "porcentaje" / 100 * "aplica"
            WHEN "tipo" = 'hijo incapacitado' THEN {monto_ded_hijo_inc} * "porcentaje" / 100 * "aplica"
        END AS "monto"
    FROM
        "f572_deducciones_familiares"
    ORDER BY
        "cuil" ASC,
        "tipo" ASC;

    CREATE TABLE "deducciones_familiares_agrupadas" AS
    SELECT
        "cuil",
        CAST(SUM("monto") AS DECIMAL(12, 2)) AS "monto"
    FROM
        "deducciones_familiares_procesadas" WHERE "aplica" = 1
    GROUP BY
        "cuil"
    ORDER by
        "cuil" ASC;

    CREATE TABLE "deducciones_generales_procesadas" AS
    SELECT
        "cuil",    
        CAST(COALESCE(SUM("monto_total") FILTER(WHERE "tipo_deduccion" = 1), 0) AS DECIMAL(14, 2)) AS "medico_asistencial",
        CAST(COALESCE(SUM("monto_total") FILTER(WHERE "tipo_deduccion" = 2), 0) AS DECIMAL(14, 2)) AS "seguro_muerte",
        CAST(COALESCE(SUM("monto_total") FILTER(WHERE "tipo_deduccion" = 3), 0) AS DECIMAL(14, 2)) AS "donaciones",
        CAST(COALESCE(SUM("monto_total") FILTER(WHERE "tipo_deduccion" = 4), 0) AS DECIMAL(14, 2)) AS "intereses_hipotecarios",
        CAST(COALESCE(SUM("monto_total") FILTER(WHERE "tipo_deduccion" = 5), 0) AS DECIMAL(14, 2)) AS "gastos_sepelio",
        CAST(COALESCE(SUM("monto_total") FILTER(WHERE "tipo_deduccion" = 7), 0) AS DECIMAL(14, 2)) AS "honorarios_medicos",
        CAST(COALESCE(SUM("monto_total") FILTER(WHERE "tipo_deduccion" = 8), 0) AS DECIMAL(14, 2)) AS "casas_particulares",   
        CAST(COALESCE(SUM("monto_total") FILTER(WHERE "tipo_deduccion" = 9), 0) AS DECIMAL(14, 2)) AS "sociedades_garantia_reciproca",
        CAST(COALESCE(SUM("monto_total") FILTER(WHERE "tipo_deduccion" = 10), 0) AS DECIMAL(14, 2)) AS "viajantes_comercio",
        CAST(COALESCE(SUM("monto_total") FILTER(WHERE "tipo_deduccion" = 11), 0) AS DECIMAL(14, 2)) AS "movilidad_viaticos",   
        CAST(COALESCE(SUM("monto_total") FILTER(WHERE "tipo_deduccion" = 21), 0) AS DECIMAL(14, 2)) AS "indumentaria_trabajo",
        CAST(COALESCE(SUM("monto_total") FILTER(WHERE "tipo_deduccion" = 22), 0) AS DECIMAL(14, 2)) AS "alquileres_40",
        CAST(COALESCE(SUM("monto_total") FILTER(WHERE "tipo_deduccion" = 23), 0) AS DECIMAL(14, 2)) AS "seguro_mixto",
        CAST(COALESCE(SUM("monto_total") FILTER(WHERE "tipo_deduccion" = 24), 0) AS DECIMAL(14, 2)) AS "seguro_retiro",   
        CAST(COALESCE(SUM("monto_total") FILTER(WHERE "tipo_deduccion" = 25), 0) AS DECIMAL(14, 2)) AS "fondos_comunes_inversion",
        CAST(COALESCE(SUM("monto_total") FILTER(WHERE "tipo_deduccion" = 32), 0) AS DECIMAL(14, 2)) AS "fines_educativos",
        CAST(COALESCE(SUM("monto_total") FILTER(WHERE "tipo_deduccion" = 33), 0) AS DECIMAL(14, 2)) AS "alquileres_10_locatario",
        CAST(COALESCE(SUM("monto_total") FILTER(WHERE "tipo_deduccion" = 34), 0) AS DECIMAL(14, 2)) AS "alquileres_10_locador", 
        CAST(COALESCE(SUM("monto_total") FILTER(WHERE "tipo_deduccion" = 99), 0) AS DECIMAL(14, 2)) AS "otras_deducciones"    
    FROM 
        "f572_deducciones_generales"
    GROUP BY
        "cuil",    
    ORDER BY
        "cuil" ASC;
    """)
    return


@app.cell
def _(con, dropdown_periodo_mes):
    def resumen_liquidaciones():
        con.raw_sql(
            f"""
            DROP TABLE IF EXISTS "resumen_liquidaciones";
            DROP TABLE IF EXISTS "resumen_liquidaciones_agrupadas";

            CREATE TABLE "resumen_liquidaciones" AS
            SELECT
                numero_documento AS cuil,
                apellido,
                --UPPER(SUBSTR(Apellido, 1, 1)) || LOWER(SUBSTR(Apellido, 2)) AS apellido,
                nombre,
                COALESCE(SUM("haber") FILTER(WHERE "categoria_item" = 'remunerativo'), 0) AS "remunerativo",
                COALESCE(SUM("haber") FILTER(WHERE "categoria_item" = 'no remunerativo'), 0) AS "no_remunerativo",
                COALESCE(SUM("dscto") FILTER(WHERE "subcategoria_item" = 'previsional'), 0) AS "previsional",
                COALESCE(SUM("dscto") FILTER(WHERE "subcategoria_item" = 'obra social'), 0) AS "obra_social",
                COALESCE(SUM("dscto") FILTER(WHERE "subcategoria_item" = 'sindical'), 0) AS "sindical",
                COALESCE(SUM("dscto") FILTER(WHERE "subcategoria_item" = 'anticipo'), 0) AS "anticipos",
                COALESCE(SUM("dscto") FILTER(WHERE "subcategoria_item" = 'embargo'), 0) AS "embargos",
                COALESCE(SUM("dscto") FILTER(WHERE "subcategoria_item" = 'descuento'), 0) AS "otros_descuentos",
                COALESCE(SUM("dscto") FILTER(WHERE "ganancias_item" = 'impuesto'), 0) AS "imp_ganancias",            
                remunerativo + no_remunerativo - previsional - obra_social - sindical - anticipos - embargos - otros_descuentos - imp_ganancias AS "neto"
            FROM
                "items_liquidados"
            WHERE
                "periodo_mes" = {dropdown_periodo_mes.value}
            GROUP BY
                "id_dato_personal", "apellido", "nombre", "cuil"
            ORDER BY
                "apellido" ASC,
                "nombre" ASC;

            CREATE TABLE "resumen_liquidaciones_agrupadas" AS
            SELECT
                SUM("remunerativo") AS "remunerativo",
                SUM("no_remunerativo") AS "no_remunerativo",
                SUM("previsional") AS "previsional",
                SUM("obra_social") AS "obra_social",
                SUM("sindical") AS "sindical",
                SUM("anticipos") AS "anticipos",
                SUM("embargos") AS "embargos",
                SUM("otros_descuentos") AS "otros_descuentos",
                SUM("imp_ganancias") AS "imp_ganancias",
                SUM("neto") AS "neto",
            FROM
                "resumen_liquidaciones"
            """
        )

        tabla_1 = con.table("resumen_liquidaciones")
        tabla_2 = con.table("resumen_liquidaciones_agrupadas")
        return tabla_1, tabla_2


    tabla_resumen_liquidaciones, tabla_resumen_liquidaciones_agrupadas = (
        resumen_liquidaciones()
    )
    return tabla_resumen_liquidaciones, tabla_resumen_liquidaciones_agrupadas


@app.cell
def _(con, dropdown_periodo_mes, tope_periodo_maximo, tope_periodo_minimo):
    def control_descuentos():
        """
        Función para el control de descuentos de ley (Jubilación + Ley 19.032, Obras sociales y sindicato)
        """

        con.raw_sql(
            f"""
            DROP TABLE IF EXISTS "control_descuentos";

            CREATE TABLE "control_descuentos" AS
            SELECT
                cuil,
                apellido,
                nombre,
                base_calculada,
                CAST(
                    base_calculada * 
                    CASE
                        WHEN "jubilado" = FALSE
                        THEN 14
                        ELSE 11
                    END / 100 - desc_suss
                    AS DECIMAL(10, 2)
                ) AS 'ctrl_suss',
                CAST(     
                    CASE
                        WHEN base_calculada < ({tope_periodo_minimo} * 2) THEN {tope_periodo_minimo} * 2 * 3
                        ELSE base_calculada *
                        CASE
                            WHEN jubilado = FALSE
                            THEN 3
                            ELSE 0
                        END
                    END / 100 - desc_os
                    AS DECIMAL(10, 2)
                ) AS 'ctrl_os',                    
                --Tengo que mnajear una logica con % variable segun el sindicato que este aplicando
                --CAST((rem_sin_sac + sac) * 2 / 100 - desc_sindical AS DECIMAL(14, 2)) AS 'ctrl_sind',
                desc_suss,
                desc_os,
                desc_sindical,
                rem_sin_sac,
                sac,
                no_rem_con_apo,
                jubilado,
                convenio,
                SUBSTRING(obra_social, 1, 6) AS obra_social,
                sindicato,
                sind_afiliado
            FROM (
                SELECT
                    numero_documento as cuil,
                    apellido,
                    nombre, 
                    COALESCE(SUM("dscto") FILTER(WHERE "subcategoria_item" = 'previsional'), 0) AS 'desc_suss',
                    COALESCE(SUM("dscto") FILTER(WHERE "subcategoria_item" = 'obra social'), 0) AS 'desc_os',
                    COALESCE(SUM("dscto") FILTER(WHERE "subcategoria_item" = 'sindical'), 0) AS 'desc_sindical',
                    COALESCE(SUM("haber") FILTER(WHERE "categoria_item" = 'remunerativo' AND "subcategoria_item" != 'sac'), 0) AS 'rem_sin_sac',
                    COALESCE(SUM("haber") FILTER(WHERE "subcategoria_item" = 'sac'), 0) AS 'sac',
                    COALESCE(SUM("haber") FILTER(WHERE "categoria_item" = 'no remunerativo'), 0) AS 'no_rem_con_apo',            
                    jubilado,
                    convenio,
                    obra_social,
                    sindicato,
                    sind_afiliado,
                    CASE
                        WHEN rem_sin_sac > {tope_periodo_maximo} THEN {tope_periodo_maximo}
                        WHEN rem_sin_sac < {tope_periodo_minimo} THEN {tope_periodo_minimo}
                        ELSE rem_sin_sac
                    END +
                    CASE
                        WHEN sac > ({tope_periodo_maximo} / 2) THEN {tope_periodo_maximo} / 2                    
                        ELSE sac
                    END AS base_calculada
                FROM
                    "items_liquidados" AS "t0"
                JOIN
                    "datos_adicionales" AS "t1" ON "t0"."id_dato_personal" = "t1"."id_dato_personal"
                WHERE
                    "periodo_mes" = {dropdown_periodo_mes.value}
                GROUP BY
                    "t0"."id_dato_personal",
                    "cuil",
                    "apellido",
                    "nombre",
                    "jubilado",
                    "convenio",
                    "obra_social",
                    "sindicato",
                    "sind_afiliado",                    
                ) AS "paso_inicial"
            GROUP BY
                "cuil",
                "apellido",
                "nombre",  
                "base_calculada",
                "desc_suss",
                "desc_os",
                "desc_sindical",            
                "rem_sin_sac",
                "sac",
                "no_rem_con_apo",
                "jubilado",
                "convenio",
                "obra_social",
                "sindicato",
                "sind_afiliado"
            ORDER BY
                CASE 
                    WHEN "ctrl_suss" AND "ctrl_os" = 0
                    THEN 1
                    ELSE 0
                END ASC,
                "ctrl_suss" DESC,
                "ctrl_os" DESC,
                "apellido" ASC,
                "nombre" ASC;
            """
        )

        resultado = con.table("control_descuentos")
        return resultado


    tabla_control_descuentos = control_descuentos()
    return (tabla_control_descuentos,)


@app.cell
def _(con, f572_path):
    def procesar_f572():
        con.raw_sql(
            """
        DROP TABLE IF EXISTS "f572_deducciones_familiares";
        DROP TABLE IF EXISTS "f572_deducciones_generales";
        DROP TABLE IF EXISTS "f572_otros_empleos";

        CREATE TABLE "f572_deducciones_familiares" (
            cuil VARCHAR,
            parentesco TINYINT,
            porcentaje TINYINT,
            mes_desde TINYINT,
            mes_hasta TINYINT
            );

        CREATE TABLE "f572_deducciones_generales" (
            cuil VARCHAR,
            tipo_deduccion TINYINT,        
            monto_total DECIMAL(14, 2)
        );

        CREATE TABLE "f572_otros_empleos" (
            cuil VARCHAR,
            concepto INT,
            monto DECIMAL(14, 2)
        );
        """
        )

        items_conyuge = [1]
        items_hijos = [3, 30]
        items_hijos_incapacitados = [31, 32]

        for xml in os.listdir(f572_path):
            if xml.endswith(".xml"):
                xml_path = os.path.join(f572_path, xml)
                with open(xml_path, "r", encoding="utf-8") as formulario:
                    contenido = formulario.read()
                    parser = bs(contenido, features="xml")

                # CUIL Empleado
                cuil = str(parser.find("empleado").find("cuit").text)

                # Tipos de deducciones
                deducciones_familiares = parser.find_all("cargaFamilia")
                deducciones_generales = parser.find_all("deduccion")

                # Obtiene deducciones familiares
                for integrante in deducciones_familiares:
                    parentesco = int(integrante.find("parentesco").text)
                    mes_desde = int(integrante.find("mesDesde").text)
                    mes_hasta = int(integrante.find("mesHasta").text)
                    porcentaje_deduccion = int(
                        integrante.find("porcentajeDeduccion").text
                    )

                    con.raw_sql(f"""
                    INSERT INTO "f572_deducciones_familiares" (
                    cuil, parentesco, porcentaje, mes_desde, mes_hasta
                    ) VALUES (
                    '{cuil}', {parentesco}, {porcentaje_deduccion}, {mes_desde}, {mes_hasta}
                    );
                    """)

                # Obtiene deducciones generales
                for deduccion in deducciones_generales:
                    tipo_deduccion = int(deduccion.get("tipo"))
                    monto_total = Decimal(deduccion.find("montoTotal").text)

                    con.raw_sql(f"""
                    INSERT INTO "f572_deducciones_generales" (
                    cuil, tipo_deduccion, monto_total
                    ) VALUES (
                    '{cuil}', {tipo_deduccion}, {monto_total}
                    );
                    """)

                # Itera por cada item de otros empleos
                # for item in parser.find_all("ingAp"):
                #     mes = int(item.get("mes"))
                #     concepto = item.find(concepto)
                #     # if concepto is not None:
                #     monto = float(concepto.text)

                #     con.raw_sql(f"""
                #     INSERT INTO "f572_otros_empleos" (
                #     cuil, concepto, monto
                #     ) VALUES (
                #     '{cuil}', {concepto}, {monto}
                #     );
                #     """)

            f572_deducciones_familiares = con.table("f572_deducciones_familiares")
            f572_otros_empleos = con.table("f572_otros_empleos")

        f572_deducciones_familiares = con.table("f572_deducciones_familiares")
        f572_deducciones_generales = con.table("f572_deducciones_generales")

        return (
            f572_deducciones_familiares,
            f572_deducciones_generales,
            items_conyuge,
            items_hijos,
            items_hijos_incapacitados,
        )


    (
        f572_deducciones_familiares,
        f572_deducciones_generales,
        items_conyuge,
        items_hijos,
        items_hijos_incapacitados,
    ) = procesar_f572()
    return (
        f572_deducciones_generales,
        items_conyuge,
        items_hijos,
        items_hijos_incapacitados,
    )


@app.cell
def _(f572_deducciones_generales):
    mo.ui.table(f572_deducciones_generales)
    return


@app.cell
def _(
    con,
    dropdown_no_habituales,
    dropdown_periodo_año,
    dropdown_periodo_mes,
    dropdown_tratamiento_sac,
    tope_alquileres_40,
    tope_casas_particulares,
    tope_fines_educativos,
    tope_gastos_sepelio,
    tope_intereses_hipotecarios,
    tope_movilidad_viaticos,
    tope_seguro_mixto,
    tope_seguro_muerte,
    tope_seguro_retiro,
):
    def control_ganancias():
        """
        Función para efectuar el control del impuesto a las ganancias
        """

        # f572_deducciones_familiares, f572_deducciones_generales, items_conyuge, items_hijos, items_hijos_incapacitados = procesar_f572()

        # Si estoy en el primer semestre...
        var_between = f"BETWEEN 1 AND {dropdown_periodo_mes.value}"

        if dropdown_no_habituales.value == "Sí":
            formula_no_habituales_1s = f"""
                SUM(
                    CASE 
                        WHEN ganancias_item = 'no habitual' AND periodo_mes BETWEEN 1 AND 6                    
                        THEN "haber" / (13 - periodo_mes) * ({dropdown_periodo_mes.value} - periodo_mes + 1)
                        ELSE 0
                    END
                )                
                """
            if dropdown_periodo_mes.value < 7:
                formula_no_habituales_2s = 0
            else:
                formula_no_habituales_2s = f"""
                    SUM(
                        CASE 
                            WHEN ganancias_item = 'no habitual' AND periodo_mes BETWEEN 7 AND 12
                            THEN "haber" / (13 - periodo_mes) * ({dropdown_periodo_mes.value} - periodo_mes + 1)
                            ELSE 0
                        END
                    )                
                    """
        else:
            formula_no_habituales_1s = f"""
                SUM("haber") FILTER(WHERE "ganancias_item" = 'no habitual' AND "periodo_mes" BETWEEN 1 AND 6)
                """
            formula_no_habituales_2s = f"""
                SUM("haber") FILTER(WHERE "ganancias_item" = 'no habitual' AND "periodo_mes" BETWEEN 7 AND 12)
                """

        # Hasta el mes 5 siempre prorrateamos independientemente del método
        if (
            dropdown_tratamiento_sac.value == "Anual"
            or dropdown_periodo_mes.value < 6
        ):
            formula_sac_1s = """
            CAST(COALESCE((habitual_gravado_1s + no_habitual_gravado_1s) / 12, 0) AS DECIMAL(16,2))
            """
            formula_sac_2s = """
            CAST(COALESCE((habitual_gravado_2s + no_habitual_gravado_2s) / 12, 0) AS DECIMAL(16,2))
            """
        else:
            formula_sac_1s = f"""       
            CAST(COALESCE(SUM("haber") FILTER(WHERE "ganancias_item" = 'sac' AND "periodo_mes" BETWEEN 1 AND 6), 0) AS DECIMAL(16, 2))
            """
            formula_sac_2s = f"""
            CAST(COALESCE(SUM("haber") FILTER(WHERE "ganancias_item" = 'sac' AND "periodo_mes" BETWEEN 7 AND 12), 0) AS DECIMAL(16, 2))
            """

        resta_deducciones_generales = """
        COALESCE((medico_asistencial - seguro_muerte - donaciones - intereses_hipotecarios - gastos_sepelio - honorarios_medicos - casas_particulares - sociedades_garantia_reciproca - viajantes_comercio - movilidad_viaticos - indumentaria_trabajo - alquileres_40 - seguro_mixto - seguro_retiro - fondos_comunes_inversion - fines_educativos - alquileres_10_locatario - alquileres_10_locador - otras_deducciones), 0)
        """

        calculo_ganancia_neta = """    
        (rem_gravada - ded_descuentos - COALESCE((seguro_muerte - intereses_hipotecarios - gastos_sepelio - casas_particulares - sociedades_garantia_reciproca - viajantes_comercio - movilidad_viaticos - indumentaria_trabajo - alquileres_40 - seguro_mixto - seguro_retiro - fondos_comunes_inversion - fines_educativos - alquileres_10_locatario - alquileres_10_locador - otras_deducciones), 0))
        """

        con.raw_sql(
            f"""
            DROP TABLE IF EXISTS "control_ganancias";

            CREATE TABLE 
                "control_ganancias" AS
            WITH
                "base" AS (
            SELECT
                "t0"."numero_documento" AS "cuil",
                "t0"."apellido",
                "t0"."nombre",            
                CAST({dropdown_periodo_mes.value} AS TINYINT) AS "mes",
                CASE
                    WHEN MONTH("t1"."fecha_baja") = {dropdown_periodo_mes.value} THEN 'Baja'
                    ELSE 'Activo'
                END AS "situacion",
                CAST(COALESCE(SUM("haber") FILTER(WHERE "ganancias_item" = 'habitual' AND "periodo_mes" BETWEEN 1 AND 6), 0) AS DECIMAL(16, 2)) AS "habitual_gravado_1s",
                CAST(COALESCE(SUM("haber") FILTER(WHERE "ganancias_item" = 'habitual' AND "periodo_mes" BETWEEN 7 AND 12), 0) AS DECIMAL(16, 2)) AS "habitual_gravado_2s", 
                CAST(COALESCE({formula_no_habituales_1s}, 0) AS DECIMAL(16, 2)) AS "no_habitual_gravado_1s",
                CAST(COALESCE({formula_no_habituales_2s}, 0) AS DECIMAL(16, 2)) AS "no_habitual_gravado_2s",                        
                CAST(COALESCE(SUM("haber") FILTER(WHERE "ganancias_item" = 'exento' AND "periodo_mes" {var_between}), 0) AS DECIMAL(16, 2)) AS "exento",            
                --{dropdown_tratamiento_sac} AS "modo_sac",
                --Al sac falta que lo tome en el mes correspondiente segun el criterio
                --CASE
                  --  WHEN "modo_sac" = S
                {formula_sac_1s} AS "sac_1s",
                {formula_sac_2s} AS "sac_2s",
                CAST(COALESCE(SUM("dscto") FILTER(WHERE "ganancias_item" IN ('previsional', 'obra social', 'sindical') AND "periodo_mes" {var_between}), 0) AS DECIMAL(16, 2)) AS "ded_descuentos",
                CAST(COALESCE(SUM("dscto") FILTER(WHERE "ganancias_item" = 'impuesto' AND "periodo_mes" {var_between}), 0) AS DECIMAL(16, 2)) AS "impuesto_retenido"
            FROM
                items_liquidados AS "t0"
            JOIN
                datos_adicionales AS "t1" ON "t0"."id_dato_personal" = "t1"."id_dato_personal"
            GROUP BY
                "t0"."id_dato_personal",
                "t0"."numero_documento",
                "t0"."apellido",
                "t0"."nombre",
                "t1"."fecha_baja"
            ),

                "intermedio" AS (
            SELECT
                "base"."cuil",
                "base"."apellido",
                "base"."nombre",            
                "base"."mes",
                "base"."situacion",
                "base"."habitual_gravado_1s",
                "base"."habitual_gravado_2s",
                "base"."no_habitual_gravado_1s",
                "base"."no_habitual_gravado_2s",
                "base"."exento",            
                "base"."sac_1s",
                "base"."sac_2s",            
                ("base"."habitual_gravado_1s" + "base"."habitual_gravado_2s" + "base"."no_habitual_gravado_1s" + "base"."no_habitual_gravado_2s" + "base"."sac_1s" + "base"."sac_2s") AS "rem_gravada",            
                "t2"."ded_gni",
                "t2"."ded_especial",                  
                COALESCE("t3"."monto", 0) AS "f572_familiares",  
                CAST(("ded_gni" + "ded_especial" + "f572_familiares") / 12 AS DECIMAL(12, 2)) AS "ded_12va",
                "base"."ded_descuentos",      
                CASE
                    WHEN "t4"."seguro_muerte" > {tope_seguro_muerte} THEN {tope_seguro_muerte}
                    ELSE COALESCE("t4"."seguro_muerte", 0)
                END *
                CASE
                    WHEN "situacion" = 'Activo' THEN 0
                    WHEN "situacion" = 'Baja' THEN 1
                END AS "seguro_muerte",                        
                CASE
                    WHEN "t4"."intereses_hipotecarios" > {tope_intereses_hipotecarios} THEN {tope_intereses_hipotecarios}
                    ELSE COALESCE("t4"."intereses_hipotecarios", 0)
                END AS "intereses_hipotecarios",
                CASE
                    WHEN "t4"."gastos_sepelio" > {tope_gastos_sepelio} THEN {tope_gastos_sepelio}
                    ELSE COALESCE("t4"."gastos_sepelio", 0)
                END AS "gastos_sepelio",            

                CASE 
                    WHEN "t4"."casas_particulares" > {tope_casas_particulares} THEN {tope_casas_particulares}
                    ELSE COALESCE("t4"."casas_particulares", 0)
                END AS "casas_particulares",   
                COALESCE("t4"."sociedades_garantia_reciproca", 0) AS "sociedades_garantia_reciproca",
                COALESCE("t4"."viajantes_comercio", 0) AS "viajantes_comercio",
                CASE
                    WHEN "t4"."movilidad_viaticos" > {tope_movilidad_viaticos} THEN {tope_movilidad_viaticos}
                    ELSE COALESCE("t4"."movilidad_viaticos", 0) 
                END AS "movilidad_viaticos",   
                COALESCE("t4"."indumentaria_trabajo", 0) AS "indumentaria_trabajo",
                CASE
                    WHEN "t4"."alquileres_40" > {tope_alquileres_40} THEN {tope_alquileres_40}
                    ELSE COALESCE("t4"."alquileres_40", 0)
                END AS "alquileres_40",
                CASE
                    WHEN "t4"."seguro_mixto" > {tope_seguro_mixto} THEN {tope_seguro_mixto}
                    ELSE COALESCE("t4"."seguro_mixto", 0)
                END *
                CASE
                    WHEN "situacion" = 'Activo' THEN 0
                    WHEN "situacion" = 'Baja' THEN 1
                END AS "seguro_mixto",   
                CASE
                    WHEN "t4"."seguro_retiro" > {tope_seguro_retiro} THEN {tope_seguro_retiro}
                    ELSE COALESCE("t4"."seguro_retiro", 0)
                END *
                CASE
                    WHEN "situacion" = 'Activo' THEN 0
                    WHEN "situacion" = 'Baja' THEN 1
                END AS "seguro_retiro",      
                COALESCE("t4"."fondos_comunes_inversion", 0) *
                CASE
                    WHEN "situacion" = 'Activo' THEN 0
                    WHEN "situacion" = 'Baja' THEN 1            
                END AS "fondos_comunes_inversion",
                CASE
                    WHEN "t4"."fines_educativos" > {tope_fines_educativos} THEN {tope_fines_educativos}
                    ELSE COALESCE("t4"."fines_educativos", 0)
                END AS "fines_educativos",
                COALESCE("t4"."alquileres_10_locatario", 0) AS "alquileres_10_locatario",
                COALESCE("t4"."alquileres_10_locador", 0) AS "alquileres_10_locador", 
                COALESCE("t4"."otras_deducciones", 0) AS "otras_deducciones",   
                CAST({calculo_ganancia_neta} * 5 / 100 AS DECIMAL(10, 2)) AS "tope_ganancia_neta",
                CASE
                    WHEN "t4"."medico_asistencial" > "tope_ganancia_neta" THEN "tope_ganancia_neta"
                    ELSE COALESCE("t4"."medico_asistencial", 0) 
                END AS "medico_asistencial",      
                CASE
                    WHEN "t4"."donaciones" > "tope_ganancia_neta" THEN "tope_ganancia_neta"
                    ELSE COALESCE("t4"."donaciones", 0) 
                END AS "donaciones",
                CASE
                    WHEN "t4"."honorarios_medicos" > "tope_ganancia_neta" THEN "tope_ganancia_neta"
                    ELSE COALESCE("t4"."honorarios_medicos", 0) 
                END *
                CASE
                    WHEN "situacion" = 'Activo' THEN 0
                    WHEN "situacion" = 'Baja' THEN 1
                END AS "honorarios_medicos",
                CASE
                    WHEN rem_gravada - ded_gni - ded_especial - ded_descuentos - f572_familiares - ded_12va - {resta_deducciones_generales} < 0 THEN 0
                    ELSE rem_gravada - ded_gni - ded_especial - ded_descuentos - f572_familiares - ded_12va - {resta_deducciones_generales}
                END AS "gnsi",                        
                "base"."impuesto_retenido"
            FROM
                "base"
            LEFT JOIN "ganancias_deducciones_personales" AS "t2"
                ON YEAR("t2"."periodo") = {dropdown_periodo_año.value}              
                AND (
                    ("base"."situacion" = 'Activo'
                    AND CAST(MONTH("t2"."periodo") AS TINYINT) = "base"."mes"
                    AND "t2"."tipo" = 0)
                OR
                    ("base"."situacion" = 'Baja'
                    AND "base"."mes" <= 6
                    AND CAST(MONTH("t2"."periodo") AS TINYINT) = 12
                    AND "t2"."tipo" = 1)
                OR
                    ("base"."situacion" = 'Baja'
                    AND "base"."mes" >= 7
                    AND CAST(MONTH("t2"."periodo") AS TINYINT) = 12
                    AND "t2"."tipo" = 2)
                )
            LEFT JOIN "deducciones_familiares_agrupadas" AS "t3"
                ON REPLACE("base"."cuil", '-', '') = "t3"."cuil"            
            LEFT JOIN "deducciones_generales_procesadas" AS "t4"
                ON REPLACE("base"."cuil", '-', '') = "t4"."cuil"      
            ORDER BY
                "apellido" ASC,
                "nombre" ASC
                )

            SELECT
                "intermedio"."cuil",
                "intermedio"."apellido",
                "intermedio"."nombre",            
                "intermedio"."mes",
                "intermedio"."situacion",
                "intermedio"."habitual_gravado_1s",
                "intermedio"."habitual_gravado_2s",
                "intermedio"."no_habitual_gravado_1s",
                "intermedio"."no_habitual_gravado_2s",
                "intermedio"."exento",            
                "intermedio"."sac_1s",
                "intermedio"."sac_2s",            
                "intermedio"."rem_gravada",
                "intermedio"."ded_gni",
                "intermedio"."ded_especial",            
                "intermedio"."f572_familiares",
                "intermedio"."ded_12va",
                "intermedio"."ded_descuentos",
                "intermedio"."medico_asistencial",
                "intermedio"."seguro_muerte",
                "intermedio"."donaciones",
                "intermedio"."intereses_hipotecarios",
                "intermedio"."gastos_sepelio",
                "intermedio"."honorarios_medicos",
                "intermedio"."casas_particulares",   
                "intermedio"."sociedades_garantia_reciproca",
                "intermedio"."viajantes_comercio",
                "intermedio"."movilidad_viaticos",   
                "intermedio"."indumentaria_trabajo",
                "intermedio"."alquileres_40",
                "intermedio"."seguro_mixto",
                "intermedio"."seguro_retiro",   
                "intermedio"."fondos_comunes_inversion",
                "intermedio"."fines_educativos",
                "intermedio"."alquileres_10_locatario",
                "intermedio"."alquileres_10_locador", 
                "intermedio"."otras_deducciones",   
                "intermedio"."tope_ganancia_neta",
                "intermedio"."gnsi",            
                COALESCE("t5"."suma_fija", 0) AS "suma_fija",           
                COALESCE("t5"."coeficiente", 0) AS "coeficiente",
                COALESCE("t5"."excedente", 0) AS "excedente",
                COALESCE(CAST(("intermedio"."gnsi" - "t5"."excedente") * "t5"."coeficiente" / 100 AS DECIMAL(14, 2)), 0) AS "suma_variable",
                COALESCE("suma_fija" + "suma_variable", 0) AS "impuesto_determinado",
                "intermedio"."impuesto_retenido",
                "impuesto_determinado" - "impuesto_retenido" AS "saldo"
            FROM
                "intermedio"
            LEFT JOIN "ganancias_art_94" AS "t5"
                ON YEAR("t5"."Periodo") = {dropdown_periodo_año.value}
                AND (
                    ("intermedio"."situacion" = 'Activo'
                    AND CAST(MONTH("t5"."periodo") AS TINYINT) = "intermedio"."mes"
                    AND "t5"."tipo" = 0)
                OR
                    ("intermedio"."situacion" = 'Baja'
                    AND "intermedio"."mes" <= 6
                    AND CAST(MONTH("t5"."periodo") AS TINYINT) = 12
                    AND "t5"."tipo" = 1)
                OR
                    ("intermedio"."situacion" = 'Baja'
                    AND "intermedio"."mes" >= 7
                    AND CAST(MONTH("t5"."periodo") AS TINYINT) = 12
                    AND "t5"."tipo" = 2)                
                )
                AND "intermedio"."GNSI" > "t5"."desde" AND "intermedio"."GNSI" <= "t5"."hasta" 
            """
        )

        resultado = con.table("control_ganancias")
        return resultado


    tabla_control_ganancias = control_ganancias()
    return (tabla_control_ganancias,)


@app.cell
def _(con, dropdown_periodo_mes, tope_periodo_maximo):
    def control_f931():
        """
        Función para efectual el control del F.931 ARCA
        """

        remuneraciones_sin_sac = f"""
        COALESCE(SUM("haber") FILTER(WHERE "categoria_item" = 'remunerativo' AND "subcategoria_item" != 'sac' AND "periodo_mes" = {dropdown_periodo_mes.value}), 0)
        """

        remuneraciones_solo_sac = f"""
        COALESCE(SUM("haber") FILTER(WHERE "categoria_item" = 'remunerativo' AND "subcategoria_item" = 'sac' AND "periodo_mes" = {dropdown_periodo_mes.value}), 0)
        """

        # De momento no tiene en cuenta los no remunerativos con aportes de obra social, pendiente de implementar
        remuneraciones_obra_social_sin_sac = f"""
        COALESCE(SUM("haber") FILTER(WHERE "categoria_item" = 'remunerativo' AND "subcategoria_item" != 'sac' AND "periodo_mes" = {dropdown_periodo_mes.value}), 0)
        """

        remuneraciones_obra_social_solo_sac = f"""
        COALESCE(SUM("haber") FILTER(WHERE "categoria_item" = 'remunerativo' AND "subcategoria_item" = 'sac' AND "periodo_mes" = {dropdown_periodo_mes.value}), 0)
        """

        remuneraciones_art = f"""
        COALESCE(SUM("haber") FILTER(WHERE "r9" = true AND "periodo_mes" = {dropdown_periodo_mes.value}), 0)
        """

        con.raw_sql(f"""
        DROP TABLE IF EXISTS "control_f931";

        CREATE TABLE "control_f931" AS
        SELECT
            "numero_documento" AS "cuil",
            "apellido",
            "nombre",
            CAST((CASE
                WHEN {remuneraciones_sin_sac} > {tope_periodo_maximo} THEN {tope_periodo_maximo}
                ELSE {remuneraciones_sin_sac}
            END +
            CASE
                WHEN {remuneraciones_solo_sac} > ({tope_periodo_maximo} / 2) THEN ({tope_periodo_maximo} / 2)
                ELSE {remuneraciones_solo_sac}
            END) AS DECIMAL(16, 2)) AS "r1",
            CAST({remuneraciones_sin_sac} + {remuneraciones_solo_sac} AS DECIMAL(16, 2)) AS "r2",
            "r2" AS "r3",
            CAST((CASE
                WHEN {remuneraciones_obra_social_sin_sac} > {tope_periodo_maximo} THEN {tope_periodo_maximo}
                ELSE {remuneraciones_obra_social_sin_sac}
            END +
            CASE
                WHEN {remuneraciones_obra_social_solo_sac} > ({tope_periodo_maximo} / 2) THEN ({tope_periodo_maximo} / 2)
                ELSE {remuneraciones_obra_social_solo_sac}
            END) AS DECIMAL(16, 2)) AS "r4",
            CAST(0 AS DECIMAL(2, 2)) as "r6",
            "r6" AS "r7",
            CAST({remuneraciones_obra_social_sin_sac} + {remuneraciones_obra_social_solo_sac} AS DECIMAL(16, 2)) AS "r8",
            CAST({remuneraciones_art} AS DECIMAL(16, 2)) AS "r9",
            "r2" - 7003.68 AS "r10",
            CAST(COALESCE(SUM("haber") FILTER(WHERE "categoria_item" = 'no remunerativo' AND "periodo_mes" = {dropdown_periodo_mes.value}), 0) AS DECIMAL(16, 2)) AS "no_rem",
            CAST(COALESCE(SUM("haber") FILTER(WHERE "periodo_mes" = {dropdown_periodo_mes.value}), 0) AS DECIMAL(16, 2)) AS "rem_total"
        FROM
            "items_liquidados"
        GROUP By
            "numero_documento",
            "apellido",
            "nombre"
        """)

        resultado = con.table("control_f931")
        return resultado


    tabla_control_f931 = control_f931()
    return (tabla_control_f931,)


@app.cell
def _(
    button_exportar_items_faltantes,
    button_exportar_items_todos,
    button_exportar_liquidaciones,
    items_parametrizados,
    items_sin_parametrizar,
    tabla_control_descuentos,
    tabla_control_f931,
    tabla_control_ganancias,
    tabla_items_liquidados,
    tabla_resumen_liquidaciones,
    tabla_resumen_liquidaciones_agrupadas,
):
    tabs_controles = mo.accordion(
        {
            "Resúmen de importes": mo.vstack(
                [
                    mo.ui.table(
                        tabla_resumen_liquidaciones,
                        page_size=15,
                        selection="multi",
                        show_column_summaries=False,
                        show_download=True,
                    ),
                ]
            ),
            "Resúmen de importes (agrupados)": mo.vstack(
                [
                    mo.ui.table(
                        tabla_resumen_liquidaciones_agrupadas,
                        page_size=1,
                        selection="single",
                        show_column_summaries=False,
                        show_download=True,
                    ),
                ]
            ),
            "Control de descuentos": mo.vstack(
                [
                    mo.ui.table(
                        tabla_control_descuentos,
                        page_size=15,
                        selection="multi",
                        show_column_summaries=False,
                        show_download=True,
                    )
                ]
            ),
            "Control Imp. Ganancias": mo.vstack(
                [
                    mo.ui.table(
                        tabla_control_ganancias,
                        page_size=15,
                        selection="multi",
                        show_column_summaries=False,
                        show_download=True,
                    )
                ]
            ),
            "Control F.931 / LSD": mo.vstack(
                [
                    mo.ui.table(
                        tabla_control_f931,
                        page_size=15,
                        selection="multi",
                        show_column_summaries=False,
                        show_download=True,
                    )
                ]
            ),
        }
    )

    tabs_items = mo.accordion(
        {
            "Items parametrizados": mo.vstack(
                [
                    mo.ui.table(
                        items_parametrizados,
                        page_size=20,
                        selection="multi",
                        show_column_summaries=False,
                        show_download=True,
                    ),
                ]
            ),
            "Items sin parametrizar": mo.vstack(
                [
                    mo.hstack(
                        [
                            button_exportar_items_todos,
                            button_exportar_items_faltantes,
                        ],
                        justify="start",
                    ),
                    mo.ui.table(
                        items_sin_parametrizar,
                        page_size=20,
                        selection="multi",
                        show_column_summaries=False,
                        show_download=True,
                    ),
                ]
            ),
            "Items condicionales": mo.ui.data_editor(items_parametrizados),
        }
    )

    tabs = mo.ui.tabs(
        {
            "Controles": tabs_controles,
            "Datos de orígen": mo.vstack(
                [
                    button_exportar_liquidaciones,
                    mo.ui.table(tabla_items_liquidados),
                ]
            ),
            "Items": tabs_items,
        },
    )

    tabs
    return


@app.cell
def _(button_exportar_items_todos, con):
    mo.stop(not button_exportar_items_todos.value)

    archivo_todos = "items_exportados_todos.xlsx"

    pd.DataFrame(con.table("items_liquidacion")).to_excel(
        archivo_todos,
        index=False,
        header=[
            "idLiquidacionItem",
            "Nombre",
        ],
        engine="xlsxwriter",
    )

    os.startfile(archivo_todos)
    return


@app.cell
def _(button_exportar_items_faltantes, con):
    mo.stop(not button_exportar_items_faltantes.value)

    archivo_faltantes = "items_exportados_faltantes.xlsx"

    try:
        pd.DataFrame(con.table("items_sin_parametrizar")).to_excel(
            archivo_faltantes,
            index=False,
            header=[
                "idLiquidacionItem",
                "Nombre",
            ],
            engine="xlsxwriter",
        )
    except Exception as e:
        f"Error {e}"

    os.startfile(archivo_faltantes)
    return


@app.cell
def _(button_exportar_liquidaciones, con):
    mo.stop(not button_exportar_liquidaciones.value)

    archivo_liquidaciones = "liquidaciones.xlsx"

    try:
        pd.DataFrame(con.table("items_liquidados")).to_excel(
            archivo_liquidaciones,
            index=False,
            header=[
                "id_dato_personal",
                "apellido",
                "nombre",
                "id_liquidacion",
                "id_liquidacion_tipo",
                "id_liquidacion_item",
                "cantidad",
                "haber",
                "dscto",
                "nombre_item",
                "periodo_mes",
                "numero_documento",
                "idLiquidacionItem",
                "nombre_item",
                "categoria_item",
                "subcategoria_item",
                "ganancias_item",
                "r4r8",
                "r9",
            ],
            engine="xlsxwriter",
        )
    except Exception as e:
        f"Error {e}"

    os.startfile(archivo_liquidaciones)
    return


@app.cell
def generacion_excel(
    button_generar_control,
    tabla_control_descuentos,
    tabla_control_ganancias,
    tabla_resumen_liquidaciones,
    tabla_resumen_liquidaciones_agrupadas,
    tope_periodo_maximo,
    tope_periodo_minimo,
):
    # Generación de hoja de cáclulo de control
    mo.stop(not button_generar_control.value)

    # Inicializa el escritor
    escritor = pd.ExcelWriter("control.xlsx", engine="xlsxwriter")

    # Inicializa el libro
    libro = escritor.book

    # Formatos a usar
    # Fuente tamaño 10
    formato_fuente_10 = libro.add_format({"font_size": 10})
    # Fuente tamaño 10 texto centrado
    formato_fuente_10_centrado = libro.add_format(
        {"font_size": 10, "align": "center"}
    )
    # Formato para los encabezados
    formato_encabezados = libro.add_format(
        {"font_size": 10, "align": "left", "bold": "true"}
    )
    # Crea formato de texto
    formato_texto = libro.add_format({"num_format": "@", "font_size": 10})
    # Crea formato de contabilidad
    formato_contable = libro.add_format(
        {"num_format": "#,##0.00", "font_size": 10}
    )
    # Crea formato de porcentaje
    formato_porcentaje = libro.add_format(
        {"num_format": "0%", "font_size": 10, "align": "center"}
    )
    # Formato alerta
    formato_alerta = libro.add_format({"bg_color": "#FFC000"})

    # Formato "es correcto" o "coincide"
    formato_coincide = libro.add_format(
        {"bg_color": "#C6EFCE", "font_color": "#006100"}
    )
    # Formato "a devolver"
    formato_devolucion = libro.add_format(
        {"bg_color": "#B7DEE8", "font_color": "#215967"}
    )
    # Formato "es erróneo" o "diferencia"
    formato_diferencia = libro.add_format(
        {"bg_color": "#FFC7CE", "font_color": "#9C0006"}
    )

    #########################################
    # EXPORTACION DE RESUMEN DE LIQUIDACIONES#
    #########################################

    # Inicializa la hoja y datos para resúmen de importes liquidados
    datos_resumen_liquidaciones = tabla_resumen_liquidaciones.to_pandas()
    datos_resumen_liquidaciones.to_excel(
        escritor,
        index=False,
        sheet_name="Resumen de importes liquidados",
        float_format="%.2f",
    )
    hoja_resumen_liquidaciones = escritor.sheets["Resumen de importes liquidados"]

    longitud_columnas = {
        "cuil": 14,
        "apellido": 16,
        "nombre": 18,
        "remunerativo": 14,
        "no_remunerativo": 16,
        "previsional": 14,
        "obra_social": 14,
        "sindical": 14,
        "anticipos": 14,
        "embargos": 14,
        "otros_descuentos": 16,
        "imp_ganancias": 14,
        "neto": 14,
    }

    nombres_encabezados = [
        "CUIL",
        "Apellido",
        "Nombres",
        "Remunerativo",
        "No Remunerativo",
        "Previsional",
        "Obra Social",
        "Sindical",
        "Anticipos",
        "Embargos",
        "Otros Descuentos",
        "Imp. Ganancias",
        "Neto",
    ]

    columnas_formato_texto = [
        "cuil",
        "apellido",
        "nombre",
    ]

    columnas_formato_contable = [
        "remunerativo",
        "no_remunerativo",
        "previsional",
        "obra_social",
        "sindical",
        "anticipos",
        "embargos",
        "otros_descuentos",
        "imp_ganancias",
        "neto",
    ]

    datos_resumen_liquidaciones_agrupadas = (
        tabla_resumen_liquidaciones_agrupadas.to_pandas().transpose().reset_index()
    )

    hoja_resumen_liquidaciones.add_table(
        0,
        14,
        10,
        16,
        {
            "name": "totales_agrupados",
            "columns": [
                {"header": "Concepto"},
                {"header": "Importe"},
                {"header": "Acumulativo"},
            ],
            "style": "Table Style Medium 21",
        },
    )

    for fila in range(len(datos_resumen_liquidaciones_agrupadas)):
        for columna in range(len(datos_resumen_liquidaciones_agrupadas.columns)):
            hoja_resumen_liquidaciones.write(
                fila + 1,
                columna + 14,
                datos_resumen_liquidaciones_agrupadas.iloc[fila, columna],
                formato_encabezados,
            )

            if columna == 1:
                hoja_resumen_liquidaciones.write(
                    fila + 1,
                    columna + 14,
                    datos_resumen_liquidaciones_agrupadas.iloc[fila, columna],
                    formato_contable,
                )

    hoja_resumen_liquidaciones.write("O1", "Concepto", formato_encabezados)
    hoja_resumen_liquidaciones.write("P1", "Importe", formato_encabezados)
    hoja_resumen_liquidaciones.write("Q1", "Acumulativo", formato_encabezados)

    hoja_resumen_liquidaciones.write("O2", "Remunerativo", formato_encabezados)
    hoja_resumen_liquidaciones.write("O3", "No remunerativo", formato_encabezados)
    hoja_resumen_liquidaciones.write("O4", "Previsional", formato_encabezados)
    hoja_resumen_liquidaciones.write("O5", "Obra social", formato_encabezados)
    hoja_resumen_liquidaciones.write("O6", "Sindical", formato_encabezados)
    hoja_resumen_liquidaciones.write("O7", "Anticipos", formato_encabezados)
    hoja_resumen_liquidaciones.write("O8", "Embargos", formato_encabezados)
    hoja_resumen_liquidaciones.write("O9", "Otros descuentos", formato_encabezados)
    hoja_resumen_liquidaciones.write("O10", "Imp. ganancias", formato_encabezados)
    hoja_resumen_liquidaciones.write("O11", "Neto", formato_encabezados)

    hoja_resumen_liquidaciones.set_column("O:O", 15)
    hoja_resumen_liquidaciones.set_column("P:P", 15)
    hoja_resumen_liquidaciones.set_column("Q:Q", 15)

    hoja_resumen_liquidaciones.write_formula("Q2", "=P2", formato_contable)
    hoja_resumen_liquidaciones.write_formula("Q3", "=Q2+P3", formato_contable)
    hoja_resumen_liquidaciones.write_formula("Q4", "=Q3-P4", formato_contable)
    hoja_resumen_liquidaciones.write_formula("Q5", "=Q4-P5", formato_contable)
    hoja_resumen_liquidaciones.write_formula("Q6", "=Q5-P6", formato_contable)
    hoja_resumen_liquidaciones.write_formula("Q7", "=Q6-P7", formato_contable)
    hoja_resumen_liquidaciones.write_formula("Q8", "=Q7-P8", formato_contable)
    hoja_resumen_liquidaciones.write_formula("Q9", "=Q8-P9", formato_contable)
    hoja_resumen_liquidaciones.write_formula("Q10", "=Q9-P10", formato_contable)
    hoja_resumen_liquidaciones.write_formula("Q11", "=Q10-P11", formato_contable)

    tabla_excel_resumen_liquidaciones = {
        "data": datos_resumen_liquidaciones.values.tolist(),
        "columns": [{"header": columna} for columna in nombres_encabezados],
        "autofilter": "true",
        "name": "resumen_liquidaciones",
        "style": "Table Style Medium 16",
    }
    hoja_resumen_liquidaciones.add_table(
        0,
        0,
        len(datos_resumen_liquidaciones),
        len(datos_resumen_liquidaciones.columns) - 1,
        tabla_excel_resumen_liquidaciones,
    )

    for columna_indice, nombre_columna in enumerate(
        datos_resumen_liquidaciones.columns
    ):
        if nombre_columna in columnas_formato_contable:
            hoja_resumen_liquidaciones.set_column(
                columna_indice,
                columna_indice,
                longitud_columnas[nombre_columna],
                formato_contable,
            )
        elif nombre_columna in columnas_formato_texto:
            hoja_resumen_liquidaciones.set_column(
                columna_indice,
                columna_indice,
                longitud_columnas[nombre_columna],
                formato_texto,
            )
        else:
            hoja_resumen_liquidaciones.set_column(
                columna_indice,
                columna_indice,
                longitud_columnas.get(nombre_columna, 15),
                formato_fuente_10,
            )

    hoja_resumen_liquidaciones.conditional_format(
        "Q11:Q11",
        {
            "type": "cell",
            "criteria": "==",
            "value": 0,
            "format": formato_coincide,
        },
    )

    hoja_resumen_liquidaciones.conditional_format(
        "Q11:Q11",
        {
            "type": "cell",
            "criteria": "<>",
            "value": 0,
            "format": formato_diferencia,
        },
    )

    hoja_resumen_liquidaciones.freeze_panes(1, 0)

    ######################################
    #EXPORTACION DE CONTROL DE DESCUENTOS#
    ######################################

    # Inicializa la hoja y datos para control de descuentos
    datos_control_descuentos = tabla_control_descuentos.to_pandas()
    datos_control_descuentos.to_excel(
        escritor,
        index=False,
        sheet_name="Control de descuentos",
        float_format="%.2f",
    )
    hoja_control_descuentos = escritor.sheets["Control de descuentos"]

    total_filas = datos_control_descuentos.shape[0] + 1

    longitud_columnas = {
        "cuil": 12,
        "apellido": 10,
        "nombre": 13,
        "base_calculada": 11,
        "ctrl_suss": 10,
        "ctrl_os": 9,
        "desc_suss": 10,
        "desc_os": 10,
        "desc_sindical": 11,
        "rem_sin_sac": 13,
        "sac": 11,
        "no_rem_con_apo": 11,
        "jubilado": 10,
        "convenio": 10,
        "obra_social": 7,
        "sindicato": 10,
        "sind_afiliado": 10,
    }

    nombres_encabezados = [
        "CUIL",
        "Apellido",
        "Nombres",
        "Base calc.",
        "Ctrl SUSS",
        "Ctrl OS",
        "SUSS Liq.",
        "OS Liq.",
        "Sind. Liq.",
        "Rem. sin SAC",
        "SAC",
        "Nr. c/apo.",
        "Jubilado?",
        "Convenio",
        "RNAS",
        "Sindicato",
        "Afiliado?",
    ]

    columnas_formato_texto = [
        "cuil",
        "apellido",
        "nombre",
        "jubilado",
        "convenio",
        "obra_social",
        "sindicato",
        "sind_afiliado",
    ]

    columnas_formato_contable = [
        "base_calculada",
        "ctrl_suss",
        "ctrl_os",
        "desc_suss",
        "desc_os",
        "desc_sindical",
        "rem_sin_sac",
        "sac",
        "no_rem_con_apo",
    ]

    tabla_excel_control_descuentos = {
        "data": datos_control_descuentos.values.tolist(),
        "columns": [{"header": columna} for columna in nombres_encabezados],
        "autofilter": "true",
        "name": "control_descuentos",
        "style": "Table Style Medium 17",
    }
    hoja_control_descuentos.add_table(
        0,
        0,
        len(datos_control_descuentos),
        len(datos_control_descuentos.columns) - 1,
        tabla_excel_control_descuentos,
    )

    for columna_indice, nombre_columna in enumerate(
        datos_control_descuentos.columns
    ):
        if nombre_columna in columnas_formato_contable:
            hoja_control_descuentos.set_column(
                columna_indice,
                columna_indice,
                longitud_columnas[nombre_columna],
                formato_contable,
            )
        elif nombre_columna in columnas_formato_texto:
            hoja_control_descuentos.set_column(
                columna_indice,
                columna_indice,
                longitud_columnas[nombre_columna],
                formato_texto,
            )
        else:
            hoja_control_descuentos.set_column(
                columna_indice,
                columna_indice,
                longitud_columnas.get(nombre_columna, 15),
                formato_fuente_10,
            )

    hoja_control_descuentos.conditional_format(
        f"E2:F{total_filas}",
        {
            "type": "cell",
            "criteria": "==",
            "value": 0,
            "format": formato_coincide,
        },
    )

    hoja_control_descuentos.conditional_format(
        f"E2:F{total_filas}",
        {
            "type": "cell",
            "criteria": "<=",
            "value": 0.1,
            "format": formato_devolucion,
        },
    )

    hoja_control_descuentos.conditional_format(
        f"E2:F{total_filas}",
        {
            "type": "cell",
            "criteria": ">=",
            "value": 0.1,
            "format": formato_diferencia,
        },
    )

    hoja_control_descuentos.freeze_panes(1, 0)

    # Imprimimos datos de referencia: tope remuneración, tope sac, tope adherentes, mínimo OS
    hoja_control_descuentos.write("S1", "Tope:", formato_encabezados)
    hoja_control_descuentos.write("U1", "Mínimo:", formato_encabezados)
    hoja_control_descuentos.write("W1", "Tope SAC:", formato_encabezados)
    hoja_control_descuentos.write("Y1", "Min. O.S.:", formato_encabezados)

    hoja_control_descuentos.write(
        "T1", Decimal(tope_periodo_maximo), formato_contable
    )
    hoja_control_descuentos.write(
        "V1", Decimal(tope_periodo_minimo), formato_contable
    )
    hoja_control_descuentos.write(
        "X1",
        Decimal(tope_periodo_maximo).quantize(Decimal("0.01"), ROUND_HALF_UP) / 2,
        formato_contable,
    )
    hoja_control_descuentos.write(
        "Z1",
        Decimal(tope_periodo_minimo).quantize(Decimal("0.01"), ROUND_HALF_UP) * 2,
        formato_contable,
    )

    hoja_control_descuentos.set_column("R:R", 0.5)
    hoja_control_descuentos.set_column("S:S", 5)
    hoja_control_descuentos.set_column("T:T", 11)
    hoja_control_descuentos.set_column("U:U", 7)
    hoja_control_descuentos.set_column("V:V", 9)
    hoja_control_descuentos.set_column("W:W", 8)
    hoja_control_descuentos.set_column("X:X", 11)
    hoja_control_descuentos.set_column("Y:Y", 9)
    hoja_control_descuentos.set_column("Z:Z", 9)

    ######################################
    #EXPORTACION DE CONTROL DE GANANCIAS #
    ######################################

    # Inicializa la hoja y datos para control de descuentos
    datos_control_ganancias = tabla_control_ganancias.to_pandas()

    orden_columnas = ["saldo", "cuil", "apellido", "nombre", "mes", "situacion",   "habitual_gravado_1s", "habitual_gravado_2s", "no_habitual_gravado_1s",    "no_habitual_gravado_2s", "exento", "sac_1s", "sac_2s", "rem_gravada", "ded_gni",
    "ded_especial", "ded_descuentos", "f572_familiares", "ded_12va", "medico_asistencial",
    "seguro_muerte", "donaciones", "intereses_hipotecarios", "gastos_sepelio",   "honorarios_medicos", "casas_particulares", "sociedades_garantia_reciproca",   "viajantes_comercio", "movilidad_viaticos", "indumentaria_trabajo", "alquileres_40",
    "seguro_mixto", "seguro_retiro", "fondos_comunes_inversion", "fines_educativos",
    "alquileres_10_locatario", "alquileres_10_locador", "otras_deducciones", "tope_ganancia_neta",
    "gnsi", "suma_fija", "coeficiente", "excedente", "suma_variable", "impuesto_determinado",
    "impuesto_retenido"]

    datos_control_ganancias[orden_columnas].to_excel(
        escritor,
        index=False,
        sheet_name="Imp. Ganancias",
        float_format="%.2f",
    )
    hoja_control_ganancias = escritor.sheets["Imp. Ganancias"]

    total_filas = datos_control_ganancias.shape[0] + 1

    longitud_columnas = {    
        "cuil": 12,
        "apellido": 10,
        "nombre": 13,            
        "mes": 3,
        "situacion": 6,
        "habitual_gravado_1s": 10,
        "habitual_gravado_2s": 10,
        "no_habitual_gravado_1s": 10,
        "no_habitual_gravado_2s": 10,
        "exento": 10,            
        "sac_1s": 10,
        "sac_2s": 10,            
        "rem_gravada": 10,
        "ded_gni": 10,
        "ded_especial": 10,
        "f572_familiares": 10,
        "ded_12va": 10,
        "ded_descuentos": 10,    
        "medico_asistencial": 10,
        "seguro_muerte": 10,
        "donaciones": 10,
        "intereses_hipotecarios": 10,
        "gastos_sepelio": 10,
        "honorarios_medicos": 10,
        "casas_particulares": 10,   
        "sociedades_garantia_reciproca": 10,
        "viajantes_comercio": 10,
        "movilidad_viaticos": 10,   
        "indumentaria_trabajo": 10,
        "alquileres_40": 10,
        "seguro_mixto": 10,
        "seguro_retiro": 10,   
        "fondos_comunes_inversion": 10,
        "fines_educativos": 10,
        "alquileres_10_locatario": 10,
        "alquileres_10_locador": 10, 
        "otras_deducciones": 10,   
        "tope_ganancia_neta": 10,
        "gnsi": 10,            
        "suma_fija": 10,           
        "coeficiente": 10,
        "excedente": 10,
        "suma_variable": 10,
        "impuesto_determinado": 10,
        "impuesto_retenido": 10,
        "saldo": 10
    }

    nombres_encabezados = [    
        "CUIL",
        "Apellido",
        "Nombre",
        "Mes",
        "Situacion",
        "Habitual_gravado_1s",
        "Habitual_gravado_2s",
        "No_habitual_gravado_1s",
        "No_habitual_gravado_2s",
        "Exento" ,
        "Sac_1s",
        "Sac_2s",
        "Total_rem_gravada",
        "Ded_gni",
        "Ded_especial",
        "Ded_descuentos",
        "F572_familiares",
        "Ded_12va",
        "Medico_asistencial",
        "Seguro_muerte",
        "Donaciones",
        "Intereses_hipotecarios",
        "Gastos_sepelio",
        "Honorarios_medicos",
        "Casas_particulares",
        "Sociedades_garantia_reciproca",
        "Viajantes_comercio",
        "Movilidad_viaticos",
        "Indumentaria_trabajo",
        "Alquileres_40",
        "Seguro_mixto",
        "Seguro_retiro",
        "Fondos_comunes_inversion",
        "Fines_educativos",
        "Alquileres_10_locatario",
        "Alquileres_10_locador",
        "Otras_deducciones",
        "Tope_ganancia_neta",
        "Gnsi",
        "Suma_fija",
        "Coeficiente",
        "Excedente",
        "Suma_variable",
        "Impuesto_determinado",
        "Impuesto_retenido",
        "Saldo"
    ]

    columnas_formato_texto = [
        "cuil",
        "apellido",
        "nombre",
        "mes",
        "situacion"
    ]

    columnas_formato_contable = [
        "habitual_gravado_1s",
        "habitual_gravado_2s",
        "no_habitual_gravado_1s",
        "no_habitual_gravado_2s",
        "exento" ,
        "sac_1s",
        "sac_2s",
        "total_rem_gravada",
        "ded_gni",
        "ded_especial",    
        "f572_familiares",
        "ded_12va",
        "ded_descuentos",
        "medico_asistencial",
        "seguro_muerte",
        "donaciones",
        "intereses_hipotecarios",
        "gastos_sepelio",
        "honorarios_medicos",
        "casas_particulares",
        "sociedades_garantia_reciproca",
        "viajantes_comercio",
        "movilidad_viaticos",
        "indumentaria_trabajo",
        "alquileres_40",
        "seguro_mixto",
        "seguro_retiro",
        "fondos_comunes_inversion",
        "fines_educativos",
        "alquileres_10_locatario",
        "alquileres_10_locador",
        "otras_deducciones",
        "tope_ganancia_neta",
        "gnsi",
        "suma_fija",
        "coeficiente",
        "excedente",
        "suma_variable",
        "impuesto_determinado",
        "impuesto_retenido",
        "saldo",
    ]

    tabla_excel_control_ganancias = {
        "data": datos_control_ganancias.values.tolist(),
        "columns": [{"header": columna} for columna in nombres_encabezados],
        "autofilter": "true",
        "name": "control_ganancias",
        "style": "Table Style Medium 18",
    }
    hoja_control_ganancias.add_table(
        0,
        0,
        len(datos_control_ganancias),
        len(datos_control_ganancias.columns) - 1,
        tabla_excel_control_ganancias,
    )

    for columna_indice, nombre_columna in enumerate(
        datos_control_ganancias.columns
    ):
        if nombre_columna in columnas_formato_contable:
            hoja_control_ganancias.set_column(
                columna_indice,
                columna_indice,
                longitud_columnas[nombre_columna],
                formato_contable,
            )
        elif nombre_columna in columnas_formato_texto:
            hoja_control_ganancias.set_column(
                columna_indice,
                columna_indice,
                longitud_columnas[nombre_columna],
                formato_texto,
            )
        else:
            hoja_control_ganancias.set_column(
                columna_indice,
                columna_indice,
                longitud_columnas.get(nombre_columna, 15),
                formato_fuente_10,
            )

    # hoja_control_ganancias.conditional_format(
    #     f"E2:F{total_filas}",
    #     {
    #         "type": "cell",
    #         "criteria": "==",
    #         "value": 0,
    #         "format": formato_coincide,
    #     },
    # )

    # hoja_control_ganancias.conditional_format(
    #     f"E2:F{total_filas}",
    #     {
    #         "type": "cell",
    #         "criteria": "<=",
    #         "value": 0.1,
    #         "format": formato_devolucion,
    #     },
    # )

    # hoja_control_ganancias.conditional_format(
    #     f"E2:F{total_filas}",
    #     {
    #         "type": "cell",
    #         "criteria": ">=",
    #         "value": 0.1,
    #         "format": formato_diferencia,
    #     },
    # )

    hoja_control_ganancias.freeze_panes(1, 0)

    # Cierra el escritor
    escritor.close()
    return


@app.cell
async def sidebar():
    # Carga la barra lateral de navegación
    from navegacion import app

    sidebar = await app.embed()

    sidebar.output
    return


if __name__ == "__main__":
    app.run()
