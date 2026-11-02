"""
CAPA 3 (PERSISTENCIA - GENERADOR SQL): consultas_sql.py
=============================================================================
Centraliza el 100% de las sentencias SQL del sistema de fiscalización de
hidrocarburos. Todos los predicados sobre FECHA_INSUMO e ID_CONTRATO son
SARGables (usan directamente los índices B-Tree idx_fecha e idx_contrato).
"""


def sql_select_usuarios() -> str:
    return (
        "SELECT cveUsuario, nombreUsuario, aPaterno, aMaterno, cvePerfil, "
        "emailUsuario, cveEstatus FROM USUARIOS ORDER BY cveUsuario;"
    )


def sql_select_contratos() -> str:
    return (
        "SELECT idContrato, cveOperador, modalidadFiscal, tasaRegalia, "
        "limiteRecuperacionCostos, cuencaPetrolera FROM CONTRATOS_HIDROCARBUROS ORDER BY idContrato;"
    )


def sql_insert_contrato() -> str:
    return (
        "INSERT INTO CONTRATOS_HIDROCARBUROS (idContrato, cveOperador, modalidadFiscal, "
        "tasaRegalia, limiteRecuperacionCostos, cuencaPetrolera) VALUES (?, ?, ?, ?, ?, ?);"
    )


def sql_delete_insumo_por_fecha(nombre_tabla: str) -> str:
    tbl = nombre_tabla.replace("]", "").strip()
    return f"DELETE FROM [{tbl}] WHERE [FECHA_INSUMO] = ?;"


def sql_delete_consolidada_por_fecha_y_regimen() -> str:
    return "DELETE FROM PADRON_CONTRAPRESTACIONES_CONSOLIDADO WHERE [FECHA_INSUMO] = ? AND [REGIMEN] = ?;"


def sql_copiar_insumo_a_consolidada(nombre_tabla_origen: str) -> str:
    tbl = nombre_tabla_origen.replace("]", "").strip()
    return (
        f"INSERT INTO PADRON_CONTRAPRESTACIONES_CONSOLIDADO "
        f"SELECT * FROM [{tbl}] WHERE [FECHA_INSUMO] = ?;"
    )


def sql_contar_tres_insumos_por_fecha() -> str:
    return (
        "SELECT 1 AS RegimenIdx, COUNT(*) AS TotalReg FROM BD_INSUMO_ASIGNACIONES WHERE [FECHA_INSUMO] = ? "
        "UNION ALL "
        "SELECT 2 AS RegimenIdx, COUNT(*) AS TotalReg FROM BD_INSUMO_LICENCIAS WHERE [FECHA_INSUMO] = ? "
        "UNION ALL "
        "SELECT 3 AS RegimenIdx, COUNT(*) AS TotalReg FROM BD_INSUMO_PRODUCCION_COMPARTIDA WHERE [FECHA_INSUMO] = ?;"
    )


def sql_fechas_disponibles() -> str:
    return (
        "SELECT DISTINCT FechaInsumo FROM ("
        "SELECT [FECHA_INSUMO] AS FechaInsumo FROM BD_INSUMO_ASIGNACIONES WHERE [FECHA_INSUMO] IS NOT NULL AND [FECHA_INSUMO] <> '' UNION "
        "SELECT [FECHA_INSUMO] AS FechaInsumo FROM BD_INSUMO_LICENCIAS WHERE [FECHA_INSUMO] IS NOT NULL AND [FECHA_INSUMO] <> '' UNION "
        "SELECT [FECHA_INSUMO] AS FechaInsumo FROM BD_INSUMO_PRODUCCION_COMPARTIDA WHERE [FECHA_INSUMO] IS NOT NULL AND [FECHA_INSUMO] <> ''"
        ") ORDER BY FechaInsumo DESC;"
    )


def sql_insert_bitacora() -> str:
    return (
        "INSERT INTO BITACORA_CARGA_INSUMOS (nombreArchivo, fechaCorte, fechaHoraCarga, totalRegistros, regimenesActualizados) "
        "VALUES (?, ?, datetime('now', 'localtime'), ?, ?);"
    )


def sql_reporte_analisis_1_asignaciones() -> str:
    return (
        "SELECT COALESCE(T.[DESCRIPCIÓN CUENTA], 'OTROS DERECHOS') AS ConceptoFiscal, "
        "MIN(COALESCE(T.[MONEDA], 'M.N.')) || ' | ' || MIN(COALESCE(T.[ID_CONTRATO], 'BLK-ASIG-0151-MARINA')) || "
        "' (' || printf('%.2f', AVG(COALESCE(C.tasaRegalia, 65.0))) || '%)' AS MonedaContrato, "
        "COUNT(*) AS NumOperaciones, "
        "SUM(COALESCE(T.[SDINMO], 0.0)) / 1000000.0 AS ValorContractualMDP, "
        "SUM(COALESCE(T.[SDFIMN], 0.0)) / 1000000.0 AS ContraprestacionMDP, "
        "SUM(COALESCE(T.[SDPRMN], 0.0)) / 1000000.0 AS DerechoPromedioMDP, "
        "SUM(COALESCE(T.[SDFIDL], 0.0)) / 1000000.0 AS EquivalenteUSD_MDD "
        "FROM BD_INSUMO_ASIGNACIONES AS T "
        "LEFT JOIN CONTRATOS_HIDROCARBUROS AS C ON T.[ID_CONTRATO] = C.idContrato "
        "WHERE T.[FECHA_INSUMO] = ? "
        "GROUP BY COALESCE(T.[DESCRIPCIÓN CUENTA], 'OTROS DERECHOS') "
        "ORDER BY SUM(COALESCE(T.[SDFIMN], 0.0)) DESC;"
    )


def sql_reporte_analisis_2_licencias() -> str:
    return (
        "SELECT COALESCE(T.[CUENCA_PETROLERA], '') || ' | ' || COALESCE(T.[AREA_CONTRACTUAL], '') AS CuencaArea, "
        "MIN(COALESCE(T.[MONEDA], 'USD')) || ' | ' || MIN(COALESCE(T.[ID_CONTRATO], 'BLK-LIC-R01-P01')) || "
        "' (' || printf('%.2f', AVG(COALESCE(C.tasaRegalia, 20.0))) || '%)' AS MonedaContrato, "
        "COUNT(*) AS NumOperaciones, "
        "SUM(COALESCE(T.[SDINMO], 0.0)) / 1000000.0 AS ValorContractualMDP, "
        "SUM(COALESCE(T.[SDFIMN], 0.0)) / 1000000.0 AS ContraprestacionMDP, "
        "SUM(COALESCE(T.[SDPRMN], 0.0)) / 1000000.0 AS DerechoPromedioMDP, "
        "SUM(COALESCE(T.[SDFIDL], 0.0)) / 1000000.0 AS EquivalenteUSD_MDD "
        "FROM BD_INSUMO_LICENCIAS AS T "
        "LEFT JOIN CONTRATOS_HIDROCARBUROS AS C ON T.[ID_CONTRATO] = C.idContrato "
        "WHERE T.[FECHA_INSUMO] = ? "
        "GROUP BY COALESCE(T.[CUENCA_PETROLERA], '') || ' | ' || COALESCE(T.[AREA_CONTRACTUAL], '') "
        "ORDER BY SUM(COALESCE(T.[SDFIMN], 0.0)) DESC;"
    )


def sql_reporte_analisis_3_produccion_compartida() -> str:
    return (
        "SELECT COALESCE(T.[AREA_CONTRACTUAL], '') || ' (' || COALESCE(T.[CAMPO_EXTRACCION], '') || ')' AS AreaCampo, "
        "MIN(COALESCE(T.[MONEDA], 'USD')) || ' | ' || MIN(COALESCE(T.[ID_CONTRATO], 'BLK-PSC-R01-A01')) || "
        "' (' || printf('%.2f', AVG(COALESCE(C.tasaRegalia, 70.0))) || '%)' AS MonedaContrato, "
        "COUNT(*) AS NumOperaciones, "
        "SUM(COALESCE(T.[SDINMO], 0.0)) / 1000000.0 AS ValorContractualMDP, "
        "SUM(COALESCE(T.[SDFIMN], 0.0)) / 1000000.0 AS ContraprestacionMDP, "
        "SUM(COALESCE(T.[SDPRMN], 0.0)) / 1000000.0 AS DerechoPromedioMDP, "
        "SUM(COALESCE(T.[SDFIDL], 0.0)) / 1000000.0 AS EquivalenteUSD_MDD "
        "FROM BD_INSUMO_PRODUCCION_COMPARTIDA AS T "
        "LEFT JOIN CONTRATOS_HIDROCARBUROS AS C ON T.[ID_CONTRATO] = C.idContrato "
        "WHERE T.[FECHA_INSUMO] = ? "
        "GROUP BY COALESCE(T.[AREA_CONTRACTUAL], '') || ' (' || COALESCE(T.[CAMPO_EXTRACCION], '') || ')' "
        "ORDER BY SUM(COALESCE(T.[SDFIMN], 0.0)) DESC;"
    )


def sql_comparativo_dos_insumos() -> str:
    return (
        "SELECT U.CLAVE AS IdContrato, "
        "COALESCE(C.modalidadFiscal, 'NO REGISTRADO') AS ModalidadFiscal, "
        "SUM(U.OPS_ASIG) AS OpsAsignaciones, "
        "SUM(U.OPS_LIC) AS OpsLicencia, "
        "SUM(U.FIN_ASIG) AS ContraprestacionAsignacionesMDP, "
        "SUM(U.FIN_LIC) AS ContraprestacionLicenciaMDP, "
        "SUM(U.FIN_ASIG) + SUM(U.FIN_LIC) AS ContraprestacionCombinadaMDP, "
        "SUM(U.USD_ASIG) + SUM(U.USD_LIC) AS EquivalenteUSD_MDD "
        "FROM ("
        "  SELECT [ID_CONTRATO] AS CLAVE, COUNT(*) AS OPS_ASIG, 0 AS OPS_LIC, "
        "         SUM(COALESCE([SDFIMN], 0.0)) / 1000000.0 AS FIN_ASIG, 0.0 AS FIN_LIC, "
        "         SUM(COALESCE([SDFIDL], 0.0)) / 1000000.0 AS USD_ASIG, 0.0 AS USD_LIC "
        "  FROM BD_INSUMO_ASIGNACIONES WHERE [FECHA_INSUMO] = ? GROUP BY [ID_CONTRATO] "
        "  UNION ALL "
        "  SELECT [ID_CONTRATO] AS CLAVE, 0 AS OPS_ASIG, COUNT(*) AS OPS_LIC, "
        "         0.0 AS FIN_ASIG, SUM(COALESCE([SDFIMN], 0.0)) / 1000000.0 AS FIN_LIC, "
        "         0.0 AS USD_ASIG, SUM(COALESCE([SDFIDL], 0.0)) / 1000000.0 AS USD_LIC "
        "  FROM BD_INSUMO_LICENCIAS WHERE [FECHA_INSUMO] = ? GROUP BY [ID_CONTRATO]"
        ") AS U "
        "LEFT JOIN CONTRATOS_HIDROCARBUROS AS C ON U.CLAVE = C.idContrato "
        "GROUP BY U.CLAVE "
        "ORDER BY SUM(U.FIN_ASIG) + SUM(U.FIN_LIC) DESC;"
    )


def sql_comparativo_tres_insumos() -> str:
    return (
        "SELECT U.CLAVE AS IdContrato, "
        "COALESCE(C.modalidadFiscal, 'NO REGISTRADO') AS ModalidadFiscal, "
        "SUM(U.OPS_ASIG) + SUM(U.OPS_LIC) + SUM(U.OPS_PC) AS OpsTotales, "
        "SUM(U.FIN_ASIG) AS ContraprestacionAsignacionesMDP, "
        "SUM(U.FIN_LIC) AS ContraprestacionLicenciaMDP, "
        "SUM(U.FIN_PC) AS ContraprestacionProdCompartidaMDP, "
        "SUM(U.FIN_ASIG) + SUM(U.FIN_LIC) + SUM(U.FIN_PC) AS ContraprestacionConsolidadaMDP, "
        "(CASE WHEN SUM(U.OPS_ASIG) > 0 THEN 1 ELSE 0 END) + "
        "(CASE WHEN SUM(U.OPS_LIC) > 0 THEN 1 ELSE 0 END) + "
        "(CASE WHEN SUM(U.OPS_PC) > 0 THEN 1 ELSE 0 END) AS RegimenesConPresencia "
        "FROM ("
        "  SELECT [ID_CONTRATO] AS CLAVE, COUNT(*) AS OPS_ASIG, 0 AS OPS_LIC, 0 AS OPS_PC, "
        "         SUM(COALESCE([SDFIMN], 0.0)) / 1000000.0 AS FIN_ASIG, 0.0 AS FIN_LIC, 0.0 AS FIN_PC "
        "  FROM BD_INSUMO_ASIGNACIONES WHERE [FECHA_INSUMO] = ? GROUP BY [ID_CONTRATO] "
        "  UNION ALL "
        "  SELECT [ID_CONTRATO] AS CLAVE, 0, COUNT(*), 0, "
        "         0.0, SUM(COALESCE([SDFIMN], 0.0)) / 1000000.0, 0.0 "
        "  FROM BD_INSUMO_LICENCIAS WHERE [FECHA_INSUMO] = ? GROUP BY [ID_CONTRATO] "
        "  UNION ALL "
        "  SELECT [ID_CONTRATO] AS CLAVE, 0, 0, COUNT(*), "
        "         0.0, 0.0, SUM(COALESCE([SDFIMN], 0.0)) / 1000000.0 "
        "  FROM BD_INSUMO_PRODUCCION_COMPARTIDA WHERE [FECHA_INSUMO] = ? GROUP BY [ID_CONTRATO]"
        ") AS U "
        "LEFT JOIN CONTRATOS_HIDROCARBUROS AS C ON U.CLAVE = C.idContrato "
        "GROUP BY U.CLAVE "
        "ORDER BY ContraprestacionConsolidadaMDP DESC;"
    )
