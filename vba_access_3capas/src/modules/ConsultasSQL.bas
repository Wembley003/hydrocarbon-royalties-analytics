Attribute VB_Name = "ConsultasSQL"
Option Compare Database
Option Explicit

' =============================================================================
' CAPA 3 (PERSISTENCIA - GENERADOR SQL): ConsultasSQL.bas
' -----------------------------------------------------------------------------
' Centraliza el 100% de las sentencias SQL Jet del aplicativo.
' Regla de Arquitectura de 3 Capas:
'   - Ningún formulario (Presentación) ni FondoBO (Negocio) puede contener SQL.
'   - FondoDAO es el único consumidor de este módulo y jamás construye cadenas
'     SQL sueltas fuera de ConsultasSQL.bas.
'   - Todas las tablas consultadas residen físicamente en BD\RIESGOSBD.accdb
'     y están vinculadas en el Frontend Modulo_Administracion_Riesgos.accdb.
' =============================================================================

' =========================================================================
' 1. CONSULTAS SQL DE SEGURIDAD Y ADMINISTRACIÓN DE USUARIOS (Tabla USUARIOS)
' =========================================================================

' Devuelve el listado completo de empleados ordenado por clave para los subformularios
Public Function sqlSelectUsuarios() As String
    sqlSelectUsuarios = "SELECT cveUsuario, nombreUsuario, aPaterno, aMaterno, cvePerfil, emailUsuario, cveEstatus FROM USUARIOS ORDER BY cveUsuario;"
End Function

' Consulta un empleado por clave y contraseña cifrada (ENC$...)
Public Function sqlSelectUsuarioLogin(ByVal cveUsuario As String, ByVal pawdCifrada As String) As String
    Dim pwdEnc As String
    pwdEnc = Utilities.cifrarContrasena(pawdCifrada, PASSWORD_PREDETERMINADA)
    sqlSelectUsuarioLogin = "SELECT cveUsuario, [password], nombreUsuario, aPaterno, aMaterno, cvePerfil, emailUsuario, cveEstatus " & _
                            "FROM USUARIOS WHERE cveUsuario = '" & Replace(cveUsuario, "'", "''") & "' AND [password] = '" & Replace(pwdEnc, "'", "''") & "';"
End Function

' Consulta los datos completos de un empleado a partir de su número de empleado (4 dígitos)
Public Function sqlSelectUsuarioPorClave(ByVal cveUsuario As String) As String
    sqlSelectUsuarioPorClave = "SELECT cveUsuario, [password], nombreUsuario, aPaterno, aMaterno, cvePerfil, emailUsuario, cveEstatus " & _
                               "FROM USUARIOS WHERE cveUsuario = '" & Replace(cveUsuario, "'", "''") & "';"
End Function

' Inserta un nuevo empleado en USUARIOS almacenando su contraseña siempre cifrada
Public Function sqlInsertUsuario(ByVal u As UsuarioDTO) As String
    Dim est As String
    est = u.getCveEstatus()
    If est = "" Then est = CLAVE_ESTATU_ACTIVO
    sqlInsertUsuario = "INSERT INTO USUARIOS (cveUsuario, [password], nombreUsuario, aPaterno, aMaterno, cvePerfil, emailUsuario, cveEstatus) VALUES ('" & _
                       Replace(u.getCveUsuario(), "'", "''") & "', '" & _
                       Replace(Utilities.cifrarContrasena(u.getPassword(), PASSWORD_PREDETERMINADA), "'", "''") & "', '" & _
                       Replace(u.getNombre(), "'", "''") & "', '" & _
                       Replace(u.getAPaterno(), "'", "''") & "', '" & _
                       Replace(u.getAMaterno(), "'", "''") & "', '" & _
                       Replace(u.getCvePerfil(), "'", "''") & "', '" & _
                       Replace(u.getEmail(), "'", "''") & "', '" & _
                       Replace(est, "'", "''") & "');"
End Function

' Actualiza nombre, apellidos, perfil y correo de un empleado existente
Public Function sqlUpdateUsuario(ByVal u As UsuarioDTO) As String
    sqlUpdateUsuario = "UPDATE USUARIOS SET " & _
                       "nombreUsuario = '" & Replace(u.getNombre(), "'", "''") & "', " & _
                       "aPaterno = '" & Replace(u.getAPaterno(), "'", "''") & "', " & _
                       "aMaterno = '" & Replace(u.getAMaterno(), "'", "''") & "', " & _
                       "cvePerfil = '" & Replace(u.getCvePerfil(), "'", "''") & "', " & _
                       "emailUsuario = '" & Replace(u.getEmail(), "'", "''") & "' " & _
                       "WHERE cveUsuario = '" & Replace(u.getCveUsuario(), "'", "''") & "';"
End Function

' Ejecuta la baja lógica de un empleado cambiando su estatus a INACTIVO (nunca DELETE físico)
Public Function sqlInactivarUsuario(ByVal cveUsuario As String) As String
    sqlInactivarUsuario = "UPDATE USUARIOS SET cveEstatus = '" & CLAVE_ESTATU_INACTIVO & "' WHERE cveUsuario = '" & Replace(cveUsuario, "'", "''") & "';"
End Function

' Actualiza la contraseña cifrada de un empleado (autoservicio o restablecimiento por Administrador)
Public Function sqlUpdatePassword(ByVal cveUsuario As String, ByVal nuevaContrasena As String) As String
    sqlUpdatePassword = "UPDATE USUARIOS SET [password] = '" & Replace(Utilities.cifrarContrasena(nuevaContrasena, PASSWORD_PREDETERMINADA), "'", "''") & "' WHERE cveUsuario = '" & Replace(cveUsuario, "'", "''") & "';"
End Function

' =========================================================================
' 2. CONSULTAS SQL DEL CATÁLOGO DE INSTRUMENTOS FINANCIEROS (Tabla INSTRUMENTOS)
' =========================================================================

' Devuelve todos los instrumentos financieros ordenados por clave para los subformularios
Public Function sqlSelectInstrumentos() As String
    sqlSelectInstrumentos = "SELECT idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento FROM INSTRUMENTOS ORDER BY idInstrumento;"
End Function

' Busca un instrumento financiero específico por su identificador único (idInstrumento)
Public Function sqlSelectInstrumentoPorId(ByVal idInstrumento As String) As String
    sqlSelectInstrumentoPorId = "SELECT idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento " & _
                                "FROM INSTRUMENTOS WHERE idInstrumento = '" & UCase(Replace(Trim(idInstrumento), "'", "''")) & "';"
End Function

' Inserta un nuevo instrumento financiero en el catálogo INSTRUMENTOS
Public Function sqlInsertInstrumento(ByVal inst As InstrumentoDTO) As String
    sqlInsertInstrumento = "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('" & _
                           Replace(inst.getIdInstrumento(), "'", "''") & "', '" & _
                           Replace(inst.getCveEmisor(), "'", "''") & "', '" & _
                           Replace(inst.getTipoInstrumento(), "'", "''") & "', " & _
                           Str(inst.getTasa()) & ", " & _
                           CStr(inst.getPlazo()) & ", Date(), Date() + " & CStr(inst.getPlazo()) & ");"
End Function

' Actualiza emisor, tipo, tasa y plazo de un instrumento financiero registrado
Public Function sqlUpdateInstrumento(ByVal inst As InstrumentoDTO) As String
    sqlUpdateInstrumento = "UPDATE INSTRUMENTOS SET " & _
                           "cveEmisor = '" & Replace(inst.getCveEmisor(), "'", "''") & "', " & _
                           "tipoInstrumento = '" & Replace(inst.getTipoInstrumento(), "'", "''") & "', " & _
                           "tasa = " & Str(inst.getTasa()) & ", " & _
                           "plazo = " & CStr(inst.getPlazo()) & " " & _
                           "WHERE idInstrumento = '" & Replace(inst.getIdInstrumento(), "'", "''") & "';"
End Function

' =========================================================================
' 3. CONSULTAS SQL DE CARGA, LIMPIEZA Y BITÁCORA DE LAS BASES DE INSUMOS
'   - BD_INSUMO_PRIMER_PISO  (Cartera de Crédito Comercial - Primer Piso)
'   - BD_INSUMO_TESORERIA    (Portafolio de Tesorería y Mercado de Dinero)
'   - BD_INSUMO_SEGUNDO_PISO (Financiamiento Sectorial y Garantías - Segundo Piso)
'   - CARTERA_ETIQUETADA     (Repositorio consolidado global de 53 columnas)
'   - BITACORA_CARGA_INSUMOS (Control histórico de archivos cargados)
' =========================================================================

' Elimina los registros previos de una fecha de corte específica en una tabla de insumo antes de recargar
Public Function sqlDeleteInsumoTablaPorFecha(ByVal nombreTabla As String, ByVal fechaCorte As String) As String
    sqlDeleteInsumoTablaPorFecha = "DELETE FROM [" & Replace(Trim(nombreTabla), "]", "") & "] WHERE [FECHA_INSUMO] = '" & Replace(Trim(fechaCorte), "'", "''") & "';"
End Function

' Elimina del repositorio consolidado CARTERA_ETIQUETADA los registros de una fecha y piso específicos
Public Function sqlDeleteCarteraPorFechaYPiso(ByVal fechaCorte As String, ByVal pisoFiltro As String) As String
    sqlDeleteCarteraPorFechaYPiso = "DELETE FROM CARTERA_ETIQUETADA WHERE [FECHA_INSUMO] = '" & Replace(Trim(fechaCorte), "'", "''") & "' AND [PISO] = '" & Replace(UCase(Trim(pisoFiltro)), "'", "''") & "';"
End Function

' Vacía por completo una tabla de insumo cuando no se especifica fecha de corte
Public Function sqlDeleteTablaCompleta(ByVal nombreTabla As String) As String
    sqlDeleteTablaCompleta = "DELETE FROM [" & Replace(Trim(nombreTabla), "]", "") & "];"
End Function

' Copia en una sola operación SQL las 53 columnas recién cargadas hacia CARTERA_ETIQUETADA
Public Function sqlCopiarInsumoACarteraEtiquetada(ByVal nombreTablaOrigen As String, ByVal fechaCorte As String) As String
    Dim tbl As String, sql As String
    tbl = Replace(Trim(nombreTablaOrigen), "]", "")
    sql = "INSERT INTO CARTERA_ETIQUETADA ( [P],[PISO],[G],[GRUPO],[ESTATA],[DESC_SUCURSAL],[SG],[DESC_SUBGRUPO],[CAMPOA],[NOMBRE]," & _
          "[RFC],[MONEDA2],[MON],[MONEDA],[EF],[ENT_FED],[CONCATENADO],[CTA],[SCATAS],[SCTA],[SSCTA],[SSSCTA],[SECTOR],[C],[CONCEPTO]," & _
          "[SC],[SCONCEPTO],[SSCONCEPTO],[DESTINO],[SDINMO],[SDFIMO],[PROMMO],[SDFIMN],[SDPRMN],[SDFIDL],[SDPRDL],[MESPRO],[ANOPRO]," & _
          "[ACT],[ACTIVIDAD],[GRUCON],[GRUPO_CONT],[GRUCONTNORM],[ARTICULO_NORMATIVO],[CODIGO_ACTIVIDAD_ECONOMICA_GRAL]," & _
          "[DESC_ACTIVIDAD_ECONOMICA_GRAL],[CODIGO_EJECUTIVO],[NOMBRE_EJECUTIVO],[DESCRIPCI" & ChrW(211) & "N CUENTA]," & _
          "[SITUACION DEL CREDITO],[CONCEPTO_ETAPA],[FECHA_INSUMO],[ID_INSTRUMENTO] ) " & _
          "SELECT [P],[PISO],[G],[GRUPO],[ESTATA],[DESC_SUCURSAL],[SG],[DESC_SUBGRUPO],[CAMPOA],[NOMBRE]," & _
          "[RFC],[MONEDA2],[MON],[MONEDA],[EF],[ENT_FED],[CONCATENADO],[CTA],[SCATAS],[SCTA],[SSCTA],[SSSCTA],[SECTOR],[C],[CONCEPTO]," & _
          "[SC],[SCONCEPTO],[SSCONCEPTO],[DESTINO],[SDINMO],[SDFIMO],[PROMMO],[SDFIMN],[SDPRMN],[SDFIDL],[SDPRDL],[MESPRO],[ANOPRO]," & _
          "[ACT],[ACTIVIDAD],[GRUCON],[GRUPO_CONT],[GRUCONTNORM],[ARTICULO_NORMATIVO],[CODIGO_ACTIVIDAD_ECONOMICA_GRAL]," & _
          "[DESC_ACTIVIDAD_ECONOMICA_GRAL],[CODIGO_EJECUTIVO],[NOMBRE_EJECUTIVO],[DESCRIPCI" & ChrW(211) & "N CUENTA]," & _
          "[SITUACION DEL CREDITO],[CONCEPTO_ETAPA],[FECHA_INSUMO],[ID_INSTRUMENTO] FROM [" & tbl & "]"
    If Trim(fechaCorte) <> "" Then
        sql = sql & " WHERE [FECHA_INSUMO] = '" & Replace(Trim(fechaCorte), "'", "''") & "'"
    End If
    sqlCopiarInsumoACarteraEtiquetada = sql & ";"
End Function

' Selecciona todas las filas de una tabla de insumo para una fecha dada (usado en respaldo por ordinal)
Public Function sqlSelectTablaInsumoPorFecha(ByVal nombreTabla As String, ByVal fechaCorte As String) As String
    Dim sql As String
    sql = "SELECT * FROM [" & Replace(Trim(nombreTabla), "]", "") & "]"
    If Trim(fechaCorte) <> "" Then
        sql = sql & " WHERE [FECHA_INSUMO] = '" & Replace(Trim(fechaCorte), "'", "''") & "'"
    End If
    sqlSelectTablaInsumoPorFecha = sql & ";"
End Function

' Cuenta cuántos registros existen en una tabla de insumo específica para una fecha de corte
Public Function sqlContarInsumoTablaPorFecha(ByVal nombreTabla As String, ByVal fechaCorte As String) As String
    sqlContarInsumoTablaPorFecha = "SELECT COUNT(*) AS TotalCreditos FROM [" & Replace(Trim(nombreTabla), "]", "") & "] WHERE [FECHA_INSUMO] = '" & Replace(Trim(fechaCorte), "'", "''") & "';"
End Function

' Obtiene en 1 sola consulta indexada el conteo de registros de las 3 tablas de insumo para una fecha
Public Function sqlContarTresInsumosPorFecha(ByVal fechaCorte As String) As String
    Dim f As String
    f = Replace(Trim(fechaCorte), "'", "''")
    sqlContarTresInsumosPorFecha = "SELECT 1 AS PisoIdx, COUNT(*) AS TotalReg FROM BD_INSUMO_PRIMER_PISO WHERE [FECHA_INSUMO] = '" & f & "' " & _
                                   "UNION ALL SELECT 2 AS PisoIdx, COUNT(*) AS TotalReg FROM BD_INSUMO_TESORERIA WHERE [FECHA_INSUMO] = '" & f & "' " & _
                                   "UNION ALL SELECT 3 AS PisoIdx, COUNT(*) AS TotalReg FROM BD_INSUMO_SEGUNDO_PISO WHERE [FECHA_INSUMO] = '" & f & "';"
End Function

' Obtiene en una sola consulta unificada el conteo de registros por fecha en las 3 bases de insumo (para el Calendario Popup)
Public Function sqlMapaFechasConteo() As String
    sqlMapaFechasConteo = "SELECT F, SUM(N) AS T FROM (" & _
                          "SELECT [FECHA_INSUMO] AS F, COUNT(*) AS N FROM BD_INSUMO_PRIMER_PISO WHERE [FECHA_INSUMO] Is Not Null AND [FECHA_INSUMO] <> '' GROUP BY [FECHA_INSUMO] " & _
                          "UNION ALL SELECT [FECHA_INSUMO] AS F, COUNT(*) AS N FROM BD_INSUMO_TESORERIA WHERE [FECHA_INSUMO] Is Not Null AND [FECHA_INSUMO] <> '' GROUP BY [FECHA_INSUMO] " & _
                          "UNION ALL SELECT [FECHA_INSUMO] AS F, COUNT(*) AS N FROM BD_INSUMO_SEGUNDO_PISO WHERE [FECHA_INSUMO] Is Not Null AND [FECHA_INSUMO] <> '' GROUP BY [FECHA_INSUMO] " & _
                          ") AS U GROUP BY F;"
End Function

' Devuelve la lista de fechas distintas con insumos cargados en cualquiera de las bases
Public Function sqlFechasDisponiblesCartera() As String
    sqlFechasDisponiblesCartera = "SELECT DISTINCT FechaInsumo FROM (" & _
                                  "SELECT [FECHA_INSUMO] AS FechaInsumo FROM BD_INSUMO_PRIMER_PISO WHERE [FECHA_INSUMO] Is Not Null AND [FECHA_INSUMO] <> '' UNION " & _
                                  "SELECT [FECHA_INSUMO] AS FechaInsumo FROM BD_INSUMO_TESORERIA WHERE [FECHA_INSUMO] Is Not Null AND [FECHA_INSUMO] <> '' UNION " & _
                                  "SELECT [FECHA_INSUMO] AS FechaInsumo FROM BD_INSUMO_SEGUNDO_PISO WHERE [FECHA_INSUMO] Is Not Null AND [FECHA_INSUMO] <> '' UNION " & _
                                  "SELECT [FECHA_INSUMO] AS FechaInsumo FROM CARTERA_ETIQUETADA WHERE [FECHA_INSUMO] Is Not Null AND [FECHA_INSUMO] <> ''" & _
                                  ") ORDER BY FechaInsumo DESC;"
End Function

' Obtiene la última fecha de corte registrada en BITACORA_CARGA_INSUMOS
Public Function sqlUltimaFechaBitacora() As String
    sqlUltimaFechaBitacora = "SELECT TOP 1 fechaCorte FROM BITACORA_CARGA_INSUMOS ORDER BY idCarga DESC;"
End Function

' Verifica en BITACORA_CARGA_INSUMOS si un archivo ya fue cargado previamente para la misma fecha
Public Function sqlSelectBitacoraPorArchivoYFecha(ByVal nombreArchivo As String, ByVal fechaCorte As String) As String
    sqlSelectBitacoraPorArchivoYFecha = "SELECT TOP 1 nombreArchivo, fechaCorte, fechaHoraCarga FROM BITACORA_CARGA_INSUMOS " & _
                                        "WHERE UCase(nombreArchivo) = '" & UCase(Replace(Trim(nombreArchivo), "'", "''")) & "' " & _
                                        "AND fechaCorte = '" & Replace(Trim(fechaCorte), "'", "''") & "' ORDER BY idCarga DESC;"
End Function

' Inserta un registro de auditoría en BITACORA_CARGA_INSUMOS al concluir la carga de un insumo
Public Function sqlInsertBitacoraInsumo(ByVal nombreArchivo As String, ByVal fechaCorte As String, ByVal totalReg As Long, ByVal fondosStr As String) As String
    sqlInsertBitacoraInsumo = "INSERT INTO BITACORA_CARGA_INSUMOS (nombreArchivo, fechaCorte, fechaHoraCarga, totalRegistros, fondosActualizados) VALUES ('" & _
                              Replace(Trim(nombreArchivo), "'", "''") & "', '" & _
                              Replace(Trim(fechaCorte), "'", "''") & "', Now(), " & CStr(totalReg) & ", '" & _
                              Replace(Left(Trim(fondosStr), 190), "'", "''") & "');"
End Function

' =========================================================================
' 4. CONSULTAS SQL DE EXTRACCIÓN ANALÍTICA POR PORTAFOLIO (REPORTES 1, 2 Y 3)
' =========================================================================

' Consulta SQL 1: Agrupa BD_INSUMO_PRIMER_PISO + INSTRUMENTOS por DESCRIPCIÓN CUENTA (en MDP/MDD)
Public Function sqlReporteAnalisis1PrimerPiso(ByVal fechaCorte As String) As String
    sqlReporteAnalisis1PrimerPiso = "SELECT Trim(CStr(Nz(T.[DESCRIPCIÓN CUENTA], 'OTROS'))) AS ConceptoCuenta, " & _
                                    "FIRST(Trim(CStr(Nz(T.[MONEDA], 'M.N.')))) & ' | ' & FIRST(Trim(CStr(Nz(T.[ID_INSTRUMENTO], 'BLK_29')))) & ' (' & Format(AVG(IIF(I.tasa Is Null, 10.6, I.tasa)), '0.00') & '%)' AS MonedaBase, " & _
                                    "COUNT(*) AS NumCreditos, " & _
                                    "SUM(IIF(T.[SDINMO] Is Null, 0, T.[SDINMO])) / 1000000 AS SaldoInicialMDP, " & _
                                    "SUM(IIF(T.[SDFIMN] Is Null, 0, T.[SDFIMN])) / 1000000 AS SaldoFinalMDP, " & _
                                    "SUM(IIF(T.[SDPRMN] Is Null, 0, T.[SDPRMN])) / 1000000 AS SaldoPromedioMDP, " & _
                                    "SUM(IIF(T.[SDFIDL] Is Null, 0, T.[SDFIDL])) / 1000000 AS EquivalenteUSD_MDD " & _
                                    "FROM BD_INSUMO_PRIMER_PISO AS T " & _
                                    "LEFT JOIN INSTRUMENTOS AS I ON T.[ID_INSTRUMENTO] = I.idInstrumento " & _
                                    "WHERE T.[FECHA_INSUMO] = '" & Replace(Trim(fechaCorte), "'", "''") & "' " & _
                                    "GROUP BY Trim(CStr(Nz(T.[DESCRIPCIÓN CUENTA], 'OTROS'))) " & _
                                    "ORDER BY SUM(IIF(T.[SDFIMN] Is Null, 0, T.[SDFIMN])) DESC;"
End Function

' Consulta SQL 2: Agrupa BD_INSUMO_TESORERIA + INSTRUMENTOS por GRUPO | DESC_SUBGRUPO (en MDP/MDD)
Public Function sqlReporteAnalisis2SectoresMonedas(ByVal fechaCorte As String) As String
    sqlReporteAnalisis2SectoresMonedas = "SELECT Trim(CStr(Nz(T.[GRUPO], ''))) & ' | ' & Trim(CStr(Nz(T.[DESC_SUBGRUPO], ''))) AS GrupoSubgrupo, " & _
                                         "FIRST(Trim(CStr(Nz(T.[MONEDA], 'M.N.')))) & ' | ' & FIRST(Trim(CStr(Nz(T.[ID_INSTRUMENTO], 'CETES_28D')))) & ' (' & Format(AVG(IIF(I.tasa Is Null, 10.75, I.tasa)), '0.00') & '%)' AS DivisaPredominante, " & _
                                         "COUNT(*) AS NumOperaciones, " & _
                                         "SUM(IIF(T.[SDINMO] Is Null, 0, T.[SDINMO])) / 1000000 AS SaldoInicialMDP, " & _
                                         "SUM(IIF(T.[SDFIMN] Is Null, 0, T.[SDFIMN])) / 1000000 AS PosicionFinalMDP, " & _
                                         "SUM(IIF(T.[SDPRMN] Is Null, 0, T.[SDPRMN])) / 1000000 AS PosicionPromedioMDP, " & _
                                         "SUM(IIF(T.[SDFIDL] Is Null, 0, T.[SDFIDL])) / 1000000 AS PosicionDivisaMDD " & _
                                         "FROM BD_INSUMO_TESORERIA AS T " & _
                                         "LEFT JOIN INSTRUMENTOS AS I ON T.[ID_INSTRUMENTO] = I.idInstrumento " & _
                                         "WHERE T.[FECHA_INSUMO] = '" & Replace(Trim(fechaCorte), "'", "''") & "' " & _
                                         "GROUP BY Trim(CStr(Nz(T.[GRUPO], ''))) & ' | ' & Trim(CStr(Nz(T.[DESC_SUBGRUPO], ''))) " & _
                                         "ORDER BY SUM(IIF(T.[SDFIMN] Is Null, 0, T.[SDFIMN])) DESC;"
End Function

' Consulta SQL 3: Agrupa BD_INSUMO_SEGUNDO_PISO + INSTRUMENTOS por DESC_SUBGRUPO (ACTIVIDAD) (en MDP/MDD)
Public Function sqlReporteAnalisis3SegundoPiso(ByVal fechaCorte As String) As String
    sqlReporteAnalisis3SegundoPiso = "SELECT Trim(CStr(Nz(T.[DESC_SUBGRUPO], ''))) & ' (' & Trim(CStr(Nz(T.[ACTIVIDAD], ''))) & ')' AS SegmentoActividad, " & _
                                     "FIRST(Trim(CStr(Nz(T.[MONEDA], 'M.N.')))) & ' | ' & FIRST(Trim(CStr(Nz(T.[ID_INSTRUMENTO], 'NAFIN_28')))) & ' (' & Format(AVG(IIF(I.tasa Is Null, 9.85, I.tasa)), '0.00') & '%)' AS EtapaNormativa, " & _
                                     "COUNT(*) AS NumCreditos, " & _
                                     "SUM(IIF(T.[SDINMO] Is Null, 0, T.[SDINMO])) / 1000000 AS SaldoInicialMDP, " & _
                                     "SUM(IIF(T.[SDFIMN] Is Null, 0, T.[SDFIMN])) / 1000000 AS SaldoVigenteMDP, " & _
                                     "SUM(IIF(T.[SDPRMN] Is Null, 0, T.[SDPRMN])) / 1000000 AS SaldoPromedioMDP, " & _
                                     "SUM(IIF(T.[SDFIDL] Is Null, 0, T.[SDFIDL])) / 1000000 AS EquivalenteUSD_MDD " & _
                                     "FROM BD_INSUMO_SEGUNDO_PISO AS T " & _
                                     "LEFT JOIN INSTRUMENTOS AS I ON T.[ID_INSTRUMENTO] = I.idInstrumento " & _
                                     "WHERE T.[FECHA_INSUMO] = '" & Replace(Trim(fechaCorte), "'", "''") & "' " & _
                                     "GROUP BY Trim(CStr(Nz(T.[DESC_SUBGRUPO], ''))) & ' (' & Trim(CStr(Nz(T.[ACTIVIDAD], ''))) & ')' " & _
                                     "ORDER BY SUM(IIF(T.[SDFIMN] Is Null, 0, T.[SDFIMN])) DESC;"
End Function

' =============================================================================
' 5. CONSULTAS CRUZADAS ENTRE BASES DE INSUMO (REPORTES COMPARATIVOS 2 Y 3 LIBROS)
' -----------------------------------------------------------------------------
' Simulan un FULL OUTER JOIN en SQL Jet mediante subconsulta UNION ALL + GROUP BY
' por ID_INSTRUMENTO para cruzar saldos y presencia entre portafolios.
' =============================================================================

' Cruce de DOS insumos: Primer Piso vs Tesorería por ID_INSTRUMENTO
Public Function sqlComparativoDosInsumos(ByVal fechaCorte As String) As String
    Dim f As String, q As String
    f = Replace(Trim(fechaCorte), "'", "''")
    q = ""
    q = q & "SELECT T.CLAVE AS Instrumento, " & _
            "FIRST(T.TIPOINS) AS TipoInstrumento, " & _
            "SUM(T.OPS_PP) AS OpsPrimerPiso, " & _
            "SUM(T.OPS_TE) AS OpsTesoreria, " & _
            "SUM(T.FIN_PP) AS FinalPrimerPisoMDP, " & _
            "SUM(T.FIN_TE) AS FinalTesoreriaMDP, " & _
            "SUM(T.FIN_PP) + SUM(T.FIN_TE) AS FinalCombinadoMDP, " & _
            "SUM(T.USD_PP) + SUM(T.USD_TE) AS EquivalenteUSD_MDD " & _
            "FROM ( " & _
            "SELECT Trim(CStr(Nz(A.[ID_INSTRUMENTO], 'SIN_CLAVE'))) AS CLAVE, "
    q = q & "Trim(CStr(Nz(I.tipoInstrumento, 'NO REGISTRADO'))) AS TIPOINS, " & _
            "1 AS OPS_PP, 0 AS OPS_TE, " & _
            "IIF(A.[SDFIMN] Is Null, 0, A.[SDFIMN]) / 1000000 AS FIN_PP, 0 AS FIN_TE, " & _
            "IIF(A.[SDFIDL] Is Null, 0, A.[SDFIDL]) / 1000000 AS USD_PP, 0 AS USD_TE " & _
            "FROM BD_INSUMO_PRIMER_PISO AS A " & _
            "LEFT JOIN INSTRUMENTOS AS I ON A.[ID_INSTRUMENTO] = I.idInstrumento " & _
            "WHERE A.[FECHA_INSUMO] = '" & f & "' " & _
            "UNION ALL " & _
            "SELECT Trim(CStr(Nz(B.[ID_INSTRUMENTO], 'SIN_CLAVE'))), " & _
            "Trim(CStr(Nz(J.tipoInstrumento, 'NO REGISTRADO'))), "
    q = q & "0, 1, " & _
            "0, IIF(B.[SDFIMN] Is Null, 0, B.[SDFIMN]) / 1000000, " & _
            "0, IIF(B.[SDFIDL] Is Null, 0, B.[SDFIDL]) / 1000000 " & _
            "FROM BD_INSUMO_TESORERIA AS B " & _
            "LEFT JOIN INSTRUMENTOS AS J ON B.[ID_INSTRUMENTO] = J.idInstrumento " & _
            "WHERE B.[FECHA_INSUMO] = '" & f & "' " & _
            ") AS T " & _
            "GROUP BY T.CLAVE " & _
            "ORDER BY SUM(T.FIN_PP) + SUM(T.FIN_TE) DESC;"
    sqlComparativoDosInsumos = q
End Function

' Cruce de TRES insumos: Primer Piso vs Tesorería vs Segundo Piso por ID_INSTRUMENTO
Public Function sqlComparativoTresInsumos(ByVal fechaCorte As String) As String
    Dim f As String, q As String
    f = Replace(Trim(fechaCorte), "'", "''")
    q = ""
    q = q & "SELECT T.CLAVE AS Instrumento, " & _
            "FIRST(T.TIPOINS) AS TipoInstrumento, " & _
            "SUM(T.OPS_PP) + SUM(T.OPS_TE) + SUM(T.OPS_SP) AS OpsTotales, " & _
            "SUM(T.FIN_PP) AS FinalPrimerPisoMDP, " & _
            "SUM(T.FIN_TE) AS FinalTesoreriaMDP, " & _
            "SUM(T.FIN_SP) AS FinalSegundoPisoMDP, " & _
            "SUM(T.FIN_PP) + SUM(T.FIN_TE) + SUM(T.FIN_SP) AS FinalConsolidadoMDP, " & _
            "IIF(SUM(T.OPS_PP) > 0, 1, 0) + IIF(SUM(T.OPS_TE) > 0, 1, 0) + IIF(SUM(T.OPS_SP) > 0, 1, 0) AS LibrosConPresencia " & _
            "FROM ( " & _
            "SELECT Trim(CStr(Nz(A.[ID_INSTRUMENTO], 'SIN_CLAVE'))) AS CLAVE, "
    q = q & "Trim(CStr(Nz(I.tipoInstrumento, 'NO REGISTRADO'))) AS TIPOINS, " & _
            "1 AS OPS_PP, 0 AS OPS_TE, 0 AS OPS_SP, " & _
            "IIF(A.[SDFIMN] Is Null, 0, A.[SDFIMN]) / 1000000 AS FIN_PP, 0 AS FIN_TE, 0 AS FIN_SP " & _
            "FROM BD_INSUMO_PRIMER_PISO AS A " & _
            "LEFT JOIN INSTRUMENTOS AS I ON A.[ID_INSTRUMENTO] = I.idInstrumento " & _
            "WHERE A.[FECHA_INSUMO] = '" & f & "' " & _
            "UNION ALL " & _
            "SELECT Trim(CStr(Nz(B.[ID_INSTRUMENTO], 'SIN_CLAVE'))), " & _
            "Trim(CStr(Nz(J.tipoInstrumento, 'NO REGISTRADO'))), " & _
            "0, 1, 0, "
    q = q & "0, IIF(B.[SDFIMN] Is Null, 0, B.[SDFIMN]) / 1000000, 0 " & _
            "FROM BD_INSUMO_TESORERIA AS B " & _
            "LEFT JOIN INSTRUMENTOS AS J ON B.[ID_INSTRUMENTO] = J.idInstrumento " & _
            "WHERE B.[FECHA_INSUMO] = '" & f & "' " & _
            "UNION ALL " & _
            "SELECT Trim(CStr(Nz(C.[ID_INSTRUMENTO], 'SIN_CLAVE'))), " & _
            "Trim(CStr(Nz(K.tipoInstrumento, 'NO REGISTRADO'))), " & _
            "0, 0, 1, " & _
            "0, 0, IIF(C.[SDFIMN] Is Null, 0, C.[SDFIMN]) / 1000000 " & _
            "FROM BD_INSUMO_SEGUNDO_PISO AS C "
    q = q & "LEFT JOIN INSTRUMENTOS AS K ON C.[ID_INSTRUMENTO] = K.idInstrumento " & _
            "WHERE C.[FECHA_INSUMO] = '" & f & "' " & _
            ") AS T " & _
            "GROUP BY T.CLAVE " & _
            "ORDER BY SUM(T.FIN_PP) + SUM(T.FIN_TE) + SUM(T.FIN_SP) DESC;"
    sqlComparativoTresInsumos = q
End Function

' Consulta simple de agregación por instrumento sobre una tabla individual (respaldo en memoria)
Public Function sqlAgregadoPorInstrumentoTabla(ByVal nombreTabla As String, ByVal fechaCorte As String) As String
    Dim f As String, t As String
    f = Replace(Trim(fechaCorte), "'", "''")
    t = Replace(Trim(nombreTabla), "]", "")
    sqlAgregadoPorInstrumentoTabla = _
        "SELECT Trim(CStr(Nz(T.[ID_INSTRUMENTO], 'SIN_CLAVE'))) AS Instrumento, " & _
        "FIRST(Trim(CStr(Nz(I.tipoInstrumento, 'NO REGISTRADO')))) AS TipoInstrumento, " & _
        "COUNT(*) AS NumOperaciones, " & _
        "SUM(IIF(T.[SDFIMN] Is Null, 0, T.[SDFIMN])) / 1000000 AS SaldoFinalMDP, " & _
        "SUM(IIF(T.[SDFIDL] Is Null, 0, T.[SDFIDL])) / 1000000 AS EquivalenteUSD_MDD " & _
        "FROM [" & t & "] AS T " & _
        "LEFT JOIN INSTRUMENTOS AS I ON T.[ID_INSTRUMENTO] = I.idInstrumento " & _
        "WHERE T.[FECHA_INSUMO] = '" & f & "' " & _
        "GROUP BY Trim(CStr(Nz(T.[ID_INSTRUMENTO], 'SIN_CLAVE'))) " & _
        "ORDER BY SUM(IIF(T.[SDFIMN] Is Null, 0, T.[SDFIMN])) DESC;"
End Function
