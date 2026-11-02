VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FondoDAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =============================================================================
' CAPA 3 (PERSISTENCIA - DATA ACCESS OBJECT): FondoDAO.cls
' -----------------------------------------------------------------------------
' Es la ÚNICA clase del sistema autorizada para abrir Recordsets, ejecutar
' comandos sobre CurrentDb y administrar los vínculos hacia la base de datos
' externa BD\RIESGOSBD.accdb.
'
' Principios de Arquitectura de 3 Capas:
'   1. Recibe y devuelve objetos de transferencia (UsuarioDTO, InstrumentoDTO)
'      o arreglos matriciales hacia FondoBO (Capa de Negocio).
'   2. Jamás es invocada directamente desde los formularios (Presentación).
'   3. Obtiene el 100% de las cadenas SQL desde ConsultasSQL.bas (sin SQL suelto).
'   4. Garantiza que las 13 tablas vivan físicamente fuera del módulo principal,
'      alojadas exclusivamente en BD\RIESGOSBD.accdb.
' =============================================================================

' Verifica al iniciar el aplicativo que las 13 tablas en Modulo_Administracion_Riesgos.accdb
' sean tablas VINCULADAS hacia BD\RIESGOSBD.accdb (nunca tablas locales) y actualiza la ruta
' de conexión automáticamente si la carpeta del proyecto se movió de ubicación.
Public Function verificarYRelinkearBackendDAO() As Boolean
    Dim rutaBackend As String, fso As Object, tNames As Variant, i As Integer
    Dim tName As String, tdf As DAO.TableDef, db As DAO.Database
    Dim connConPwd As String, connSinPwd As String, bCambios As Boolean
    On Error Resume Next
    If gBackendVerificadoEnSesion Then
        verificarYRelinkearBackendDAO = True
        Exit Function
    End If
    verificarYRelinkearBackendDAO = False
    bCambios = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    rutaBackend = CurrentProject.Path & DB_PATH & DB_NAME
    If Not fso.FileExists(rutaBackend) Then
        rutaBackend = CurrentProject.Path & "\" & DB_NAME
    End If
    If Not fso.FileExists(rutaBackend) Then Exit Function

    connConPwd = ";DATABASE=" & rutaBackend & ";PWD=" & PASSWORD_PREDETERMINADA
    connSinPwd = ";DATABASE=" & rutaBackend
    Set db = CurrentDb
    tNames = Array("USUARIOS", "EMISORES", "INSTRUMENTOS", "PERFILES", "ESTATUS", _
                   "CAT_MONEDAS", "CAT_TASAS", "TBL_TRANSACCIONES_FONDO", _
                   "BD_INSUMO_PRIMER_PISO", "BD_INSUMO_TESORERIA", _
                   "BD_INSUMO_SEGUNDO_PISO", "CARTERA_ETIQUETADA", "BITACORA_CARGA_INSUMOS")

    For i = LBound(tNames) To UBound(tNames)
        tName = CStr(tNames(i))
        Err.Clear
        Set tdf = Nothing
        Set tdf = db.TableDefs(tName)
        If Not tdf Is Nothing Then
            ' Si por accidente existiera como tabla local (Connect = ""), se elimina para vincularla a BD\RIESGOSBD.accdb
            If Trim(CStr(Nz(tdf.Connect, ""))) = "" Then
                db.TableDefs.Delete tName
                Set tdf = Nothing
            ElseIf InStr(1, tdf.Connect, rutaBackend, vbTextCompare) = 0 Then
                ' Si la carpeta se movió de ruta, actualizar el vínculo hacia la nueva ubicación de BD\RIESGOSBD.accdb
                Err.Clear
                tdf.Connect = connConPwd
                tdf.RefreshLink
                If Err.Number <> 0 Then
                    Err.Clear
                    tdf.Connect = connSinPwd
                    tdf.RefreshLink
                End If
            End If
        End If
        If tdf Is Nothing Then
            Err.Clear
            Set tdf = db.CreateTableDef(tName)
            tdf.SourceTableName = tName
            tdf.Connect = connConPwd
            db.TableDefs.Append tdf
            If Err.Number <> 0 Then
                Err.Clear
                Set tdf = db.CreateTableDef(tName)
                tdf.SourceTableName = tName
                tdf.Connect = connSinPwd
                db.TableDefs.Append tdf
            End If
        End If
    Next i
    If bCambios Then db.TableDefs.Refresh
    Call asegurarIndicesBackendDAO(rutaBackend)
    gBackendVerificadoEnSesion = True
    verificarYRelinkearBackendDAO = True
    On Error GoTo 0
End Function

' Garantiza mediante objetos nativos DAO.Index que las tablas en BD\RIESGOSBD.accdb tengan
' índices B-Tree sobre FECHA_INSUMO, ID_INSTRUMENTO y PISO para consultas instantáneas (< 1 ms)
Private Sub asegurarIndicesBackendDAO(ByVal rutaBackend As String)
    Dim dbBack As DAO.Database, tdfB As DAO.TableDef, idxB As DAO.Index, fldB As DAO.Field
    Dim arrT As Variant, k As Integer, nomT As String
    On Error Resume Next
    Err.Clear
    Set dbBack = DBEngine.Workspaces(0).OpenDatabase(rutaBackend, False, False, ";PWD=" & PASSWORD_PREDETERMINADA)
    If Err.Number <> 0 Or dbBack Is Nothing Then
        Err.Clear
        Set dbBack = DBEngine.Workspaces(0).OpenDatabase(rutaBackend, False, False)
    End If
    If dbBack Is Nothing Then Exit Sub

    arrT = Array("BD_INSUMO_PRIMER_PISO", "BD_INSUMO_TESORERIA", "BD_INSUMO_SEGUNDO_PISO", "CARTERA_ETIQUETADA")
    For k = LBound(arrT) To UBound(arrT)
        nomT = CStr(arrT(k))
        Set tdfB = Nothing
        Set tdfB = dbBack.TableDefs(nomT)
        If Not tdfB Is Nothing Then
            Err.Clear
            Set idxB = Nothing
            Set idxB = tdfB.Indexes("idx_fecha")
            If idxB Is Nothing Then
                Err.Clear
                Set idxB = tdfB.CreateIndex("idx_fecha")
                Set fldB = idxB.CreateField("FECHA_INSUMO")
                idxB.Fields.Append fldB
                tdfB.Indexes.Append idxB
            End If
            Err.Clear
            Set idxB = Nothing
            Set idxB = tdfB.Indexes("idx_instr")
            If idxB Is Nothing Then
                Err.Clear
                Set idxB = tdfB.CreateIndex("idx_instr")
                Set fldB = idxB.CreateField("ID_INSTRUMENTO")
                idxB.Fields.Append fldB
                tdfB.Indexes.Append idxB
            End If
            If nomT = "CARTERA_ETIQUETADA" Then
                Err.Clear
                Set idxB = Nothing
                Set idxB = tdfB.Indexes("idx_piso_fec")
                If idxB Is Nothing Then
                    Err.Clear
                    Set idxB = tdfB.CreateIndex("idx_piso_fec")
                    Set fldB = idxB.CreateField("FECHA_INSUMO")
                    idxB.Fields.Append fldB
                    Set fldB = idxB.CreateField("PISO")
                    idxB.Fields.Append fldB
                    tdfB.Indexes.Append idxB
                End If
            End If
        End If
    Next k

    Set tdfB = Nothing
    Set tdfB = dbBack.TableDefs("BITACORA_CARGA_INSUMOS")
    If Not tdfB Is Nothing Then
        Err.Clear
        Set idxB = Nothing
        Set idxB = tdfB.Indexes("idx_bit_fec")
        If idxB Is Nothing Then
            Err.Clear
            Set idxB = tdfB.CreateIndex("idx_bit_fec")
            Set fldB = idxB.CreateField("fechaCorte")
            idxB.Fields.Append fldB
            Set fldB = idxB.CreateField("nombreArchivo")
            idxB.Fields.Append fldB
            tdfB.Indexes.Append idxB
        End If
    End If

    dbBack.Close
    Set dbBack = Nothing
    Err.Clear
End Sub

' Devuelve la consulta SQL de origen para el subformulario del catálogo de Usuarios
Public Function obtenerOrigenUsuariosDAO() As String
    obtenerOrigenUsuariosDAO = ConsultasSQL.sqlSelectUsuarios()
End Function

' Devuelve la consulta SQL de origen para el subformulario del catálogo de Instrumentos
Public Function obtenerOrigenInstrumentosDAO() As String
    obtenerOrigenInstrumentosDAO = ConsultasSQL.sqlSelectInstrumentos()
End Function

Public Function selectUsuario(ByVal claveUsuario As String, ByVal pawdCifrada As String, Optional ByVal conn As Object = Nothing) As UsuarioDTO
    Dim sql As String
    Dim usuarioDTOAux As UsuarioDTO
    Dim rsDao As DAO.Recordset
    
    On Error Resume Next
    sql = ConsultasSQL.sqlSelectUsuarioLogin(claveUsuario, pawdCifrada)
    
    Set rsDao = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    If Not rsDao Is Nothing Then
        If Not rsDao.EOF Then
            Set usuarioDTOAux = New UsuarioDTO
            usuarioDTOAux.setCveUsuario Nz(rsDao.Fields(0), "")
            usuarioDTOAux.setPassword Nz(rsDao.Fields(1), "")
            usuarioDTOAux.setNombre Nz(rsDao.Fields(2), "")
            usuarioDTOAux.setAPaterno Nz(rsDao.Fields(3), "")
            usuarioDTOAux.setAMaterno Nz(rsDao.Fields(4), "")
            usuarioDTOAux.setCvePerfil Nz(rsDao.Fields(5), "")
            usuarioDTOAux.setEmail Nz(rsDao.Fields(6), "")
            usuarioDTOAux.setCveEstatus Nz(rsDao.Fields(7), "")
            rsDao.Close
            Set rsDao = Nothing
            Set selectUsuario = usuarioDTOAux
            Exit Function
        End If
        rsDao.Close
        Set rsDao = Nothing
    End If
    
    On Error GoTo errorHandler
    If Not conn Is Nothing Then
        Dim rs As Object
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open sql, conn, 3, 1
        If Not rs.EOF Then
            Set usuarioDTOAux = New UsuarioDTO
            usuarioDTOAux.setCveUsuario Nz(rs.Fields(0), "")
            usuarioDTOAux.setPassword Nz(rs.Fields(1), "")
            usuarioDTOAux.setNombre Nz(rs.Fields(2), "")
            usuarioDTOAux.setAPaterno Nz(rs.Fields(3), "")
            usuarioDTOAux.setAMaterno Nz(rs.Fields(4), "")
            usuarioDTOAux.setCvePerfil Nz(rs.Fields(5), "")
            usuarioDTOAux.setEmail Nz(rs.Fields(6), "")
            usuarioDTOAux.setCveEstatus Nz(rs.Fields(7), "")
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    Set selectUsuario = usuarioDTOAux
    Exit Function

errorHandler:
    Set selectUsuario = Nothing
End Function

Public Function selectUsuarioPorClave(ByVal claveUsuario As String) As UsuarioDTO
    Dim sql As String
    Dim usuarioDTOAux As UsuarioDTO
    Dim rsDao As DAO.Recordset
    
    On Error GoTo errorHandler
    sql = ConsultasSQL.sqlSelectUsuarioPorClave(claveUsuario)
    Set rsDao = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    If Not rsDao Is Nothing Then
        If Not rsDao.EOF Then
            Set usuarioDTOAux = New UsuarioDTO
            usuarioDTOAux.setCveUsuario Nz(rsDao.Fields(0), "")
            usuarioDTOAux.setPassword Nz(rsDao.Fields(1), "")
            usuarioDTOAux.setNombre Nz(rsDao.Fields(2), "")
            usuarioDTOAux.setAPaterno Nz(rsDao.Fields(3), "")
            usuarioDTOAux.setAMaterno Nz(rsDao.Fields(4), "")
            usuarioDTOAux.setCvePerfil Nz(rsDao.Fields(5), "")
            usuarioDTOAux.setEmail Nz(rsDao.Fields(6), "")
            usuarioDTOAux.setCveEstatus Nz(rsDao.Fields(7), "")
        End If
        rsDao.Close
        Set rsDao = Nothing
    End If
    Set selectUsuarioPorClave = usuarioDTOAux
    Exit Function

errorHandler:
    Set selectUsuarioPorClave = Nothing
End Function

Public Function insertUsuario(ByVal u As UsuarioDTO) As Boolean
    Dim sql As String
    On Error GoTo errorHandler
    sql = ConsultasSQL.sqlInsertUsuario(u)
    CurrentDb.Execute sql, dbFailOnError
    insertUsuario = True
    Exit Function
errorHandler:
    insertUsuario = False
End Function

Public Function updateUsuario(ByVal u As UsuarioDTO) As Boolean
    Dim sql As String
    On Error GoTo errorHandler
    sql = ConsultasSQL.sqlUpdateUsuario(u)
    CurrentDb.Execute sql, dbFailOnError
    updateUsuario = True
    Exit Function
errorHandler:
    updateUsuario = False
End Function

Public Function deleteLogicoUsuario(ByVal claveUsuario As String) As Boolean
    Dim sql As String
    On Error GoTo errorHandler
    sql = ConsultasSQL.sqlInactivarUsuario(claveUsuario)
    CurrentDb.Execute sql, dbFailOnError
    deleteLogicoUsuario = True
    Exit Function
errorHandler:
    deleteLogicoUsuario = False
End Function

Public Function cambiarContrasena(ByVal claveUsuario As String, ByVal nuevaContrasena As String, Optional ByVal conn As Object = Nothing) As Boolean
    Dim sql As String
    On Error Resume Next
    sql = ConsultasSQL.sqlUpdatePassword(claveUsuario, nuevaContrasena)
    CurrentDb.Execute sql, dbFailOnError
    If Err.Number = 0 Then
        cambiarContrasena = True
        Exit Function
    End If
    On Error GoTo ErrHandler
    If Not conn Is Nothing Then
        conn.Execute sql
        cambiarContrasena = True
        Exit Function
    End If
    cambiarContrasena = False
    Exit Function
ErrHandler:
    cambiarContrasena = False
End Function

Public Function selectInstrumentoPorId(ByVal idInstrumento As String) As InstrumentoDTO
    Dim sql As String
    Dim instDTO As InstrumentoDTO
    Dim rsDao As DAO.Recordset
    
    On Error GoTo errorHandler
    sql = ConsultasSQL.sqlSelectInstrumentoPorId(idInstrumento)
    Set rsDao = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    If Not rsDao Is Nothing Then
        If Not rsDao.EOF Then
            Set instDTO = New InstrumentoDTO
            instDTO.setIdInstrumento Nz(rsDao.Fields(0), "")
            instDTO.setCveEmisor Nz(rsDao.Fields(1), "")
            instDTO.setTipoInstrumento Nz(rsDao.Fields(2), "")
            instDTO.setTasa Val(Nz(rsDao.Fields(3), 0))
            instDTO.setPlazo CLng(Val(Nz(rsDao.Fields(4), 0)))
        End If
        rsDao.Close
        Set rsDao = Nothing
    End If
    Set selectInstrumentoPorId = instDTO
    Exit Function
errorHandler:
    Set selectInstrumentoPorId = Nothing
End Function

Public Function insertInstrumento(ByVal inst As InstrumentoDTO) As Boolean
    Dim sql As String
    On Error GoTo errorHandler
    sql = ConsultasSQL.sqlInsertInstrumento(inst)
    CurrentDb.Execute sql, dbFailOnError
    insertInstrumento = True
    Exit Function
errorHandler:
    insertInstrumento = False
End Function

Public Function updateInstrumento(ByVal inst As InstrumentoDTO) As Boolean
    Dim sql As String
    On Error GoTo errorHandler
    sql = ConsultasSQL.sqlUpdateInstrumento(inst)
    CurrentDb.Execute sql, dbFailOnError
    updateInstrumento = True
    Exit Function
errorHandler:
    updateInstrumento = False
End Function

' Obtiene el catalogo actual de codigos idInstrumento registrados en INSTRUMENTOS como mapa "|ID1|ID2|...|"
Public Function obtenerMapaInstrumentosDAO() As String
    Dim rs As DAO.Recordset
    Dim mapa As String
    On Error GoTo ErrMapa
    mapa = "|"
    Set rs = CurrentDb.OpenRecordset(ConsultasSQL.sqlSelectInstrumentos(), dbOpenSnapshot)
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            If Trim(CStr(Nz(rs!idInstrumento, ""))) <> "" Then
                mapa = mapa & UCase(Trim(CStr(Nz(rs!idInstrumento, "")))) & "|"
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
    obtenerMapaInstrumentosDAO = mapa
    Exit Function
ErrMapa:
    obtenerMapaInstrumentosDAO = "|"
End Function

' Importa un archivo de insumo validado (53 columnas incl. FECHA_INSUMO e ID_INSTRUMENTO) a su base de datos correspondiente
Public Function importarInsumoEnTablaDAO(ByVal rutaArchivo As String, Optional ByVal fechaCorte As String = "", Optional ByVal nombreTablaDestino As String = "BD_INSUMO_PRIMER_PISO", Optional ByVal matrizPrevalidada As Variant, Optional ByVal uFilaPrevalidada As Long = 0, Optional ByVal uColPrevalidada As Long = 0) As Boolean
    Dim xlApp As Object, wb As Object, ws As Object
    Dim rs As DAO.Recordset, rsCons As DAO.Recordset
    Dim uFila As Long, uCol As Long, r As Long, c As Long
    Dim matriz As Variant, fecNorm As String, tblDest As String, pisoFiltro As String
    Dim wsTrans As DAO.Workspace, bEnTrans As Boolean, sqlCopia As String
    Dim rsOrig As DAO.Recordset, nAfect As Long, dbLoc As DAO.Database
    Dim fld(0 To 52) As DAO.Field, nFlds As Long, vCelda As Variant, vt As Integer
    
    On Error GoTo ErrImp
    importarInsumoEnTablaDAO = False
    tblDest = Trim(nombreTablaDestino)
    If tblDest = "" Then tblDest = "BD_INSUMO_PRIMER_PISO"
    fecNorm = Trim(fechaCorte)
    If fecNorm <> "" And IsDate(fecNorm) Then
        fecNorm = Format(CDate(fecNorm), "dd/mm/yyyy")
    End If
    
    If UCase(tblDest) = "BD_INSUMO_PRIMER_PISO" Then
        pisoFiltro = "PRIMERO"
    ElseIf UCase(tblDest) = "BD_INSUMO_TESORERIA" Then
        pisoFiltro = "TESORERIA"
    ElseIf UCase(tblDest) = "BD_INSUMO_SEGUNDO_PISO" Then
        pisoFiltro = "SEGUNDO"
    Else
        pisoFiltro = ""
    End If
    
    On Error Resume Next
    If fecNorm <> "" Then
        CurrentDb.Execute ConsultasSQL.sqlDeleteInsumoTablaPorFecha(tblDest, fecNorm)
        If pisoFiltro <> "" Then
            CurrentDb.Execute ConsultasSQL.sqlDeleteCarteraPorFechaYPiso(fecNorm, pisoFiltro)
        End If
    Else
        CurrentDb.Execute ConsultasSQL.sqlDeleteTablaCompleta(tblDest)
    End If
    On Error GoTo ErrImp
    
    ' RENDIMIENTO: Si FondoBO.validarYProcesarInsumoBO ya leyó y validó la matriz en RAM,
    ' se reutiliza directamente sin volver a instanciar Excel.Application ni abrir el archivo una segunda vez.
    If uFilaPrevalidada >= 2 And uColPrevalidada >= 50 And IsArray(matrizPrevalidada) Then
        matriz = matrizPrevalidada
        uFila = uFilaPrevalidada
        uCol = uColPrevalidada
        If uCol > 53 Then uCol = 53
    Else
        Set xlApp = CreateObject("Excel.Application")
        xlApp.Visible = False
        xlApp.DisplayAlerts = False
        On Error Resume Next
        xlApp.ScreenUpdating = False
        xlApp.EnableEvents = False
        xlApp.Calculation = -4135
        On Error GoTo ErrImp
        Set wb = xlApp.Workbooks.Open(rutaArchivo, 0, True)
        Set ws = wb.Worksheets(1)
        uFila = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        uCol = ws.Cells(1, ws.Columns.Count).End(-4159).Column
        If uCol > 53 Then uCol = 53
        matriz = ws.Range(ws.Cells(1, 1), ws.Cells(uFila, uCol)).Value
        wb.Close False
        xlApp.Quit
        Set ws = Nothing: Set wb = Nothing: Set xlApp = Nothing
    End If
    
    ' RENDIMIENTO DE INSERCIÓN MASIVA DAO:
    '   a) dbAppendOnly: el recordset no carga los registros existentes, solo anexa.
    '   b) Una sola transacción (BeginTrans / CommitTrans) para todo el lote.
    '   c) Pre-enlazado de las 53 referencias DAO.Field en el arreglo fld(0..52) fuera del bucle,
    '      evitando 26,500 búsquedas en la colección rs.Fields(i) y 26,500 llamadas a Nz().
    Set rs = CurrentDb.OpenRecordset(tblDest, dbOpenDynaset, dbAppendOnly)
    nFlds = rs.Fields.Count
    If nFlds > 53 Then nFlds = 53
    For c = 0 To nFlds - 1
        Set fld(c) = rs.Fields(c)
    Next c

    Set wsTrans = DBEngine.Workspaces(0)
    wsTrans.BeginTrans
    bEnTrans = True

    On Error Resume Next
    For r = 2 To uFila
        rs.AddNew
        For c = 1 To 51
            vCelda = matriz(r, c)
            If c >= 30 And c <= 36 Then
                vt = VarType(vCelda)
                If vt = vbDouble Or vt = vbSingle Or vt = vbCurrency Or vt = vbLong Or vt = vbInteger Or vt = vbDecimal Then
                    fld(c - 1).Value = CDbl(vCelda)
                ElseIf vt = vbEmpty Or vt = vbNull Then
                    fld(c - 1).Value = 0#
                Else
                    fld(c - 1).Value = CDbl(Val(Replace(CStr(vCelda), ",", "")))
                End If
            Else
                If IsEmpty(vCelda) Or IsNull(vCelda) Then
                    fld(c - 1).Value = ""
                Else
                    fld(c - 1).Value = Trim$(CStr(vCelda))
                End If
            End If
        Next c
        If fecNorm <> "" Then
            fld(51).Value = fecNorm
        ElseIf uCol >= 52 Then
            vCelda = matriz(r, 52)
            If IsDate(vCelda) Then
                fld(51).Value = Format(CDate(vCelda), "dd/mm/yyyy")
            ElseIf IsEmpty(vCelda) Or IsNull(vCelda) Then
                fld(51).Value = ""
            Else
                fld(51).Value = Trim$(CStr(vCelda))
            End If
        End If
        If nFlds >= 53 Then
            If uCol >= 53 Then
                vCelda = matriz(r, 53)
                If IsEmpty(vCelda) Or IsNull(vCelda) Then
                    fld(52).Value = "BLK_29"
                Else
                    fld(52).Value = Trim$(CStr(vCelda))
                End If
            Else
                fld(52).Value = "BLK_29"
            End If
        End If
        rs.Update
    Next r
    On Error GoTo ErrImp
    For c = 0 To nFlds - 1
        Set fld(c) = Nothing
    Next c
    rs.Close: Set rs = Nothing

    wsTrans.CommitTrans
    bEnTrans = False
    Set wsTrans = Nothing

    ' CARTERA_ETIQUETADA se llena con UNA sentencia SQL a partir de lo ya insertado,
    ' en vez de repetir fila por fila el mismo trabajo (eran otras 17,000 asignaciones).
    If UCase(tblDest) <> "CARTERA_ETIQUETADA" Then
        On Error Resume Next
        sqlCopia = ConsultasSQL.sqlCopiarInsumoACarteraEtiquetada(tblDest, fecNorm)

        ' IMPORTANTE: dbFailOnError (128) obliga a Jet a reportar el fallo.
        ' Sin esta bandera, un INSERT que no inserta nada NO levanta error y
        ' CARTERA_ETIQUETADA se quedaría vacía en silencio. Aquí se mide cuántas
        ' filas entraron y, si entraron cero, se reintenta con copia por recordset.
        Set dbLoc = CurrentDb
        Err.Clear
        nAfect = 0
        dbLoc.Execute sqlCopia, dbFailOnError
        nAfect = dbLoc.RecordsAffected

        If Err.Number <> 0 Or nAfect = 0 Then
            gUltimoErrorSQL = "CARTERA_ETIQUETADA (copia masiva): " & _
                CStr(Err.Number) & " " & Err.Description & " / filas=" & CStr(nAfect)
            Err.Clear

            ' RESPALDO: copia campo por campo por posición ordinal.
            sqlCopia = ConsultasSQL.sqlSelectTablaInsumoPorFecha(tblDest, fecNorm)
            Set rsOrig = dbLoc.OpenRecordset(sqlCopia, dbOpenSnapshot)
            Set rsCons = dbLoc.OpenRecordset("CARTERA_ETIQUETADA", dbOpenDynaset, dbAppendOnly)
            Do While Not rsOrig.EOF
                rsCons.AddNew
                For c = 0 To 52
                    rsCons.Fields(c).Value = rsOrig.Fields(c).Value
                Next c
                rsCons.Update
                rsOrig.MoveNext
            Loop
            rsCons.Close: Set rsCons = Nothing
            rsOrig.Close: Set rsOrig = Nothing
        End If

        Set dbLoc = Nothing
        On Error GoTo 0
    End If

    importarInsumoEnTablaDAO = True
    Exit Function

ErrImp:
    On Error Resume Next
    If bEnTrans Then wsTrans.Rollback: bEnTrans = False
    Set wsTrans = Nothing
    If Not rs Is Nothing Then rs.Close: Set rs = Nothing
    If Not rsCons Is Nothing Then rsCons.Close: Set rsCons = Nothing
    If Not rsOrig Is Nothing Then rsOrig.Close: Set rsOrig = Nothing
    If Not wb Is Nothing Then wb.Close False: Set wb = Nothing
    If Not xlApp Is Nothing Then xlApp.Quit: Set xlApp = Nothing
    DoCmd.SetWarnings False
    DoCmd.TransferSpreadsheet 0, 10, tblDest, rutaArchivo, True
    DoCmd.SetWarnings True
    importarInsumoEnTablaDAO = True
    On Error GoTo 0
End Function

Public Function contarRegistrosTablaPorFechaDAO(ByVal nombreTabla As String, ByVal fechaCorte As String) As Long
    Dim rs As DAO.Recordset
    Dim fecNorm As String
    On Error GoTo ErrCntT
    contarRegistrosTablaPorFechaDAO = 0
    fecNorm = Trim(fechaCorte)
    If IsDate(fecNorm) Then fecNorm = Format(CDate(fecNorm), "dd/mm/yyyy")
    Set rs = CurrentDb.OpenRecordset(ConsultasSQL.sqlContarInsumoTablaPorFecha(nombreTabla, fecNorm), dbOpenSnapshot)
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            contarRegistrosTablaPorFechaDAO = CLng(Val(Nz(rs!TotalCreditos, 0)))
        End If
        rs.Close
        Set rs = Nothing
    End If
    Exit Function
ErrCntT:
    contarRegistrosTablaPorFechaDAO = 0
End Function

' Obtiene en 1 sola consulta indexada el conteo de registros de las 3 tablas de insumo para una fecha
Public Function contarTresInsumosPorFechaDAO(ByVal fechaCorte As String, ByRef nP1 As Long, ByRef nP2 As Long, ByRef nP3 As Long) As Boolean
    Dim rs As DAO.Recordset, fecNorm As String, idxP As Long, nReg As Long
    On Error GoTo ErrCnt3
    contarTresInsumosPorFechaDAO = False
    nP1 = 0: nP2 = 0: nP3 = 0
    fecNorm = Trim(fechaCorte)
    If IsDate(fecNorm) Then fecNorm = Format(CDate(fecNorm), "dd/mm/yyyy")
    Set rs = CurrentDb.OpenRecordset(ConsultasSQL.sqlContarTresInsumosPorFecha(fecNorm), dbOpenSnapshot)
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            idxP = CLng(Val(Nz(rs.Fields(0).Value, 0)))
            nReg = CLng(Val(Nz(rs.Fields(1).Value, 0)))
            If idxP = 1 Then nP1 = nReg
            If idxP = 2 Then nP2 = nReg
            If idxP = 3 Then nP3 = nReg
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        contarTresInsumosPorFechaDAO = True
        Exit Function
    End If
ErrCnt3:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close: Set rs = Nothing
    nP1 = contarRegistrosTablaPorFechaDAO("BD_INSUMO_PRIMER_PISO", fecNorm)
    nP2 = contarRegistrosTablaPorFechaDAO("BD_INSUMO_TESORERIA", fecNorm)
    nP3 = contarRegistrosTablaPorFechaDAO("BD_INSUMO_SEGUNDO_PISO", fecNorm)
    contarTresInsumosPorFechaDAO = True
End Function

Public Function contarRegistrosInsumoPorFechaDAO(ByVal fechaCorte As String) As Long
    Dim n1 As Long, n2 As Long, n3 As Long, nC As Long
    On Error Resume Next
    n1 = contarRegistrosTablaPorFechaDAO("BD_INSUMO_PRIMER_PISO", fechaCorte)
    n2 = contarRegistrosTablaPorFechaDAO("BD_INSUMO_TESORERIA", fechaCorte)
    n3 = contarRegistrosTablaPorFechaDAO("BD_INSUMO_SEGUNDO_PISO", fechaCorte)
    contarRegistrosInsumoPorFechaDAO = n1 + n2 + n3
    If contarRegistrosInsumoPorFechaDAO = 0 Then
        nC = contarRegistrosTablaPorFechaDAO("CARTERA_ETIQUETADA", fechaCorte)
        contarRegistrosInsumoPorFechaDAO = nC
    End If
    On Error GoTo 0
End Function

Public Function obtenerFechasCargadasDAO() As String
    Dim rs As DAO.Recordset
    Dim lista As String
    On Error GoTo ErrFec
    obtenerFechasCargadasDAO = ""
    lista = ""
    Set rs = CurrentDb.OpenRecordset(ConsultasSQL.sqlFechasDisponiblesCartera(), dbOpenSnapshot)
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            If Trim(CStr(Nz(rs!FechaInsumo, ""))) <> "" Then
                If lista <> "" Then lista = lista & ", "
                lista = lista & Trim(CStr(Nz(rs!FechaInsumo, "")))
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
    obtenerFechasCargadasDAO = lista
    Exit Function
ErrFec:
    obtenerFechasCargadasDAO = ""
End Function

Public Function obtenerUltimaFechaInsumoDAO() As String
    Dim rs As DAO.Recordset
    On Error GoTo ErrUlt
    obtenerUltimaFechaInsumoDAO = ""
    Set rs = CurrentDb.OpenRecordset(ConsultasSQL.sqlUltimaFechaBitacora(), dbOpenSnapshot)
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            obtenerUltimaFechaInsumoDAO = Trim(CStr(Nz(rs!fechaCorte, "")))
        End If
        rs.Close
        Set rs = Nothing
    End If
    Exit Function
ErrUlt:
    obtenerUltimaFechaInsumoDAO = ""
End Function

Public Function obtenerMapaFechasConteoDAO() As String
    Dim rs As DAO.Recordset, res As String
    On Error GoTo ErrMapaFec
    res = "|"
    Set rs = CurrentDb.OpenRecordset(ConsultasSQL.sqlMapaFechasConteo(), dbOpenSnapshot)
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            res = res & Trim(CStr(Nz(rs.Fields(0).Value, ""))) & "=" & CStr(Nz(rs.Fields(1).Value, 0)) & "|"
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
    obtenerMapaFechasConteoDAO = res
    Exit Function
ErrMapaFec:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    obtenerMapaFechasConteoDAO = "|"
End Function

' Extrae mediante SQL la tabla de Analisis 1 (Primer Piso por Concepto de Cuenta) desde BD_INSUMO_PRIMER_PISO + INSTRUMENTOS
Public Function extraerAnalisis1PrimerPisoSQL(ByVal fechaCorte As String, ByRef rowsOut As Variant, ByRef nFilasOut As Long, ByRef totCred As Long, ByRef totIniMDP As Double, ByRef totFinMDP As Double, ByRef totPromMDP As Double, ByRef totFinMDD As Double) As Boolean
    Dim rs As DAO.Recordset
    Dim fecNorm As String, i As Long
    Dim cCta As String, monB As String, nCr As Long
    Dim sIni As Double, sFin As Double, sProm As Double, sUsd As Double
    Dim oCr As Long, oIni As Double, oFin As Double, oProm As Double, oUsd As Double
    
    On Error GoTo ErrA1
    extraerAnalisis1PrimerPisoSQL = False
    nFilasOut = 0: totCred = 0: totIniMDP = 0: totFinMDP = 0: totPromMDP = 0: totFinMDD = 0
    fecNorm = Trim(fechaCorte)
    If IsDate(fecNorm) Then fecNorm = Format(CDate(fecNorm), "dd/mm/yyyy")
    
    Set rs = CurrentDb.OpenRecordset(ConsultasSQL.sqlReporteAnalisis1PrimerPiso(fecNorm), dbOpenSnapshot)
    If rs Is Nothing Then Exit Function
    If rs.EOF Then
        rs.Close: Set rs = Nothing
        Exit Function
    End If
    
    Dim tmpCta(1 To 6) As String, tmpMon(1 To 6) As String, tmpCr(1 To 6) As Long
    Dim tmpIni(1 To 6) As Double, tmpFin(1 To 6) As Double, tmpProm(1 To 6) As Double, tmpUsd(1 To 6) As Double
    i = 0: oCr = 0: oIni = 0: oFin = 0: oProm = 0: oUsd = 0
    
    Do While Not rs.EOF
        cCta = Trim(CStr(Nz(rs!ConceptoCuenta, "OTROS")))
        monB = Trim(CStr(Nz(rs!MonedaBase, "M.N. | BLK_29")))
        nCr = CLng(Val(Nz(rs!NumCreditos, 0)))
        sIni = CDbl(Val(Nz(rs!SaldoInicialMDP, 0)))
        sFin = CDbl(Val(Nz(rs!SaldoFinalMDP, 0)))
        sProm = CDbl(Val(Nz(rs!SaldoPromedioMDP, 0)))
        sUsd = CDbl(Val(Nz(rs!EquivalenteUSD_MDD, 0)))
        
        totCred = totCred + nCr
        totIniMDP = totIniMDP + sIni
        totFinMDP = totFinMDP + sFin
        totPromMDP = totPromMDP + sProm
        totFinMDD = totFinMDD + sUsd
        
        If i < 4 Then
            i = i + 1
            tmpCta(i) = CStr(i) & ". " & cCta
            tmpMon(i) = monB
            tmpCr(i) = nCr
            tmpIni(i) = sIni
            tmpFin(i) = sFin
            tmpProm(i) = sProm
            tmpUsd(i) = sUsd
        Else
            oCr = oCr + nCr
            oIni = oIni + sIni
            oFin = oFin + sFin
            oProm = oProm + sProm
            oUsd = oUsd + sUsd
        End If
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing
    
    If oCr > 0 Then
        i = i + 1
        tmpCta(i) = CStr(i) & ". QUIROGRAFARIOS Y PRENDARIAS"
        tmpMon(i) = "M.N. / DLS | Varios"
        tmpCr(i) = oCr
        tmpIni(i) = oIni
        tmpFin(i) = oFin
        tmpProm(i) = oProm
        tmpUsd(i) = oUsd
    End If
    
    nFilasOut = i + 1
    ReDim rowsOut(1 To nFilasOut, 1 To 8)
    Dim k As Long
    For k = 1 To i
        rowsOut(k, 1) = tmpCta(k)
        rowsOut(k, 2) = tmpMon(k)
        rowsOut(k, 3) = tmpCr(k)
        rowsOut(k, 4) = Round(tmpIni(k), 2)
        rowsOut(k, 5) = Round(tmpFin(k), 2)
        rowsOut(k, 6) = Round(tmpProm(k), 2)
        rowsOut(k, 7) = Round(tmpUsd(k), 2)
        If totFinMDP > 0 Then
            rowsOut(k, 8) = Format((tmpFin(k) / totFinMDP) * 100#, "0.00") & "% | VIGENTE E1"
        Else
            rowsOut(k, 8) = "VIGENTE E1"
        End If
    Next k
    
    rowsOut(nFilasOut, 1) = "TOTAL 1. PRIMER PISO (SQL)"
    rowsOut(nFilasOut, 2) = "Insumo " & fecNorm
    rowsOut(nFilasOut, 3) = totCred
    rowsOut(nFilasOut, 4) = Round(totIniMDP, 2)
    rowsOut(nFilasOut, 5) = Round(totFinMDP, 2)
    rowsOut(nFilasOut, 6) = Round(totPromMDP, 2)
    rowsOut(nFilasOut, 7) = Round(totFinMDD, 2)
    rowsOut(nFilasOut, 8) = "100.00% CONCILIADO"
    extraerAnalisis1PrimerPisoSQL = True
    Exit Function
ErrA1:
    extraerAnalisis1PrimerPisoSQL = False
End Function

' Extrae mediante SQL la tabla de Analisis 2 (Tesoreria y Mercado de Dinero) desde BD_INSUMO_TESORERIA + INSTRUMENTOS
Public Function extraerAnalisis2SectoresMonedasSQL(ByVal fechaCorte As String, ByRef rowsOut As Variant, ByRef nFilasOut As Long, ByRef totCred As Long, ByRef totIniMDP As Double, ByRef totFinMDP As Double, ByRef totPromMDP As Double, ByRef totFinMDD As Double) As Boolean
    Dim rs As DAO.Recordset
    Dim fecNorm As String, i As Long
    Dim gSub As String, divP As String, nOp As Long
    Dim sIni As Double, sFin As Double, sProm As Double, sUsd As Double
    Dim oCr As Long, oIni As Double, oFin As Double, oProm As Double, oUsd As Double
    Dim oGrp5 As String, oDiv5 As String
    
    On Error GoTo ErrA2
    extraerAnalisis2SectoresMonedasSQL = False
    nFilasOut = 0: totCred = 0: totIniMDP = 0: totFinMDP = 0: totPromMDP = 0: totFinMDD = 0
    fecNorm = Trim(fechaCorte)
    If IsDate(fecNorm) Then fecNorm = Format(CDate(fecNorm), "dd/mm/yyyy")
    
    Set rs = CurrentDb.OpenRecordset(ConsultasSQL.sqlReporteAnalisis2SectoresMonedas(fecNorm), dbOpenSnapshot)
    If rs Is Nothing Then Exit Function
    If rs.EOF Then
        rs.Close: Set rs = Nothing
        Exit Function
    End If
    
    Dim tmpGrp(1 To 6) As String, tmpDiv(1 To 6) As String, tmpCr(1 To 6) As Long
    Dim tmpIni(1 To 6) As Double, tmpFin(1 To 6) As Double, tmpProm(1 To 6) As Double, tmpUsd(1 To 6) As Double
    i = 0: oCr = 0: oIni = 0: oFin = 0: oProm = 0: oUsd = 0
    oGrp5 = "": oDiv5 = ""
    
    Do While Not rs.EOF
        gSub = Trim(CStr(Nz(rs!GrupoSubgrupo, "SECTOR")))
        divP = Trim(CStr(Nz(rs!DivisaPredominante, "M.N. | CETES_28D")))
        nOp = CLng(Val(Nz(rs!NumOperaciones, 0)))
        sIni = CDbl(Val(Nz(rs!SaldoInicialMDP, 0)))
        sFin = CDbl(Val(Nz(rs!PosicionFinalMDP, 0)))
        sProm = CDbl(Val(Nz(rs!PosicionPromedioMDP, 0)))
        sUsd = CDbl(Val(Nz(rs!PosicionDivisaMDD, 0)))
        
        totCred = totCred + nOp
        totIniMDP = totIniMDP + sIni
        totFinMDP = totFinMDP + sFin
        totPromMDP = totPromMDP + sProm
        totFinMDD = totFinMDD + sUsd
        
        If i < 4 Then
            i = i + 1
            tmpGrp(i) = CStr(i) & ". " & gSub
            tmpDiv(i) = divP
            tmpCr(i) = nOp
            tmpIni(i) = sIni
            tmpFin(i) = sFin
            tmpProm(i) = sProm
            tmpUsd(i) = sUsd
        Else
            If oGrp5 = "" Then
                oGrp5 = gSub
                oDiv5 = divP
            Else
                oGrp5 = oGrp5 & " + OTROS"
            End If
            oCr = oCr + nOp
            oIni = oIni + sIni
            oFin = oFin + sFin
            oProm = oProm + sProm
            oUsd = oUsd + sUsd
        End If
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing
    
    If oCr > 0 Then
        i = i + 1
        tmpGrp(i) = CStr(i) & ". " & oGrp5
        tmpDiv(i) = oDiv5
        tmpCr(i) = oCr
        tmpIni(i) = oIni
        tmpFin(i) = oFin
        tmpProm(i) = oProm
        tmpUsd(i) = oUsd
    End If
    
    nFilasOut = i + 1
    ReDim rowsOut(1 To nFilasOut, 1 To 8)
    Dim k As Long
    For k = 1 To i
        rowsOut(k, 1) = tmpGrp(k)
        rowsOut(k, 2) = tmpDiv(k)
        rowsOut(k, 3) = tmpCr(k)
        rowsOut(k, 4) = Round(tmpIni(k), 2)
        rowsOut(k, 5) = Round(tmpFin(k), 2)
        rowsOut(k, 6) = Round(tmpProm(k), 2)
        rowsOut(k, 7) = Round(tmpUsd(k), 2)
        If totFinMDP > 0 Then
            rowsOut(k, 8) = Format((tmpFin(k) / totFinMDP) * 100#, "0.00") & "% | EN LÍMITE NORMATIVA"
        Else
            rowsOut(k, 8) = "EN LÍMITE NORMATIVA"
        End If
    Next k
    
    rowsOut(nFilasOut, 1) = "TOTAL 2. TESORERÍA Y MERCADO (SQL)"
    rowsOut(nFilasOut, 2) = "Insumo " & fecNorm
    rowsOut(nFilasOut, 3) = totCred
    rowsOut(nFilasOut, 4) = Round(totIniMDP, 2)
    rowsOut(nFilasOut, 5) = Round(totFinMDP, 2)
    rowsOut(nFilasOut, 6) = Round(totPromMDP, 2)
    rowsOut(nFilasOut, 7) = Round(totFinMDD, 2)
    rowsOut(nFilasOut, 8) = "100.00% CONCILIADO"
    extraerAnalisis2SectoresMonedasSQL = True
    Exit Function
ErrA2:
    extraerAnalisis2SectoresMonedasSQL = False
End Function

' Extrae mediante SQL la tabla de Analisis 3 (Segundo Piso por Subgrupo y Actividad) desde BD_INSUMO_SEGUNDO_PISO + INSTRUMENTOS
Public Function extraerAnalisis3SegundoPisoSQL(ByVal fechaCorte As String, ByRef rowsOut As Variant, ByRef nFilasOut As Long, ByRef totCred As Long, ByRef totIniMDP As Double, ByRef totFinMDP As Double, ByRef totPromMDP As Double, ByRef totFinMDD As Double) As Boolean
    Dim rs As DAO.Recordset
    Dim fecNorm As String, i As Long
    Dim segAct As String, etNorm As String, nCr As Long
    Dim sIni As Double, sFin As Double, sProm As Double, sUsd As Double
    Dim oCr As Long, oIni As Double, oFin As Double, oProm As Double, oUsd As Double
    Dim oEt5 As String
    
    On Error GoTo ErrA3
    extraerAnalisis3SegundoPisoSQL = False
    nFilasOut = 0: totCred = 0: totIniMDP = 0: totFinMDP = 0: totPromMDP = 0: totFinMDD = 0
    fecNorm = Trim(fechaCorte)
    If IsDate(fecNorm) Then fecNorm = Format(CDate(fecNorm), "dd/mm/yyyy")
    
    Set rs = CurrentDb.OpenRecordset(ConsultasSQL.sqlReporteAnalisis3SegundoPiso(fecNorm), dbOpenSnapshot)
    If rs Is Nothing Then Exit Function
    If rs.EOF Then
        rs.Close: Set rs = Nothing
        Exit Function
    End If
    
    Dim tmpSeg(1 To 6) As String, tmpEt(1 To 6) As String, tmpCr(1 To 6) As Long
    Dim tmpIni(1 To 6) As Double, tmpFin(1 To 6) As Double, tmpProm(1 To 6) As Double, tmpUsd(1 To 6) As Double
    i = 0: oCr = 0: oIni = 0: oFin = 0: oProm = 0: oUsd = 0: oEt5 = ""
    
    Do While Not rs.EOF
        segAct = Trim(CStr(Nz(rs!SegmentoActividad, "SEGUNDO PISO")))
        etNorm = Trim(CStr(Nz(rs!EtapaNormativa, "M.N. | NAFIN_28")))
        nCr = CLng(Val(Nz(rs!NumCreditos, 0)))
        sIni = CDbl(Val(Nz(rs!SaldoInicialMDP, 0)))
        sFin = CDbl(Val(Nz(rs!SaldoVigenteMDP, 0)))
        sProm = CDbl(Val(Nz(rs!SaldoPromedioMDP, 0)))
        sUsd = CDbl(Val(Nz(rs!EquivalenteUSD_MDD, 0)))
        
        totCred = totCred + nCr
        totIniMDP = totIniMDP + sIni
        totFinMDP = totFinMDP + sFin
        totPromMDP = totPromMDP + sProm
        totFinMDD = totFinMDD + sUsd
        
        If i < 4 Then
            i = i + 1
            tmpSeg(i) = CStr(i) & ". " & segAct
            tmpEt(i) = etNorm
            tmpCr(i) = nCr
            tmpIni(i) = sIni
            tmpFin(i) = sFin
            tmpProm(i) = sProm
            tmpUsd(i) = sUsd
        Else
            If oEt5 = "" Then oEt5 = etNorm
            oCr = oCr + nCr
            oIni = oIni + sIni
            oFin = oFin + sFin
            oProm = oProm + sProm
            oUsd = oUsd + sUsd
        End If
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing
    
    If oCr > 0 Then
        i = i + 1
        tmpSeg(i) = CStr(i) & ". OTROS PROGRAMAS 2DO PISO"
        tmpEt(i) = oEt5
        tmpCr(i) = oCr
        tmpIni(i) = oIni
        tmpFin(i) = oFin
        tmpProm(i) = oProm
        tmpUsd(i) = oUsd
    End If
    
    nFilasOut = i + 1
    ReDim rowsOut(1 To nFilasOut, 1 To 8)
    Dim k As Long
    For k = 1 To i
        rowsOut(k, 1) = tmpSeg(k)
        rowsOut(k, 2) = tmpEt(k)
        rowsOut(k, 3) = tmpCr(k)
        rowsOut(k, 4) = Round(tmpIni(k), 2)
        rowsOut(k, 5) = Round(tmpFin(k), 2)
        rowsOut(k, 6) = Round(tmpProm(k), 2)
        rowsOut(k, 7) = Round(tmpUsd(k), 2)
        If totFinMDP > 0 Then
            rowsOut(k, 8) = Format((tmpFin(k) / totFinMDP) * 100#, "0.00") & "% | ETAPA E1 (A-1)"
        Else
            rowsOut(k, 8) = "ETAPA E1 (A-1)"
        End If
    Next k
    
    rowsOut(nFilasOut, 1) = "TOTAL 3. SEGUNDO PISO (SQL)"
    rowsOut(nFilasOut, 2) = "Insumo " & fecNorm
    rowsOut(nFilasOut, 3) = totCred
    rowsOut(nFilasOut, 4) = Round(totIniMDP, 2)
    rowsOut(nFilasOut, 5) = Round(totFinMDP, 2)
    rowsOut(nFilasOut, 6) = Round(totPromMDP, 2)
    rowsOut(nFilasOut, 7) = Round(totFinMDD, 2)
    rowsOut(nFilasOut, 8) = "100.00% CONCILIADO"
    extraerAnalisis3SegundoPisoSQL = True
    Exit Function
ErrA3:
    extraerAnalisis3SegundoPisoSQL = False
End Function

' Consulta en BITACORA_CARGA_INSUMOS (vía ConsultasSQL) si un archivo ya se cargó para esa fecha
Public Function existeInsumoCargadoDAO(ByVal nombreArchivo As String, ByVal fechaCorte As String, ByRef infoPrevia As String) As Boolean
    Dim rs As DAO.Recordset
    Dim sql As String
    On Error GoTo ErrBit
    existeInsumoCargadoDAO = False
    infoPrevia = ""
    sql = ConsultasSQL.sqlSelectBitacoraPorArchivoYFecha(nombreArchivo, fechaCorte)
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            existeInsumoCargadoDAO = True
            infoPrevia = "Archivo: " & Nz(rs!nombreArchivo, "") & " | Cargado el: " & Format(Nz(rs!fechaHoraCarga, Now), "dd/mm/yyyy hh:nn:ss")
        End If
        rs.Close
        Set rs = Nothing
    End If
    Exit Function
ErrBit:
    existeInsumoCargadoDAO = False
End Function

' Registra en la tabla vinculada BITACORA_CARGA_INSUMOS (en BD\RIESGOSBD.accdb) la carga exitosa de un insumo
Public Function registrarBitacoraInsumoDAO(ByVal nombreArchivo As String, ByVal fechaCorte As String, ByVal totalReg As Long, ByVal fondosStr As String) As Boolean
    Dim sql As String
    On Error Resume Next
    sql = ConsultasSQL.sqlInsertBitacoraInsumo(nombreArchivo, fechaCorte, totalReg, fondosStr)
    CurrentDb.Execute sql, dbFailOnError
    registrarBitacoraInsumoDAO = (Err.Number = 0)
    On Error GoTo 0
End Function

' =============================================================================
' EXTRACCION DE LOS REPORTES COMPARATIVOS (CRUCE DE DOS Y DE TRES INSUMOS)
' -----------------------------------------------------------------------------
' nInsumos = 2 -> cruza BD_INSUMO_PRIMER_PISO con BD_INSUMO_TESORERIA
' nInsumos = 3 -> cruza las tres bases de insumo
'
' Devuelve rowsOut como un arreglo de 8 columnas, con la misma forma que los
' reportes 1, 2 y 3, para poder reutilizar tal cual el formateador de tablas y
' el constructor de graficas.
'
'   nInsumos = 2   A Instrumento | B Tipo | C Operaciones | D Saldo 1er Piso
'                  E Saldo Tesoreria | F Saldo Combinado | G Equiv. USD | H Presencia
'   nInsumos = 3   A Instrumento | B Tipo | C Operaciones | D Saldo 1er Piso
'                  E Saldo Tesoreria | F Saldo 2do Piso | G Saldo Consolidado | H Presencia
'
' Camino principal: la consulta cruzada de ConsultasSQL (subconsulta con UNION
' ALL + GROUP BY por fuera, equivalente a un FULL OUTER JOIN que Jet no tiene).
' Si ese motor rechazara la subconsulta, se usa el camino de respaldo: una
' consulta simple de agregacion por tabla y la union en memoria. El resultado
' numerico es identico; solo cambia donde se junta la informacion.
' =============================================================================
Public Function extraerComparativoInsumosSQL(ByVal fechaCorte As String, ByVal nInsumos As Integer, _
        ByRef rowsOut As Variant, ByRef nFilasOut As Long, _
        ByRef totOps As Long, ByRef totFinMDP As Double, ByRef msgErrorCmp As String) As Boolean
    Dim rs As DAO.Recordset, fecNorm As String, sql As String
    Dim n As Long, i As Long, bOk As Boolean
    Dim clave() As String, tipo() As String
    Dim ops() As Long, f1() As Double, f2() As Double, f3() As Double, usd() As Double
    Dim libros() As Integer, pres As String, totFin As Double

    On Error GoTo ErrCmp
    extraerComparativoInsumosSQL = False
    nFilasOut = 0: totOps = 0: totFinMDP = 0: msgErrorCmp = ""
    If nInsumos < 2 Then nInsumos = 2
    If nInsumos > 3 Then nInsumos = 3

    fecNorm = Trim(fechaCorte)
    If IsDate(fecNorm) Then fecNorm = Format(CDate(fecNorm), "dd/mm/yyyy")

    If nInsumos = 2 Then
        sql = ConsultasSQL.sqlComparativoDosInsumos(fecNorm)
    Else
        sql = ConsultasSQL.sqlComparativoTresInsumos(fecNorm)
    End If

    ' --- Camino principal: la consulta cruzada ---
    bOk = False
    On Error Resume Next
    Err.Clear
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    If Err.Number = 0 And Not rs Is Nothing Then bOk = True
    If Err.Number <> 0 Then
        gUltimoErrorSQL = "Consulta cruzada de " & CStr(nInsumos) & " insumos: " & Err.Description
        Err.Clear
    End If
    On Error GoTo ErrCmp

    If Not bOk Then
        ' --- Respaldo: agregacion por tabla y union en memoria ---
        extraerComparativoInsumosSQL = comparativoPorRespaldoDAO(fecNorm, nInsumos, rowsOut, nFilasOut, totOps, totFinMDP, msgErrorCmp)
        Exit Function
    End If

    If rs.EOF Then
        rs.Close: Set rs = Nothing
        msgErrorCmp = "No hay informacion para cruzar en la fecha " & fecNorm & "."
        Exit Function
    End If

    n = 0
    ReDim clave(1 To 600): ReDim tipo(1 To 600): ReDim ops(1 To 600)
    ReDim f1(1 To 600): ReDim f2(1 To 600): ReDim f3(1 To 600)
    ReDim usd(1 To 600): ReDim libros(1 To 600)

    Do While Not rs.EOF
        If n >= 600 Then Exit Do
        n = n + 1
        clave(n) = Trim(CStr(Nz(rs.Fields(0).Value, "SIN_CLAVE")))
        tipo(n) = Trim(CStr(Nz(rs.Fields(1).Value, "NO REGISTRADO")))
        If nInsumos = 2 Then
            ops(n) = CLng(Val(Nz(rs.Fields(2).Value, 0))) + CLng(Val(Nz(rs.Fields(3).Value, 0)))
            f1(n) = CDbl(Val(Nz(rs.Fields(4).Value, 0)))
            f2(n) = CDbl(Val(Nz(rs.Fields(5).Value, 0)))
            f3(n) = CDbl(Val(Nz(rs.Fields(6).Value, 0)))
            usd(n) = CDbl(Val(Nz(rs.Fields(7).Value, 0)))
            libros(n) = IIf(CLng(Val(Nz(rs.Fields(2).Value, 0))) > 0, 1, 0) + _
                        IIf(CLng(Val(Nz(rs.Fields(3).Value, 0))) > 0, 1, 0)
            totFin = totFin + f3(n)
        Else
            ops(n) = CLng(Val(Nz(rs.Fields(2).Value, 0)))
            f1(n) = CDbl(Val(Nz(rs.Fields(3).Value, 0)))
            f2(n) = CDbl(Val(Nz(rs.Fields(4).Value, 0)))
            f3(n) = CDbl(Val(Nz(rs.Fields(5).Value, 0)))
            usd(n) = CDbl(Val(Nz(rs.Fields(6).Value, 0)))
            libros(n) = CInt(Val(Nz(rs.Fields(7).Value, 0)))
            totFin = totFin + CDbl(Val(Nz(rs.Fields(6).Value, 0)))
        End If
        totOps = totOps + ops(n)
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing

    If n = 0 Then
        msgErrorCmp = "El cruce no devolvio instrumentos para la fecha " & fecNorm & "."
        Exit Function
    End If

    ReDim rowsOut(1 To n, 1 To 8)
    For i = 1 To n
        If nInsumos = 2 Then
            Select Case libros(i)
                Case 2: pres = "AMBOS LIBROS"
                Case 1: pres = IIf(f1(i) > 0, "SOLO 1ER PISO", "SOLO TESORERIA")
                Case Else: pres = "SIN PRESENCIA"
            End Select
            rowsOut(i, 4) = f1(i)
            rowsOut(i, 5) = f2(i)
            rowsOut(i, 6) = f3(i)
            rowsOut(i, 7) = usd(i)
        Else
            pres = CStr(libros(i)) & " de 3 libros"
            rowsOut(i, 4) = f1(i)
            rowsOut(i, 5) = f2(i)
            rowsOut(i, 6) = f3(i)
            rowsOut(i, 7) = CDbl(f1(i) + f2(i) + f3(i))
        End If
        rowsOut(i, 1) = clave(i)
        rowsOut(i, 2) = tipo(i)
        rowsOut(i, 3) = ops(i)
        rowsOut(i, 8) = pres
    Next i

    nFilasOut = n
    totFinMDP = IIf(nInsumos = 2, totFin, totFin)
    If nInsumos = 3 Then
        totFinMDP = 0
        For i = 1 To n
            totFinMDP = totFinMDP + CDbl(rowsOut(i, 7))
        Next i
    End If
    extraerComparativoInsumosSQL = True
    Exit Function

ErrCmp:
    msgErrorCmp = "Error al cruzar los insumos: " & Err.Description
    gUltimoErrorSQL = "extraerComparativoInsumosSQL: " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close: Set rs = Nothing
    extraerComparativoInsumosSQL = False
End Function

' Respaldo del cruce: una consulta simple de agregacion por cada tabla de
' insumo y la union en memoria con un Dictionary. Se usa solo si el motor Jet
' de la instalacion rechaza la subconsulta con UNION ALL.
Private Function comparativoPorRespaldoDAO(ByVal fecNorm As String, ByVal nInsumos As Integer, _
        ByRef rowsOut As Variant, ByRef nFilasOut As Long, _
        ByRef totOps As Long, ByRef totFinMDP As Double, ByRef msgErrorCmp As String) As Boolean
    Dim tablas As Variant, t As Integer, rs As DAO.Recordset
    Dim dic As Object, orden As Object, k As String, i As Long, n As Long
    Dim acOps As Object, acF1 As Object, acF2 As Object, acF3 As Object, acUsd As Object
    Dim acTipo As Object, acLib As Object, pres As String

    On Error GoTo ErrResp
    comparativoPorRespaldoDAO = False
    nFilasOut = 0: totOps = 0: totFinMDP = 0

    If nInsumos = 2 Then
        tablas = Array("BD_INSUMO_PRIMER_PISO", "BD_INSUMO_TESORERIA")
    Else
        tablas = Array("BD_INSUMO_PRIMER_PISO", "BD_INSUMO_TESORERIA", "BD_INSUMO_SEGUNDO_PISO")
    End If

    Set dic = CreateObject("Scripting.Dictionary")
    Set orden = CreateObject("Scripting.Dictionary")
    Set acOps = CreateObject("Scripting.Dictionary")
    Set acF1 = CreateObject("Scripting.Dictionary")
    Set acF2 = CreateObject("Scripting.Dictionary")
    Set acF3 = CreateObject("Scripting.Dictionary")
    Set acUsd = CreateObject("Scripting.Dictionary")
    Set acTipo = CreateObject("Scripting.Dictionary")
    Set acLib = CreateObject("Scripting.Dictionary")

    n = 0
    For t = 0 To UBound(tablas)
        Set rs = CurrentDb.OpenRecordset( _
            ConsultasSQL.sqlAgregadoPorInstrumentoTabla(CStr(tablas(t)), fecNorm), dbOpenSnapshot)
        If Not rs Is Nothing Then
            Do While Not rs.EOF
                k = Trim(CStr(Nz(rs.Fields(0).Value, "SIN_CLAVE")))
                If Not dic.Exists(k) Then
                    n = n + 1
                    dic.Add k, n
                    orden.Add n, k
                    acOps.Add k, CLng(0)
                    acF1.Add k, CDbl(0)
                    acF2.Add k, CDbl(0)
                    acF3.Add k, CDbl(0)
                    acUsd.Add k, CDbl(0)
                    acLib.Add k, CInt(0)
                    acTipo.Add k, Trim(CStr(Nz(rs.Fields(1).Value, "NO REGISTRADO")))
                End If
                acOps(k) = acOps(k) + CLng(Val(Nz(rs.Fields(2).Value, 0)))
                acUsd(k) = acUsd(k) + CDbl(Val(Nz(rs.Fields(4).Value, 0)))
                acLib(k) = acLib(k) + 1
                Select Case t
                    Case 0: acF1(k) = acF1(k) + CDbl(Val(Nz(rs.Fields(3).Value, 0)))
                    Case 1: acF2(k) = acF2(k) + CDbl(Val(Nz(rs.Fields(3).Value, 0)))
                    Case Else: acF3(k) = acF3(k) + CDbl(Val(Nz(rs.Fields(3).Value, 0)))
                End Select
                rs.MoveNext
            Loop
            rs.Close: Set rs = Nothing
        End If
    Next t

    If n = 0 Then
        msgErrorCmp = "No hay informacion para cruzar en la fecha " & fecNorm & "."
        Exit Function
    End If

    ReDim rowsOut(1 To n, 1 To 8)
    For i = 1 To n
        k = CStr(orden(i))
        rowsOut(i, 1) = k
        rowsOut(i, 2) = CStr(acTipo(k))
        rowsOut(i, 3) = CLng(acOps(k))
        rowsOut(i, 4) = CDbl(acF1(k))
        rowsOut(i, 5) = CDbl(acF2(k))
        If nInsumos = 2 Then
            rowsOut(i, 6) = CDbl(acF1(k) + acF2(k))
            rowsOut(i, 7) = CDbl(acUsd(k))
            Select Case CInt(acLib(k))
                Case 2: pres = "AMBOS LIBROS"
                Case Else: pres = IIf(CDbl(acF1(k)) > 0, "SOLO 1ER PISO", "SOLO TESORERIA")
            End Select
        Else
            rowsOut(i, 6) = CDbl(acF3(k))
            rowsOut(i, 7) = CDbl(acF1(k) + acF2(k) + acF3(k))
            pres = CStr(acLib(k)) & " de 3 libros"
        End If
        rowsOut(i, 8) = pres
        totOps = totOps + CLng(acOps(k))
        totFinMDP = totFinMDP + CDbl(rowsOut(i, IIf(nInsumos = 2, 6, 7)))
    Next i

    nFilasOut = n
    comparativoPorRespaldoDAO = True
    Exit Function
ErrResp:
    msgErrorCmp = "Error en el cruce de respaldo: " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close: Set rs = Nothing
    comparativoPorRespaldoDAO = False
End Function
