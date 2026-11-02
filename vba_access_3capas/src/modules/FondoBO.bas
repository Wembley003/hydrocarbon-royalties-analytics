Attribute VB_Name = "FondoBO"
Option Compare Database
Option Explicit

' =============================================================================
' VARIABLES GLOBALES DE DIAGNOSTICO (seccion de declaraciones del modulo)
' VBA exige que TODA declaracion a nivel de modulo este aqui arriba, antes del
' primer Sub/Function. Si se coloca entre procedimientos, Access responde
' "Solo pueden aparecer comentarios despues de End Sub / End Function" y el
' proyecto no compila.
'   gUltimoErrorCorreo : ultimo fallo al enviar el acuse de cambio de contrasena.
'   gUltimoErrorSQL    : ultimo fallo de una sentencia SQL de carga de insumos.
' Ninguna de las dos detiene la operacion; sirven para que el area tecnica
' pueda consultar que paso sin tener que reproducir el error.
' =============================================================================
Public gUltimoErrorCorreo As String
Public gUltimoErrorSQL As String

' -----------------------------------------------------------------------------
' COORDENADAS DEL TEMPLATE Generico.xlsx
' Nombres de hoja y filas clave. Estan aqui, en la cabecera del modulo, porque
' VBA solo admite declaraciones a nivel de modulo antes del primer Sub/Function.
' Si el template se reestructura, se cambian AQUI y en ninguna otra parte.
' -----------------------------------------------------------------------------
Public Const HOJA_TPL_REPORTE As String = "REPORTE"
Public Const HOJA_TPL_DATOS As String = "DATOS"
Public Const HOJA_TPL_CALCULO As String = "CALCULO"
Public Const HOJA_TPL_CALCINSTR As String = "CALC_INSTR"
Public Const TPL_CALC_FILA_INI As Long = 2
Public Const TPL_CALC_FILA_TOT As Long = 5
Public Const TPL_INSTR_FILA_INI As Long = 2
Public Const TPL_INSTR_FILA_FIN As Long = 201

Public Const TPL_PORTAFOLIO_1 As String = "1. Cartera de Credito Comercial Primer Piso"
Public Const TPL_PORTAFOLIO_2 As String = "2. Portafolio de Tesoreria y Mercado de Dinero"
Public Const TPL_PORTAFOLIO_3 As String = "3. Financiamiento Sectorial Segundo Piso"

' -----------------------------------------------------------------------------
' CLAVES DE PANTALLA PARA EL CONTROL DE ACCESO
' Cada pantalla restringida tiene una clave. La regla de quien puede entrar y el
' texto que se le muestra al que no puede viven en autorizarPantallaBO, en la
' capa de Negocio, NO en los formularios.
' Estan aqui, en la cabecera del modulo, porque VBA solo admite declaraciones a
' nivel de modulo antes del primer Sub o Function.
' -----------------------------------------------------------------------------
Public Const PANT_MENU_GENERAL As String = "MENU_GENERAL"
Public Const PANT_GESTION_USUARIOS As String = "GESTION_USUARIOS"
Public Const PANT_CONSULTA_USUARIOS As String = "CONSULTA_USUARIOS"
Public Const PANT_ALTA_USUARIOS As String = "ALTA_USUARIOS"
Public Const PANT_CAMBIO_USUARIOS As String = "CAMBIO_USUARIOS"
Public Const PANT_BAJA_USUARIOS As String = "BAJA_USUARIOS"
Public Const PANT_RESTABLECER_CONTRASENA As String = "RESTABLECER_CONTRASENA"
Public Const PANT_CAMBIO_MI_CONTRASENA As String = "CAMBIO_MI_CONTRASENA"
Public Const PANT_GESTION_INSTRUMENTOS As String = "GESTION_INSTRUMENTOS"
Public Const PANT_CONSULTA_INSTRUMENTOS As String = "CONSULTA_INSTRUMENTOS"
Public Const PANT_ALTA_INSTRUMENTOS As String = "ALTA_INSTRUMENTOS"
Public Const PANT_CAMBIO_INSTRUMENTOS As String = "CAMBIO_INSTRUMENTOS"
Public Const PANT_CARGA_INSUMOS As String = "CARGA_INSUMOS"
Public Const PANT_GENERAR_REPORTE As String = "GENERAR_REPORTE"

' -----------------------------------------------------------------------------
' CACHÉ EN MEMORIA RAM Y SESIÓN REUTILIZABLE DE EXCEL (ACELERACIÓN DE PROCESOS)
' Permite:
'   1) Reutilizar la matriz 2D recién validada en RAM al guardar en BD sin abrir
'      Excel una segunda vez por archivo.
'   2) Mantener una única instancia oculta de Excel.Application durante cargas o
'      exportaciones múltiples de reportes, ahorrando 2-3 seg por cada archivo.
'   3) Cachear conteos y agregaciones SQL de la fecha activa (invalidándose en
'      automático al cargar cualquier insumo o modificar instrumentos).
' -----------------------------------------------------------------------------
Private mXlAppSesion As Object
Private mSesionExcelActiva As Boolean
Private mCacheRutaValidada As String
Private mCacheFechaValidada As String
Private mCacheMatrizValidada As Variant
Private mCacheUFilaValidada As Long
Private mCacheUColValidada As Long
Private mCacheFecConteo As String
Private mCacheNP1 As Long
Private mCacheNP2 As Long
Private mCacheNP3 As Long
Private mCacheMapaInst As String
Private mCacheFecAnalisis As String
Private mCacheRowsBD As Variant
Private mCacheNFilasBD As Long
Private mCacheCredBD As Long
Private mCacheIniBD As Double
Private mCacheFinBD As Double
Private mCachePromBD As Double
Private mCacheUsdBD As Double
Private mCacheRowsCD As Variant
Private mCacheNFilasCD As Long
Private mCacheCredCD As Long
Private mCacheIniCD As Double
Private mCacheFinCD As Double
Private mCachePromCD As Double
Private mCacheUsdCD As Double
Private mCacheRowsPEA As Variant
Private mCacheNFilasPEA As Long
Private mCacheCredPEA As Long
Private mCacheIniPEA As Double
Private mCacheFinPEA As Double
Private mCachePromPEA As Double
Private mCacheUsdPEA As Double


' =============================================================================
' CAPA 2 (REGLAS DE NEGOCIO - BUSINESS OBJECT): FondoBO.bas
' -----------------------------------------------------------------------------
' Concentra el 100% de las reglas de negocio, validaciones normativas NORMATIVA,
' seguridad de contraseñas, autorización por perfil de usuario y generación de
' reportes ejecutivos y comparativos en Excel (.xlsx) y PDF (.pdf).
'
' Principios de Arquitectura de 3 Capas:
'   1. Es el único puente entre los formularios (Presentación) y FondoDAO (Datos).
'   2. Jamás ejecuta SQL directo ni abre Recordsets (delega toda persistencia a FondoDAO).
' =============================================================================

' Verifica que todas las tablas vinculadas del Frontend apunten a BD\RIESGOSBD.accdb (fuera del módulo principal)
Public Function verificarConexionBackendBO() As Boolean
    Dim dao As FondoDAO
    On Error Resume Next
    Set dao = New FondoDAO
    verificarConexionBackendBO = dao.verificarYRelinkearBackendDAO()
    Set dao = Nothing
    On Error GoTo 0
End Function

' Limpia la caché en memoria cuando se cargan insumos o se modifican instrumentos
Public Sub InvalidarCacheDatosBO()
    On Error Resume Next
    mCacheRutaValidada = ""
    mCacheFechaValidada = ""
    Erase mCacheMatrizValidada
    mCacheUFilaValidada = 0
    mCacheUColValidada = 0
    mCacheFecConteo = ""
    mCacheNP1 = 0: mCacheNP2 = 0: mCacheNP3 = 0
    mCacheMapaInst = ""
    mCacheFecAnalisis = ""
    Err.Clear
End Sub

' Inicia una sesión oculta de Excel reutilizable para lotes de validación o generación de reportes
Public Sub IniciarSesionExcelRapidaBO()
    On Error Resume Next
    mSesionExcelActiva = True
    If mXlAppSesion Is Nothing Then
        Set mXlAppSesion = CreateObject("Excel.Application")
        If Not mXlAppSesion Is Nothing Then
            mXlAppSesion.Visible = False
            mXlAppSesion.DisplayAlerts = False
            mXlAppSesion.ScreenUpdating = False
            mXlAppSesion.EnableEvents = False
            mXlAppSesion.AutomationSecurity = 3
        End If
    End If
    Err.Clear
End Sub

' Cierra y libera la sesión oculta de Excel al terminar un lote de operaciones
Public Sub FinalizarSesionExcelRapidaBO()
    On Error Resume Next
    mSesionExcelActiva = False
    If Not mXlAppSesion Is Nothing Then
        mXlAppSesion.DisplayAlerts = False
        mXlAppSesion.Quit
        Set mXlAppSesion = Nothing
    End If
    Err.Clear
End Sub

' Devuelve una instancia lista de Excel.Application (reutilizando la de lote si está activa)
Private Function ObtenerExcelMotorBO(ByRef xlOut As Object, ByRef cerrarAlSalir As Boolean) As Boolean
    On Error Resume Next
    ObtenerExcelMotorBO = False
    cerrarAlSalir = True
    If mSesionExcelActiva Then
        If mXlAppSesion Is Nothing Then
            Set mXlAppSesion = CreateObject("Excel.Application")
        End If
        If Not mXlAppSesion Is Nothing Then
            Err.Clear
            mXlAppSesion.Visible = False
            If Err.Number = 0 Then
                mXlAppSesion.DisplayAlerts = False
                mXlAppSesion.ScreenUpdating = False
                mXlAppSesion.EnableEvents = False
                Set xlOut = mXlAppSesion
                cerrarAlSalir = False
                ObtenerExcelMotorBO = True
                Exit Function
            Else
                Err.Clear
                Set mXlAppSesion = Nothing
            End If
        End If
    End If
    Err.Clear
    Set xlOut = CreateObject("Excel.Application")
    If Not xlOut Is Nothing Then
        xlOut.Visible = False
        xlOut.DisplayAlerts = False
        xlOut.ScreenUpdating = False
        xlOut.EnableEvents = False
        cerrarAlSalir = True
        ObtenerExcelMotorBO = True
    End If
End Function

' Obtiene en 1 sola llamada (con caché en RAM para la fecha activa) los conteos de las 3 tablas de insumo
Public Sub obtenerConteosTresTablasFechaBO(ByVal fechaCorte As String, ByRef nP1 As Long, ByRef nP2 As Long, ByRef nP3 As Long)
    Dim dao As FondoDAO, fecNorm As String
    On Error Resume Next
    nP1 = 0: nP2 = 0: nP3 = 0
    If Not esFechaFormatoCorrectoBO(fechaCorte) Then Exit Sub
    fecNorm = normalizarFechaBO(fechaCorte)
    If mCacheFecConteo = fecNorm And fecNorm <> "" Then
        nP1 = mCacheNP1
        nP2 = mCacheNP2
        nP3 = mCacheNP3
        Exit Sub
    End If
    Set dao = New FondoDAO
    Call dao.contarTresInsumosPorFechaDAO(fecNorm, nP1, nP2, nP3)
    Set dao = Nothing
    mCacheFecConteo = fecNorm
    mCacheNP1 = nP1
    mCacheNP2 = nP2
    mCacheNP3 = nP3
End Sub

' Extrae (o recupera de caché en RAM si ya se consultó la misma fecha) las 3 tablas analíticas SQL
Private Sub obtenerAnalisisTresPortafoliosBO(ByVal fecNorm As String, _
        ByRef rowsBD As Variant, ByRef nFilasBD As Long, ByRef credBD As Long, ByRef iniBD As Double, ByRef finBD As Double, ByRef promBD As Double, ByRef usdBD As Double, _
        ByRef rowsCD As Variant, ByRef nFilasCD As Long, ByRef credCD As Long, ByRef iniCD As Double, ByRef finCD As Double, ByRef promCD As Double, ByRef usdCD As Double, _
        ByRef rowsPEA As Variant, ByRef nFilasPEA As Long, ByRef credPEA As Long, ByRef iniPEA As Double, ByRef finPEA As Double, ByRef promPEA As Double, ByRef usdPEA As Double)
    Dim dao As FondoDAO
    On Error Resume Next
    If mCacheFecAnalisis = fecNorm And fecNorm <> "" Then
        rowsBD = mCacheRowsBD: nFilasBD = mCacheNFilasBD: credBD = mCacheCredBD
        iniBD = mCacheIniBD: finBD = mCacheFinBD: promBD = mCachePromBD: usdBD = mCacheUsdBD
        rowsCD = mCacheRowsCD: nFilasCD = mCacheNFilasCD: credCD = mCacheCredCD
        iniCD = mCacheIniCD: finCD = mCacheFinCD: promCD = mCachePromCD: usdCD = mCacheUsdCD
        rowsPEA = mCacheRowsPEA: nFilasPEA = mCacheNFilasPEA: credPEA = mCacheCredPEA
        iniPEA = mCacheIniPEA: finPEA = mCacheFinPEA: promPEA = mCachePromPEA: usdPEA = mCacheUsdPEA
        Exit Sub
    End If
    Set dao = New FondoDAO
    Call dao.extraerAnalisis1PrimerPisoSQL(fecNorm, rowsBD, nFilasBD, credBD, iniBD, finBD, promBD, usdBD)
    Call dao.extraerAnalisis2SectoresMonedasSQL(fecNorm, rowsCD, nFilasCD, credCD, iniCD, finCD, promCD, usdCD)
    Call dao.extraerAnalisis3SegundoPisoSQL(fecNorm, rowsPEA, nFilasPEA, credPEA, iniPEA, finPEA, promPEA, usdPEA)
    Set dao = Nothing
    mCacheFecAnalisis = fecNorm
    mCacheRowsBD = rowsBD: mCacheNFilasBD = nFilasBD: mCacheCredBD = credBD
    mCacheIniBD = iniBD: mCacheFinBD = finBD: mCachePromBD = promBD: mCacheUsdBD = usdBD
    mCacheRowsCD = rowsCD: mCacheNFilasCD = nFilasCD: mCacheCredCD = credCD
    mCacheIniCD = iniCD: mCacheFinCD = finCD: mCachePromCD = promCD: mCacheUsdCD = usdCD
    mCacheRowsPEA = rowsPEA: mCacheNFilasPEA = nFilasPEA: mCacheCredPEA = credPEA
    mCacheIniPEA = iniPEA: mCacheFinPEA = finPEA: mCachePromPEA = promPEA: mCacheUsdPEA = usdPEA
End Sub

' Devuelve el origen de datos del catálogo de Usuarios para los subformularios de Presentación
Public Function obtenerOrigenUsuariosBO() As String
    Dim dao As FondoDAO
    On Error Resume Next
    Set dao = New FondoDAO
    obtenerOrigenUsuariosBO = dao.obtenerOrigenUsuariosDAO()
    Set dao = Nothing
    On Error GoTo 0
End Function

' Devuelve el origen de datos del catálogo de Instrumentos para los subformularios de Presentación
Public Function obtenerOrigenInstrumentosBO() As String
    Dim dao As FondoDAO
    On Error Resume Next
    Set dao = New FondoDAO
    obtenerOrigenInstrumentosBO = dao.obtenerOrigenInstrumentosDAO()
    Set dao = Nothing
    On Error GoTo 0
End Function

' Valida exclusivamente el Numero de Usuario (usado al presionar Enter en txtUsuario aunque no haya escrito contrasena)
Public Function validarSoloUsuarioBO(ByVal cveCapturada As String, ByRef msgErrorUsuario As String) As Boolean
    Dim dao As FondoDAO
    Dim uPorClave As UsuarioDTO
    On Error GoTo ErrValU
    validarSoloUsuarioBO = False
    msgErrorUsuario = ""
    cveCapturada = Trim(cveCapturada)
    
    If cveCapturada = "" Then
        msgErrorUsuario = "Falta el Número de Empleado:" & vbCrLf & _
                          "- El campo 'Número de empleado' está vacío." & vbCrLf & vbCrLf & _
                          "Favor de ingresar un número de usuario de 4 dígitos."
        Exit Function
    End If
    If Utilities.tieneLetras(cveCapturada) Then
        msgErrorUsuario = "Formato incorrecto en Número de Empleado:" & vbCrLf & _
                          "- El valor capturado ('" & cveCapturada & "') contiene letras o caracteres especiales no válidos." & vbCrLf & _
                          "- El número de usuario debe aceptar únicamente números y ser de 4 dígitos."
        Exit Function
    End If
    If Len(cveCapturada) <> 4 Then
        msgErrorUsuario = "Extensión incorrecta en Número de Empleado:" & vbCrLf & _
                          "- El valor capturado ('" & cveCapturada & "') tiene " & Len(cveCapturada) & " dígitos." & vbCrLf & _
                          "- El número de usuario debe ser de 4 dígitos."
        Exit Function
    End If
    
    Set dao = New FondoDAO
    Set uPorClave = dao.selectUsuarioPorClave(cveCapturada)
    Set dao = Nothing
    
    If uPorClave Is Nothing Then
        msgErrorUsuario = "Número de Empleado inexistente:" & vbCrLf & _
                          "- El usuario '" & cveCapturada & "' NO se encuentra registrado en el catálogo USUARIOS." & vbCrLf & _
                          "- Verifique que el número de usuario de 4 dígitos sea correcto."
        Exit Function
    End If
    If UCase(Trim(uPorClave.getCveEstatus())) <> CLAVE_ESTATU_ACTIVO And UCase(Trim(uPorClave.getCveEstatus())) <> "A" Then
        msgErrorUsuario = "Cuenta Inactiva:" & vbCrLf & _
                          "- El empleado " & cveCapturada & " (" & uPorClave.getNombre() & " " & uPorClave.getAPaterno() & ") tiene estatus INACTIVO." & vbCrLf & _
                          "- Solicite al Administrador la reactivación de la cuenta."
        Exit Function
    End If
    validarSoloUsuarioBO = True
    Exit Function
ErrValU:
    Set dao = Nothing
    msgErrorUsuario = "Error al validar el usuario: " & Err.Description
    validarSoloUsuarioBO = False
End Function

' Validacion detallada de credenciales en Form_Acceso:
' SIEMPRE valida primero que el Usuario sea correcto (incluso si el campo Contrasena esta vacio)
Public Function validarAccesoDetalladoBO(ByVal cveCapturada As String, ByVal pwdCapturada As String, ByRef msgErrorDetallado As String) As UsuarioDTO
    Dim dao As FondoDAO
    Dim uPorClave As UsuarioDTO
    Dim pwdCifrada As String
    Dim descRol As String
    Dim obsFormato As String
    Dim msgPol As String
    
    On Error GoTo ErrLogin
    Set validarAccesoDetalladoBO = Nothing
    msgErrorDetallado = ""
    cveCapturada = Trim(cveCapturada)
    pwdCapturada = Trim(pwdCapturada)
    
    ' 1. Verificar si ambos campos estan vacios o falta el usuario
    If cveCapturada = "" And pwdCapturada = "" Then
        msgErrorDetallado = "Credenciales vacías:" & vbCrLf & _
                            "- No capturó el Número de Empleado (4 dígitos)." & vbCrLf & _
                            "- No capturó la Contraseña (8 caracteres)." & vbCrLf & vbCrLf & _
                            "Favor de ingresar ambos datos para iniciar sesión."
        Exit Function
    End If
    If cveCapturada = "" Then
        msgErrorDetallado = "Falta el Número de Empleado:" & vbCrLf & _
                            "- El campo 'Número de empleado' está vacío." & vbCrLf & vbCrLf & _
                            "Favor de ingresar un número de usuario de 4 dígitos."
        Exit Function
    End If
    
    ' 2. PRIMERO validar formato, extension de 4 digitos y existencia del Usuario en BD (aunque no haya puesto contrasena)
    If Utilities.tieneLetras(cveCapturada) Then
        msgErrorDetallado = "Formato incorrecto en Número de Empleado:" & vbCrLf & _
                            "- El valor capturado ('" & cveCapturada & "') contiene letras o caracteres especiales no válidos." & vbCrLf & _
                            "- El número de usuario debe aceptar únicamente números y ser de 4 dígitos."
        Exit Function
    End If
    If Len(cveCapturada) <> 4 Then
        msgErrorDetallado = "Extensión incorrecta en Número de Empleado:" & vbCrLf & _
                            "- El valor capturado ('" & cveCapturada & "') tiene " & Len(cveCapturada) & " dígitos." & vbCrLf & _
                            "- El número de usuario debe ser de 4 dígitos."
        Exit Function
    End If
    
    Set dao = New FondoDAO
    Set uPorClave = dao.selectUsuarioPorClave(cveCapturada)
    Set dao = Nothing
    
    If uPorClave Is Nothing Then
        msgErrorDetallado = "Número de Empleado inexistente:" & vbCrLf & _
                            "- El empleado '" & cveCapturada & "' NO se encuentra registrado en el catálogo USUARIOS." & vbCrLf & _
                            "- Verifique que el número de usuario de 4 dígitos sea correcto."
        Exit Function
    End If
    
    If UCase(Trim(uPorClave.getCvePerfil())) = PERFIL_ADMINISTRADOR Then
        descRol = "ADMINISTRADOR"
    Else
        descRol = "USUARIO OPERATIVO"
    End If
    
    If UCase(Trim(uPorClave.getCveEstatus())) <> CLAVE_ESTATU_ACTIVO And UCase(Trim(uPorClave.getCveEstatus())) <> "A" Then
        msgErrorDetallado = "Cuenta Inactiva (" & descRol & "):" & vbCrLf & _
                            "- El empleado " & cveCapturada & " (" & uPorClave.getNombre() & " " & uPorClave.getAPaterno() & ") tiene estatus INACTIVO." & vbCrLf & _
                            "- Solicite al Administrador la reactivación de la cuenta."
        Exit Function
    End If
    
    ' 3. Una vez confirmado que el usuario es valido y existe, validar el campo Contrasena
    If pwdCapturada = "" Then
        msgErrorDetallado = "Falta la Contraseña:" & vbCrLf & _
                            "- El usuario '" & cveCapturada & "' (" & uPorClave.getNombre() & " " & uPorClave.getAPaterno() & ") es válido, pero dejó vacío el campo 'Contraseña'." & vbCrLf & vbCrLf & _
                            "Favor de ingresar su contraseña de exactamente 8 caracteres."
        Exit Function
    End If
    If Len(pwdCapturada) <> 8 Then
        msgErrorDetallado = "Extensión de Contraseña no permitida:" & vbCrLf & _
                            "- La contraseña capturada tiene " & Len(pwdCapturada) & " caracteres." & vbCrLf & _
                            "- Las contraseñas deben ser necesariamente de una extensión exacta de 8 caracteres."
        Exit Function
    End If
    
    If UCase(Trim(uPorClave.getCvePerfil())) = PERFIL_ADMINISTRADOR Then
        descRol = "ADMINISTRADOR"
    Else
        descRol = "USUARIO OPERATIVO"
    End If
    
    ' 4. Verificar si la contrasena coincide con la credencial encriptada del empleado
    Dim pwdPlanaBD As String, coincidePwd As Boolean
    pwdCifrada = Utilities.cifrarContrasena(pwdCapturada, PASSWORD_PREDETERMINADA)
    pwdPlanaBD = Utilities.descifrarContrasena(uPorClave.getPassword(), PASSWORD_PREDETERMINADA)
    coincidePwd = False
    If StrComp(Trim(uPorClave.getPassword()), Trim(pwdCifrada), vbTextCompare) = 0 Then coincidePwd = True
    If StrComp(Trim(pwdPlanaBD), Trim(pwdCapturada), vbTextCompare) = 0 Then coincidePwd = True
    If Not coincidePwd Then
        obsFormato = ""
        If Not validaPoliticaContrasenaBO(pwdCapturada, msgPol) Then
            obsFormato = vbCrLf & "- Observación de formato: " & msgPol
        End If
        msgErrorDetallado = "Contraseña incorrecta para el perfil " & descRol & ":" & vbCrLf & _
                            "- Empleado identificado: " & cveCapturada & " (" & uPorClave.getNombre() & " " & uPorClave.getAPaterno() & ")" & vbCrLf & _
                            "- La contraseña capturada NO coincide con la credencial encriptada registrada." & obsFormato & vbCrLf & vbCrLf & _
                            "Verifique mayúsculas, minúsculas y números e intente de nuevo."
        Exit Function
    End If
    
    ' 5. Verificar estatus ACTIVO de la cuenta
    If UCase(Trim(uPorClave.getCveEstatus())) <> CLAVE_ESTATU_ACTIVO And UCase(Trim(uPorClave.getCveEstatus())) <> "A" Then
        msgErrorDetallado = "Cuenta Inactiva (" & descRol & "):" & vbCrLf & _
                            "- El empleado " & cveCapturada & " (" & uPorClave.getNombre() & " " & uPorClave.getAPaterno() & ") tiene estatus INACTIVO." & vbCrLf & _
                            "- Solicite al Administrador la reactivación de la cuenta."
        Exit Function
    End If
    
    Set validarAccesoDetalladoBO = uPorClave
    Exit Function
ErrLogin:
    Set dao = Nothing
    msgErrorDetallado = "Error al validar credenciales: " & Err.Description
    Set validarAccesoDetalladoBO = Nothing
End Function

' Regla de Negocio: Politica institucional de seguridad de contrasenas
Public Function validaPoliticaContrasenaBO(ByVal pwd As String, ByRef msgError As String) As Boolean
    Dim tieneMayus As Boolean, tieneMinus As Boolean, tieneNum As Boolean
    Dim i As Long, ch As String
    
    tieneMayus = False: tieneMinus = False: tieneNum = False
    If Len(pwd) <> 8 Then
        msgError = "La contraseña debe tener necesariamente una extensión exacta de 8 caracteres (capturó " & Len(pwd) & " caracteres)."
        validaPoliticaContrasenaBO = False
        Exit Function
    End If
    For i = 1 To Len(pwd)
        ch = Mid(pwd, i, 1)
        If ch >= "A" And ch <= "Z" Then tieneMayus = True
        If ch >= "a" And ch <= "z" Then tieneMinus = True
        If ch >= "0" And ch <= "9" Then tieneNum = True
    Next i
    If Not tieneMayus Then
        msgError = "La contraseña debe incluir al menos una letra mayúscula (A-Z)."
        validaPoliticaContrasenaBO = False
        Exit Function
    End If
    If Not tieneMinus Then
        msgError = "La contraseña debe incluir al menos una letra minúscula (a-z)."
        validaPoliticaContrasenaBO = False
        Exit Function
    End If
    If Not tieneNum Then
        msgError = "La contraseña debe incluir al menos un número (0-9)."
        validaPoliticaContrasenaBO = False
        Exit Function
    End If
    validaPoliticaContrasenaBO = True
End Function

' Regla de Negocio: Cambio de contrasena por autoservicio de empleado (4 campos)
' =============================================================================
' NOTIFICACION POR CORREO DE CAMBIO DE CONTRASENA
' Envia un acuse en HTML al correo institucional registrado del usuario.
' Intenta primero con Outlook (instalado en el equipo del area) y, si no esta
' disponible, con CDO sobre el relay SMTP corporativo.
' El envio NUNCA bloquea ni revierte el cambio de contrasena: si falla, se
' registra en gUltimoErrorCorreo y la operacion se da por buena igualmente.
' =============================================================================
Private Function CuerpoHtmlCambioBO(ByVal nombreCompleto As String, ByVal cveUsuario As String, _
                                    ByVal tipoCambio As String, ByVal quienAplico As String) As String
    Dim h As String
    h = "<html><body style=""margin:0;padding:0;background:#F1F3F6;font-family:Segoe UI,Arial,sans-serif;"">"
    h = h & "<table width=""100%"" cellpadding=""0"" cellspacing=""0""><tr><td align=""center"" style=""padding:24px 12px;"">"
    h = h & "<table width=""620"" cellpadding=""0"" cellspacing=""0"" style=""background:#FFFFFF;border:1px solid #D6DFEC;"">"
    h = h & "<tr><td style=""background:#D6DFEC;padding:18px 24px;"">"
    h = h & "<div style=""font-family:Cambria,Georgia,serif;font-size:18px;font-weight:bold;color:#1F4E78;"">Banco Nacional de Comercio Exterior, S.N.C.</div>"
    h = h & "<div style=""font-family:Cambria,Georgia,serif;font-size:13px;color:#004E4C;padding-top:4px;"">M" & ChrW(243) & "dulo de Administraci" & ChrW(243) & "n Integral de Riesgos</div>"
    h = h & "</td></tr>"
    h = h & "<tr><td style=""padding:24px;"">"
    h = h & "<div style=""font-size:15px;font-weight:bold;color:#004E4C;padding-bottom:10px;"">Confirmaci" & ChrW(243) & "n de cambio de contrase" & ChrW(241) & "a</div>"
    h = h & "<div style=""font-size:13px;color:#333333;line-height:1.5;"">Estimado(a) <b>" & nombreCompleto & "</b>,<br><br>"
    h = h & "Le confirmamos que la contrase" & ChrW(241) & "a de su cuenta en el M" & ChrW(243) & "dulo de Administraci" & ChrW(243) & "n Integral de Riesgos fue modificada correctamente.</div>"
    h = h & "<table cellpadding=""7"" cellspacing=""0"" style=""margin-top:16px;border-collapse:collapse;font-size:12.5px;width:100%;"">"
    h = h & "<tr><td style=""border:1px solid #D6DFEC;background:#F7F9FC;width:38%;color:#1F4E78;""><b>N" & ChrW(250) & "mero de empleado</b></td><td style=""border:1px solid #D6DFEC;"">" & cveUsuario & "</td></tr>"
    h = h & "<tr><td style=""border:1px solid #D6DFEC;background:#F7F9FC;color:#1F4E78;""><b>Tipo de movimiento</b></td><td style=""border:1px solid #D6DFEC;"">" & tipoCambio & "</td></tr>"
    h = h & "<tr><td style=""border:1px solid #D6DFEC;background:#F7F9FC;color:#1F4E78;""><b>Aplicado por</b></td><td style=""border:1px solid #D6DFEC;"">" & quienAplico & "</td></tr>"
    h = h & "<tr><td style=""border:1px solid #D6DFEC;background:#F7F9FC;color:#1F4E78;""><b>Fecha y hora</b></td><td style=""border:1px solid #D6DFEC;"">" & Format(Now, "dd/mm/yyyy hh:nn:ss") & "</td></tr>"
    h = h & "</table>"
    h = h & "<div style=""margin-top:18px;padding:12px;background:#FFF7E6;border-left:4px solid #D9822B;font-size:12.5px;color:#5C4317;"">"
    h = h & "<b>Si usted no solicit" & ChrW(243) & " este cambio</b>, rep" & ChrW(243) & "rtelo de inmediato a la Direcci" & ChrW(243) & "n de Administraci" & ChrW(243) & "n Integral de Riesgos.</div>"
    h = h & "<div style=""margin-top:18px;font-size:11px;color:#8A8A8A;border-top:1px solid #E6EAF0;padding-top:12px;"">"
    h = h & "Mensaje autom" & ChrW(225) & "tico. Por seguridad, este correo no incluye la contrase" & ChrW(241) & "a. No responda a esta direcci" & ChrW(243) & "n.</div>"
    h = h & "</td></tr></table></td></tr></table></body></html>"
    CuerpoHtmlCambioBO = h
End Function

Public Function notificarCambioContrasenaBO(ByVal cveUsuario As String, ByVal tipoCambio As String, _
                                            ByVal quienAplico As String) As Boolean
    Dim dao As FondoDAO, u As UsuarioDTO
    Dim correo As String, nombreC As String, asunto As String, cuerpo As String
    Dim olApp As Object, olMail As Object, cdo As Object

    On Error GoTo ErrMail
    notificarCambioContrasenaBO = False
    gUltimoErrorCorreo = ""

    Set dao = New FondoDAO
    Set u = dao.selectUsuarioPorClave(Trim(cveUsuario))
    Set dao = Nothing
    If u Is Nothing Then
        gUltimoErrorCorreo = "No se encontro el usuario " & cveUsuario
        Exit Function
    End If
    correo = Trim(u.getEmail())
    nombreC = Trim(u.getNombre() & " " & u.getAPaterno() & " " & u.getAMaterno())
    If correo = "" Or InStr(correo, "@") = 0 Then
        gUltimoErrorCorreo = "El usuario " & cveUsuario & " no tiene correo registrado."
        Exit Function
    End If

    asunto = "Hidrocarburos | Confirmacion de cambio de contrasena - Empleado " & cveUsuario
    cuerpo = CuerpoHtmlCambioBO(nombreC, cveUsuario, tipoCambio, quienAplico)

    ' --- Via 1: Outlook ---
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    If Not olApp Is Nothing Then
        Set olMail = olApp.CreateItem(0)
        olMail.To = correo
        olMail.Subject = asunto
        olMail.HTMLBody = cuerpo
        olMail.Send
        If Err.Number = 0 Then
            Set olMail = Nothing: Set olApp = Nothing
            notificarCambioContrasenaBO = True
            Exit Function
        End If
        gUltimoErrorCorreo = "Outlook: " & Err.Description
        Err.Clear
    End If
    Set olMail = Nothing: Set olApp = Nothing

    ' --- Via 2: CDO sobre el relay SMTP corporativo ---
    ' Ajustar servidor y puerto a los del area de sistemas si fuera necesario.
    Set cdo = CreateObject("CDO.Message")
    If Not cdo Is Nothing Then
        cdo.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        cdo.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.energia-upstream.org"
        cdo.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        cdo.Configuration.Fields.Update
        cdo.From = "notificaciones.riesgos@energia-upstream.org"
        cdo.To = correo
        cdo.Subject = asunto
        cdo.HTMLBody = cuerpo
        cdo.Send
        If Err.Number = 0 Then
            Set cdo = Nothing
            notificarCambioContrasenaBO = True
            Exit Function
        End If
        gUltimoErrorCorreo = gUltimoErrorCorreo & " | CDO: " & Err.Description
        Err.Clear
    End If
    Set cdo = Nothing
    Exit Function

ErrMail:
    On Error Resume Next
    gUltimoErrorCorreo = Err.Description
    Set olMail = Nothing: Set olApp = Nothing: Set cdo = Nothing
    notificarCambioContrasenaBO = False
End Function

Public Function cambiarContrasenaOperativoBO(ByVal cveUsuario As String, ByVal pwdActual As String, ByVal pwdNueva As String, ByVal pwdConfirmar As String, ByRef msgResultado As String) As Boolean
    Dim dao As FondoDAO
    Dim uActual As UsuarioDTO
    
    On Error GoTo ErrH
    cambiarContrasenaOperativoBO = False
    
    If Trim(cveUsuario) = "" Then
        msgResultado = "Por favor capture el número de empleado de 4 dígitos."
        Exit Function
    End If
    If Utilities.tieneLetras(Trim(cveUsuario)) Or Len(Trim(cveUsuario)) <> 4 Then
        msgResultado = "El número de usuario debe ser estrictamente numérico y de 4 dígitos."
        Exit Function
    End If
    If Trim(pwdActual) = "" Then
        msgResultado = "Por favor capture la contraseña actual del empleado."
        Exit Function
    End If
    If Len(Trim(pwdActual)) <> 8 Then
        msgResultado = "La contraseña actual debe tener necesariamente una extensión exacta de 8 caracteres."
        Exit Function
    End If
    If Trim(pwdNueva) = "" Then
        msgResultado = "Por favor capture la nueva contraseña."
        Exit Function
    End If
    If Trim(pwdConfirmar) = "" Then
        msgResultado = "Por favor confirme la nueva contraseña."
        Exit Function
    End If
    If pwdNueva <> pwdConfirmar Then
        msgResultado = "La confirmación de la contraseña no coincide con la nueva contraseña capturada." & vbCrLf & "Favor de verificar ambos campos."
        Exit Function
    End If
    If pwdNueva = pwdActual Then
        msgResultado = "La nueva contraseña no puede ser igual a la contraseña actual." & vbCrLf & "Favor de capturar una contraseña distinta."
        Exit Function
    End If
    If Not validaPoliticaContrasenaBO(pwdNueva, msgResultado) Then
        Exit Function
    End If
    
    Set dao = New FondoDAO
    Set uActual = dao.selectUsuarioPorClave(Trim(cveUsuario))
    If uActual Is Nothing Then
        msgResultado = "El empleado " & cveUsuario & " no existe en el sistema."
        Set dao = Nothing
        Exit Function
    End If
    If StrComp(Trim(uActual.getPassword()), Trim(Utilities.cifrarContrasena(pwdActual, PASSWORD_PREDETERMINADA)), vbTextCompare) <> 0 Then
        msgResultado = "La contraseña actual capturada no coincide con la registrada en el sistema."
        Set dao = Nothing
        Exit Function
    End If
    
    If dao.cambiarContrasena(Trim(cveUsuario), pwdNueva) Then
        msgResultado = "Contraseña actualizada exitosamente para el empleado " & cveUsuario & "." & vbCrLf & vbCrLf & "A partir de este momento deberá ingresar con su nueva credencial."
        cambiarContrasenaOperativoBO = True
    ' Acuse por correo. No condiciona el resultado: si falla, el cambio ya quedo hecho.
    Call notificarCambioContrasenaBO(cveUsuario, "Cambio solicitado por el propio usuario", "El titular de la cuenta")
    Else
        msgResultado = "No fue posible actualizar la contraseña en la base de datos."
    End If
    Set dao = Nothing
    Exit Function
ErrH:
    Set dao = Nothing
    msgResultado = "Error en capa de negocio: " & Err.Description
    cambiarContrasenaOperativoBO = False
End Function

' Regla de Negocio: Restablecimiento de contrasena por Administrador (3 campos)
Public Function restablecerContrasenaAdminBO(ByVal cveUsuario As String, ByVal pwdNueva As String, ByVal pwdConfirmar As String, ByRef msgResultado As String) As Boolean
    Dim dao As FondoDAO
    Dim uExistente As UsuarioDTO
    
    On Error GoTo ErrH
    restablecerContrasenaAdminBO = False
    
    If Trim(cveUsuario) = "" Then
        msgResultado = "Capture el número de empleado de 4 dígitos."
        Exit Function
    End If
    If Utilities.tieneLetras(Trim(cveUsuario)) Or Len(Trim(cveUsuario)) <> 4 Then
        msgResultado = "El número de usuario debe ser estrictamente numérico y de 4 dígitos."
        Exit Function
    End If
    If Trim(pwdNueva) = "" Then
        msgResultado = "Capture la nueva contraseña."
        Exit Function
    End If
    If Trim(pwdConfirmar) = "" Then
        msgResultado = "Confirme la nueva contraseña."
        Exit Function
    End If
    If pwdNueva <> pwdConfirmar Then
        msgResultado = "La confirmación no coincide con la nueva contraseña."
        Exit Function
    End If
    If Not validaPoliticaContrasenaBO(pwdNueva, msgResultado) Then
        Exit Function
    End If
    
    Set dao = New FondoDAO
    Set uExistente = dao.selectUsuarioPorClave(Trim(cveUsuario))
    If uExistente Is Nothing Then
        msgResultado = "El empleado " & cveUsuario & " no existe en el sistema."
        Set dao = Nothing
        Exit Function
    End If
    
    If dao.cambiarContrasena(Trim(cveUsuario), pwdNueva) Then
        msgResultado = "Contraseña restablecida exitosamente para el empleado " & cveUsuario & " (Modo Administrador)."
        restablecerContrasenaAdminBO = True
    ' Acuse por correo. No condiciona el resultado: si falla, el cambio ya quedo hecho.
    Call notificarCambioContrasenaBO(cveUsuario, "Restablecimiento administrativo", "Un usuario con perfil Administrador")
    Else
        msgResultado = "No fue posible restablecer la contraseña en la base de datos."
    End If
    Set dao = Nothing
    Exit Function
ErrH:
    Set dao = Nothing
    msgResultado = "Error al restablecer contraseña: " & Err.Description
    restablecerContrasenaAdminBO = False
End Function

' CRUD Usuarios en Capa de Negocio (BO)
Public Function consultarUsuarioBO(ByVal claveUsuario As String) As UsuarioDTO
    Dim dao As FondoDAO
    On Error GoTo ErrH
    Set dao = New FondoDAO
    Set consultarUsuarioBO = dao.selectUsuarioPorClave(Trim(claveUsuario))
    Set dao = Nothing
    Exit Function
ErrH:
    Set dao = Nothing
    Set consultarUsuarioBO = Nothing
End Function

Public Function registrarUsuarioBO(ByVal u As UsuarioDTO, ByRef msgResultado As String) As Boolean
    Dim dao As FondoDAO
    Dim uExistente As UsuarioDTO
    On Error GoTo ErrH
    registrarUsuarioBO = False
    
    If u Is Nothing Then
        msgResultado = "Datos de usuario inválidos."
        Exit Function
    End If
    If Trim(u.getCveUsuario()) = "" Then
        msgResultado = "Capture el número de empleado de 4 dígitos."
        Exit Function
    End If
    If Utilities.tieneLetras(Trim(u.getCveUsuario())) Or Len(Trim(u.getCveUsuario())) <> 4 Then
        msgResultado = "El número de usuario debe ser estrictamente numérico y de 4 dígitos."
        Exit Function
    End If
    If Trim(u.getNombre()) = "" Then
        msgResultado = "Capture el nombre del empleado."
        Exit Function
    End If
    If Trim(u.getPassword()) = "" Then
        msgResultado = "Capture la contraseña del nuevo usuario (extensión obligatoria de 8 caracteres)."
        Exit Function
    End If
    If Not validaPoliticaContrasenaBO(Trim(u.getPassword()), msgResultado) Then
        Exit Function
    End If
    If Trim(u.getCvePerfil()) = "" Then
        u.setCvePerfil PERFIL_EMPLEADO
    End If
    If Trim(u.getCveEstatus()) = "" Then
        u.setCveEstatus CLAVE_ESTATU_ACTIVO
    End If
    
    Set dao = New FondoDAO
    Set uExistente = dao.selectUsuarioPorClave(u.getCveUsuario())
    If Not uExistente Is Nothing Then
        msgResultado = "El número de empleado " & u.getCveUsuario() & " ya se encuentra registrado."
        Set dao = Nothing
        Exit Function
    End If
    
    If dao.insertUsuario(u) Then
        msgResultado = "Usuario registrado exitosamente."
        registrarUsuarioBO = True
    Else
        msgResultado = "No fue posible registrar el usuario en la base de datos."
    End If
    Set dao = Nothing
    Exit Function
ErrH:
    Set dao = Nothing
    msgResultado = "Error al registrar usuario: " & Err.Description
    registrarUsuarioBO = False
End Function

Public Function actualizarUsuarioBO(ByVal u As UsuarioDTO, ByRef msgResultado As String) As Boolean
    Dim dao As FondoDAO
    On Error GoTo ErrH
    actualizarUsuarioBO = False
    If u Is Nothing Or Trim(u.getCveUsuario()) = "" Then
        msgResultado = "Capture el número de empleado de 4 dígitos a actualizar."
        Exit Function
    End If
    If Utilities.tieneLetras(Trim(u.getCveUsuario())) Or Len(Trim(u.getCveUsuario())) <> 4 Then
        msgResultado = "El número de usuario debe ser estrictamente numérico y de 4 dígitos."
        Exit Function
    End If
    Set dao = New FondoDAO
    If dao.updateUsuario(u) Then
        msgResultado = "Usuario actualizado exitosamente."
        actualizarUsuarioBO = True
    Else
        msgResultado = "No fue posible actualizar los datos del usuario."
    End If
    Set dao = Nothing
    Exit Function
ErrH:
    Set dao = Nothing
    msgResultado = "Error al actualizar usuario: " & Err.Description
    actualizarUsuarioBO = False
End Function

Public Function inactivarUsuarioBO(ByVal claveUsuario As String, ByRef msgResultado As String) As Boolean
    Dim dao As FondoDAO
    On Error GoTo ErrH
    inactivarUsuarioBO = False
    If Trim(claveUsuario) = "" Then
        msgResultado = "Capture el número de empleado de 4 dígitos."
        Exit Function
    End If
    If Utilities.tieneLetras(Trim(claveUsuario)) Or Len(Trim(claveUsuario)) <> 4 Then
        msgResultado = "El número de usuario debe ser estrictamente numérico y de 4 dígitos."
        Exit Function
    End If
    Set dao = New FondoDAO
    If dao.deleteLogicoUsuario(Trim(claveUsuario)) Then
        msgResultado = "Usuario marcado como INACTIVO exitosamente."
        inactivarUsuarioBO = True
    Else
        msgResultado = "No fue posible dar de baja al usuario."
    End If
    Set dao = Nothing
    Exit Function
ErrH:
    Set dao = Nothing
    msgResultado = "Error al dar de baja al usuario: " & Err.Description
    inactivarUsuarioBO = False
End Function

' CRUD Instrumentos en Capa de Negocio (BO)
Public Function consultarInstrumentoBO(ByVal idInstrumento As String) As InstrumentoDTO
    Dim dao As FondoDAO
    On Error GoTo ErrH
    Set dao = New FondoDAO
    Set consultarInstrumentoBO = dao.selectInstrumentoPorId(Trim(idInstrumento))
    Set dao = Nothing
    Exit Function
ErrH:
    Set dao = Nothing
    Set consultarInstrumentoBO = Nothing
End Function

Public Function registrarInstrumentoBO(ByVal inst As InstrumentoDTO, ByRef msgResultado As String) As Boolean
    Dim dao As FondoDAO
    On Error GoTo ErrH
    registrarInstrumentoBO = False
    If inst Is Nothing Or Trim(inst.getIdInstrumento()) = "" Then
        msgResultado = "Capture el ID del instrumento."
        Exit Function
    End If
    Set dao = New FondoDAO
    If dao.insertInstrumento(inst) Then
        msgResultado = "Instrumento registrado exitosamente."
        registrarInstrumentoBO = True
    Else
        msgResultado = "No fue posible registrar el instrumento."
    End If
    Set dao = Nothing
    Exit Function
ErrH:
    Set dao = Nothing
    msgResultado = "Error al registrar instrumento: " & Err.Description
    registrarInstrumentoBO = False
End Function

Public Function actualizarInstrumentoBO(ByVal inst As InstrumentoDTO, ByRef msgResultado As String) As Boolean
    Dim dao As FondoDAO
    Dim instPrevio As InstrumentoDTO
    On Error GoTo ErrH
    actualizarInstrumentoBO = False
    If inst Is Nothing Or Trim(inst.getIdInstrumento()) = "" Then
        msgResultado = "Capture el ID del instrumento a actualizar."
        Exit Function
    End If
    Set dao = New FondoDAO
    Set instPrevio = dao.selectInstrumentoPorId(Trim(inst.getIdInstrumento()))
    If instPrevio Is Nothing Then
        msgResultado = "El instrumento '" & Trim(inst.getIdInstrumento()) & "' no existe en el catálogo INSTRUMENTOS."
        Set dao = Nothing
        Exit Function
    End If
    If inst.getPlazo() <= 0 Then
        inst.setPlazo instPrevio.getPlazo()
    End If
    Set instPrevio = Nothing
    If dao.updateInstrumento(inst) Then
        msgResultado = "Instrumento actualizado exitosamente."
        actualizarInstrumentoBO = True
    Else
        msgResultado = "No fue posible actualizar el instrumento."
    End If
    Set dao = Nothing
    Exit Function
ErrH:
    Set dao = Nothing
    msgResultado = "Error al actualizar instrumento: " & Err.Description
    actualizarInstrumentoBO = False
End Function

Public Function obtenerRutaDescargasPorDefectoBO() As String
    Dim rutaDescargas As String
    On Error Resume Next
    rutaDescargas = Environ("USERPROFILE") & "\Downloads"
    If rutaDescargas = "" Or Dir(rutaDescargas, vbDirectory) = "" Then
        rutaDescargas = Environ("HOMEDRIVE") & Environ("HOMEPATH") & "\Downloads"
    End If
    If rutaDescargas = "" Or Dir(rutaDescargas, vbDirectory) = "" Then
        rutaDescargas = CurrentProject.Path
    End If
    obtenerRutaDescargasPorDefectoBO = rutaDescargas
    On Error GoTo 0
End Function

' ==============================================================================
' VALIDACION EXHAUSTIVA DATO POR DATO, ESTRUCTURA, FECHAS Y SOBRESCRITURA (BO)
' ==============================================================================
Private Sub AgregarErrorInsumo(ByRef detalleErrores As String, ByVal lineaErr As String)
    If Len(detalleErrores) + Len(lineaErr) <= 550 Then
        detalleErrores = detalleErrores & lineaErr & vbCrLf
    End If
End Sub

' Verifica que la cadena tenga formato dd/mm/aaaa valido (independiente de configuracion regional)
Public Function esFechaFormatoCorrectoBO(ByVal strFecha As String) As Boolean
    Dim partes() As String
    Dim d As Integer, m As Integer, a As Integer
    Dim dt As Date
    On Error GoTo ErrF
    esFechaFormatoCorrectoBO = False
    strFecha = Trim(strFecha)
    If Len(strFecha) < 6 Or Len(strFecha) > 10 Then Exit Function
    If InStr(strFecha, "/") = 0 Then Exit Function
    partes = Split(strFecha, "/")
    If UBound(partes) <> 2 Then Exit Function
    If Not (IsNumeric(partes(0)) And IsNumeric(partes(1)) And IsNumeric(partes(2))) Then Exit Function
    d = CInt(partes(0)): m = CInt(partes(1)): a = CInt(partes(2))
    If Len(Trim(partes(2))) = 2 And a >= 0 And a <= 99 Then a = 2000 + a
    If a < 1900 Or a > 9999 Then Exit Function
    If m < 1 Or m > 12 Then Exit Function
    If d < 1 Or d > 31 Then Exit Function
    dt = DateSerial(a, m, d)
    If Day(dt) <> d Or Month(dt) <> m Or Year(dt) <> a Then Exit Function
    esFechaFormatoCorrectoBO = True
    Exit Function
ErrF:
    esFechaFormatoCorrectoBO = False
End Function

' Convierte de forma segura una cadena dd/mm/aaaa a Date usando DateSerial (sin depender de CDate regional)
Public Function convertirFechaBO(ByVal strFecha As String) As Date
    Dim partes() As String
    Dim d As Integer, m As Integer, a As Integer
    On Error GoTo ErrConv
    partes = Split(Trim(strFecha), "/")
    d = CInt(partes(0)): m = CInt(partes(1)): a = CInt(partes(2))
    If Len(Trim(partes(2))) = 2 And a >= 0 And a <= 99 Then a = 2000 + a
    convertirFechaBO = DateSerial(a, m, d)
    Exit Function
ErrConv:
    convertirFechaBO = Date
End Function

' Normaliza cualquier fecha valida d/m/aaaa o dd/mm/aaaa al formato exacto dd/mm/aaaa
Public Function normalizarFechaBO(ByVal strFecha As String) As String
    On Error GoTo ErrNorm
    If Not esFechaFormatoCorrectoBO(strFecha) Then
        normalizarFechaBO = Trim(strFecha)
        Exit Function
    End If
    normalizarFechaBO = Format(convertirFechaBO(strFecha), "dd/mm/yyyy")
    Exit Function
ErrNorm:
    normalizarFechaBO = Trim(strFecha)
End Function

' Evalua si la fecha ingresada (dd/mm/aaaa) es futura (posterior al dia de la consulta actual Date)
Public Function esFechaFuturaBO(ByVal strFecha As String) As Boolean
    On Error GoTo ErrFut
    esFechaFuturaBO = False
    If Not esFechaFormatoCorrectoBO(strFecha) Then Exit Function
    Dim dtEval As Date
    dtEval = convertirFechaBO(strFecha)
    If dtEval > Date Then esFechaFuturaBO = True
    Exit Function
ErrFut:
    esFechaFuturaBO = False
End Function

Public Function extraerFechaDeTituloInsumoBO(ByVal nombreArchivo As String) As String
    Dim base As String, norm As String, partes() As String
    Dim i As Long, d As Integer, m As Integer, a As Integer
    On Error GoTo ErrExt
    extraerFechaDeTituloInsumoBO = ""
    base = Trim(nombreArchivo)
    If InStrRev(base, ".") > 0 Then base = Left(base, InStrRev(base, ".") - 1)
    If UCase(Trim(base)) = "CARTERA_IMPORTAR" Then
        extraerFechaDeTituloInsumoBO = "30/09/2026"
        Exit Function
    End If
    norm = Replace(Replace(Replace(base, "-", "_"), ".", "_"), " ", "_")
    partes = Split(norm, "_")
    If UBound(partes) >= 2 Then
        For i = 0 To UBound(partes) - 2
            If Len(partes(i)) = 2 And Len(partes(i + 1)) = 2 And Len(partes(i + 2)) = 4 Then
                If IsNumeric(partes(i)) And IsNumeric(partes(i + 1)) And IsNumeric(partes(i + 2)) Then
                    d = CInt(partes(i)): m = CInt(partes(i + 1)): a = CInt(partes(i + 2))
                    If d >= 1 And d <= 31 And m >= 1 And m <= 12 And a >= 2000 And a <= 2035 Then
                        extraerFechaDeTituloInsumoBO = Format(DateSerial(a, m, d), "dd/mm/yyyy")
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If
    Exit Function
ErrExt:
    extraerFechaDeTituloInsumoBO = ""
End Function

' Verifica en BITACORA_CARGA_INSUMOS si el archivo o fecha de corte ya fue cargado previamente
Public Function verificarInsumoYaCargadoBO(ByVal nombreArchivo As String, ByVal fechaCorte As String, ByRef infoPrevia As String) As Boolean
    Dim dao As FondoDAO
    On Error GoTo ErrV
    Set dao = New FondoDAO
    verificarInsumoYaCargadoBO = dao.existeInsumoCargadoDAO(nombreArchivo, fechaCorte, infoPrevia)
    Set dao = Nothing
    Exit Function
ErrV:
    Set dao = Nothing
    verificarInsumoYaCargadoBO = False
End Function

' Obtiene la ultima fecha de corte cargada en la base de datos RIESGOSBD.accdb mediante SQL
Public Function obtenerUltimaFechaCargadaBO() As String
    Dim dao As FondoDAO
    On Error GoTo ErrU
    Set dao = New FondoDAO
    obtenerUltimaFechaCargadaBO = dao.obtenerUltimaFechaInsumoDAO()
    Set dao = Nothing
    Exit Function
ErrU:
    Set dao = Nothing
    obtenerUltimaFechaCargadaBO = ""
End Function

' Cuenta cuantos registros existen en las 3 bases de datos de insumos para una fecha dd/mm/aaaa (usado por el Calendario)
' -----------------------------------------------------------------------------
' obtenerMapaFechasConteoBO
' Devuelve en UNA sola consulta el total de registros por fecha de corte,
' con el formato  |dd/mm/aaaa=n|dd/mm/aaaa=n|...
' Existe para el calendario: antes pintaba el mes llamando 31 veces a
' contarRegistrosFechaBO (31 consultas y 31 objetos DAO por cada cambio de mes),
' que es lo que lo hacia tardar. Ahora se consulta una vez y se busca en memoria.
' -----------------------------------------------------------------------------
Public Function obtenerMapaFechasConteoBO() As String
    Dim dao As FondoDAO
    On Error GoTo ErrMapa
    Set dao = New FondoDAO
    obtenerMapaFechasConteoBO = dao.obtenerMapaFechasConteoDAO()
    Set dao = Nothing
    Exit Function
ErrMapa:
    Set dao = Nothing
    obtenerMapaFechasConteoBO = "|"
End Function

' -----------------------------------------------------------------------------
' conteoDeFechaEnMapaBO
' Busca una fecha dentro del mapa devuelto por obtenerMapaFechasConteoBO.
' Es una busqueda en memoria (InStr), sin tocar la base de datos.
' -----------------------------------------------------------------------------
Public Function conteoDeFechaEnMapaBO(ByVal mapa As String, ByVal fecha As String) As Long
    Dim p As Long, q As Long
    On Error Resume Next
    conteoDeFechaEnMapaBO = 0
    If mapa = "" Or fecha = "" Then Exit Function
    p = InStr(1, mapa, "|" & Trim(fecha) & "=", vbTextCompare)
    If p = 0 Then Exit Function
    p = p + Len(Trim(fecha)) + 2
    q = InStr(p, mapa, "|")
    If q = 0 Then Exit Function
    conteoDeFechaEnMapaBO = CLng(Val(Mid(mapa, p, q - p)))
End Function

Public Function contarRegistrosFechaBO(ByVal fechaCorte As String) As Long
    Dim n1 As Long, n2 As Long, n3 As Long, dao As FondoDAO
    On Error Resume Next
    contarRegistrosFechaBO = 0
    If Not esFechaFormatoCorrectoBO(fechaCorte) Then Exit Function
    Call obtenerConteosTresTablasFechaBO(fechaCorte, n1, n2, n3)
    contarRegistrosFechaBO = n1 + n2 + n3
    If contarRegistrosFechaBO = 0 Then
        Set dao = New FondoDAO
        contarRegistrosFechaBO = dao.contarRegistrosTablaPorFechaDAO("CARTERA_ETIQUETADA", normalizarFechaBO(fechaCorte))
        Set dao = Nothing
    End If
    On Error GoTo 0
End Function

' Cuenta los registros cargados en una tabla de insumo especifica (Primer Piso, Tesoreria o Segundo Piso) para una fecha dada
Public Function contarRegistrosTablaFechaBO(ByVal nombreTabla As String, ByVal fechaCorte As String) As Long
    Dim n1 As Long, n2 As Long, n3 As Long, tUp As String, dao As FondoDAO
    On Error Resume Next
    contarRegistrosTablaFechaBO = 0
    If Not esFechaFormatoCorrectoBO(fechaCorte) Then Exit Function
    tUp = UCase$(Trim$(nombreTabla))
    If tUp = "BD_INSUMO_PRIMER_PISO" Or tUp = "BD_INSUMO_TESORERIA" Or tUp = "BD_INSUMO_SEGUNDO_PISO" Then
        Call obtenerConteosTresTablasFechaBO(fechaCorte, n1, n2, n3)
        If tUp = "BD_INSUMO_PRIMER_PISO" Then contarRegistrosTablaFechaBO = n1: Exit Function
        If tUp = "BD_INSUMO_TESORERIA" Then contarRegistrosTablaFechaBO = n2: Exit Function
        If tUp = "BD_INSUMO_SEGUNDO_PISO" Then contarRegistrosTablaFechaBO = n3: Exit Function
    End If
    Set dao = New FondoDAO
    contarRegistrosTablaFechaBO = dao.contarRegistrosTablaPorFechaDAO(nombreTabla, normalizarFechaBO(fechaCorte))
    Set dao = Nothing
    On Error GoTo 0
End Function

' Devuelve la lista de fechas dd/mm/aaaa que ya cuentan con informacion cargada en cualquiera de las 3 bases de datos
Public Function obtenerFechasCargadasBO() As String
    Dim dao As FondoDAO
    On Error GoTo ErrF
    Set dao = New FondoDAO
    obtenerFechasCargadasBO = dao.obtenerFechasCargadasDAO()
    Set dao = Nothing
    Exit Function
ErrF:
    Set dao = Nothing
    obtenerFechasCargadasBO = ""
End Function

' Devuelve el texto explicativo con el FORMATO UNICO OBLIGATORIO para nombrar los 3 archivos de insumo por fecha
Public Function obtenerFormatoCorrectoInsumoBO(ByVal fechaCorte As String, Optional ByVal bBD As Boolean = True, Optional ByVal bCD As Boolean = True, Optional ByVal bPEA As Boolean = True) As String
    Dim fecNorm As String, tagFec As String, msg As String
    On Error Resume Next
    If esFechaFormatoCorrectoBO(fechaCorte) Then
        fecNorm = normalizarFechaBO(fechaCorte)
    Else
        fecNorm = Format(Date, "dd/mm/yyyy")
    End If
    tagFec = Replace(fecNorm, "/", "_")
    msg = "FORMATO ÚNICO OBLIGATORIO PARA NOMBRAR Y SUBIR LOS INSUMOS:" & vbCrLf & _
          "Patrón oficial: INSUMO_<PORTAFOLIO>_DD_MM_AAAA.xlsx (53 columnas incl. ID_INSTRUMENTO)" & vbCrLf
    If bBD Or (Not bBD And Not bCD And Not bPEA) Then
        msg = msg & "  • Opción 1 (Primer Piso -> BD_INSUMO_PRIMER_PISO):" & vbCrLf & _
                    "    INSUMO_PRIMER_PISO_" & tagFec & ".xlsx" & vbCrLf
    End If
    If bCD Or (Not bBD And Not bCD And Not bPEA) Then
        msg = msg & "  • Opción 2 (Tesorería -> BD_INSUMO_TESORERIA):" & vbCrLf & _
                    "    INSUMO_TESORERIA_" & tagFec & ".xlsx" & vbCrLf
    End If
    If bPEA Or (Not bBD And Not bCD And Not bPEA) Then
        msg = msg & "  • Opción 3 (Segundo Piso -> BD_INSUMO_SEGUNDO_PISO):" & vbCrLf & _
                    "    INSUMO_SEGUNDO_PISO_" & tagFec & ".xlsx" & vbCrLf
    End If
    obtenerFormatoCorrectoInsumoBO = msg
    On Error GoTo 0
End Function

' Valida que el nombre del archivo cumpla estrictamente con el formato unico oficial INSUMO_<PORTAFOLIO>_DD_MM_AAAA.xlsx
Public Function validarFormatoUnicoNombreInsumoBO(ByVal nombreArchivo As String, ByVal fechaCorte As String, ByVal tipoEsperado As String, ByRef tipoDetectado As String, ByRef tablaDestino As String, ByRef msgErrorFormato As String) As Boolean
    Dim nomUp As String, fecNorm As String, tagFec As String, ext As String, sinExt As String
    Dim esp1 As String, esp2 As String, esp3 As String, sufijoFecha As String
    On Error GoTo ErrNom
    validarFormatoUnicoNombreInsumoBO = False
    tipoDetectado = ""
    tablaDestino = ""
    msgErrorFormato = ""
    
    If esFechaFormatoCorrectoBO(fechaCorte) Then
        fecNorm = normalizarFechaBO(fechaCorte)
    Else
        fecNorm = Format(Date, "dd/mm/yyyy")
    End If
    tagFec = Replace(fecNorm, "/", "_")
    
    nomUp = UCase(Trim(nombreArchivo))
    If InStrRev(nomUp, "\") > 0 Then nomUp = Mid(nomUp, InStrRev(nomUp, "\") + 1)
    
    If InStrRev(nomUp, ".") = 0 Then
        msgErrorFormato = "- [Formato de Nombre Inválido]: El archivo '" & nombreArchivo & "' no tiene extensión (.xlsx)."
        Exit Function
    End If
    ext = Mid(nomUp, InStrRev(nomUp, ".") + 1)
    sinExt = Left(nomUp, InStrRev(nomUp, ".") - 1)
    
    If ext <> "XLSX" And ext <> "XLS" And ext <> "XLSM" Then
        msgErrorFormato = "- [Extensión No Permitida]: El archivo '" & nombreArchivo & "' tiene extensión '." & LCase(ext) & "' (debe ser .xlsx)."
        Exit Function
    End If
    
    esp1 = "INSUMO_PRIMER_PISO_" & tagFec
    esp2 = "INSUMO_TESORERIA_" & tagFec
    esp3 = "INSUMO_SEGUNDO_PISO_" & tagFec
    
    If Left(sinExt, 19) = "INSUMO_PRIMER_PISO_" Then
        tipoDetectado = "PRIMER_PISO"
        tablaDestino = "BD_INSUMO_PRIMER_PISO"
        sufijoFecha = Mid(sinExt, 20)
    ElseIf Left(sinExt, 17) = "INSUMO_TESORERIA_" Then
        tipoDetectado = "TESORERIA"
        tablaDestino = "BD_INSUMO_TESORERIA"
        sufijoFecha = Mid(sinExt, 18)
    ElseIf Left(sinExt, 20) = "INSUMO_SEGUNDO_PISO_" Then
        tipoDetectado = "SEGUNDO_PISO"
        tablaDestino = "BD_INSUMO_SEGUNDO_PISO"
        sufijoFecha = Mid(sinExt, 21)
    Else
        msgErrorFormato = "- [Nombre de Archivo Fuera del Formato Único]: El nombre '" & nombreArchivo & "' no inicia con ninguno de los 3 prefijos oficiales ('INSUMO_PRIMER_PISO_', 'INSUMO_TESORERIA_' o 'INSUMO_SEGUNDO_PISO_')."
        Exit Function
    End If
    
    If InStr(sufijoFecha, "INCORRECTO") > 0 Or Len(sufijoFecha) <> 10 Then
        msgErrorFormato = "- [Estructura de Nombre Incorrecta]: El archivo '" & nombreArchivo & "' contiene texto adicional no permitido ('" & sufijoFecha & "').Después del prefijo del portafolio debe ir únicamente la fecha exacta '" & tagFec & ".xlsx'."
        Exit Function
    End If
    
    If sufijoFecha <> tagFec Then
        msgErrorFormato = "- [Fecha en Nombre de Archivo No Coincide]: El archivo '" & nombreArchivo & "' indica la fecha '" & Replace(sufijoFecha, "_", "/") & "', pero la fecha seleccionada en pantalla es '" & fecNorm & "' (se requiere '_" & tagFec & ".xlsx')."
        Exit Function
    End If
    
    If Trim(tipoEsperado) <> "" And UCase(Trim(tipoEsperado)) <> tipoDetectado Then
        Dim nomEsp As String
        If UCase(Trim(tipoEsperado)) = "PRIMER_PISO" Then nomEsp = esp1 & ".xlsx"
        If UCase(Trim(tipoEsperado)) = "TESORERIA" Then nomEsp = esp2 & ".xlsx"
        If UCase(Trim(tipoEsperado)) = "SEGUNDO_PISO" Then nomEsp = esp3 & ".xlsx"
        msgErrorFormato = "- [Portafolio Incorrecto en Nombre de Archivo]: El archivo '" & nombreArchivo & "' corresponde a '" & tipoDetectado & "', pero la opción marcada requiere el archivo '" & nomEsp & "'."
        Exit Function
    End If
    
    validarFormatoUnicoNombreInsumoBO = True
    Exit Function
ErrNom:
    msgErrorFormato = "- [Error en Nombre de Archivo]: No fue posible validar el nombre '" & nombreArchivo & "'."
    validarFormatoUnicoNombreInsumoBO = False
End Function

' Dada la ruta de una carpeta completa, busca automaticamente los archivos correctos de insumo para la fecha y opciones seleccionadas
Public Function buscarInsumosCorrectosEnCarpetaBO(ByVal rutaCarpeta As String, ByVal fechaCorte As String, ByVal bBD As Boolean, ByVal bCD As Boolean, ByVal bPEA As Boolean, ByRef rutasEncontradas As String, ByRef resumenHallazgos As String, ByRef msgErrorBusqueda As String) As Boolean
    Dim fso As Object, carp As String, fecNorm As String, tagFec As String
    Dim arch1 As String, arch2 As String, arch3 As String
    Dim r1 As String, r2 As String, r3 As String
    Dim faltantes As String, nReq As Integer, nHay As Integer
    On Error GoTo ErrBusq
    buscarInsumosCorrectosEnCarpetaBO = False
    rutasEncontradas = ""
    resumenHallazgos = ""
    msgErrorBusqueda = ""
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    carp = Trim(rutaCarpeta)
    If Right(carp, 1) = "\" Then carp = Left(carp, Len(carp) - 1)
    If Not fso.FolderExists(carp) Then
        msgErrorBusqueda = "- La carpeta especificada no existe: " & rutaCarpeta & vbCrLf & vbCrLf & obtenerFormatoCorrectoInsumoBO(fechaCorte, bBD, bCD, bPEA)
        Exit Function
    End If
    
    If Not esFechaFormatoCorrectoBO(fechaCorte) Then
        msgErrorBusqueda = "- La fecha capturada ('" & fechaCorte & "') no tiene formato válido dd/mm/aaaa." & vbCrLf & vbCrLf & obtenerFormatoCorrectoInsumoBO(Format(Date, "dd/mm/yyyy"), bBD, bCD, bPEA)
        Exit Function
    End If
    fecNorm = normalizarFechaBO(fechaCorte)
    tagFec = Replace(fecNorm, "/", "_")
    
    arch1 = "INSUMO_PRIMER_PISO_" & tagFec & ".xlsx"
    arch2 = "INSUMO_TESORERIA_" & tagFec & ".xlsx"
    arch3 = "INSUMO_SEGUNDO_PISO_" & tagFec & ".xlsx"
    r1 = carp & "\" & arch1
    r2 = carp & "\" & arch2
    r3 = carp & "\" & arch3
    
    ' Si el usuario aun no ha marcado ninguna de las 3 casillas al elegir una carpeta,
    ' buscar todos los documentos de esa fecha contenidos en la carpeta de manera automatica
    If Not (bBD Or bCD Or bPEA) Then
        If fso.FileExists(r1) Then bBD = True
        If fso.FileExists(r2) Then bCD = True
        If fso.FileExists(r3) Then bPEA = True
        If Not (bBD Or bCD Or bPEA) Then
            bBD = True: bCD = True: bPEA = True
        End If
    End If
    
    nReq = 0: nHay = 0: faltantes = ""
    If bBD Then
        nReq = nReq + 1
        If fso.FileExists(r1) Then
            nHay = nHay + 1
            If rutasEncontradas <> "" Then rutasEncontradas = rutasEncontradas & ";"
            rutasEncontradas = rutasEncontradas & r1
            resumenHallazgos = resumenHallazgos & "  • [Opción 1 -> BD_INSUMO_PRIMER_PISO]: " & arch1 & vbCrLf
        Else
            faltantes = faltantes & "  - Falta archivo de Opción 1 (Primer Piso): " & arch1 & vbCrLf
        End If
    End If
    If bCD Then
        nReq = nReq + 1
        If fso.FileExists(r2) Then
            nHay = nHay + 1
            If rutasEncontradas <> "" Then rutasEncontradas = rutasEncontradas & ";"
            rutasEncontradas = rutasEncontradas & r2
            resumenHallazgos = resumenHallazgos & "  • [Opción 2 -> BD_INSUMO_TESORERIA]: " & arch2 & vbCrLf
        Else
            faltantes = faltantes & "  - Falta archivo de Opción 2 (Tesorería): " & arch2 & vbCrLf
        End If
    End If
    If bPEA Then
        nReq = nReq + 1
        If fso.FileExists(r3) Then
            nHay = nHay + 1
            If rutasEncontradas <> "" Then rutasEncontradas = rutasEncontradas & ";"
            rutasEncontradas = rutasEncontradas & r3
            resumenHallazgos = resumenHallazgos & "  • [Opción 3 -> BD_INSUMO_SEGUNDO_PISO]: " & arch3 & vbCrLf
        Else
            faltantes = faltantes & "  - Falta archivo de Opción 3 (Segundo Piso): " & arch3 & vbCrLf
        End If
    End If
    
    If nReq = 0 Then
        msgErrorBusqueda = "- Debe marcar al menos una de las 3 opciones de portafolio de riesgo."
        Exit Function
    End If
    
    If nHay < nReq Then
        msgErrorBusqueda = "No se encontraron todos los insumos requeridos en la carpeta:" & vbCrLf & _
                           "  " & carp & vbCrLf & vbCrLf & _
                           "Archivos faltantes para la fecha " & fecNorm & ":" & vbCrLf & _
                           faltantes & vbCrLf & _
                           obtenerFormatoCorrectoInsumoBO(fecNorm, bBD, bCD, bPEA)
        Exit Function
    End If
    
    buscarInsumosCorrectosEnCarpetaBO = True
    Exit Function
ErrBusq:
    msgErrorBusqueda = "- Error al buscar insumos en la carpeta: " & Err.Description & vbCrLf & vbCrLf & obtenerFormatoCorrectoInsumoBO(fechaCorte, bBD, bCD, bPEA)
    buscarInsumosCorrectosEnCarpetaBO = False
End Function

' Ejecuta la carga definitiva a la tabla correspondiente (BD_INSUMO_PRIMER_PISO, BD_INSUMO_TESORERIA o BD_INSUMO_SEGUNDO_PISO) y registra en BITACORA_CARGA_INSUMOS
Public Function guardarInsumoValidadoBO(ByVal rutaArchivo As String, ByVal nombreArchivo As String, ByVal fechaCorte As String, ByVal totalFilas As Long, ByVal fondosStr As String, Optional ByVal nombreTablaDestino As String = "") As Boolean
    Dim dao As FondoDAO, tbl As String, nomUp As String
    On Error Resume Next
    tbl = Trim(nombreTablaDestino)
    If tbl = "" Then
        nomUp = UCase(Trim(nombreArchivo))
        If InStr(nomUp, "TESORERIA") > 0 Then
            tbl = "BD_INSUMO_TESORERIA"
        ElseIf InStr(nomUp, "SEGUNDO_PISO") > 0 Then
            tbl = "BD_INSUMO_SEGUNDO_PISO"
        Else
            tbl = "BD_INSUMO_PRIMER_PISO"
        End If
    End If
    Set dao = New FondoDAO
    ' Si el archivo acaba de ser validado en RAM por validarYProcesarInsumoBO, se pasa la matriz
    ' directamente para evitar abrir una segunda instancia de Excel.Application.
    If StrComp(mCacheRutaValidada, rutaArchivo, vbTextCompare) = 0 And mCacheUFilaValidada >= 2 Then
        Call dao.importarInsumoEnTablaDAO(rutaArchivo, fechaCorte, tbl, mCacheMatrizValidada, mCacheUFilaValidada, mCacheUColValidada)
    Else
        Call dao.importarInsumoEnTablaDAO(rutaArchivo, fechaCorte, tbl)
    End If
    Call dao.registrarBitacoraInsumoDAO(nombreArchivo, fechaCorte, totalFilas, tbl & " | " & fondosStr)
    Set dao = Nothing
    Call InvalidarCacheDatosBO
    guardarInsumoValidadoBO = True
    On Error GoTo 0
End Function

' Valida exhaustivamente celda por celda en memoria RAM (lectura matricial ultra-rapida en 1 llamada COM):
' 1) Formato unico obligatorio del nombre de archivo (INSUMO_<PORTAFOLIO>_DD_MM_AAAA.xlsx),
' 2) Orden y nombre exacto de las 53 columnas (incluyendo ID_INSTRUMENTO en columna 53),
' 3) 0 celdas vacias en las 53 columnas x N filas, 4) Tipos de dato y saldos >= 0,
' 5) FECHA_INSUMO en cada fila y 6) Deteccion de NUEVOS INSTRUMENTOS en columna ID_INSTRUMENTO contra tabla INSTRUMENTOS
Public Function validarYProcesarInsumoBO(ByVal rutaArchivo As String, ByVal fechaCorte As String, ByRef totalFilasRevisadas As Long, ByRef totalErrores As Long, ByRef detalleErrores As String, Optional ByVal tipoEsperado As String = "", Optional ByRef nuevosInstrumentosDetalle As String = "", Optional ByRef nuevosInstrumentosCodigos As String = "") As Boolean
    Dim xlApp As Object, wb As Object, ws As Object, cerrarXl As Boolean
    Dim fso As Object, dao As FondoDAO
    Dim uFila As Long, uCol As Long, r As Long, c As Long
    Dim arrCols As Variant, arrColsEnteras As Variant, idxEnt As Long, colEnt As Long
    Dim matrizDatos As Variant, vCelda As Variant
    Dim filaStr(1 To 53) As String
    Dim hActual As String, hEsperado As String, vVal As String, vValUp As String
    Dim strNomArch As String, fechaPantallaNorm As String, fecFilaNorm As String
    Dim maxColCheck As Long
    Dim tipoDet As String, tablaDest As String, msgErrNom As String
    Dim mapaInstBD As String, codInstFila As String, codUp As String, conteoNuevos As Long
    Dim bBD As Boolean, bCD As Boolean, bPEA As Boolean

    On Error GoTo ErrValidacion
    validarYProcesarInsumoBO = False
    totalFilasRevisadas = 0
    totalErrores = 0
    detalleErrores = ""
    nuevosInstrumentosDetalle = ""
    nuevosInstrumentosCodigos = ""
    mCacheRutaValidada = ""
    mCacheUFilaValidada = 0

    bBD = (tipoEsperado = "" Or UCase(tipoEsperado) = "PRIMER_PISO")
    bCD = (tipoEsperado = "" Or UCase(tipoEsperado) = "TESORERIA")
    bPEA = (tipoEsperado = "" Or UCase(tipoEsperado) = "SEGUNDO_PISO")

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(rutaArchivo) Then
        totalErrores = 1
        detalleErrores = "- La ruta de archivo especificada no existe en el equipo." & vbCrLf & vbCrLf & obtenerFormatoCorrectoInsumoBO(fechaCorte, bBD, bCD, bPEA)
        Exit Function
    End If

    strNomArch = fso.GetFileName(rutaArchivo)
    If Not esFechaFormatoCorrectoBO(fechaCorte) Then
        totalErrores = 1
        detalleErrores = "- El campo Fecha ('" & fechaCorte & "') no tiene un formato válido dd/mm/aaaa." & vbCrLf & vbCrLf & obtenerFormatoCorrectoInsumoBO(Format(Date, "dd/mm/yyyy"), bBD, bCD, bPEA)
        Exit Function
    End If
    If esFechaFuturaBO(fechaCorte) Then
        totalErrores = 1
        detalleErrores = "- No se puede escoger fechas futuras ('" & normalizarFechaBO(fechaCorte) & "' es posterior al día de la consulta actual: " & Format(Date, "dd/mm/yyyy") & ")."
        Exit Function
    End If
    fechaPantallaNorm = normalizarFechaBO(fechaCorte)

    ' 1. Validar que el nombre del archivo cumpla cabalmente con el FORMATO UNICO OFICIAL
    If Not validarFormatoUnicoNombreInsumoBO(strNomArch, fechaPantallaNorm, tipoEsperado, tipoDet, tablaDest, msgErrNom) Then
        totalErrores = totalErrores + 1
        AgregarErrorInsumo detalleErrores, msgErrNom
    End If

    ' Lista exacta y ordenada de las 53 columnas institucionales (51 NORMATIVA + FECHA_INSUMO + ID_INSTRUMENTO)
    arrCols = Array("P", "PISO", "G", "GRUPO", "ESTATA", "DESC_SUCURSAL", "SG", "DESC_SUBGRUPO", "CAMPOA", "NOMBRE", _
                    "RFC", "MONEDA2", "MON", "MONEDA", "EF", "ENT_FED", "CONCATENADO", "CTA", "SCATAS", "SCTA", _
                    "SSCTA", "SSSCTA", "SECTOR", "C", "CONCEPTO", "SC", "SCONCEPTO", "SSCONCEPTO", "DESTINO", "SDINMO", _
                    "SDFIMO", "PROMMO", "SDFIMN", "SDPRMN", "SDFIDL", "SDPRDL", "MESPRO", "ANOPRO", "ACT", "ACTIVIDAD", _
                    "GRUCON", "GRUPO_CONT", "GRUCONTNORM", "ARTICULO_NORMATIVO", "CODIGO_ACTIVIDAD_ECONOMICA_GRAL", _
                    "DESC_ACTIVIDAD_ECONOMICA_GRAL", "CODIGO_EJECUTIVO", "NOMBRE_EJECUTIVO", "DESCRIPCIÓN CUENTA", _
                    "SITUACION DEL CREDITO", "CONCEPTO_ETAPA", "FECHA_INSUMO", "ID_INSTRUMENTO")

    ' Columnas que exigen codigo numerico entero estricto
    arrColsEnteras = Array(1, 3, 7, 9, 13, 15, 18, 19, 20, 21, 22, 24, 26, 37, 38, 39, 41, 43, 45, 47)

    If mCacheMapaInst <> "" Then
        mapaInstBD = mCacheMapaInst
    Else
        Set dao = New FondoDAO
        mapaInstBD = dao.obtenerMapaInstrumentosDAO()
        Set dao = Nothing
        mCacheMapaInst = mapaInstBD
    End If

    If Not ObtenerExcelMotorBO(xlApp, cerrarXl) Then
        totalErrores = totalErrores + 1
        AgregarErrorInsumo detalleErrores, "- No se pudo iniciar el motor de lectura Excel."
        Exit Function
    End If
    On Error GoTo ErrValidacion

    Set wb = xlApp.Workbooks.Open(rutaArchivo, 0, True)
    Set ws = wb.Worksheets(1)

    uFila = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    uCol = ws.Cells(1, ws.Columns.Count).End(-4159).Column
    If uFila < 2 Then uFila = ws.UsedRange.Rows.Count
    If uCol < 2 Then uCol = ws.UsedRange.Columns.Count

    If uFila < 2 Or uCol < 1 Then
        wb.Close False
        Set ws = Nothing: Set wb = Nothing
        If cerrarXl Then xlApp.Quit: Set xlApp = Nothing
        totalErrores = totalErrores + 1
        AgregarErrorInsumo detalleErrores, "- [Archivo Vacío]: El insumo no contiene registros de datos para validar."
        detalleErrores = detalleErrores & vbCrLf & obtenerFormatoCorrectoInsumoBO(fechaPantallaNorm, bBD, bCD, bPEA)
        Exit Function
    End If

    ' Volcado matricial en 1 sola llamada COM a memoria RAM de VBA
    matrizDatos = ws.Range(ws.Cells(1, 1), ws.Cells(uFila, uCol)).Value

    wb.Close False
    Set ws = Nothing: Set wb = Nothing
    If cerrarXl Then xlApp.Quit: Set xlApp = Nothing

    ' 2. Verificar estructura interna: mismo numero (53 columnas incl. ID_INSTRUMENTO), mismo orden y mismos nombres textuales
    If uCol <> 53 Then
        totalErrores = totalErrores + 1
        AgregarErrorInsumo detalleErrores, "- [Estructura Interna]: El archivo tiene " & uCol & " columnas (se requieren exactamente las 53 columnas oficiales incluyendo ID_INSTRUMENTO)."
    End If

    maxColCheck = IIf(uCol < 53, uCol, 53)
    For c = 1 To maxColCheck
        vCelda = matrizDatos(1, c)
        If IsEmpty(vCelda) Or IsNull(vCelda) Then
            hActual = ""
        Else
            hActual = UCase$(Trim$(CStr(vCelda)))
        End If
        hEsperado = UCase$(CStr(arrCols(c - 1)))
        If c = 53 And (hActual = "INSTRUMENTO" Or hActual = "INSTRUMENTOS") Then hActual = "ID_INSTRUMENTO"
        If Replace(hActual, "Ó", "O") <> Replace(hEsperado, "Ó", "O") Then
            totalErrores = totalErrores + 1
            AgregarErrorInsumo detalleErrores, "- [Columna " & c & " Desordenada/Incorrecta]: Es '" & hActual & "' y debe ser '" & hEsperado & "'."
        End If
    Next c

    If uCol < 50 Then
        detalleErrores = detalleErrores & vbCrLf & obtenerFormatoCorrectoInsumoBO(fechaPantallaNorm, bBD, bCD, bPEA)
        validarYProcesarInsumoBO = False
        Exit Function
    End If

    conteoNuevos = 0

    ' 3. Verificar EXHAUSTIVAMENTE fila por fila y celda por celda en memoria RAM (las 53 columnas x todas las filas)
    '    Optimizado: se extrae y recorta cada celda una sola vez por fila en filaStr(1..53) con funciones nativas VBA
    '    (IsEmpty/Trim$/CStr), evitando 45,000 llamadas tardias a Application.Nz sin omitir ninguna validacion.
    For r = 2 To uFila
        totalFilasRevisadas = totalFilasRevisadas + 1

        ' 3.a Verificar que NINGUNA de las 53 celdas de la fila este vacia y cachear en filaStr(c)
        For c = 1 To maxColCheck
            vCelda = matrizDatos(r, c)
            If IsEmpty(vCelda) Or IsNull(vCelda) Then
                vVal = ""
            Else
                vVal = Trim$(CStr(vCelda))
            End If
            filaStr(c) = vVal
            If vVal = "" Then
                totalErrores = totalErrores + 1
                AgregarErrorInsumo detalleErrores, "- Fila " & r & " [" & arrCols(c - 1) & "]: Celda vacía (no se permiten campos incompletos)."
            End If
        Next c

        ' 3.b Validar tipo de dato entero estricto en las 20 columnas de claves y codigos numericos
        For idxEnt = LBound(arrColsEnteras) To UBound(arrColsEnteras)
            colEnt = CLng(arrColsEnteras(idxEnt))
            If colEnt <= maxColCheck Then
                vVal = filaStr(colEnt)
                If InStr(vVal, "'") > 0 Then vVal = Replace(vVal, "'", "")
                If vVal <> "" Then
                    If Not IsNumeric(vVal) Or Utilities.tieneLetras(vVal) Then
                        totalErrores = totalErrores + 1
                        AgregarErrorInsumo detalleErrores, "- Fila " & r & " [" & arrCols(colEnt - 1) & "]: '" & vVal & "' formato incorrecto (requiere entero numérico)."
                    End If
                End If
            End If
        Next idxEnt

        ' 3.c Validar RFC (Columna 11: 9 a 13 caracteres alfanumericos)
        vVal = filaStr(11)
        If vVal <> "" And (Len(vVal) < 9 Or Len(vVal) > 13) Then
            totalErrores = totalErrores + 1
            AgregarErrorInsumo detalleErrores, "- Fila " & r & " [RFC]: '" & vVal & "' inválido (long. " & Len(vVal) & "; requiere 9 a 13 caracteres)."
        End If

        ' 3.d Validar MON (13) y MONEDA (14)
        vVal = filaStr(13)
        If vVal <> "" Then
            If Not IsNumeric(vVal) Or (Val(vVal) <> 1 And Val(vVal) <> 2 And Val(vVal) <> 3) Then
                totalErrores = totalErrores + 1
                AgregarErrorInsumo detalleErrores, "- Fila " & r & " [MON]: Código '" & vVal & "' inválido (use 1=M.N. o 2=M.E.)."
            End If
        End If

        vValUp = UCase$(filaStr(14))
        If vValUp <> "" And vValUp <> "M.N." And vValUp <> "DLS" And vValUp <> "MXN" And vValUp <> "USD" And vValUp <> "UDI" And vValUp <> "M.E." Then
            totalErrores = totalErrores + 1
            AgregarErrorInsumo detalleErrores, "- Fila " & r & " [MONEDA]: '" & vValUp & "' fuera de catálogo CAT_MONEDAS."
        End If

        ' 3.e Validar las 7 columnas de saldos financieros (Columnas 30 a 36: SDINMO..SDPRDL)
        For c = 30 To 36
            vVal = filaStr(c)
            If vVal <> "" Then
                If Not IsNumeric(vVal) Then
                    totalErrores = totalErrores + 1
                    AgregarErrorInsumo detalleErrores, "- Fila " & r & " [" & arrCols(c - 1) & "]: Texto '" & vVal & "' (requiere importe numérico >= 0)."
                ElseIf CDbl(vVal) < 0# Then
                    totalErrores = totalErrores + 1
                    AgregarErrorInsumo detalleErrores, "- Fila " & r & " [" & arrCols(c - 1) & "]: Saldo negativo (" & Format(CDbl(vVal), "$#,##0.00") & ") inválido."
                End If
            End If
        Next c

        ' 3.f Validar rangos de MESPRO (37) y ANOPRO (38)
        vVal = filaStr(37)
        If vVal <> "" And IsNumeric(vVal) Then
            If Val(vVal) < 1 Or Val(vVal) > 12 Then
                totalErrores = totalErrores + 1
                AgregarErrorInsumo detalleErrores, "- Fila " & r & " [MESPRO]: Mes '" & vVal & "' fuera de rango (1 a 12)."
            End If
        End If

        vVal = filaStr(38)
        If vVal <> "" And IsNumeric(vVal) Then
            If Val(vVal) < 2000 Or Val(vVal) > 2035 Then
                totalErrores = totalErrores + 1
                AgregarErrorInsumo detalleErrores, "- Fila " & r & " [ANOPRO]: Año '" & vVal & "' fuera de rango (2000-2035)."
            End If
        End If

        ' 3.g Validar SITUACION DEL CREDITO (50)
        vValUp = UCase$(filaStr(50))
        If vValUp <> "" And vValUp <> "E1" And vValUp <> "E2" And vValUp <> "E3" And vValUp <> "VIGENTE" And vValUp <> "VENCIDO" Then
            totalErrores = totalErrores + 1
            AgregarErrorInsumo detalleErrores, "- Fila " & r & " [SITUACION]: '" & vValUp & "' inválido (use E1, E2, E3)."
        End If

        ' 3.h Validar FECHA_INSUMO (Columna 52) en cada fila contra la fecha del titulo y del campo Fecha
        If uCol >= 52 Then
            vVal = filaStr(52)
            If vVal <> "" Then
                If vVal <> fechaPantallaNorm Then
                    If Not esFechaFormatoCorrectoBO(vVal) And Not IsDate(vVal) Then
                        totalErrores = totalErrores + 1
                        AgregarErrorInsumo detalleErrores, "- Fila " & r & " [FECHA_INSUMO]: '" & vVal & "' no es una fecha válida (dd/mm/aaaa)."
                    Else
                        If esFechaFormatoCorrectoBO(vVal) Then
                            fecFilaNorm = normalizarFechaBO(vVal)
                        Else
                            fecFilaNorm = Format(CDate(vVal), "dd/mm/yyyy")
                        End If
                        If fecFilaNorm <> fechaPantallaNorm Then
                            totalErrores = totalErrores + 1
                            AgregarErrorInsumo detalleErrores, "- Fila " & r & " [FECHA_INSUMO]: Fecha de fila (" & fecFilaNorm & ") no coincide con " & fechaPantallaNorm & "."
                        End If
                    End If
                End If
            End If
        End If

        ' 3.i Validar ID_INSTRUMENTO (Columna 53) y detectar si hay NUEVOS INSTRUMENTOS no registrados en INSTRUMENTOS
        If uCol >= 53 Then
            codInstFila = filaStr(53)
            If codInstFila <> "" Then
                codUp = UCase$(codInstFila)
                If InStr(1, mapaInstBD, "|" & codUp & "|", vbBinaryCompare) = 0 Then
                    If InStr(1, ";" & nuevosInstrumentosCodigos & ";", ";" & codUp & ";", vbTextCompare) = 0 Then
                        conteoNuevos = conteoNuevos + 1
                        If nuevosInstrumentosCodigos <> "" Then nuevosInstrumentosCodigos = nuevosInstrumentosCodigos & ";"
                        nuevosInstrumentosCodigos = nuevosInstrumentosCodigos & codInstFila
                        nuevosInstrumentosDetalle = nuevosInstrumentosDetalle & "  • " & codInstFila & " (detectado por primera vez en Fila " & r & ", Grupo: " & filaStr(4) & ")" & vbCrLf
                    End If
                End If
            End If
        End If
    Next r

    ' Si existen nuevos instrumentos en la columna ID_INSTRUMENTO que no esten contemplados en el catalogo INSTRUMENTOS,
    ' bloquear internamente la subida del insumo para proteger la integridad de la base de datos
    If conteoNuevos > 0 Or Trim(nuevosInstrumentosCodigos) <> "" Then
        totalErrores = totalErrores + conteoNuevos
        AgregarErrorInsumo detalleErrores, "- [Validación Interna de Nuevos Instrumentos]: Se detectaron " & conteoNuevos & " nuevo(s) instrumento(s) en la columna [ID_INSTRUMENTO] que NO están contemplados en el catálogo INSTRUMENTOS:" & vbCrLf & nuevosInstrumentosDetalle & "  No se permite subir el insumo hasta que todos sus instrumentos estén dados de alta en el apartado 'Gestionar Instrumentos'."
    End If

    If totalErrores > 0 Then
        detalleErrores = detalleErrores & vbCrLf & obtenerFormatoCorrectoInsumoBO(fechaPantallaNorm, bBD, bCD, bPEA)
        validarYProcesarInsumoBO = False
        Exit Function
    End If

    ' Guardar en cache la matriz validada para que guardarInsumoValidadoBO la inserte sin reabrir Excel
    mCacheRutaValidada = rutaArchivo
    mCacheFechaValidada = fechaPantallaNorm
    mCacheMatrizValidada = matrizDatos
    mCacheUFilaValidada = uFila
    mCacheUColValidada = uCol

    validarYProcesarInsumoBO = True
    Exit Function

ErrValidacion:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    Set ws = Nothing: Set wb = Nothing
    If cerrarXl And Not xlApp Is Nothing Then xlApp.Quit: Set xlApp = Nothing
    If totalErrores = 0 Then totalErrores = 1
    AgregarErrorInsumo detalleErrores, "- [Error Crítico de Lectura/Estructura]: No fue posible validar la integridad del archivo (" & Err.Description & ")."
    detalleErrores = detalleErrores & vbCrLf & obtenerFormatoCorrectoInsumoBO(fechaCorte, bBD, bCD, bPEA)
    validarYProcesarInsumoBO = False
End Function

' ==============================================================================================
' GENERADOR EJECUTIVO INSTITUCIONAL ALIMENTADO 100% POR CONSULTAS SQL DESDE RIESGOSBD.ACCDB
' Extrae los datos reales cargados desde los insumos en CARTERA_ETIQUETADA mediante ConsultasSQL
' y FondoDAO, con margenes minimos (0.10 pulg / ~2.5 mm) y centrado vertical y horizontal 1:1.
' ==============================================================================================
Private Sub FormatearTablaEjecutivaBO(ByRef ws As Object, ByVal filaHdr As Long, ByVal filaIni As Long, ByVal filaFin As Long, ByVal hdrArr As Variant, ByVal dataArr As Variant, Optional ByVal modoIndividual As Boolean = False, Optional ByVal etiquetaPortafolio As String = "")
    ' etiquetaPortafolio: si se indica, las columnas numericas de la tabla (C a G)
    ' se dejan como formulas SUMIFS contra la hoja DATOS del template (rango acotado
    ' $2:$5000 para recalculo instantaneo).
    Dim r As Long
    Dim sPort As String, k As Integer
    Dim colRep As Variant, colDat As Variant
    ws.Rows(filaHdr).RowHeight = IIf(modoIndividual, 30, 26)
    ws.Range("A" & filaHdr & ":H" & filaHdr).Value = hdrArr
    With ws.Range("A" & filaHdr & ":H" & filaHdr)
        .Interior.Color = RGB(214, 223, 236)
        .Font.Name = "Cambria"
        .Font.Size = 9.6
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 120)
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
        .WrapText = True
        .Borders.LineStyle = 1
        .Borders.Color = RGB(184, 196, 212)
    End With

    ws.Range("A" & filaIni & ":H" & filaFin).Value = dataArr

    If Trim(etiquetaPortafolio) <> "" Then
        On Error Resume Next
        colRep = Array("C", "D", "E", "F", "G")
        colDat = Array("D", "E", "F", "G", "H")
        sPort = Replace(etiquetaPortafolio, Chr(34), Chr(34) & Chr(34))
        For k = 0 To 4
            If filaFin > filaIni Then
                ws.Range(colRep(k) & filaIni & ":" & colRep(k) & (filaFin - 1)).Formula = _
                    "=SUMIFS(DATOS!$" & colDat(k) & "$2:$" & colDat(k) & "$5000" & _
                    ",DATOS!$A$2:$A$5000," & Chr(34) & sPort & Chr(34) & _
                    ",DATOS!$B$2:$B$5000,$A" & filaIni & ")"
                ws.Cells(filaFin, 3 + k).Formula = "=SUM(" & colRep(k) & filaIni & ":" & colRep(k) & (filaFin - 1) & ")"
            Else
                ws.Range(colRep(k) & filaIni & ":" & colRep(k) & filaFin).Formula = _
                    "=SUMIFS(DATOS!$" & colDat(k) & "$2:$" & colDat(k) & "$5000" & _
                    ",DATOS!$A$2:$A$5000," & Chr(34) & sPort & Chr(34) & _
                    ",DATOS!$B$2:$B$5000,$A" & filaIni & ")"
            End If
        Next k
        Err.Clear
        On Error GoTo 0
    End If

    With ws.Range("A" & filaIni & ":H" & filaFin)
        .Font.Name = "Cambria"
        .Font.Size = 9.2
        .Font.Color = RGB(30, 41, 59)
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
        .ShrinkToFit = True
        .Borders.LineStyle = 1
        .Borders.Color = RGB(184, 196, 212)
    End With

    ' RENDIMIENTO: Altura de filas y fondo blanco en 1 sola llamada COM por rango;
    ' solo se colorean las filas impares (Step 2), reduciendo llamadas COM a la mitad.
    If filaFin > filaIni Then
        ws.Range("A" & filaIni & ":A" & (filaFin - 1)).RowHeight = IIf(modoIndividual, 23.5, 20.5)
        ws.Range("A" & filaIni & ":G" & (filaFin - 1)).Interior.Color = RGB(255, 255, 255)
        For r = filaIni + 1 To filaFin - 1 Step 2
            ws.Range("A" & r & ":G" & r).Interior.Color = RGB(245, 248, 252)
        Next r
    End If

    ws.Range("A" & filaIni & ":A" & filaFin).Font.Bold = True
    ws.Range("H" & filaIni & ":H" & filaFin).Font.Bold = True

    If filaFin > filaIni Then
        With ws.Range("H" & filaIni & ":H" & (filaFin - 1))
            .Interior.Color = RGB(232, 245, 233)
            .Font.Color = RGB(0, 90, 60)
        End With
    End If

    ws.Rows(filaFin).RowHeight = IIf(modoIndividual, 25, 22)
    With ws.Range("A" & filaFin & ":H" & filaFin)
        .Font.Bold = True
        .Font.Size = 9.6
        .Font.Color = RGB(31, 78, 120)
        .Interior.Color = RGB(218, 229, 242)
    End With
End Sub

Private Sub PrepararMarcoDictamenYGraficaBO(ByRef ws As Object, ByRef curRow As Long, ByVal textoDictamen As String, ByVal modoIndividual As Boolean, ByRef filaIniImg As Long, ByRef filaFinImg As Long)
    Dim nFilasImg As Long

    ws.Range("A" & curRow & ":H" & curRow).Merge
    ws.Cells(curRow, 1).Value = textoDictamen
    With ws.Range("A" & curRow & ":H" & curRow)
        .Interior.Color = RGB(242, 246, 251)
        .Borders.LineStyle = 1
        .Borders.Color = RGB(184, 196, 212)
        .Font.Name = "Cambria"
        .Font.Italic = True
        .Font.Size = 9.2
        .Font.Color = RGB(31, 58, 82)
        .HorizontalAlignment = -4131
        .VerticalAlignment = -4108
        .WrapText = True
    End With
    ws.Rows(curRow).RowHeight = IIf(modoIndividual, 34, 28)
    curRow = curRow + 1

    nFilasImg = IIf(modoIndividual, 15, 12)
    filaIniImg = curRow
    filaFinImg = curRow + nFilasImg - 1
    ' 1 sola llamada COM para ajustar la altura de todas las filas del contenedor de grafica
    ws.Range("A" & filaIniImg & ":A" & filaFinImg).RowHeight = IIf(modoIndividual, 22, 19)

    ws.Range("A" & filaIniImg & ":H" & filaFinImg).Merge
    With ws.Range("A" & filaIniImg & ":H" & filaFinImg)
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = 1
        .Borders.Color = RGB(184, 196, 212)
    End With

    curRow = filaFinImg + 1
    ws.Rows(curRow).RowHeight = 6
    curRow = curRow + 1
End Sub

Private Sub ColocarGraficaDatosBO(ByRef ws As Object, ByRef fso As Object, _
        ByVal filaIniImg As Long, ByVal filaFinImg As Long, _
        ByVal filaIniDat As Long, ByVal filaFinDat As Long, _
        ByVal titulo As String, ByVal rutaImgRespaldo As String, _
        ByVal tipoGrafica As String, _
        Optional ByRef wsFuente As Object, _
        Optional ByVal colCat As String = "A", _
        Optional ByVal colVal As String = "E")
    ' Construye una grafica NATIVA de Excel (objeto ChartObject real, no imagen).
    '
    ' wsFuente / colCat / colVal indican DE DONDE toma la grafica sus numeros.
    ' Si se omiten, se comporta exactamente como antes: lee la propia hoja del
    ' reporte, categorias en la columna A y valores en la columna E.
    ' El generador de reportes si los usa, para que la grafica lea las hojas
    ' CALCULO y CALC_INSTR del template, es decir el RESULTADO de las formulas
    ' que calcula Excel, y no los datos crudos.
    '
    ' Si por cualquier motivo no se puede crear el objeto, cae al PNG de respaldo
    ' para no entregar un reporte incompleto.
    On Error GoTo Respaldo
    Dim topBox As Double, altoBox As Double, anchoTabla As Double, leftBox As Double
    Dim co As Object, ch As Object, sc As Object, nCat As Long

    If wsFuente Is Nothing Then Set wsFuente = ws
    If Trim(colCat) = "" Then colCat = "A"
    If Trim(colVal) = "" Then colVal = "E"

    If filaIniImg <= 0 Or filaFinImg < filaIniImg Then Exit Sub
    If filaIniDat <= 0 Or filaFinDat < filaIniDat Then GoTo Respaldo
    nCat = filaFinDat - filaIniDat + 1
    If nCat < 1 Then GoTo Respaldo

    topBox = ws.Rows(filaIniImg).Top
    altoBox = (ws.Rows(filaFinImg).Top + ws.Rows(filaFinImg).Height) - topBox
    anchoTabla = ws.Range("A" & filaIniImg & ":H" & filaIniImg).Width
    leftBox = ws.Columns(1).Left

    Set co = ws.ChartObjects.Add(leftBox + 3, topBox + 2, anchoTabla - 6, altoBox - 4)
    If co Is Nothing Then GoTo Respaldo
    Set ch = co.Chart

    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop
    ' El tipo de grafica se elige segun la naturaleza de los datos de cada reporte:
    '   BARRA   -> Primer Piso. Los conceptos de cuenta tienen nombres largos;
    '              en barras horizontales se leen completos.
    '   ANILLO  -> Tesoreria. Un portafolio se entiende mejor como composicion
    '              (que peso tiene cada grupo) que como comparacion de magnitudes.
    '   COLUMNA -> Segundo Piso. Comparacion directa entre sectores economicos.
    Select Case UCase(Trim(tipoGrafica))
        Case "BARRA":  ch.ChartType = 57          ' xlBarClustered
        Case "ANILLO": ch.ChartType = -4120       ' xlDoughnut
        Case Else:     ch.ChartType = 51          ' xlColumnClustered
    End Select
    Set sc = ch.SeriesCollection.NewSeries
    sc.Values = wsFuente.Range(colVal & filaIniDat & ":" & colVal & filaFinDat)
    sc.XValues = wsFuente.Range(colCat & filaIniDat & ":" & colCat & filaFinDat)
    sc.Name = "Saldo Final M.N. ($ MDP)"

    On Error Resume Next
    If UCase(Trim(tipoGrafica)) <> "ANILLO" Then
        sc.Format.Fill.ForeColor.RGB = RGB(0, 78, 76)
        sc.Format.Line.Visible = False
    End If
    ch.HasTitle = True
    ch.ChartTitle.Text = titulo
    ch.ChartTitle.Font.Name = "Cambria"
    ch.ChartTitle.Font.Size = 9
    ch.ChartTitle.Font.Color = RGB(0, 78, 76)
    ch.ChartArea.Interior.ColorIndex = -4142           ' sin relleno
    ch.ChartArea.Border.LineStyle = 0
    ch.PlotArea.Interior.ColorIndex = -4142

    If UCase(Trim(tipoGrafica)) = "ANILLO" Then
        ' Composicion: leyenda a la derecha y porcentaje sobre cada anillo.
        ch.HasLegend = True
        ch.Legend.Position = -4152                     ' xlLegendPositionRight
        ch.Legend.Font.Size = 6.5
        ch.Legend.Font.Name = "Segoe UI"
        sc.HasDataLabels = True
        sc.DataLabels.ShowPercentage = True
        sc.DataLabels.ShowValue = False
        sc.DataLabels.Font.Size = 6
        sc.ApplyDataLabels 5                           ' xlDataLabelsShowPercent
        ch.ChartGroups(1).DoughnutHoleSize = 45
    Else
        ch.HasLegend = False
        With ch.Axes(1)                                ' categorias
            .TickLabels.Font.Size = 6.5
            .TickLabels.Font.Name = "Segoe UI"
            If UCase(Trim(tipoGrafica)) <> "BARRA" Then .TickLabels.Orientation = 45
        End With
        With ch.Axes(2)                                ' valores
            .TickLabels.Font.Size = 6.5
            .TickLabels.Font.Name = "Segoe UI"
            .TickLabels.NumberFormat = "$#,##0"
            .MajorGridlines.Border.Color = RGB(214, 223, 236)
        End With
        If nCat <= 12 Then
            sc.HasDataLabels = True
            sc.DataLabels.Font.Size = 6
            sc.DataLabels.NumberFormat = "$#,##0"
            sc.DataLabels.Font.Color = RGB(60, 60, 60)
        End If
    End If
    co.Name = "GraficaDatos_" & CStr(filaIniImg)
    Exit Sub

Respaldo:
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    Call ColocarGraficaCentradaBO(ws, fso, filaIniImg, filaFinImg, rutaImgRespaldo)
End Sub

Private Sub ColocarGraficaCentradaBO(ByRef ws As Object, ByRef fso As Object, ByVal filaIniImg As Long, ByVal filaFinImg As Long, ByVal rutaImg As String)
    Dim topBox As Double, altoBox As Double, anchoTabla As Double, anchoImg As Double, altoImg As Double, leftImg As Double
    If filaIniImg <= 0 Or filaFinImg < filaIniImg Then Exit Sub
    If Not fso.FileExists(rutaImg) Then Exit Sub

    topBox = ws.Rows(filaIniImg).Top
    altoBox = (ws.Rows(filaFinImg).Top + ws.Rows(filaFinImg).Height) - topBox
    anchoTabla = ws.Range("A" & filaIniImg & ":H" & filaIniImg).Width

    altoImg = altoBox - 4
    anchoImg = altoImg * 2.22
    If anchoImg > anchoTabla - 6 Then
        anchoImg = anchoTabla - 6
        altoImg = anchoImg / 2.22
        If altoImg > altoBox - 4 Then altoImg = altoBox - 4
    End If
    leftImg = ws.Columns(1).Left + ((anchoTabla - anchoImg) / 2)
    ws.Shapes.AddPicture rutaImg, False, True, leftImg, topBox + ((altoBox - altoImg) / 2), anchoImg, altoImg
End Sub

Private Sub CalibrarRangoVerticalCartaBO(ByRef ws As Object, ByVal filaIniPag As Long, ByVal filaFinPag As Long, ByVal filaIniImg As Long, ByVal filaFinImg As Long, ByVal incluyeRepeticionEncabezado As Boolean)
    On Error Resume Next
    Dim anchoTablaPt As Double, sumaColUnidades As Double, altoEncabezadoPt As Double, altoActualPt As Double
    Dim altoMetaTotalPt As Double, altoMetaRangoPt As Double, altoPostPt As Double
    Dim factorEscala As Double, rCal As Long, cCal As Long, nuevoAlto As Double, diffPt As Double, nFilasImg As Long

    If filaFinPag < filaIniPag Then Exit Sub

    anchoTablaPt = ws.Range("A1:H1").Width
    If anchoTablaPt <= 100 Then
        sumaColUnidades = 0
        For cCal = 1 To 8
            sumaColUnidades = sumaColUnidades + ws.Columns(cCal).ColumnWidth
        Next cCal
        anchoTablaPt = sumaColUnidades * 5.35
    End If

    altoMetaTotalPt = anchoTablaPt * 1.296

    altoEncabezadoPt = 0
    If incluyeRepeticionEncabezado Then
        altoEncabezadoPt = ws.Range("A1:A5").Height
    End If

    altoMetaRangoPt = altoMetaTotalPt - altoEncabezadoPt
    altoActualPt = ws.Range("A" & filaIniPag & ":A" & filaFinPag).Height
    If altoActualPt <= 50 Or altoMetaRangoPt <= 50 Then Exit Sub

    factorEscala = altoMetaRangoPt / altoActualPt
    If factorEscala < 0.55 Then factorEscala = 0.55
    If factorEscala > 1.85 Then factorEscala = 1.85

    For rCal = filaIniPag To filaFinPag
        nuevoAlto = Round(ws.Rows(rCal).RowHeight * factorEscala, 1)
        If nuevoAlto < 5 Then nuevoAlto = 5
        If nuevoAlto > 90 Then nuevoAlto = 90
        ws.Rows(rCal).RowHeight = nuevoAlto
    Next rCal

    If filaIniImg >= filaIniPag And filaFinImg <= filaFinPag And filaFinImg >= filaIniImg Then
        altoPostPt = ws.Range("A" & filaIniPag & ":A" & filaFinPag).Height
        diffPt = altoMetaRangoPt - altoPostPt
        nFilasImg = (filaFinImg - filaIniImg) + 1
        If Abs(diffPt) > 0.2 And nFilasImg > 0 Then
            nuevoAlto = Round(ws.Rows(filaIniImg).RowHeight + (diffPt / nFilasImg), 1)
            If nuevoAlto >= 8 And nuevoAlto <= 95 Then
                ws.Range("A" & filaIniImg & ":A" & filaFinImg).RowHeight = nuevoAlto
            End If
        End If
    End If
End Sub

' Devuelve la ruta fisica del template generico, probando las ubicaciones posibles
Public Function rutaTemplateGenericoBO() As String
    Dim fso As Object, raiz As String, r As String
    rutaTemplateGenericoBO = ""
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    raiz = CurrentProject.Path
    r = raiz & PATH_TEMPLATES & TEMPLATE_GENERICO
    If fso.FileExists(r) Then rutaTemplateGenericoBO = r: Exit Function
    r = raiz & "" & TEMPLATE_GENERICO
    If fso.FileExists(r) Then rutaTemplateGenericoBO = r: Exit Function
    r = fso.GetParentFolderName(raiz) & PATH_TEMPLATES & TEMPLATE_GENERICO
    If fso.FileExists(r) Then rutaTemplateGenericoBO = r: Exit Function
End Function

' Copia el template al destino con FileCopy sin abrir ni tocar jamas el original
Public Function prepararLibroDesdeTemplateBO(ByVal rutaDestino As String, ByRef msgError As String) As Boolean
    Dim fso As Object, rutaTpl As String
    On Error GoTo ErrPrep
    prepararLibroDesdeTemplateBO = False
    msgError = ""
    rutaTpl = rutaTemplateGenericoBO()
    If rutaTpl = "" Then
        msgError = "No se encontro el template " & TEMPLATE_GENERICO & " en la carpeta " & _
                   CurrentProject.Path & PATH_TEMPLATES & ". El reporte se generara sin template."
        Exit Function
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(rutaDestino) Then fso.DeleteFile rutaDestino, True
    fso.CopyFile rutaTpl, rutaDestino, True
    If Not fso.FileExists(rutaDestino) Then
        msgError = "No se pudo crear la copia de trabajo del template en: " & rutaDestino
        Exit Function
    End If
    prepararLibroDesdeTemplateBO = True
    Exit Function
ErrPrep:
    msgError = "Error al preparar la copia del template: " & Err.Description
    prepararLibroDesdeTemplateBO = False
End Function

Private Function claveInstrumentoDesdeTextoBO(ByVal txt As String) As String
    Dim t As String, pos As Long
    t = Trim(CStr(Nz(txt, "")))
    pos = InStr(t, "|")
    If pos > 0 Then
        t = Trim(Mid(t, pos + 1))
    Else
        pos = InStr(t, "/")
        If pos > 0 Then t = Trim(Mid(t, pos + 1))
    End If
    pos = InStr(t, "(")
    If pos > 0 Then t = Trim(Left(t, pos - 1))
    t = Replace(Replace(Replace(t, Chr(10), " "), Chr(13), " "), vbTab, " ")
    claveInstrumentoDesdeTextoBO = Trim(t)
End Function

Public Sub volcarDatosCrudosEnTemplateBO(ByRef wb As Object, ByVal etiquetaPortafolio As String, _
        ByVal dataArr As Variant, ByVal nFilas As Long, ByRef filaDestino As Long)
    Dim wd As Object, salida As Variant, r As Long, c As Long, nFilasDet As Long
    On Error Resume Next
    If nFilas <= 0 Then Exit Sub
    nFilasDet = nFilas
    If nFilasDet > 1 Then
        If Left(UCase(Trim(CStr(Nz(dataArr(nFilasDet, 1), "")))), 5) = "TOTAL" Then
            nFilasDet = nFilasDet - 1
        End If
    End If
    If nFilasDet <= 0 Then Exit Sub
    Set wd = wb.Worksheets(HOJA_TPL_DATOS)
    If wd Is Nothing Then Exit Sub
    If filaDestino < 2 Then filaDestino = 2

    ReDim salida(1 To nFilasDet, 1 To 10)
    For r = 1 To nFilasDet
        salida(r, 1) = etiquetaPortafolio
        For c = 1 To 8
            salida(r, c + 1) = dataArr(r, c)
        Next c
        salida(r, 10) = claveInstrumentoDesdeTextoBO(CStr(Nz(dataArr(r, 2), "")))
    Next r

    wd.Range(wd.Cells(filaDestino, 1), wd.Cells(filaDestino + nFilasDet - 1, 10)).Value = salida
    filaDestino = filaDestino + nFilasDet
End Sub

' Escribe en CALC_INSTR!A la lista de instrumentos distintos leyendo DATOS!J en 1 sola llamada matricial COM
Public Function sembrarInstrumentosEnCalcInstrBO(ByRef wb As Object) As Long
    Dim wd As Object, wi As Object, dic As Object
    Dim uFila As Long, r As Long, k As String, n As Long, salida As Variant, colJArr As Variant
    On Error GoTo ErrSiembra
    sembrarInstrumentosEnCalcInstrBO = 0
    Set wd = wb.Worksheets(HOJA_TPL_DATOS)
    Set wi = wb.Worksheets(HOJA_TPL_CALCINSTR)
    If wd Is Nothing Or wi Is Nothing Then Exit Function

    wi.Range(wi.Cells(TPL_INSTR_FILA_INI, 1), wi.Cells(TPL_INSTR_FILA_FIN, 1)).ClearContents

    uFila = wd.Cells(wd.Rows.Count, 10).End(-4162).Row
    If uFila < 2 Then Exit Function

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = 1
    ReDim salida(1 To (TPL_INSTR_FILA_FIN - TPL_INSTR_FILA_INI + 1), 1 To 1)
    n = 0
    If uFila = 2 Then
        k = Trim$(CStr(Nz(wd.Cells(2, 10).Value, "")))
        If k <> "" Then
            n = 1
            salida(1, 1) = k
        End If
    Else
        colJArr = wd.Range(wd.Cells(2, 10), wd.Cells(uFila, 10)).Value
        For r = 1 To (uFila - 1)
            If IsEmpty(colJArr(r, 1)) Or IsNull(colJArr(r, 1)) Then
                k = ""
            Else
                k = Trim$(CStr(colJArr(r, 1)))
            End If
            If k <> "" Then
                If Not dic.Exists(k) Then
                    dic.Add k, 1
                    If n < (TPL_INSTR_FILA_FIN - TPL_INSTR_FILA_INI + 1) Then
                        n = n + 1
                        salida(n, 1) = k
                    End If
                End If
            End If
        Next r
    End If
    Set dic = Nothing

    If n > 0 Then
        wi.Range(wi.Cells(TPL_INSTR_FILA_INI, 1), _
                 wi.Cells(TPL_INSTR_FILA_INI + n - 1, 1)).Value = salida
    End If
    sembrarInstrumentosEnCalcInstrBO = n
    Exit Function
ErrSiembra:
    gUltimoErrorSQL = "sembrarInstrumentosEnCalcInstrBO: " & Err.Description
    sembrarInstrumentosEnCalcInstrBO = 0
End Function

' Activa el calculo automatico y recalcula el libro rapidamente sin reconstruir el arbol de 1M de filas
Public Sub recalcularTemplateBO(ByRef xlApp As Object, ByRef wb As Object)
    On Error Resume Next
    xlApp.Calculation = -4105                  ' xlCalculationAutomatic
    xlApp.Calculate
    If Err.Number <> 0 Then
        Err.Clear
        wb.Worksheets(HOJA_TPL_CALCULO).Calculate
        wb.Worksheets(HOJA_TPL_CALCINSTR).Calculate
        wb.Worksheets(HOJA_TPL_REPORTE).Calculate
    End If
    Err.Clear
End Sub

Private Sub RenderizarFirmasYPieInstitucionalBO(ByRef ws As Object, ByRef curRow As Long, ByVal strElaboro As String, ByVal strOrigenBD As String)
    ws.Range("A" & curRow & ":D" & curRow).Merge
    ws.Range("E" & curRow & ":H" & curRow).Merge
    ws.Range("A" & (curRow + 1) & ":D" & (curRow + 1)).Merge
    ws.Range("E" & (curRow + 1) & ":H" & (curRow + 1)).Merge

    ws.Cells(curRow, 1).Value = "ELABORACIÓN Y VALIDACIÓN TÉCNICA"
    ws.Cells(curRow, 5).Value = "AUTORIZACIÓN Y DICTAMEN EJECUTIVO"
    ws.Cells(curRow + 1, 1).Value = "Gerencia de Verificación de Volúmenes y Regalías — " & strElaboro
    ws.Cells(curRow + 1, 5).Value = "Dirección de Fiscalización y Contraprestaciones Petroleras — Hidrocarburos S.N.C."

    ws.Rows(curRow).RowHeight = 18
    ws.Rows(curRow + 1).RowHeight = 21.5
    With ws.Range("A" & curRow & ":H" & curRow)
        .Interior.Color = RGB(238, 243, 249)
        .Font.Name = "Cambria"
        .Font.Size = 8.3
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 120)
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
        .Borders.LineStyle = 1
        .Borders.Color = RGB(184, 196, 212)
    End With
    With ws.Range("A" & (curRow + 1) & ":H" & (curRow + 1))
        .Interior.Color = RGB(255, 255, 255)
        .Font.Name = "Cambria"
        .Font.Size = 8.6
        .Font.Bold = True
        .Font.Color = RGB(0, 78, 76)
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
        .Borders.LineStyle = 1
        .Borders.Color = RGB(184, 196, 212)
    End With
    curRow = curRow + 2
    ws.Rows(curRow).RowHeight = 7
    curRow = curRow + 1

    ws.Range("A" & curRow & ":H" & curRow).Merge
    ws.Cells(curRow, 1).Value = "DIRECCIÓN CORPORATIVA DE ECONOMÍA Y FISCALIZACIÓN DE HIDROCARBUROS  •  DIRECCIÓN DE FISCALIZACIÓN Y CONTRAPRESTACIONES PETROLERAS"
    With ws.Range("A" & curRow & ":H" & curRow)
        .Interior.Color = RGB(0, 78, 76)
        .Font.Name = "Cambria"
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 8.5
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
    End With
    ws.Rows(curRow).RowHeight = 18.5
    curRow = curRow + 1

    ws.Range("A" & curRow & ":H" & curRow).Merge
    ws.Cells(curRow, 1).Value = "Información Extraída por SQL de " & strOrigenBD & "  |  Gerencia de Verificación de Volúmenes y Regalías  |  Emisión: " & Format(Now, "dd/mm/yyyy hh:nn")
    With ws.Range("A" & curRow & ":H" & curRow)
        .Interior.Color = RGB(214, 223, 236)
        .Font.Name = "Cambria"
        .Font.Color = RGB(31, 78, 120)
        .Font.Size = 8.1
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
    End With
    ws.Rows(curRow).RowHeight = 17
End Sub

Public Function generarReporteEjecutivoBO(ByVal fechaCorte As String, ByVal formatoSalida As String, ByVal incBD As Boolean, ByVal incCD As Boolean, ByVal incPEA As Boolean, ByVal incConsolidado As Boolean, ByRef rutaSalida As String, ByRef msgError As String, Optional ByVal carpetaDestinoUsuario As String = "") As Boolean
    Dim fso As Object, dao As FondoDAO
    Dim rutaDescargas As String, fechaCorteStr As String, fechaPantallaNorm As String, tagSeleccion As String, tituloAlcance As String
    Dim xlApp As Object, wb As Object, ws As Object
    Dim curRow As Long, c As Long, r As Long, filaHdr As Long, filaIni As Long, filaFin As Long
    Dim strImgDir As String, strLogoPath As String, rutaImg1 As String, rutaImg2 As String, rutaImg3 As String
    Dim numSeleccionados As Integer, bloquesRenderizados As Integer, strElaboro As String
    Dim modoIndividual As Boolean
    Dim kpiAct As String, kpiRiesgo As String, kpiRend As String, kpiDict As String
    Dim strOrigenBD As String, strRegistrosSQL As String
    Dim hBoxLogo As Double, hLogo As Double, wLogo As Double
    Dim img1Ini As Long, img1Fin As Long, img2Ini As Long, img2Fin As Long, img3Ini As Long, img3Fin As Long
    Dim dat1Ini As Long, dat1Fin As Long, dat2Ini As Long, dat2Fin As Long, dat3Ini As Long, dat3Fin As Long
    Dim pagStart(1 To 3) As Long, pagEnd(1 To 3) As Long, pagImgIni(1 To 3) As Long, pagImgFin(1 To 3) As Long, numPags As Integer, pIdx As Integer
    Dim totalRegFecha As Long, fechasCargadas As String
    Dim totGlobalCred As Long, totGlobalIniMDP As Double, totGlobalFinMDP As Double, totGlobalPromMDP As Double, totGlobalFinMDD As Double
    Dim rowsBD As Variant, nFilasBD As Long, credBD As Long, iniBD As Double, finBD As Double, promBD As Double, usdBD As Double
    Dim rowsCD As Variant, nFilasCD As Long, credCD As Long, iniCD As Double, finCD As Double, promCD As Double, usdCD As Double
    Dim rowsPEA As Variant, nFilasPEA As Long, credPEA As Long, iniPEA As Double, finPEA As Double, promPEA As Double, usdPEA As Double
    ' --- Motor basado en el template Generico.xlsx ---
    Dim bDesdeTemplate As Boolean, msgTpl As String, filaDatos As Long, nInstrSembrados As Long
    Dim rutaTrabajoTpl As String
    Dim dblCalcExcel As Double, dblCalcVBA As Double
    Dim wsCalc As Object, wsCalcInstr As Object

    On Error GoTo ErrGen
    generarReporteEjecutivoBO = False
    msgError = ""

    If Not esFechaFormatoCorrectoBO(fechaCorte) Then
        msgError = "La fecha solicitada ('" & fechaCorte & "') no tiene un formato válido dd/mm/aaaa."
        Exit Function
    End If
    If esFechaFuturaBO(fechaCorte) Then
        msgError = "No se puede escoger fechas futuras (" & normalizarFechaBO(fechaCorte) & " es posterior al día de la consulta actual: " & Format(Date, "dd/mm/yyyy") & ")."
        Exit Function
    End If
    fechaPantallaNorm = normalizarFechaBO(fechaCorte)

    ' Si se solicita el reporte Consolidado, incluye internamente los 3 portafolios sin obligar a marcar sus casillas en pantalla
    If incConsolidado Then
        incBD = True: incCD = True: incPEA = True
    End If

    numSeleccionados = 0
    If incBD Then numSeleccionados = numSeleccionados + 1
    If incCD Then numSeleccionados = numSeleccionados + 1
    If incPEA Then numSeleccionados = numSeleccionados + 1

    If numSeleccionados = 0 Then
        msgError = "Debe seleccionar al menos un portafolio de riesgo."
        Exit Function
    End If

    ' ==============================================================================================
    ' EXTRACCION OBLIGATORIA DESDE LA BASE DE DATOS RIESGOSBD.ACCDB MEDIANTE CONSULTAS SQL (DAO)
    ' ==============================================================================================
    Dim nRegP1 As Long, nRegP2 As Long, nRegP3 As Long, msgBasesFaltantes As String, tagFecRep As String
    Dim cerrarXlRep As Boolean
    Call obtenerConteosTresTablasFechaBO(fechaPantallaNorm, nRegP1, nRegP2, nRegP3)
    totalRegFecha = nRegP1 + nRegP2 + nRegP3
    tagFecRep = Replace(fechaPantallaNorm, "/", "_")
    msgBasesFaltantes = ""
    If incBD And nRegP1 <= 0 Then msgBasesFaltantes = msgBasesFaltantes & "  • Opción 1 (BD_INSUMO_PRIMER_PISO): Falta cargar INSUMO_PRIMER_PISO_" & tagFecRep & ".xlsx" & vbCrLf
    If incCD And nRegP2 <= 0 Then msgBasesFaltantes = msgBasesFaltantes & "  • Opción 2 (BD_INSUMO_TESORERIA): Falta cargar INSUMO_TESORERIA_" & tagFecRep & ".xlsx" & vbCrLf
    If incPEA And nRegP3 <= 0 Then msgBasesFaltantes = msgBasesFaltantes & "  • Opción 3 (BD_INSUMO_SEGUNDO_PISO): Falta cargar INSUMO_SEGUNDO_PISO_" & tagFecRep & ".xlsx" & vbCrLf
    If totalRegFecha <= 0 Or msgBasesFaltantes <> "" Then
        Set dao = New FondoDAO
        fechasCargadas = dao.obtenerFechasCargadasDAO()
        Set dao = Nothing
        If fechasCargadas = "" Then
            msgError = "No existen insumos cargados en las 3 bases de datos de RIESGOSBD.accdb para la fecha " & fechaPantallaNorm & ":" & vbCrLf & _
                       msgBasesFaltantes & vbCrLf & _
                       "Ingrese primero a 'Cargar Insumos' y suba el(los) archivo(s) correspondiente(s)."
        Else
            msgError = "Faltan insumos por cargar en su base de datos correspondiente para la fecha de corte " & fechaPantallaNorm & ":" & vbCrLf & _
                       msgBasesFaltantes & vbCrLf & _
                       "Fechas con insumos actualmente cargados en RIESGOSBD.accdb: " & fechasCargadas
        End If
        Exit Function
    End If

    ' Extraer (o reutilizar desde caché en RAM si ya se consultó la misma fecha en este lote) las 3 consultas SQL
    Call obtenerAnalisisTresPortafoliosBO(fechaPantallaNorm, _
        rowsBD, nFilasBD, credBD, iniBD, finBD, promBD, usdBD, _
        rowsCD, nFilasCD, credCD, iniCD, finCD, promCD, usdCD, _
        rowsPEA, nFilasPEA, credPEA, iniPEA, finPEA, promPEA, usdPEA)

    ' Conciliacion matematica exacta: El total consolidado es exactamente la suma de los 3 reportes individuales
    totGlobalCred = credBD + credCD + credPEA
    totGlobalIniMDP = Round(iniBD, 2) + Round(iniCD, 2) + Round(iniPEA, 2)
    totGlobalFinMDP = Round(finBD, 2) + Round(finCD, 2) + Round(finPEA, 2)
    totGlobalPromMDP = Round(promBD, 2) + Round(promCD, 2) + Round(promPEA, 2)
    totGlobalFinMDD = Round(usdBD, 2) + Round(usdCD, 2) + Round(usdPEA, 2)

    img1Ini = 0: img1Fin = 0: img2Ini = 0: img2Fin = 0: img3Ini = 0: img3Fin = 0
    numPags = 1
    pagStart(1) = 1
    modoIndividual = (numSeleccionados = 1 And Not incConsolidado)

    If incConsolidado Or numSeleccionados = 3 Then
        tagSeleccion = "Consolidado_Integral_Riesgos"
        tituloAlcance = "INFORME EJECUTIVO CONSOLIDADO (SQL) — ADMINISTRACIÓN INTEGRAL DE RIESGOS (" & CStr(totGlobalCred) & " OPERACIONES)"
        kpiAct = "$" & Format(totGlobalFinMDP, "#,##0.00") & " MDP"
        kpiRiesgo = CStr(totGlobalCred) & " Op. | $" & Format(totGlobalFinMDD, "#,##0.00") & " MDD"
        kpiRend = "$" & Format(totGlobalPromMDP, "#,##0.00") & " MDP (Prom.)"
        kpiDict = "CONSOLIDADO SQL (" & fechaPantallaNorm & ")"
        strOrigenBD = "RIESGOSBD (3 Bases Insumos + INSTRUMENTOS)"
        strRegistrosSQL = CStr(totGlobalCred) & " operaciones vía SQL"
    ElseIf incBD And Not incCD And Not incPEA Then
        tagSeleccion = "Cartera_Comercial_1er_Piso"
        tituloAlcance = "INFORME EJECUTIVO DE RIESGOS (SQL): 1. CARTERA DE CRÉDITO COMERCIAL PRIMER PISO (" & CStr(credBD) & " CRÉDITOS)"
        kpiAct = "$" & Format(Round(finBD, 2), "#,##0.00") & " MDP"
        kpiRiesgo = CStr(credBD) & " Créd. | $" & Format(Round(usdBD, 2), "#,##0.00") & " MDD"
        kpiRend = "$" & Format(Round(promBD, 2), "#,##0.00") & " MDP (Prom.)"
        kpiDict = "PRIMER PISO SQL (" & fechaPantallaNorm & ")"
        strOrigenBD = "BD_INSUMO_PRIMER_PISO + INSTRUMENTOS"
        strRegistrosSQL = CStr(credBD) & " créditos vía SQL"
    ElseIf incCD And Not incBD And Not incPEA Then
        tagSeleccion = "Tesoreria_Y_Mercado_Dinero"
        tituloAlcance = "INFORME EJECUTIVO DE RIESGOS (SQL): 2. PORTAFOLIO DE TESORERÍA Y MERCADO DE DINERO (" & CStr(credCD) & " OPERACIONES)"
        kpiAct = "$" & Format(Round(finCD, 2), "#,##0.00") & " MDP"
        kpiRiesgo = CStr(credCD) & " Op. | $" & Format(Round(usdCD, 2), "#,##0.00") & " MDD"
        kpiRend = "$" & Format(Round(promCD, 2), "#,##0.00") & " MDP (Prom.)"
        kpiDict = "TESORERÍA SQL (" & fechaPantallaNorm & ")"
        strOrigenBD = "BD_INSUMO_TESORERIA + INSTRUMENTOS"
        strRegistrosSQL = CStr(credCD) & " operaciones vía SQL"
    ElseIf incPEA And Not incBD And Not incCD Then
        tagSeleccion = "Financiamiento_Sectorial_2do_Piso"
        tituloAlcance = "INFORME EJECUTIVO DE RIESGOS (SQL): 3. FINANCIAMIENTO SECTORIAL SEGUNDO PISO (" & CStr(credPEA) & " CRÉDITOS)"
        kpiAct = "$" & Format(Round(finPEA, 2), "#,##0.00") & " MDP"
        kpiRiesgo = CStr(credPEA) & " Créd. | $" & Format(Round(usdPEA, 2), "#,##0.00") & " MDD"
        kpiRend = "$" & Format(Round(promPEA, 2), "#,##0.00") & " MDP (Prom.)"
        kpiDict = "SEGUNDO PISO SQL (" & fechaPantallaNorm & ")"
        strOrigenBD = "BD_INSUMO_SEGUNDO_PISO + INSTRUMENTOS"
        strRegistrosSQL = CStr(credPEA) & " créditos vía SQL"
    Else
        tagSeleccion = "Seleccion_Parcial"
        tituloAlcance = "INFORME EJECUTIVO DE PORTAFOLIOS DE RIESGO SELECCIONADOS (CONSULTA SQL)"
        Dim credParc As Long, expParc As Double, promParc As Double, usdParc As Double
        credParc = IIf(incBD, credBD, 0) + IIf(incCD, credCD, 0) + IIf(incPEA, credPEA, 0)
        expParc = IIf(incBD, Round(finBD, 2), 0#) + IIf(incCD, Round(finCD, 2), 0#) + IIf(incPEA, Round(finPEA, 2), 0#)
        promParc = IIf(incBD, Round(promBD, 2), 0#) + IIf(incCD, Round(promCD, 2), 0#) + IIf(incPEA, Round(promPEA, 2), 0#)
        usdParc = IIf(incBD, Round(usdBD, 2), 0#) + IIf(incCD, Round(usdCD, 2), 0#) + IIf(incPEA, Round(usdPEA, 2), 0#)
        kpiAct = "$" & Format(expParc, "#,##0.00") & " MDP"
        kpiRiesgo = CStr(credParc) & " Op. | $" & Format(usdParc, "#,##0.00") & " MDD"
        kpiRend = "$" & Format(promParc, "#,##0.00") & " MDP (Prom.)"
        kpiDict = "EXTRAÍDO POR SQL (" & fechaPantallaNorm & ")"
        strOrigenBD = "RIESGOSBD (Bases Seleccionadas + INSTRUMENTOS)"
        strRegistrosSQL = CStr(credParc) & " operaciones vía SQL"
    End If

    strElaboro = "1000 - Wembley Ibarra (Dir. Economía de Hidrocarburos)"
    On Error Resume Next
    If Not usuarioSesionDTO Is Nothing Then
        strElaboro = usuarioSesionDTO.getCveUsuario() & " - " & usuarioSesionDTO.getNombre() & " " & usuarioSesionDTO.getAPaterno() & " (Dir. Economía de Hidrocarburos)"
    End If
    On Error GoTo ErrGen

    Set fso = CreateObject("Scripting.FileSystemObject")

    rutaDescargas = Trim(carpetaDestinoUsuario)
    If rutaDescargas <> "" Then
        If Right(rutaDescargas, 1) = "\" Then rutaDescargas = Left(rutaDescargas, Len(rutaDescargas) - 1)
        If Not fso.FolderExists(rutaDescargas) Then
            On Error Resume Next
            fso.CreateFolder rutaDescargas
            On Error GoTo ErrGen
        End If
    End If
    If rutaDescargas = "" Or Not fso.FolderExists(rutaDescargas) Then
        rutaDescargas = obtenerRutaDescargasPorDefectoBO()
    End If

    fechaCorteStr = Format(convertirFechaBO(fechaPantallaNorm), "yyyymmdd")

    strImgDir = CurrentProject.Path & "\Archivos\Imagenes"
    strLogoPath = strImgDir & "\LogoCorporativoEnergia.png"
    If Not fso.FileExists(strLogoPath) Then strLogoPath = strImgDir & "\LogoCorporativoEnergia.jpg"
    If Not fso.FileExists(strLogoPath) Then strLogoPath = CurrentProject.Path & "\LogoCorporativoEnergia.png"
    If Not fso.FileExists(strLogoPath) Then strLogoPath = CurrentProject.Path & "\LogoCorporativoEnergia.jpg"

    rutaImg1 = strImgDir & "\Grafica_Analisis_1.png"
    If Not fso.FileExists(rutaImg1) Then rutaImg1 = strImgDir & "\Grafica_Analisis_1.bmp"
    rutaImg2 = strImgDir & "\Grafica_Analisis_2.png"
    If Not fso.FileExists(rutaImg2) Then rutaImg2 = strImgDir & "\Grafica_Analisis_2.bmp"
    rutaImg3 = strImgDir & "\Grafica_Analisis_3.png"
    If Not fso.FileExists(rutaImg3) Then rutaImg3 = strImgDir & "\Grafica_Analisis_3.bmp"

    If UCase(formatoSalida) = "EXCEL" Then
        rutaSalida = rutaDescargas & "\Reporte_Institucional_" & tagSeleccion & "_" & fechaCorteStr & "_" & Format(Now, "hhmmss") & ".xlsx"
    Else
        rutaSalida = rutaDescargas & "\Reporte_Institucional_" & tagSeleccion & "_" & fechaCorteStr & "_" & Format(Now, "hhmmss") & ".pdf"
    End If

    If Not ObtenerExcelMotorBO(xlApp, cerrarXlRep) Then
        msgError = "No se pudo iniciar Microsoft Excel para construir el reporte."
        Exit Function
    End If
    On Error Resume Next
    xlApp.ScreenUpdating = False
    xlApp.EnableEvents = False
    xlApp.Calculation = -4135                  ' xlCalculationManual
    On Error GoTo ErrGen

    ' ------------------------------------------------------------------------
    ' EL REPORTE NACE DE UNA COPIA DEL TEMPLATE Generico.xlsx
    ' Se copia primero con FileCopy y se abre la COPIA: el template original
    ' nunca se abre, por lo que es imposible modificarlo por accidente.
    ' Si el template no estuviera disponible, se cae al libro en blanco de
    ' siempre para no dejar al usuario sin reporte, y se avisa en msgError.
    ' ------------------------------------------------------------------------
    bDesdeTemplate = False
    If UCase(formatoSalida) = "EXCEL" Then
        bDesdeTemplate = prepararLibroDesdeTemplateBO(rutaSalida, msgTpl)
    Else
        ' Para PDF la copia de trabajo se hace en la carpeta temporal del equipo.
        rutaTrabajoTpl = rutaDescargas & "\~tmp_" & Format(Now, "yyyymmddhhmmss") & "_" & TEMPLATE_GENERICO
        bDesdeTemplate = prepararLibroDesdeTemplateBO(rutaTrabajoTpl, msgTpl)
    End If

    If bDesdeTemplate Then
        If UCase(formatoSalida) = "EXCEL" Then
            Set wb = xlApp.Workbooks.Open(rutaSalida)
        Else
            Set wb = xlApp.Workbooks.Open(rutaTrabajoTpl)
        End If
        Set ws = wb.Worksheets(HOJA_TPL_REPORTE)
        On Error Resume Next
        ws.Cells.UnMerge
        ws.Cells.Clear
        On Error GoTo ErrGen
        Set wsCalc = wb.Worksheets(HOJA_TPL_CALCULO)
        Set wsCalcInstr = wb.Worksheets(HOJA_TPL_CALCINSTR)

        ' --------------------------------------------------------------------
        ' Los datos crudos del SQL se depositan en la hoja DATOS. De ahi en
        ' adelante todo lo numerico del reporte son formulas que leen esta hoja,
        ' por lo que no existen dos caminos de calculo que puedan discrepar.
        ' --------------------------------------------------------------------
        filaDatos = 2
        If incBD And nFilasBD > 0 Then
            Call volcarDatosCrudosEnTemplateBO(wb, TPL_PORTAFOLIO_1, rowsBD, nFilasBD, filaDatos)
        End If
        If incCD And nFilasCD > 0 Then
            Call volcarDatosCrudosEnTemplateBO(wb, TPL_PORTAFOLIO_2, rowsCD, nFilasCD, filaDatos)
        End If
        If incPEA And nFilasPEA > 0 Then
            Call volcarDatosCrudosEnTemplateBO(wb, TPL_PORTAFOLIO_3, rowsPEA, nFilasPEA, filaDatos)
        End If

        ' Lista de instrumentos que realmente aparecieron. Solo se escribe la
        ' columna A: las formulas de CALC_INSTR ya vienen en el template y se
        ' activan solas, por eso agregar o quitar instrumentos no obliga a
        ' agregar ni quitar formulas en ningun lado.
        nInstrSembrados = sembrarInstrumentosEnCalcInstrBO(wb)
    Else
        Set wb = xlApp.Workbooks.Add
        Set ws = wb.Worksheets(1)
        ws.Name = "Reporte_Institucional"
    End If

    On Error Resume Next
    xlApp.ActiveWindow.DisplayGridlines = False
    On Error GoTo ErrGen

    ' Columnas agrandadas para cubrir el ancho total de la hoja con las 8 columnas unificadas
    ws.Columns(1).ColumnWidth = 30.5
    ws.Columns(2).ColumnWidth = 26#
    ws.Columns(3).ColumnWidth = 17.5
    ws.Columns(4).ColumnWidth = 18.5
    ws.Columns(5).ColumnWidth = 18.5
    ws.Columns(6).ColumnWidth = 18#
    ws.Columns(7).ColumnWidth = 17#
    ws.Columns(8).ColumnWidth = 21#

    ' ==============================================================================================
    ' ENCABEZADO EJECUTIVO SIMETRICO DE DIRECCION GENERAL EN 3 LINEAS INDEPENDIENTES (A1:H2)
    ' ==============================================================================================
    ws.Range("A1:H2").Merge
    ws.Range("A1").Value = "DIRECCIÓN CORPORATIVA DE ECONOMÍA Y FISCALIZACIÓN DE HIDROCARBUROS" & vbLf & _
                           "DIRECCIÓN DE FISCALIZACIÓN Y CONTRAPRESTACIONES PETROLERAS" & vbLf & _
                           "GERENCIA DE VERIFICACIÓN DE VOLÚMENES Y REGALÍAS"
    With ws.Range("A1:H2")
        .Interior.Color = RGB(214, 223, 236)
        .Font.Name = "Cambria"
        .Font.Color = RGB(31, 78, 120)
        .Font.Bold = True
        .Font.Size = 11.5
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
        .WrapText = True
        .Borders.LineStyle = 1
        .Borders.Color = RGB(0, 78, 76)
    End With
    ws.Rows(1).RowHeight = 32
    ws.Rows(2).RowHeight = 32

    ws.Range("A3:H3").Merge
    ws.Range("A3").Value = tituloAlcance
    With ws.Range("A3:H3")
        .Interior.Color = RGB(0, 78, 76)
        .Font.Name = "Cambria"
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
    End With
    ws.Rows(3).RowHeight = 23

    ws.Range("F4:H4").Merge
    ws.Range("F5:H5").Merge

    Dim arrMeta(1 To 2, 1 To 6) As Variant
    arrMeta(1, 1) = "Fecha de Corte SQL:": arrMeta(1, 2) = fechaPantallaNorm: arrMeta(1, 3) = "Folio Institucional:": arrMeta(1, 4) = "UPST-SQL-" & fechaCorteStr: arrMeta(1, 5) = "Responsable:": arrMeta(1, 6) = strElaboro
    arrMeta(2, 1) = "Origen de Datos:": arrMeta(2, 2) = strOrigenBD: arrMeta(2, 3) = "Registros Extraídos:": arrMeta(2, 4) = strRegistrosSQL: arrMeta(2, 5) = "Órgano Colegiado:": arrMeta(2, 6) = "Comité Técnico de Verificación de Contraprestaciones"
    ws.Range("A4:F5").Value = arrMeta

    ws.Rows(4).RowHeight = 20.5
    ws.Rows(5).RowHeight = 20.5
    With ws.Range("A4:H5")
        .Font.Name = "Cambria"
        .Font.Size = 8.2
        .VerticalAlignment = -4108
        .ShrinkToFit = True
        .Borders.LineStyle = 1
        .Borders.Color = RGB(184, 196, 212)
        .Interior.Color = RGB(255, 255, 255)
    End With
    For c = 1 To 5 Step 2
        With ws.Range(ws.Cells(4, c), ws.Cells(5, c))
            .Interior.Color = RGB(238, 243, 249)
            .Font.Bold = True
            .Font.Color = RGB(31, 78, 120)
            .HorizontalAlignment = -4108
        End With
    Next c
    ws.Range("B4:B5").HorizontalAlignment = -4108
    ws.Range("D4:D5").HorizontalAlignment = -4108
    ws.Cells(4, 4).Font.Bold = True
    ws.Cells(4, 4).Font.Color = RGB(0, 78, 76)
    ws.Range("F4:H5").HorizontalAlignment = -4108

    ws.Rows(6).RowHeight = 8

    ' ==============================================================================================
    ' TABLERO EJECUTIVO DE INDICADORES CLAVE EXTRAIDOS POR SQL (FILAS 7 Y 8)
    ' ==============================================================================================
    ws.Range("A7:B7").Merge: ws.Range("A8:B8").Merge
    ws.Range("C7:D7").Merge: ws.Range("C8:D8").Merge
    ws.Range("E7:F7").Merge: ws.Range("E8:F8").Merge
    ws.Range("G7:H7").Merge: ws.Range("G8:H8").Merge

    ws.Cells(7, 1).Value = "SALDO FINAL VALUADO (SQL SDFIMN)": ws.Cells(8, 1).Value = kpiAct
    ws.Cells(7, 3).Value = "OPERACIONES Y EQUIVALENTE USD (SQL)": ws.Cells(8, 3).Value = kpiRiesgo
    ws.Cells(7, 5).Value = "SALDO PROMEDIO CONCILIADO (SQL SDPRMN)": ws.Cells(8, 5).Value = kpiRend
    ws.Cells(7, 7).Value = "TRAZABILIDAD DE BASE DE DATOS": ws.Cells(8, 7).Value = kpiDict

    ws.Rows(7).RowHeight = 17.5
    ws.Rows(8).RowHeight = 23.5
    With ws.Range("A7:H7")
        .Interior.Color = RGB(227, 236, 247)
        .Font.Name = "Cambria"
        .Font.Size = 8
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 120)
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
        .Borders.LineStyle = 1
        .Borders.Color = RGB(184, 196, 212)
    End With
    With ws.Range("A8:H8")
        .Interior.Color = RGB(255, 255, 255)
        .Font.Name = "Cambria"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(0, 78, 76)
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
        .Borders.LineStyle = 1
        .Borders.Color = RGB(184, 196, 212)
    End With

    ws.Rows(9).RowHeight = 9
    curRow = 10
    bloquesRenderizados = 0

    ' ==============================================================================================
    ' ANALISIS 1 (SQL): CARTERA COMERCIAL 1ER PISO EXTRAIDA DE BD_INSUMO_PRIMER_PISO + INSTRUMENTOS
    ' ==============================================================================================
    If incBD And nFilasBD > 0 Then
        bloquesRenderizados = bloquesRenderizados + 1
        ws.Range("A" & curRow & ":H" & curRow).Merge
        ws.Cells(curRow, 1).Value = "  ANÁLISIS SQL 1 — CARTERA DE CRÉDITO COMERCIAL (1ER PISO): BD_INSUMO_PRIMER_PISO + INSTRUMENTOS"
        With ws.Range("A" & curRow & ":H" & curRow)
            .Interior.Color = RGB(0, 78, 76)
            .Font.Name = "Cambria"
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .Font.Size = 9.5
            .HorizontalAlignment = -4131
            .VerticalAlignment = -4108
        End With
        ws.Rows(curRow).RowHeight = 22
        curRow = curRow + 1

        Dim hdrBD As Variant
        hdrBD = Array("Concepto de Cuenta (SQL)", "Moneda / Instrumento" & vbLf & "Asociado (SQL)", "Núm." & vbLf & "Créditos", "Saldo Inicial" & vbLf & "($ MDP)", "Saldo Final MN" & vbLf & "($ MDP)", "Saldo Promedio" & vbLf & "($ MDP)", "Equivalente USD" & vbLf & "($ MDD)", "% Part. /" & vbLf & "Estatus NORMATIVA")

        filaHdr = curRow: filaIni = curRow + 1: filaFin = curRow + nFilasBD
        dat1Ini = filaIni: dat1Fin = filaFin
        Call FormatearTablaEjecutivaBO(ws, filaHdr, filaIni, filaFin, hdrBD, rowsBD, modoIndividual, IIf(bDesdeTemplate, TPL_PORTAFOLIO_1, ""))
        ws.Range("C" & filaIni & ":C" & filaFin).NumberFormat = "#,##0"
        ws.Range("D" & filaIni & ":G" & filaFin).NumberFormat = "$#,##0.00"
        curRow = filaFin + 1

        Call PrepararMarcoDictamenYGraficaBO(ws, curRow, _
            "Dictamen Técnico SQL (1. Cartera Comercial Primer Piso — Consulta sqlReporteAnalisis1PrimerPiso): Se extrajeron " & CStr(credBD) & " créditos desde BD_INSUMO_PRIMER_PISO + INSTRUMENTOS (Fecha: " & fechaPantallaNorm & ") con Saldo Inicial de $" & Format(Round(iniBD, 2), "#,##0.00") & " MDP, Saldo Final Valuado de $" & Format(Round(finBD, 2), "#,##0.00") & " MDP, Saldo Promedio de $" & Format(Round(promBD, 2), "#,##0.00") & " MDP y Equivalente en Divisas de $" & Format(Round(usdBD, 2), "#,##0.00") & " MDD.", _
            modoIndividual, img1Ini, img1Fin)
        pagImgIni(numPags) = img1Ini: pagImgFin(numPags) = img1Fin
    End If

    ' ==============================================================================================
    ' ANALISIS 2 (SQL): PORTAFOLIO DE TESORERIA Y MERCADO DE DINERO EXTRAIDO DE BD_INSUMO_TESORERIA
    ' ==============================================================================================
    If incCD And nFilasCD > 0 Then
        If bloquesRenderizados >= 1 Then
            pagEnd(numPags) = curRow - 1
            On Error Resume Next
            ws.HPageBreaks.Add Before:=ws.Rows(curRow)
            On Error GoTo ErrGen
            numPags = numPags + 1
            pagStart(numPags) = curRow
        End If
        bloquesRenderizados = bloquesRenderizados + 1

        ws.Range("A" & curRow & ":H" & curRow).Merge
        ws.Cells(curRow, 1).Value = "  ANÁLISIS SQL 2 — PORTAFOLIO DE TESORERÍA Y MERCADO DE DINERO: BD_INSUMO_TESORERIA + INSTRUMENTOS"
        With ws.Range("A" & curRow & ":H" & curRow)
            .Interior.Color = RGB(31, 78, 120)
            .Font.Name = "Cambria"
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .Font.Size = 9.5
            .HorizontalAlignment = -4131
            .VerticalAlignment = -4108
        End With
        ws.Rows(curRow).RowHeight = 22
        curRow = curRow + 1

        Dim hdrCD As Variant
        hdrCD = Array("Grupo y Subgrupo Institucional (SQL)", "Moneda / Instrumento" & vbLf & "Asociado (SQL)", "Núm." & vbLf & "Operaciones", "Saldo Inicial" & vbLf & "($ MDP)", "Saldo Final MN" & vbLf & "($ MDP)", "Saldo Promedio" & vbLf & "($ MDP)", "Equivalente USD" & vbLf & "($ MDD)", "% Part. /" & vbLf & "Estatus NORMATIVA")

        filaHdr = curRow: filaIni = curRow + 1: filaFin = curRow + nFilasCD
        dat2Ini = filaIni: dat2Fin = filaFin
        Call FormatearTablaEjecutivaBO(ws, filaHdr, filaIni, filaFin, hdrCD, rowsCD, modoIndividual, IIf(bDesdeTemplate, TPL_PORTAFOLIO_2, ""))
        ws.Range("C" & filaIni & ":C" & filaFin).NumberFormat = "#,##0"
        ws.Range("D" & filaIni & ":G" & filaFin).NumberFormat = "$#,##0.00"
        curRow = filaFin + 1

        Call PrepararMarcoDictamenYGraficaBO(ws, curRow, _
            "Dictamen Técnico SQL (2. Tesorería y Mercado de Dinero — Consulta sqlReporteAnalisis2SectoresMonedas): Se extrajeron " & CStr(credCD) & " operaciones desde BD_INSUMO_TESORERIA + INSTRUMENTOS (Fecha: " & fechaPantallaNorm & ") con Saldo Inicial de $" & Format(Round(iniCD, 2), "#,##0.00") & " MDP, Saldo Final Valuado de $" & Format(Round(finCD, 2), "#,##0.00") & " MDP, Saldo Promedio de $" & Format(Round(promCD, 2), "#,##0.00") & " MDP y Equivalente en Divisas de $" & Format(Round(usdCD, 2), "#,##0.00") & " MDD.", _
            modoIndividual, img2Ini, img2Fin)
        pagImgIni(numPags) = img2Ini: pagImgFin(numPags) = img2Fin
    End If

    ' ==============================================================================================
    ' ANALISIS 3 (SQL): FINANCIAMIENTO SECTORIAL 2DO PISO EXTRAIDO DE BD_INSUMO_SEGUNDO_PISO
    ' ==============================================================================================
    If incPEA And nFilasPEA > 0 Then
        If bloquesRenderizados >= 1 Then
            pagEnd(numPags) = curRow - 1
            On Error Resume Next
            ws.HPageBreaks.Add Before:=ws.Rows(curRow)
            On Error GoTo ErrGen
            numPags = numPags + 1
            pagStart(numPags) = curRow
        End If
        bloquesRenderizados = bloquesRenderizados + 1

        ws.Range("A" & curRow & ":H" & curRow).Merge
        ws.Cells(curRow, 1).Value = "  ANÁLISIS SQL 3 — FINANCIAMIENTO SECTORIAL Y GARANTÍAS (2DO PISO): BD_INSUMO_SEGUNDO_PISO + INSTRUMENTOS"
        With ws.Range("A" & curRow & ":H" & curRow)
            .Interior.Color = RGB(0, 110, 106)
            .Font.Name = "Cambria"
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .Font.Size = 9.5
            .HorizontalAlignment = -4131
            .VerticalAlignment = -4108
        End With
        ws.Rows(curRow).RowHeight = 22
        curRow = curRow + 1

        Dim hdrPEA As Variant
        hdrPEA = Array("Subgrupo y Actividad Económica (SQL)", "Moneda / Instrumento" & vbLf & "Asociado (SQL)", "Núm." & vbLf & "Créditos", "Saldo Inicial" & vbLf & "($ MDP)", "Saldo Final MN" & vbLf & "($ MDP)", "Saldo Promedio" & vbLf & "($ MDP)", "Equivalente USD" & vbLf & "($ MDD)", "% Part. /" & vbLf & "Etapa NORMATIVA")

        filaHdr = curRow: filaIni = curRow + 1: filaFin = curRow + nFilasPEA
        dat3Ini = filaIni: dat3Fin = filaFin
        Call FormatearTablaEjecutivaBO(ws, filaHdr, filaIni, filaFin, hdrPEA, rowsPEA, modoIndividual, IIf(bDesdeTemplate, TPL_PORTAFOLIO_3, ""))
        ws.Range("C" & filaIni & ":C" & filaFin).NumberFormat = "#,##0"
        ws.Range("D" & filaIni & ":G" & filaFin).NumberFormat = "$#,##0.00"
        curRow = filaFin + 1

        Call PrepararMarcoDictamenYGraficaBO(ws, curRow, _
            "Dictamen Técnico SQL (3. Financiamiento Sectorial Segundo Piso — Consulta sqlReporteAnalisis3SegundoPiso): Se extrajeron " & CStr(credPEA) & " créditos desde BD_INSUMO_SEGUNDO_PISO + INSTRUMENTOS (Fecha: " & fechaPantallaNorm & ") con Saldo Inicial de $" & Format(Round(iniPEA, 2), "#,##0.00") & " MDP, Saldo Final Valuado de $" & Format(Round(finPEA, 2), "#,##0.00") & " MDP, Saldo Promedio de $" & Format(Round(promPEA, 2), "#,##0.00") & " MDP y Equivalente en Divisas de $" & Format(Round(usdPEA, 2), "#,##0.00") & " MDD.", _
            modoIndividual, img3Ini, img3Fin)
        pagImgIni(numPags) = img3Ini: pagImgFin(numPags) = img3Fin
    End If

    ' ==============================================================================================
    ' SECCION INFERIOR UNIFICADA (8 COLUMNAS IDENTICAS EN INDIVIDUALES Y EN CONSOLIDADO)
    ' ==============================================================================================
    Dim hdrCons As Variant
    hdrCons = Array("Portafolio Institucional Extraído por SQL", "Base de Datos /" & vbLf & "Instrumentos (SQL)", "Núm." & vbLf & "Operaciones", "Saldo Inicial" & vbLf & "($ MDP)", "Saldo Final MN" & vbLf & "($ MDP)", "Saldo Promedio" & vbLf & "($ MDP)", "Equivalente USD" & vbLf & "($ MDD)", "% Part. /" & vbLf & "Dictamen SQL")

    If incConsolidado Or numSeleccionados > 1 Then
        ws.Range("A" & curRow & ":H" & curRow).Merge
        ws.Cells(curRow, 1).Value = "  RESUMEN EJECUTIVO CONSOLIDADO EXTRAÍDO POR SQL DESDE LAS 3 BASES DE RIESGOSBD.ACCDB"
        With ws.Range("A" & curRow & ":H" & curRow)
            .Interior.Color = RGB(31, 78, 120)
            .Font.Name = "Cambria"
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .Font.Size = 9.2
            .HorizontalAlignment = -4131
            .VerticalAlignment = -4108
        End With
        ws.Rows(curRow).RowHeight = 20
        curRow = curRow + 1

        Dim nCons As Long, idxC As Long
        nCons = 0
        If incBD Then nCons = nCons + 1
        If incCD Then nCons = nCons + 1
        If incPEA Then nCons = nCons + 1
        nCons = nCons + 1
        Dim rowsCons() As Variant
        ReDim rowsCons(1 To nCons, 1 To 8)
        idxC = 0

        Dim sCredC As Long, sIniC As Double, sFinC As Double, sPromC As Double, sUsdC As Double
        sCredC = 0: sIniC = 0: sFinC = 0: sPromC = 0: sUsdC = 0

        If incBD Then
            idxC = idxC + 1
            rowsCons(idxC, 1) = "1. Cartera de Crédito Comercial (Primer Piso)"
            rowsCons(idxC, 2) = "BD_INSUMO_PRIMER_PISO + INSTRUMENTOS"
            rowsCons(idxC, 3) = credBD
            rowsCons(idxC, 4) = Round(iniBD, 2)
            rowsCons(idxC, 5) = Round(finBD, 2)
            rowsCons(idxC, 6) = Round(promBD, 2)
            rowsCons(idxC, 7) = Round(usdBD, 2)
            rowsCons(idxC, 8) = IIf(totGlobalFinMDP > 0, Format((Round(finBD, 2) / totGlobalFinMDP) * 100#, "0.00") & "% | VIGENTE E1", "VIGENTE E1")
            sCredC = sCredC + credBD
            sIniC = sIniC + Round(iniBD, 2)
            sFinC = sFinC + Round(finBD, 2)
            sPromC = sPromC + Round(promBD, 2)
            sUsdC = sUsdC + Round(usdBD, 2)
        End If
        If incCD Then
            idxC = idxC + 1
            rowsCons(idxC, 1) = "2. Portafolio de Tesorería y Mercado de Dinero"
            rowsCons(idxC, 2) = "BD_INSUMO_TESORERIA + INSTRUMENTOS"
            rowsCons(idxC, 3) = credCD
            rowsCons(idxC, 4) = Round(iniCD, 2)
            rowsCons(idxC, 5) = Round(finCD, 2)
            rowsCons(idxC, 6) = Round(promCD, 2)
            rowsCons(idxC, 7) = Round(usdCD, 2)
            rowsCons(idxC, 8) = IIf(totGlobalFinMDP > 0, Format((Round(finCD, 2) / totGlobalFinMDP) * 100#, "0.00") & "% | EN LÍMITE NORMATIVA", "EN LÍMITE NORMATIVA")
            sCredC = sCredC + credCD
            sIniC = sIniC + Round(iniCD, 2)
            sFinC = sFinC + Round(finCD, 2)
            sPromC = sPromC + Round(promCD, 2)
            sUsdC = sUsdC + Round(usdCD, 2)
        End If
        If incPEA Then
            idxC = idxC + 1
            rowsCons(idxC, 1) = "3. Financiamiento Sectorial y Garantías (Segundo Piso)"
            rowsCons(idxC, 2) = "BD_INSUMO_SEGUNDO_PISO + INSTRUMENTOS"
            rowsCons(idxC, 3) = credPEA
            rowsCons(idxC, 4) = Round(iniPEA, 2)
            rowsCons(idxC, 5) = Round(finPEA, 2)
            rowsCons(idxC, 6) = Round(promPEA, 2)
            rowsCons(idxC, 7) = Round(usdPEA, 2)
            rowsCons(idxC, 8) = IIf(totGlobalFinMDP > 0, Format((Round(finPEA, 2) / totGlobalFinMDP) * 100#, "0.00") & "% | ETAPA E1 (A-1)", "ETAPA E1 (A-1)")
            sCredC = sCredC + credPEA
            sIniC = sIniC + Round(iniPEA, 2)
            sFinC = sFinC + Round(finPEA, 2)
            sPromC = sPromC + Round(promPEA, 2)
            sUsdC = sUsdC + Round(usdPEA, 2)
        End If

        idxC = idxC + 1
        rowsCons(idxC, 1) = "TOTAL CONSOLIDADO EXTRAÍDO POR SQL (" & fechaPantallaNorm & ")"
        rowsCons(idxC, 2) = "3 Bases Conciliadas (" & fechaPantallaNorm & ")"
        rowsCons(idxC, 3) = sCredC
        rowsCons(idxC, 4) = Round(sIniC, 2)
        rowsCons(idxC, 5) = Round(sFinC, 2)
        rowsCons(idxC, 6) = Round(sPromC, 2)
        rowsCons(idxC, 7) = Round(sUsdC, 2)
        rowsCons(idxC, 8) = "100.00% CONCILIADO"

        filaHdr = curRow: filaIni = curRow + 1: filaFin = curRow + nCons
        Call FormatearTablaEjecutivaBO(ws, filaHdr, filaIni, filaFin, hdrCons, rowsCons, False)
        If bDesdeTemplate Then
            On Error Resume Next
            Dim rCons As Long, idxPortCalc As Long
            idxPortCalc = 0
            For rCons = filaIni To filaFin - 1
                If Left(CStr(Nz(ws.Cells(rCons, 1).Value, "")), 2) = "1." Then idxPortCalc = 2
                If Left(CStr(Nz(ws.Cells(rCons, 1).Value, "")), 2) = "2." Then idxPortCalc = 3
                If Left(CStr(Nz(ws.Cells(rCons, 1).Value, "")), 2) = "3." Then idxPortCalc = 4
                If idxPortCalc >= 2 Then
                    ws.Cells(rCons, 3).Formula = "=CALCULO!B" & idxPortCalc
                    ws.Cells(rCons, 4).Formula = "=CALCULO!C" & idxPortCalc
                    ws.Cells(rCons, 5).Formula = "=CALCULO!D" & idxPortCalc
                    ws.Cells(rCons, 6).Formula = "=CALCULO!E" & idxPortCalc
                    ws.Cells(rCons, 7).Formula = "=CALCULO!F" & idxPortCalc
                End If
            Next rCons
            ws.Cells(filaFin, 3).Formula = "=SUM(C" & filaIni & ":C" & (filaFin - 1) & ")"
            ws.Cells(filaFin, 4).Formula = "=SUM(D" & filaIni & ":D" & (filaFin - 1) & ")"
            ws.Cells(filaFin, 5).Formula = "=SUM(E" & filaIni & ":E" & (filaFin - 1) & ")"
            ws.Cells(filaFin, 6).Formula = "=SUM(F" & filaIni & ":F" & (filaFin - 1) & ")"
            ws.Cells(filaFin, 7).Formula = "=SUM(G" & filaIni & ":G" & (filaFin - 1) & ")"
            On Error GoTo ErrGen
        End If
        ws.Range("C" & filaIni & ":C" & filaFin).NumberFormat = "#,##0"
        ws.Range("D" & filaIni & ":G" & filaFin).NumberFormat = "$#,##0.00"
        ws.Range("H" & filaIni & ":H" & filaFin).HorizontalAlignment = -4108
        curRow = filaFin + 1
        ws.Rows(curRow).RowHeight = 7
        curRow = curRow + 1
    Else
        ws.Range("A" & curRow & ":H" & curRow).Merge
        ws.Cells(curRow, 1).Value = "  RESUMEN DE CONCILIACIÓN SQL Y TRAZABILIDAD DEL PORTAFOLIO (COMITÉ TÉCNICO DE CONTRAPRESTACIONES)"
        With ws.Range("A" & curRow & ":H" & curRow)
            .Interior.Color = RGB(0, 78, 76)
            .Font.Name = "Cambria"
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .Font.Size = 9
            .HorizontalAlignment = -4131
            .VerticalAlignment = -4108
        End With
        ws.Rows(curRow).RowHeight = 22
        curRow = curRow + 1

        Dim rowsInd(1 To 2, 1 To 8) As Variant
        If incBD Then
            rowsInd(1, 1) = "1. Cartera de Crédito Comercial (Primer Piso)"
            rowsInd(1, 2) = "BD_INSUMO_PRIMER_PISO + INSTRUMENTOS"
            rowsInd(1, 3) = credBD
            rowsInd(1, 4) = Round(iniBD, 2)
            rowsInd(1, 5) = Round(finBD, 2)
            rowsInd(1, 6) = Round(promBD, 2)
            rowsInd(1, 7) = Round(usdBD, 2)
            rowsInd(1, 8) = IIf(totGlobalFinMDP > 0, Format((Round(finBD, 2) / totGlobalFinMDP) * 100#, "0.00") & "% del Total Global", "100.00% VIGENTE")
        ElseIf incCD Then
            rowsInd(1, 1) = "2. Portafolio de Tesorería y Mercado de Dinero"
            rowsInd(1, 2) = "BD_INSUMO_TESORERIA + INSTRUMENTOS"
            rowsInd(1, 3) = credCD
            rowsInd(1, 4) = Round(iniCD, 2)
            rowsInd(1, 5) = Round(finCD, 2)
            rowsInd(1, 6) = Round(promCD, 2)
            rowsInd(1, 7) = Round(usdCD, 2)
            rowsInd(1, 8) = IIf(totGlobalFinMDP > 0, Format((Round(finCD, 2) / totGlobalFinMDP) * 100#, "0.00") & "% del Total Global", "100.00% EN LÍMITE")
        Else
            rowsInd(1, 1) = "3. Financiamiento Sectorial y Garantías (Segundo Piso)"
            rowsInd(1, 2) = "BD_INSUMO_SEGUNDO_PISO + INSTRUMENTOS"
            rowsInd(1, 3) = credPEA
            rowsInd(1, 4) = Round(iniPEA, 2)
            rowsInd(1, 5) = Round(finPEA, 2)
            rowsInd(1, 6) = Round(promPEA, 2)
            rowsInd(1, 7) = Round(usdPEA, 2)
            rowsInd(1, 8) = IIf(totGlobalFinMDP > 0, Format((Round(finPEA, 2) / totGlobalFinMDP) * 100#, "0.00") & "% del Total Global", "100.00% ETAPA E1")
        End If
        rowsInd(2, 1) = "REFERENCIA TOTAL CONSOLIDADO (3 BASES SQL " & fechaPantallaNorm & ")"
        rowsInd(2, 2) = "RIESGOSBD (3 Bases Conciliadas)"
        rowsInd(2, 3) = totGlobalCred
        rowsInd(2, 4) = Round(totGlobalIniMDP, 2)
        rowsInd(2, 5) = Round(totGlobalFinMDP, 2)
        rowsInd(2, 6) = Round(totGlobalPromMDP, 2)
        rowsInd(2, 7) = Round(totGlobalFinMDD, 2)
        rowsInd(2, 8) = "100.00% CONCILIADO"

        filaHdr = curRow: filaIni = curRow + 1: filaFin = curRow + 2
        Call FormatearTablaEjecutivaBO(ws, filaHdr, filaIni, filaFin, hdrCons, rowsInd, False)
        ws.Range("C" & filaIni & ":C" & filaFin).NumberFormat = "#,##0"
        ws.Range("D" & filaIni & ":G" & filaFin).NumberFormat = "$#,##0.00"
        ws.Range("H" & filaIni & ":H" & filaFin).HorizontalAlignment = -4108
        curRow = filaFin + 1
        ws.Rows(curRow).RowHeight = 8
        curRow = curRow + 1

        ws.Range("A" & curRow & ":H" & curRow).Merge
        ws.Cells(curRow, 1).Value = "  NOTA METODOLÓGICA DE EXTRACCIÓN SQL Y CONCILIACIÓN CON REPORTE CONSOLIDADO (RIESGOSBD.ACCDB)"
        With ws.Range("A" & curRow & ":H" & curRow)
            .Interior.Color = RGB(31, 78, 120)
            .Font.Name = "Cambria"
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .Font.Size = 8.8
            .HorizontalAlignment = -4131
            .VerticalAlignment = -4108
        End With
        ws.Rows(curRow).RowHeight = 20.5
        curRow = curRow + 1

        ws.Range("A" & curRow & ":H" & (curRow + 1)).Merge
        ws.Cells(curRow, 1).Value = "1. Todas las cifras de este reporte individual fueron extraídas mediante consultas SQL (ConsultasSQL.bas / FondoDAO.cls) desde " & strOrigenBD & " (" & strRegistrosSQL & ") para la Fecha de Corte " & fechaPantallaNorm & "." & vbLf & _
                                    "2. Las 8 columnas y totales de este reporte coinciden exactamente al 100% con su sección y renglón correspondiente dentro del Reporte Ejecutivo Consolidado."
        With ws.Range("A" & curRow & ":H" & (curRow + 1))
            .Interior.Color = RGB(245, 248, 252)
            .Font.Name = "Cambria"
            .Font.Size = 8.4
            .Font.Color = RGB(30, 41, 59)
            .HorizontalAlignment = -4131
            .VerticalAlignment = -4108
            .WrapText = True
            .Borders.LineStyle = 1
            .Borders.Color = RGB(184, 196, 212)
        End With
        ws.Rows(curRow).RowHeight = 20.5
        ws.Rows(curRow + 1).RowHeight = 20.5
        curRow = curRow + 2
        ws.Rows(curRow).RowHeight = 8
        curRow = curRow + 1
    End If

    ' Bloque ejecutivo de firmas, autorización y pie de página institucional
    Call RenderizarFirmasYPieInstitucionalBO(ws, curRow, strElaboro, strOrigenBD)
    pagEnd(numPags) = curRow

    For pIdx = 1 To numPags
        Call CalibrarRangoVerticalCartaBO(ws, pagStart(pIdx), pagEnd(pIdx), pagImgIni(pIdx), pagImgFin(pIdx), (pIdx > 1))
    Next pIdx

    If fso.FileExists(strLogoPath) Then
        hBoxLogo = ws.Rows(1).Height + ws.Rows(2).Height
        hLogo = hBoxLogo - 8
        wLogo = hLogo * 1.4306
        ws.Shapes.AddPicture strLogoPath, False, True, _
            ws.Columns(1).Left + 14, _
            ws.Rows(1).Top + ((hBoxLogo - hLogo) / 2), _
            wLogo, hLogo
    End If

    ' ------------------------------------------------------------------------
    ' RECALCULO. Hasta aqui el calculo estuvo en manual para que Excel no
    ' recalculara en cada escritura. Ahora se devuelve a automatico y se fuerza
    ' un recalculo completo, de modo que las formulas del template YA tengan su
    ' resultado cuando las graficas los lean. Las graficas leen la columna E de
    ' cada tabla, que a partir de esta version es una formula SUMIFS: por eso la
    ' grafica se dibuja sobre el resultado del calculo y no sobre el dato crudo.
    ' ------------------------------------------------------------------------
    If bDesdeTemplate Then
        Call recalcularTemplateBO(xlApp, wb)

        ' AUTOCONCILIACION. Ahora hay dos caminos para el mismo numero: el que
        ' calculo el VBA al leer el SQL (totGlobalFinMDP) y el que calculo EXCEL
        ' con las formulas de la hoja CALCULO. Tienen que coincidir. Si no, algo
        ' esta mal y conviene que se sepa en vez de entregar un reporte cuadrado
        ' por casualidad. No se detiene la generacion: se deja constancia.
        On Error Resume Next
        If Not wsCalc Is Nothing Then
            dblCalcExcel = CDbl(Val(Nz(wsCalc.Cells(TPL_CALC_FILA_TOT, 4).Value, 0)))
            dblCalcVBA = 0
            If incBD Then dblCalcVBA = dblCalcVBA + Round(finBD, 2)
            If incCD Then dblCalcVBA = dblCalcVBA + Round(finCD, 2)
            If incPEA Then dblCalcVBA = dblCalcVBA + Round(finPEA, 2)
            If Abs(dblCalcExcel - dblCalcVBA) > 0.011 Then
                gUltimoErrorSQL = "DESCUADRE en el reporte del " & fechaPantallaNorm & _
                    ": Excel calculo " & Format(dblCalcExcel, "#,##0.00") & " MDP y el modulo " & _
                    Format(dblCalcVBA, "#,##0.00") & " MDP (diferencia " & _
                    Format(Abs(dblCalcExcel - dblCalcVBA), "#,##0.00") & " MDP)."
            End If
        End If
        Err.Clear
        On Error GoTo ErrGen
    End If

    If incBD And img1Ini > 0 Then Call ColocarGraficaDatosBO(ws, fso, img1Ini, img1Fin, _
        dat1Ini, IIf(dat1Fin > dat1Ini, dat1Fin - 1, dat1Fin), "1. Cartera Comercial Primer Piso — Saldo Final M.N. por concepto ($ MDP)", rutaImg1, "BARRA")
    If incCD And img2Ini > 0 Then Call ColocarGraficaDatosBO(ws, fso, img2Ini, img2Fin, _
        dat2Ini, IIf(dat2Fin > dat2Ini, dat2Fin - 1, dat2Fin), "2. Tesorería y Mercado de Dinero — Posición Final por grupo ($ MDP)", rutaImg2, "ANILLO")
    If incPEA And img3Ini > 0 Then Call ColocarGraficaDatosBO(ws, fso, img3Ini, img3Fin, _
        dat3Ini, IIf(dat3Fin > dat3Ini, dat3Fin - 1, dat3Fin), "3. Financiamiento Sectorial Segundo Piso — Saldo Vigente por segmento ($ MDP)", rutaImg3, "COLUMNA")

    ' RENDIMIENTO: Un solo bloque con PrintCommunication = False evita consultar el spooler
    ' de impresora de Windows 14 veces seguidas (ahorra ~1.8 segundos por reporte).
    On Error Resume Next
    xlApp.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = "$A$1:$H$" & curRow
        .PrintTitleRows = "$1:$5"
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .PaperSize = 1
        .Orientation = 1
        .Zoom = False
        .FitToPagesWide = 1
        If modoIndividual Then
            .FitToPagesTall = 1
        Else
            .FitToPagesTall = numPags
        End If
        .LeftMargin = xlApp.InchesToPoints(0.03)
        .RightMargin = xlApp.InchesToPoints(0.03)
        .TopMargin = xlApp.InchesToPoints(0.03)
        .BottomMargin = xlApp.InchesToPoints(0.03)
        .HeaderMargin = xlApp.InchesToPoints(0.02)
        .FooterMargin = xlApp.InchesToPoints(0.02)
        .CenterHorizontally = True
        .CenterVertically = True
    End With
    xlApp.PrintCommunication = True
    On Error GoTo ErrGen

    If UCase(formatoSalida) = "EXCEL" Then
        If bDesdeTemplate Then
            On Error Resume Next
            wb.Save
            If Err.Number <> 0 Then
                Err.Clear
                wb.SaveAs rutaSalida, 51
            End If
            Err.Clear
            On Error GoTo ErrGen
        Else
            wb.SaveAs rutaSalida, 51
        End If
        wb.Close False
        Set ws = Nothing: Set wb = Nothing
        If cerrarXlRep Then xlApp.Quit: Set xlApp = Nothing
        generarReporteEjecutivoBO = fso.FileExists(rutaSalida)
        Exit Function
    ElseIf UCase(formatoSalida) = "PDF" Then
        On Error Resume Next
        ws.ExportAsFixedFormat 0, rutaSalida, 0, True, False
        If Not fso.FileExists(rutaSalida) Then
            wb.ExportAsFixedFormat 0, rutaSalida, 0, True, False
        End If
        wb.Close False
        Set ws = Nothing: Set wb = Nothing
        If cerrarXlRep Then xlApp.Quit: Set xlApp = Nothing

        ' La copia de trabajo del template ya cumplio su funcion: se elimina
        ' para no dejar basura en la carpeta de destino del usuario.
        If bDesdeTemplate And rutaTrabajoTpl <> "" Then
            On Error Resume Next
            If fso.FileExists(rutaTrabajoTpl) Then fso.DeleteFile rutaTrabajoTpl, True
            Err.Clear
            On Error GoTo ErrGen
        End If

        If Not fso.FileExists(rutaSalida) Then
            DoCmd.OpenReport "Rpt_Analisis_Integral_Riesgos", 1
            Reports("Rpt_Analisis_Integral_Riesgos").Controls("lblTitA1").Visible = incBD
            Reports("Rpt_Analisis_Integral_Riesgos").Controls("lblTxtA1").Visible = incBD
            Reports("Rpt_Analisis_Integral_Riesgos").Controls("imgA1").Visible = incBD
            Reports("Rpt_Analisis_Integral_Riesgos").Controls("lblTitA2").Visible = incCD
            Reports("Rpt_Analisis_Integral_Riesgos").Controls("lblTxtA2").Visible = incCD
            Reports("Rpt_Analisis_Integral_Riesgos").Controls("imgA2").Visible = incCD
            Reports("Rpt_Analisis_Integral_Riesgos").Controls("lblTitA3").Visible = incPEA
            Reports("Rpt_Analisis_Integral_Riesgos").Controls("lblTxtA3").Visible = incPEA
            Reports("Rpt_Analisis_Integral_Riesgos").Controls("imgA3").Visible = incPEA
            DoCmd.Close 3, "Rpt_Analisis_Integral_Riesgos", 1
            DoCmd.OutputTo 3, "Rpt_Analisis_Integral_Riesgos", "PDFFormat(*.pdf)", rutaSalida, False
        End If
        On Error GoTo ErrGen
        generarReporteEjecutivoBO = fso.FileExists(rutaSalida)
        Exit Function
    End If
    Exit Function

ErrGen:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If cerrarXlRep And Not xlApp Is Nothing Then xlApp.Quit
    On Error Resume Next
    If rutaTrabajoTpl <> "" Then
        If fso.FileExists(rutaTrabajoTpl) Then fso.DeleteFile rutaTrabajoTpl, True
    End If
    Err.Clear
    msgError = Err.Description
    generarReporteEjecutivoBO = False
End Function

' =============================================================================
' REPORTE COMPARATIVO ENTRE INSUMOS (CRUCE DE DOS Y DE TRES BASES)
' -----------------------------------------------------------------------------
' nInsumos = 2 -> cruza Primer Piso con Tesoreria
' nInsumos = 3 -> cruza los tres libros
'
' Es un reporte distinto a los tres individuales y al consolidado: no suma
' portafolios, los CRUZA por instrumento. Responde que instrumento aparece en
' cual libro, con que saldo en cada uno, y cuales aparecen en uno solo, que es
' donde suelen estar los errores de clasificacion.
'
' Igual que el reporte ejecutivo, nace de una COPIA de Generico.xlsx, y la fila
' de totales y las columnas derivadas son FORMULAS de Excel, no numeros que el
' VBA haya calculado: el usuario puede abrir el archivo, cambiar un dato y ver
' el total y la grafica actualizarse.
' =============================================================================
Public Function generarReporteComparativoBO(ByVal fechaCorte As String, ByVal formatoSalida As String, _
        ByVal nInsumos As Integer, ByRef rutaSalida As String, ByRef msgError As String, _
        Optional ByVal carpetaDestinoUsuario As String = "") As Boolean
    Dim fso As Object, dao As FondoDAO, xlApp As Object, wb As Object, ws As Object
    Dim rows As Variant, nFilas As Long, totOps As Long, totFin As Double, msgCmp As String
    Dim fecNorm As String, rutaDescargas As String, fechaCorteStr As String, tagSel As String
    Dim bTpl As Boolean, msgTpl As String, rutaTmp As String
    Dim hdr As Variant, filaHdr As Long, filaIni As Long, filaFin As Long, filaTot As Long
    Dim strLogoPath As String, strImgDir As String, curRow As Long, k As Integer
    Dim colTot As String, tituloGraf As String, nReq As Integer
    Dim nP1 As Long, nP2 As Long, nP3 As Long, faltan As String, tagFec As String
    Dim strElaboro As String

    On Error GoTo ErrCmp
    generarReporteComparativoBO = False
    msgError = ""
    If nInsumos < 2 Then nInsumos = 2
    If nInsumos > 3 Then nInsumos = 3

    If Not esFechaFormatoCorrectoBO(fechaCorte) Then
        msgError = "La fecha solicitada ('" & fechaCorte & "') no tiene un formato valido dd/mm/aaaa."
        Exit Function
    End If
    If esFechaFuturaBO(fechaCorte) Then
        msgError = "No se puede escoger fechas futuras."
        Exit Function
    End If
    fecNorm = normalizarFechaBO(fechaCorte)
    tagFec = Replace(fecNorm, "/", "_")

    ' Un cruce exige que TODOS los insumos que participan esten cargados: si
    ' falta uno, el cruce daria la falsa impresion de que ese libro esta vacio.
    Dim cerrarXlCmp As Boolean
    Call obtenerConteosTresTablasFechaBO(fecNorm, nP1, nP2, nP3)
    Set dao = New FondoDAO
    faltan = ""
    If nP1 <= 0 Then faltan = faltan & "  - Primer Piso: falta INSUMO_PRIMER_PISO_" & tagFec & ".xlsx" & vbCrLf
    If nP2 <= 0 Then faltan = faltan & "  - Tesoreria: falta INSUMO_TESORERIA_" & tagFec & ".xlsx" & vbCrLf
    If nInsumos = 3 And nP3 <= 0 Then faltan = faltan & "  - Segundo Piso: falta INSUMO_SEGUNDO_PISO_" & tagFec & ".xlsx" & vbCrLf
    If faltan <> "" Then
        msgError = "El reporte comparativo de " & CStr(nInsumos) & " insumos necesita que TODOS los " & _
                   "insumos que participan esten cargados para la fecha " & fecNorm & "." & vbCrLf & vbCrLf & _
                   faltan & vbCrLf & "Fechas con insumos cargados: " & dao.obtenerFechasCargadasDAO()
        Set dao = Nothing
        Exit Function
    End If

    If Not dao.extraerComparativoInsumosSQL(fecNorm, nInsumos, rows, nFilas, totOps, totFin, msgCmp) Then
        Set dao = Nothing
        msgError = IIf(msgCmp = "", "No fue posible ejecutar el cruce de insumos.", msgCmp)
        Exit Function
    End If
    Set dao = Nothing
    If nFilas <= 0 Then
        msgError = "El cruce no devolvio instrumentos para la fecha " & fecNorm & "."
        Exit Function
    End If

    strElaboro = "1000 - Direccion de Fiscalizacion y Contraprestaciones Petroleras"
    On Error Resume Next
    If Not usuarioSesionDTO Is Nothing Then
        strElaboro = usuarioSesionDTO.getCveUsuario() & " - " & usuarioSesionDTO.getNombre() & " " & _
                     usuarioSesionDTO.getAPaterno() & " (Dir. Economia de Hidrocarburos)"
    End If
    On Error GoTo ErrCmp

    Set fso = CreateObject("Scripting.FileSystemObject")
    rutaDescargas = Trim(carpetaDestinoUsuario)
    If rutaDescargas <> "" Then
        If Right(rutaDescargas, 1) = "\" Then rutaDescargas = Left(rutaDescargas, Len(rutaDescargas) - 1)
        If Not fso.FolderExists(rutaDescargas) Then
            On Error Resume Next
            fso.CreateFolder rutaDescargas
            On Error GoTo ErrCmp
        End If
    End If
    If rutaDescargas = "" Or Not fso.FolderExists(rutaDescargas) Then
        rutaDescargas = obtenerRutaDescargasPorDefectoBO()
    End If

    fechaCorteStr = Format(convertirFechaBO(fecNorm), "yyyymmdd")
    tagSel = IIf(nInsumos = 2, "Comparativo_2_Insumos", "Comparativo_3_Insumos")
    If UCase(formatoSalida) = "EXCEL" Then
        rutaSalida = rutaDescargas & "\Reporte_" & tagSel & "_" & fechaCorteStr & "_" & Format(Now, "hhmmss") & ".xlsx"
    Else
        rutaSalida = rutaDescargas & "\Reporte_" & tagSel & "_" & fechaCorteStr & "_" & Format(Now, "hhmmss") & ".pdf"
    End If

    If Not ObtenerExcelMotorBO(xlApp, cerrarXlCmp) Then
        msgError = "No se pudo iniciar Microsoft Excel para construir el reporte comparativo."
        Exit Function
    End If
    On Error Resume Next
    xlApp.ScreenUpdating = False
    xlApp.EnableEvents = False
    xlApp.Calculation = -4135                  ' manual mientras se escribe
    On Error GoTo ErrCmp

    rutaTmp = ""
    If UCase(formatoSalida) = "EXCEL" Then
        bTpl = prepararLibroDesdeTemplateBO(rutaSalida, msgTpl)
        If bTpl Then Set wb = xlApp.Workbooks.Open(rutaSalida)
    Else
        rutaTmp = rutaDescargas & "\~tmp_" & Format(Now, "yyyymmddhhmmss") & "_" & TEMPLATE_GENERICO
        bTpl = prepararLibroDesdeTemplateBO(rutaTmp, msgTpl)
        If bTpl Then Set wb = xlApp.Workbooks.Open(rutaTmp)
    End If
    If bTpl Then
        Set ws = wb.Worksheets(HOJA_TPL_REPORTE)
        On Error Resume Next
        ws.Cells.UnMerge
        ws.Cells.Clear
        ' Poblar tambien DATOS y CALC_INSTR en la copia del template con los 2 o 3 portafolios que participan en el cruce
        Dim rBD As Variant, nBD As Long, rCD As Variant, nCD As Long, rPEA As Variant, nPEA As Long
        Dim c1 As Long, i1 As Double, f1 As Double, p1 As Double, u1 As Double
        Dim c2 As Long, i2 As Double, f2 As Double, p2 As Double, u2 As Double
        Dim c3 As Long, i3 As Double, f3 As Double, p3 As Double, u3 As Double, fDatCmp As Long
        fDatCmp = 2
        Call obtenerAnalisisTresPortafoliosBO(fecNorm, rBD, nBD, c1, i1, f1, p1, u1, rCD, nCD, c2, i2, f2, p2, u2, rPEA, nPEA, c3, i3, f3, p3, u3)
        If nBD > 0 Then Call volcarDatosCrudosEnTemplateBO(wb, TPL_PORTAFOLIO_1, rBD, nBD, fDatCmp)
        If nCD > 0 Then Call volcarDatosCrudosEnTemplateBO(wb, TPL_PORTAFOLIO_2, rCD, nCD, fDatCmp)
        If nInsumos = 3 And nPEA > 0 Then Call volcarDatosCrudosEnTemplateBO(wb, TPL_PORTAFOLIO_3, rPEA, nPEA, fDatCmp)
        Call sembrarInstrumentosEnCalcInstrBO(wb)
        On Error GoTo ErrCmp
    Else
        Set wb = xlApp.Workbooks.Add
        Set ws = wb.Worksheets(1)
        ws.Name = "Comparativo_Insumos"
    End If

    On Error Resume Next
    xlApp.ActiveWindow.DisplayGridlines = False
    On Error GoTo ErrCmp

    ws.Columns(1).ColumnWidth = 24#
    ws.Columns(2).ColumnWidth = 30#
    ws.Columns(3).ColumnWidth = 15#
    ws.Columns(4).ColumnWidth = 19#
    ws.Columns(5).ColumnWidth = 19#
    ws.Columns(6).ColumnWidth = 19#
    ws.Columns(7).ColumnWidth = 20#
    ws.Columns(8).ColumnWidth = 19#

    Call EncabezadoInstitucionalExcelBO(ws, _
        IIf(nInsumos = 2, _
            "REPORTE COMPARATIVO (SQL CRUZADO): PRIMER PISO vs. TESORERIA POR INSTRUMENTO", _
            "REPORTE COMPARATIVO (SQL CRUZADO): PRIMER PISO vs. TESORERIA vs. SEGUNDO PISO POR INSTRUMENTO"), _
        fecNorm, fechaCorteStr, strElaboro, _
        IIf(nInsumos = 2, "RIESGOSBD (2 bases de insumo + INSTRUMENTOS)", _
                          "RIESGOSBD (3 bases de insumo + INSTRUMENTOS)"), _
        CStr(totOps) & " operaciones cruzadas", CStr(nFilas) & " instrumentos", _
        "$" & Format(totFin, "#,##0.00") & " MDP")

    curRow = 10
    ws.Range("A" & curRow & ":H" & curRow).Merge
    ws.Cells(curRow, 1).Value = IIf(nInsumos = 2, _
        "  CRUCE DE 2 INSUMOS - BD_INSUMO_PRIMER_PISO x BD_INSUMO_TESORERIA (llave: ID_INSTRUMENTO)", _
        "  CRUCE DE 3 INSUMOS - BD_INSUMO_PRIMER_PISO x BD_INSUMO_TESORERIA x BD_INSUMO_SEGUNDO_PISO (llave: ID_INSTRUMENTO)")
    With ws.Range("A" & curRow & ":H" & curRow)
        .Interior.Color = RGB(0, 78, 76)
        .Font.Name = "Cambria"
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 9.5
        .HorizontalAlignment = -4131
        .VerticalAlignment = -4108
    End With
    ws.Rows(curRow).RowHeight = 22
    curRow = curRow + 1

    If nInsumos = 2 Then
        hdr = Array("Instrumento", "Tipo de Instrumento" & vbLf & "(catalogo INSTRUMENTOS)", _
                    "Num." & vbLf & "Operaciones", "Saldo Final" & vbLf & "1er Piso ($ MDP)", _
                    "Saldo Final" & vbLf & "Tesoreria ($ MDP)", "Saldo Final" & vbLf & "Combinado ($ MDP)", _
                    "Equivalente USD" & vbLf & "($ MDD)", "Presencia" & vbLf & "en los libros")
        colTot = "F"
        tituloGraf = "Cruce de 2 insumos - Saldo final combinado por instrumento ($ MDP)"
    Else
        hdr = Array("Instrumento", "Tipo de Instrumento" & vbLf & "(catalogo INSTRUMENTOS)", _
                    "Num." & vbLf & "Operaciones", "Saldo Final" & vbLf & "1er Piso ($ MDP)", _
                    "Saldo Final" & vbLf & "Tesoreria ($ MDP)", "Saldo Final" & vbLf & "2do Piso ($ MDP)", _
                    "Saldo Final" & vbLf & "Consolidado ($ MDP)", "Presencia" & vbLf & "en los libros")
        colTot = "G"
        tituloGraf = "Cruce de 3 insumos - Saldo final consolidado por instrumento ($ MDP)"
    End If

    filaHdr = curRow
    filaIni = curRow + 1
    filaFin = curRow + nFilas
    Call FormatearTablaEjecutivaBO(ws, filaHdr, filaIni, filaFin, hdr, rows, False)
    ws.Range("C" & filaIni & ":C" & filaFin).NumberFormat = "#,##0"
    ws.Range("D" & filaIni & ":G" & filaFin).NumberFormat = "$#,##0.00"

    ' La columna de total por instrumento es una FORMULA: suma las columnas de
    ' cada libro. Si el usuario corrige un saldo, el total se recalcula solo.
    If nInsumos = 2 Then
        ws.Range("F" & filaIni & ":F" & filaFin).Formula = "=D" & filaIni & "+E" & filaIni
    Else
        ws.Range("G" & filaIni & ":G" & filaFin).Formula = "=D" & filaIni & "+E" & filaIni & "+F" & filaIni
    End If

    ' Fila de totales, tambien con formulas SUM sobre el rango de la tabla.
    filaTot = filaFin + 1
    ws.Cells(filaTot, 1).Value = "TOTAL DEL CRUCE (" & CStr(nFilas) & " instrumentos)"
    ws.Range("A" & filaTot & ":B" & filaTot).Merge
    For k = 3 To 7
        ws.Cells(filaTot, k).Formula = "=SUM(" & Chr(64 + k) & filaIni & ":" & Chr(64 + k) & filaFin & ")"
    Next k
    ws.Cells(filaTot, 8).Value = "CONCILIADO"
    ws.Cells(filaTot, 3).NumberFormat = "#,##0"
    ws.Range("D" & filaTot & ":G" & filaTot).NumberFormat = "$#,##0.00"
    With ws.Range("A" & filaTot & ":H" & filaTot)
        .Interior.Color = RGB(214, 223, 236)
        .Font.Name = "Cambria"
        .Font.Size = 9.5
        .Font.Bold = True
        .Font.Color = RGB(0, 78, 76)
        .VerticalAlignment = -4108
        .Borders.LineStyle = 1
        .Borders.Color = RGB(184, 196, 212)
    End With
    ws.Rows(filaTot).RowHeight = 22

    ' Recalcular ANTES de la grafica, para que la grafica lea el resultado.
    On Error Resume Next
    xlApp.Calculation = -4105
    xlApp.Calculate
    Err.Clear
    On Error GoTo ErrCmp

    ' Grafica sobre la columna de totales, que es la columna con formula.
    Call ColocarGraficaDatosBO(ws, fso, filaTot + 2, filaTot + 22, filaIni, filaFin, _
        tituloGraf, "", "COLUMNA", ws, "A", colTot)

    strImgDir = CurrentProject.Path & "\Archivos\Imagenes"
    strLogoPath = strImgDir & "\LogoCorporativoEnergia.png"
    If Not fso.FileExists(strLogoPath) Then strLogoPath = strImgDir & "\LogoCorporativoEnergia.jpg"
    If fso.FileExists(strLogoPath) Then
        On Error Resume Next
        ws.Shapes.AddPicture strLogoPath, False, True, ws.Columns(1).Left + 14, ws.Rows(1).Top + 4, _
            (ws.Rows(1).Height + ws.Rows(2).Height - 8) * 1.4306, (ws.Rows(1).Height + ws.Rows(2).Height - 8)
        Err.Clear
        On Error GoTo ErrCmp
    End If

    On Error Resume Next
    xlApp.PrintCommunication = False
    ws.PageSetup.Orientation = 2
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    xlApp.PrintCommunication = True
    Err.Clear
    On Error GoTo ErrCmp

    If UCase(formatoSalida) = "EXCEL" Then
        On Error Resume Next
        If bTpl Then wb.Save Else wb.SaveAs rutaSalida, 51
        If Err.Number <> 0 Then
            Err.Clear
            wb.SaveAs rutaSalida, 51
        End If
        Err.Clear
        On Error GoTo ErrCmp
        wb.Close False
        If cerrarXlCmp Then xlApp.Quit
    Else
        On Error Resume Next
        ws.ExportAsFixedFormat 0, rutaSalida, 0, True, False
        If Not fso.FileExists(rutaSalida) Then wb.ExportAsFixedFormat 0, rutaSalida, 0, True, False
        Err.Clear
        On Error GoTo ErrCmp
        wb.Close False
        If cerrarXlCmp Then xlApp.Quit
        On Error Resume Next
        If rutaTmp <> "" Then
            If fso.FileExists(rutaTmp) Then fso.DeleteFile rutaTmp, True
        End If
        Err.Clear
        On Error GoTo ErrCmp
    End If
    Set ws = Nothing: Set wb = Nothing
    If cerrarXlCmp Then Set xlApp = Nothing

    generarReporteComparativoBO = fso.FileExists(rutaSalida)
    If Not generarReporteComparativoBO Then
        msgError = "No se pudo escribir el archivo del reporte comparativo en: " & rutaSalida
    End If
    Exit Function

ErrCmp:
    msgError = "Error al generar el reporte comparativo: " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If cerrarXlCmp And Not xlApp Is Nothing Then xlApp.Quit
    If rutaTmp <> "" Then
        If fso.FileExists(rutaTmp) Then fso.DeleteFile rutaTmp, True
    End If
    Err.Clear
    generarReporteComparativoBO = False
End Function

' Encabezado institucional comun de las hojas de Excel (filas 1 a 8). Se extrajo
' a su propio procedimiento para que el reporte ejecutivo y el comparativo usen
' exactamente el mismo encabezado y no se desfasen al modificar uno de los dos.
Private Sub EncabezadoInstitucionalExcelBO(ByRef ws As Object, ByVal titulo As String, _
        ByVal fecNorm As String, ByVal folioFecha As String, ByVal elaboro As String, _
        ByVal origen As String, ByVal registros As String, ByVal kpiIzq As String, ByVal kpiDer As String)
    Dim c As Long
    ws.Range("A1:H2").Merge
    ws.Range("A1").Value = "DIRECCIÓN CORPORATIVA DE ECONOMÍA Y FISCALIZACIÓN DE HIDROCARBUROS" & vbLf & _
                           "DIRECCION DE FISCALIZACION Y CONTRAPRESTACIONES PETROLERAS" & vbLf & _
                           "GERENCIA DE VERIFICACION DE VOLUMENES Y REGALIAS"
    With ws.Range("A1:H2")
        .Interior.Color = RGB(214, 223, 236)
        .Font.Name = "Cambria"
        .Font.Color = RGB(31, 78, 120)
        .Font.Bold = True
        .Font.Size = 11.5
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
        .WrapText = True
        .Borders.LineStyle = 1
        .Borders.Color = RGB(0, 78, 76)
    End With
    ws.Rows(1).RowHeight = 32
    ws.Rows(2).RowHeight = 32

    ws.Range("A3:H3").Merge
    ws.Range("A3").Value = titulo
    With ws.Range("A3:H3")
        .Interior.Color = RGB(0, 78, 76)
        .Font.Name = "Cambria"
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
    End With
    ws.Rows(3).RowHeight = 23

    ws.Range("F4:H4").Merge
    ws.Range("F5:H5").Merge
    ws.Cells(4, 1).Value = "Fecha de Corte SQL:": ws.Cells(4, 2).Value = fecNorm
    ws.Cells(4, 3).Value = "Folio Institucional:": ws.Cells(4, 4).Value = "UPST-SQL-" & folioFecha
    ws.Cells(4, 5).Value = "Responsable:": ws.Cells(4, 6).Value = elaboro
    ws.Cells(5, 1).Value = "Origen de Datos:": ws.Cells(5, 2).Value = origen
    ws.Cells(5, 3).Value = "Registros Extraidos:": ws.Cells(5, 4).Value = registros
    ws.Cells(5, 5).Value = "Organo Colegiado:": ws.Cells(5, 6).Value = "Comite Tecnico de Verificacion de Contraprestaciones"
    ws.Rows(4).RowHeight = 20.5
    ws.Rows(5).RowHeight = 20.5
    With ws.Range("A4:H5")
        .Font.Name = "Cambria"
        .Font.Size = 8.2
        .VerticalAlignment = -4108
        .ShrinkToFit = True
        .Borders.LineStyle = 1
        .Borders.Color = RGB(184, 196, 212)
        .Interior.Color = RGB(255, 255, 255)
    End With
    For c = 1 To 5 Step 2
        With ws.Range(ws.Cells(4, c), ws.Cells(5, c))
            .Interior.Color = RGB(238, 243, 249)
            .Font.Bold = True
            .Font.Color = RGB(31, 78, 120)
            .HorizontalAlignment = -4108
        End With
    Next c
    ws.Rows(6).RowHeight = 8

    ws.Range("A7:D7").Merge: ws.Range("A8:D8").Merge
    ws.Range("E7:H7").Merge: ws.Range("E8:H8").Merge
    ws.Cells(7, 1).Value = "INSTRUMENTOS CRUZADOS": ws.Cells(8, 1).Value = kpiIzq
    ws.Cells(7, 5).Value = "SALDO FINAL DEL CRUCE": ws.Cells(8, 5).Value = kpiDer
    ws.Rows(7).RowHeight = 17.5
    ws.Rows(8).RowHeight = 23.5
    With ws.Range("A7:H7")
        .Interior.Color = RGB(227, 236, 247)
        .Font.Name = "Cambria"
        .Font.Size = 8
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 120)
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
        .Borders.LineStyle = 1
        .Borders.Color = RGB(184, 196, 212)
    End With
    With ws.Range("A8:H8")
        .Interior.Color = RGB(255, 255, 255)
        .Font.Name = "Cambria"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(0, 78, 76)
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
        .Borders.LineStyle = 1
        .Borders.Color = RGB(184, 196, 212)
    End With
    ws.Rows(9).RowHeight = 9
End Sub

' =============================================================================
' AVISO DE INFORMACION YA PRECARGADA (SE PREGUNTA AL EXAMINAR, NO AL FINAL)
' -----------------------------------------------------------------------------
' Antes la pregunta de sobrescritura aparecia hasta el final, despues de toda la
' validacion, y se basaba en la BITACORA_CARGA_INSUMOS, es decir en el nombre
' del archivo. Eso tenia dos defectos:
'   - el usuario esperaba toda la validacion para enterarse de que iba a
'     reemplazar informacion;
'   - si la bitacora no tenia el registro, o el archivo venia con otro nombre,
'     no preguntaba aunque la tabla si tuviera datos de esa fecha.
'
' Esta funcion mira lo que REALMENTE hay en las tablas de insumo para la fecha,
' y se invoca desde los botones Archivo(s) y Carpeta, de modo que la pregunta
' aparece en el momento de examinar.
' =============================================================================
Public Function existeInformacionPrecargadaBO(ByVal fechaCorte As String, _
        ByVal bBD As Boolean, ByVal bCD As Boolean, ByVal bPEA As Boolean, _
        ByRef detalle As String) As Boolean
    Dim dao As FondoDAO, fecNorm As String
    Dim n1 As Long, n2 As Long, n3 As Long, total As Long
    On Error GoTo ErrPre
    existeInformacionPrecargadaBO = False
    detalle = ""
    If Not esFechaFormatoCorrectoBO(fechaCorte) Then Exit Function
    fecNorm = normalizarFechaBO(fechaCorte)

    ' Si no se ha marcado ninguna casilla, se revisan las tres: el usuario
    ' todavia no decidio y conviene avisarle de todo lo que hay en esa fecha.
    If Not (bBD Or bCD Or bPEA) Then
        bBD = True: bCD = True: bPEA = True
    End If

    Dim cP1 As Long, cP2 As Long, cP3 As Long
    Call obtenerConteosTresTablasFechaBO(fecNorm, cP1, cP2, cP3)
    If bBD Then n1 = cP1
    If bCD Then n2 = cP2
    If bPEA Then n3 = cP3

    total = n1 + n2 + n3
    If total <= 0 Then Exit Function

    If n1 > 0 Then detalle = detalle & "  - " & TABLA_INSUMO_PRIMER_PISO & ": " & _
                                        Format(n1, "#,##0") & " registros" & vbCrLf
    If n2 > 0 Then detalle = detalle & "  - " & TABLA_INSUMO_TESORERIA & ": " & _
                                        Format(n2, "#,##0") & " registros" & vbCrLf
    If n3 > 0 Then detalle = detalle & "  - " & TABLA_INSUMO_SEGUNDO_PISO & ": " & _
                                        Format(n3, "#,##0") & " registros" & vbCrLf
    existeInformacionPrecargadaBO = True
    Exit Function
ErrPre:
    On Error Resume Next
    Set dao = Nothing
    detalle = ""
    existeInformacionPrecargadaBO = False
End Function

' Arma el texto de la pregunta de reemplazo. Se deja aqui y no en el formulario
' para que los dos botones de examinar y la carga final usen el mismo texto.
Public Function textoPreguntaReemplazoBO(ByVal fechaCorte As String, ByVal detalle As String) As String
    textoPreguntaReemplazoBO = _
        "INFORMACION YA PRECARGADA PARA LA FECHA " & normalizarFechaBO(fechaCorte) & vbCrLf & vbCrLf & _
        "Las siguientes bases de datos YA tienen informacion de esa fecha de corte:" & vbCrLf & vbCrLf & _
        detalle & vbCrLf & _
        "Si continua, los registros de esa fecha se REEMPLAZARAN por los del archivo" & vbCrLf & _
        "que seleccione. Los registros de las demas fechas no se tocan." & vbCrLf & vbCrLf & _
        "Desea reemplazar la informacion de esta fecha?"
End Function

' =============================================================================
' CONTROL DE ACCESO Y DATOS DE LA SESIÓN ACTIVA  (CAPA DE NEGOCIO)
' -----------------------------------------------------------------------------
' Marco de Desarrollo, sección I.2: "La presentación debe evitar contener
' consultas SQL y reglas de negocios complejas".
'
' Antes, la regla de quién puede entrar a cada pantalla estaba escrita DENTRO
' del Form_Load de cada formulario: 14 copias del mismo encabezado de sesión y
' 7 copias de la regla de autorización, cada una con su propio texto. Eso es
' justo lo que el Marco advierte: duplicidad de código y dificultad para
' corregir errores, porque cambiar la regla obligaba a tocar 7 formularios y
' bastaba olvidar uno para dejar un hueco. De hecho así se había quedado sin
' protección la pantalla de restablecer contraseñas de otros usuarios.
'
' Ahora la regla vive AQUÍ, en un solo lugar. El formulario solo pregunta y
' muestra; no decide.
' =============================================================================

' Devuelve True si hay una sesión de usuario cargada.
Public Function haySesionActivaBO() As Boolean
    On Error Resume Next
    haySesionActivaBO = Not (usuarioSesionDTO Is Nothing)
    Err.Clear
End Function

' Traduce la clave de perfil a su descripción. Es una regla de negocio: la
' presentación no tiene por qué saber que "A" significa ADMINISTRADOR.
Public Function descripcionPerfilBO(ByVal cvePerfil As String) As String
    Select Case UCase(Trim(cvePerfil))
        Case UCase(PERFIL_ADMINISTRADOR)
            descripcionPerfilBO = "ADMINISTRADOR"
        Case UCase(PERFIL_EMPLEADO)
            descripcionPerfilBO = "USUARIO OPERATIVO"
        Case ""
            descripcionPerfilBO = "-"
        Case Else
            descripcionPerfilBO = "PERFIL NO RECONOCIDO (" & Trim(cvePerfil) & ")"
    End Select
End Function

' True si el usuario de la sesión tiene perfil de administrador.
Public Function esAdministradorSesionBO() As Boolean
    On Error Resume Next
    esAdministradorSesionBO = False
    If usuarioSesionDTO Is Nothing Then Exit Function
    esAdministradorSesionBO = (Trim(usuarioSesionDTO.getCvePerfil()) = PERFIL_ADMINISTRADOR)
    Err.Clear
End Function

' Texto listo para la etiqueta de identificación del encabezado.
Public Function etiquetaUsuarioSesionBO() As String
    On Error Resume Next
    etiquetaUsuarioSesionBO = "No. Empleado: -"
    If usuarioSesionDTO Is Nothing Then Exit Function
    etiquetaUsuarioSesionBO = "No. Empleado: " & usuarioSesionDTO.getCveUsuario() & _
                              " - " & usuarioSesionDTO.getNombre() & _
                              " " & usuarioSesionDTO.getAPaterno()
    Err.Clear
End Function

' Texto listo para la etiqueta de perfil del encabezado.
Public Function etiquetaPerfilSesionBO() As String
    On Error Resume Next
    etiquetaPerfilSesionBO = "Perfil: -"
    If usuarioSesionDTO Is Nothing Then Exit Function
    etiquetaPerfilSesionBO = "Perfil: " & descripcionPerfilBO(usuarioSesionDTO.getCvePerfil())
    Err.Clear
End Function

' CATÁLOGO DE AUTORIZACIÓN. Devuelve el perfil que exige la pantalla, o cadena
' vacía si no exige ninguno, y deja en 'motivo' el texto que se le mostrará al
' usuario que no cumple. Para cambiar quién entra a dónde, se cambia AQUÍ.
Public Function perfilRequeridoPantallaBO(ByVal clavePantalla As String, ByRef motivo As String) As String
    motivo = ""
    perfilRequeridoPantallaBO = ""
    Select Case UCase(Trim(clavePantalla))
        Case PANT_GESTION_USUARIOS
            motivo = "Su perfil de Usuario Operativo no cuenta con autorización para gestionar usuarios del sistema."
            perfilRequeridoPantallaBO = PERFIL_ADMINISTRADOR
        Case PANT_CONSULTA_USUARIOS
            motivo = "La consulta del directorio de usuarios requiere perfil de Administrador."
            perfilRequeridoPantallaBO = PERFIL_ADMINISTRADOR
        Case PANT_ALTA_USUARIOS
            motivo = "El alta de nuevos usuarios requiere perfil de Administrador."
            perfilRequeridoPantallaBO = PERFIL_ADMINISTRADOR
        Case PANT_CAMBIO_USUARIOS
            motivo = "La modificación de usuarios requiere perfil de Administrador."
            perfilRequeridoPantallaBO = PERFIL_ADMINISTRADOR
        Case PANT_BAJA_USUARIOS
            motivo = "La baja o inactivación de usuarios requiere perfil de Administrador."
            perfilRequeridoPantallaBO = PERFIL_ADMINISTRADOR
        Case PANT_RESTABLECER_CONTRASENA
            ' Esta pantalla NO tenía restricción. Restablecer la contraseña de
            ' otro usuario es una operación de administrador, igual que el alta
            ' o la baja, así que se le aplica la misma regla.
            motivo = "El restablecimiento de contraseñas de otros usuarios requiere perfil de Administrador."
            perfilRequeridoPantallaBO = PERFIL_ADMINISTRADOR
        Case PANT_ALTA_INSTRUMENTOS
            motivo = "El alta de instrumentos requiere perfil de Administrador."
            perfilRequeridoPantallaBO = PERFIL_ADMINISTRADOR
        Case PANT_CAMBIO_INSTRUMENTOS
            motivo = "La modificación de instrumentos requiere perfil de Administrador."
            perfilRequeridoPantallaBO = PERFIL_ADMINISTRADOR
        Case Else
            ' Menú, cambio de la contraseña propia, gestión y consulta de
            ' instrumentos, carga de insumos y generación de reportes: abiertas
            ' a cualquier perfil con sesión activa.
            perfilRequeridoPantallaBO = ""
    End Select
End Function

' LA DECISIÓN. El formulario pregunta y obedece; no evalúa perfiles ni arma el
' mensaje. Devuelve True si se puede entrar; si no, deja el texto en msgDenegado.
Public Function autorizarPantallaBO(ByVal clavePantalla As String, ByRef msgDenegado As String) As Boolean
    Dim requerido As String, motivo As String
    On Error GoTo ErrAut
    autorizarPantallaBO = True
    msgDenegado = ""

    requerido = perfilRequeridoPantallaBO(clavePantalla, motivo)
    If requerido = "" Then Exit Function          ' pantalla sin restricción

    ' Sin sesión activa no se entra a una pantalla restringida. Antes toda la
    ' verificación vivía dentro de un If Not usuarioSesionDTO Is Nothing, por lo
    ' que al no haber sesión la pantalla quedaba ABIERTA. Aquí se deniega.
    If usuarioSesionDTO Is Nothing Then
        msgDenegado = "Acceso Denegado: no hay una sesión activa." & vbCrLf & vbCrLf & _
                      "Ingrese al sistema desde la pantalla de acceso para continuar."
        autorizarPantallaBO = False
        Exit Function
    End If

    If Trim(usuarioSesionDTO.getCvePerfil()) <> requerido Then
        msgDenegado = "Acceso Denegado: " & motivo
        autorizarPantallaBO = False
    End If
    Exit Function
ErrAut:
    ' Ante cualquier duda, no se autoriza.
    msgDenegado = "Acceso Denegado: no fue posible verificar los permisos de la sesión."
    autorizarPantallaBO = False
End Function

' Que pantalla de cambio de contrasena le corresponde al usuario de la sesion.
' Es una regla de negocio: el administrador entra a la pantalla que ademas
' permite restablecer la contrasena de OTROS usuarios; el usuario operativo
' solo puede cambiar la suya. El menu no tiene por que saber esto ni comparar
' perfiles contra literales.
Public Function pantallaCambioContrasenaBO() As String
    If esAdministradorSesionBO() Then
        pantallaCambioContrasenaBO = "Form_Cambiar_Contrasena_Administrador"
    Else
        pantallaCambioContrasenaBO = "Form_Cambiar_Contrasena"
    End If
End Function
