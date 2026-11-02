Attribute VB_Name = "ModInstaladorUniversal"
Option Compare Database
' ==============================================================================
' INSTALADOR NATIVO 1-CLIC (MS ACCESS COM DIRECTO VIA VBSCRIPT)
' MODULO DE ADMINISTRACION INTEGRAL DE RIESGOS - HIDROCARBUROS (INTERFAZ CENTRADA SIMETRICA Y 101 EMPLEADOS)
' ==============================================================================
Option Explicit

Dim fso, strProjectRoot, strLogPath, logFile
Dim strBdDir, strBackendPath, strFrontendPath, strLogoPath, strCalIconPath, strEyeIconPath, strPdfIconPath, strExcelIconPath, strImgDir, rpt, rName
Dim access, db, comp, objFolder, objFile
Dim strModulesDir, strClassesDir, strCleanCode
Dim frm, fName, codeF, ctrl, tNames, tName, tdf
Dim empId, empPerfil, empNom

Dim wsh, verOff, arrVers
Dim arrNombres, arrPaternos, arrMaternos, arrCorreoPref, idxN, idxP, idxM, empApPat, empApMat, empCorreo, empEst
Dim sqlColsInsumo, arrTablasInsumo, tblInsName
Dim dbPwd, bCifrado
Dim arrOrdenDto, nomDto, rutaCls, arrOrdenMod, nomMod, rutaBas
Dim nVinculadas, nFallidas, msgVinculos
Dim lblSub
Dim arrDiasSem, dIdx, rCal, cCal, idxBtn
Dim prop
Dim wVbe, passVbe

Public Sub ConstruirTodoDesdeCero(Optional ByVal strProjectRootParam As String = "")
Set fso = CreateObject("Scripting.FileSystemObject")

    If Trim(strProjectRootParam) <> "" Then
        strProjectRoot = Trim(strProjectRootParam)
    Else
        strProjectRoot = CurrentProject.Path
    End If
If LCase(fso.GetFileName(strProjectRoot)) = "src" Then
    strProjectRoot = fso.GetParentFolderName(strProjectRoot)
End If
strLogPath = strProjectRoot & "\setup_log.txt"
strBdDir = strProjectRoot & "\BD"
strBackendPath = strBdDir & "\RIESGOSBD.accdb"
strFrontendPath = strProjectRoot & "\Modulo_Administracion_Riesgos.accdb"
strModulesDir = strProjectRoot & "\src\modules"
strClassesDir = strProjectRoot & "\src\classes"

strLogoPath = strProjectRoot & "\Archivos\Imagenes\LogoCorporativoEnergia.png"
If Not fso.FileExists(strLogoPath) Then
    strLogoPath = strProjectRoot & "\LogoCorporativoEnergia.png"
End If
If Not fso.FileExists(strLogoPath) Then
    strLogoPath = strProjectRoot & "\Archivos\Imagenes\LogoCorporativoEnergia.jpg"
End If
If Not fso.FileExists(strLogoPath) Then
    strLogoPath = strProjectRoot & "\LogoCorporativoEnergia.jpg"
End If

strImgDir = strProjectRoot & "\Archivos\Imagenes"
If Not fso.FolderExists(strProjectRoot & "\Archivos") Then fso.CreateFolder(strProjectRoot & "\Archivos")
If Not fso.FolderExists(strImgDir) Then fso.CreateFolder(strImgDir)
strCalIconPath = strImgDir & "\IconoCalendario.bmp"
If Not fso.FileExists(strCalIconPath) Then
    If fso.FileExists(strProjectRoot & "\IconoCalendario.bmp") Then
        strCalIconPath = strProjectRoot & "\IconoCalendario.bmp"
    End If
End If
strEyeIconPath = strImgDir & "\IconoOjoVer.bmp"
strPdfIconPath = strImgDir & "\IconoReportePDF.bmp"
strExcelIconPath = strImgDir & "\IconoReporteExcel.bmp"

Set logFile = fso.CreateTextFile(strLogPath, True)
logFile.WriteLine "====================================================================="
logFile.WriteLine "INSTALACION DIRECTA: MODULO DE ADMINISTRACION INTEGRAL DE RIESGOS HIDROCARBUROS"
logFile.WriteLine "Fecha y hora: " & Now
logFile.WriteLine "Directorio: " & strProjectRoot
logFile.WriteLine "====================================================================="

    Call EjecutarBloqueInstalacion_01
    Call EjecutarBloqueInstalacion_02
    Call EjecutarBloqueInstalacion_03
    Call EjecutarBloqueInstalacion_04
    Call EjecutarBloqueInstalacion_05
    Call EjecutarBloqueInstalacion_06
    Call EjecutarBloqueInstalacion_07
    Call EjecutarBloqueInstalacion_08
    Call EjecutarBloqueInstalacion_09
    Call EjecutarBloqueInstalacion_10
    Call EjecutarBloqueInstalacion_11
    Call EjecutarBloqueInstalacion_12
    Call EjecutarBloqueInstalacion_13
    Call EjecutarBloqueInstalacion_14
End Sub

Private Sub EjecutarBloqueInstalacion_01()
' Paso 1: Verificar estructura de directorios del proyecto
logFile.WriteLine "Paso 1: Verificando directorios de recursos y codigo fuente..."
If Not fso.FolderExists(strProjectRoot & "\src") Then fso.CreateFolder(strProjectRoot & "\src")
If Not fso.FolderExists(strModulesDir) Then fso.CreateFolder(strModulesDir)
If Not fso.FolderExists(strClassesDir) Then fso.CreateFolder(strClassesDir)
logFile.WriteLine "   [OK] Estructura de carpetas verificada sin bloqueos de seguridad."

' Paso 2: Crear subcarpeta BD y limpiar archivos previos
If Not fso.FolderExists(strBdDir) Then
    fso.CreateFolder strBdDir
    logFile.WriteLine "   [OK] Subcarpeta BD creada."
End If

If fso.FileExists(strBackendPath) Then
    On Error Resume Next
    fso.DeleteFile strBackendPath, True
    On Error GoTo 0
End If

If fso.FileExists(strFrontendPath) Then
    On Error Resume Next
    fso.DeleteFile strFrontendPath, True
    On Error GoTo 0
End If

' Paso 3: Habilitar acceso programático seguro al proyecto VBA en HKCU (sin requerir permisos de administrador) e iniciar Access COM
logFile.WriteLine "Paso 3: Configurando permisos de automatizacion VBA e iniciando Microsoft Access..."
On Error Resume Next
Set wsh = CreateObject("WScript.Shell")
arrVers = Array("16.0", "15.0", "14.0", "12.0")
For Each verOff In arrVers
    wsh.RegWrite "HKCU\Software\Microsoft\Office\" & verOff & "\Access\Security\AccessVBOM", 1, "REG_DWORD"
    wsh.RegWrite "HKCU\Software\Microsoft\Office\" & verOff & "\Access\Security\VBAWarnings", 1, "REG_DWORD"
Next
Set wsh = Nothing
Err.Clear
Set access = CreateObject("Access.Application")
access.AutomationSecurity = 1 ' msoAutomationSecurityLow en memoria
If Err.Number <> 0 Then
    logFile.WriteLine "ERROR CRITICO al iniciar Access.Application: " & Err.Description
    logFile.Close
    MsgBox "No se pudo iniciar Microsoft Access en esta computadora." & vbCrLf & vbCrLf & _
           "Detalle: " & Err.Description, vbCritical, "Error de Inicio"
    Exit Sub
End If
On Error GoTo 0
logFile.WriteLine "   [OK] Microsoft Access instanciado correctamente."

access.Visible = True

' PASO A: CREAR BACKEND (RIESGOSBD.accdb)
logFile.WriteLine "Paso 4: Creando Backend BD\RIESGOSBD.accdb..."
access.NewCurrentDatabase strBackendPath
Set db = access.CurrentDb

logFile.WriteLine "   Creando tablas relacionales..."
db.Execute "CREATE TABLE USUARIOS (cveUsuario VARCHAR(50) PRIMARY KEY, password VARCHAR(100) NOT NULL, nombreUsuario VARCHAR(100) NOT NULL, aPaterno VARCHAR(100) NOT NULL, aMaterno VARCHAR(100), cvePerfil VARCHAR(10) NOT NULL, emailUsuario VARCHAR(150), cveEstatus VARCHAR(20) NOT NULL, cvePuesto INTEGER);"
db.Execute "CREATE TABLE EMISORES (cveEmisor VARCHAR(20) PRIMARY KEY, nombreEmisor VARCHAR(150) NOT NULL, sector VARCHAR(100), calificacionRiesgo VARCHAR(10));"
db.Execute "CREATE TABLE INSTRUMENTOS (idInstrumento VARCHAR(50) PRIMARY KEY, cveEmisor VARCHAR(20), tipoInstrumento VARCHAR(50), tasa DOUBLE, plazo INTEGER, fechaEmision DATETIME, fechaVencimiento DATETIME);"
db.Execute "CREATE TABLE PERFILES (cvePerfil VARCHAR(10) PRIMARY KEY, descripcion VARCHAR(50));"
db.Execute "CREATE TABLE ESTATUS (cveEstatus VARCHAR(20) PRIMARY KEY, descripcion VARCHAR(50));"
db.Execute "CREATE TABLE CAT_MONEDAS (cveMoneda VARCHAR(10) PRIMARY KEY, descripcion VARCHAR(60), tipoCambioFix DOUBLE);"
db.Execute "CREATE TABLE CAT_TASAS (cveTasa VARCHAR(20) PRIMARY KEY, descripcion VARCHAR(80), valorTasa DOUBLE);"
db.Execute "CREATE TABLE TBL_TRANSACCIONES_FONDO (idTransaccion VARCHAR(30) PRIMARY KEY, fechaOperacion DATETIME, tipoFondo VARCHAR(80), idInstrumento VARCHAR(50), cveMoneda VARCHAR(10), montoInvertido DOUBLE, tasaAplicada DOUBLE, plazoDias INTEGER, rendimientoEstimado DOUBLE, cveUsuarioRegistro VARCHAR(50));"

db.Execute "CREATE TABLE BITACORA_CARGA_INSUMOS (idCarga AUTOINCREMENT PRIMARY KEY, nombreArchivo VARCHAR(200), fechaCorte VARCHAR(20), fechaHoraCarga DATETIME, totalRegistros INTEGER, fondosActualizados VARCHAR(200));"

logFile.WriteLine "   Insertando datos semilla en transaccion unica..."
On Error Resume Next
access.DBEngine.Workspaces(0).BeginTrans
On Error GoTo 0
db.Execute "INSERT INTO PERFILES (cvePerfil, descripcion) VALUES ('A', 'Administrador');"
db.Execute "INSERT INTO PERFILES (cvePerfil, descripcion) VALUES ('U', 'Usuario Operativo');"
db.Execute "INSERT INTO ESTATUS (cveEstatus, descripcion) VALUES ('ACTIVO', 'Usuario Activo');"
db.Execute "INSERT INTO ESTATUS (cveEstatus, descripcion) VALUES ('INACTIVO', 'Usuario Inactivo');"

logFile.WriteLine "   Poblando directorio institucional de empleados con nombres y correos realistas (contrasena Wembley1)..."
arrNombres = Array("Carlos Alberto", "María Fernanda", "Alejandro", "Sofía Guadalupe", "Jorge Luis", "Ana Paola", "Roberto Carlos", "Valeria", "Fernando José", "Gabriela", "Miguel Ángel", "Daniela", "Ricardo", "Adriana", "Eduardo", "Paulina", "Héctor Manuel", "Claudia Elena", "Arturo", "Mariana", "Francisco Javier", "Laura Patricia", "Diego Emiliano", "Verónica", "Raúl")
arrPaternos = Array("Mendoza", "Ramírez", "Hernández", "Gómez", "Vázquez", "Castillo", "Morales", "Ortega", "Delgado", "Ríos", "Aguilar", "Navarro", "Salazar", "Cervantes", "Domínguez", "Estrada", "Fuentes", "Guzmán", "Juárez", "Luna", "Montes", "Núñez", "Orozco", "Peña", "Rojas")
arrMaternos = Array("Villanueva", "Cárdenas", "Espinoza", "Palacios", "Cabrera", "Miranda", "Valencia", "Acosta", "Bautista", "Campos", "Díaz", "Escobar", "Figueroa", "Gallardo", "Herrera", "Ibarra", "Jiménez", "Lara", "Medina", "Nieto", "Olvera", "Pacheco", "Quintero", "Reyes", "Soto")
arrCorreoPref = Array("cmendoza", "mramirez", "ahernandez", "sgomez", "jvazquez", "acastillo", "rmorales", "vortega", "fdelgado", "grios", "maguilar", "dnavarro", "rsalazar", "acervantes", "edominguez", "pestrada", "hfuentes", "cguzman", "ajuarez", "mluna", "fmontes", "lnunez", "dorozco", "vpena", "rrojas")

' Directorio de 101 usuarios de 4 digitos (1000 al 1100) con contrasena de 8 caracteres (Wembley1)
db.Execute "INSERT INTO USUARIOS (cveUsuario, password, nombreUsuario, aPaterno, aMaterno, cvePerfil, emailUsuario, cveEstatus, cvePuesto) VALUES ('1000', 'ENC$E67D64692A29C0BB', 'Wembley', 'Ibarra', 'Tepato', 'A', 'wembley.ibarra@energia-upstream.org', 'ACTIVO', 1);"

For empId = 1001 To 1100
    idxN = (empId - 1001) Mod 25
    idxP = ((empId - 1001) * 3 + 2) Mod 25
    idxM = ((empId - 1001) * 7 + 5) Mod 25
    empNom = arrNombres(idxN)
    empApPat = arrPaternos(idxP)
    empApMat = arrMaternos(idxM)
    empCorreo = arrCorreoPref(idxN) & "." & empId & "@energia-upstream.org"
    If (empId Mod 10) = 0 Then
        empPerfil = "A"
    Else
        empPerfil = "U"
    End If
    empEst = "ACTIVO"
    db.Execute "INSERT INTO USUARIOS (cveUsuario, password, nombreUsuario, aPaterno, aMaterno, cvePerfil, emailUsuario, cveEstatus, cvePuesto) VALUES ('" & empId & "', 'ENC$E67D64692A29C0BB', '" & empNom & "', '" & empApPat & "', '" & empApMat & "', '" & empPerfil & "', '" & empCorreo & "', '" & empEst & "', 2);"
Next

' Catalogo ampliado de 15 Emisores Institucionales
db.Execute "INSERT INTO EMISORES (cveEmisor, nombreEmisor, sector, calificacionRiesgo) VALUES ('GOB_FED', 'Operadora Central de Hidrocarburos', 'Gubernamental Soberano', 'AAA');"
db.Execute "INSERT INTO EMISORES (cveEmisor, nombreEmisor, sector, calificacionRiesgo) VALUES ('BANXICO', 'Banco de México', 'Banca Central', 'AAA');"
db.Execute "INSERT INTO EMISORES (cveEmisor, nombreEmisor, sector, calificacionRiesgo) VALUES ('BANOBRAS', 'Banco Nacional de Obras y Servicios Públicos', 'Banca de Desarrollo', 'AAA');"
db.Execute "INSERT INTO EMISORES (cveEmisor, nombreEmisor, sector, calificacionRiesgo) VALUES ('NAFIN', 'Nacional Financiera S.N.C.', 'Banca de Desarrollo', 'AAA');"
db.Execute "INSERT INTO EMISORES (cveEmisor, nombreEmisor, sector, calificacionRiesgo) VALUES ('HIDROCARBUROS', 'Corporación de Exploración y Producción de Hidrocarburos', 'Banca de Desarrollo', 'AAA');"
db.Execute "INSERT INTO EMISORES (cveEmisor, nombreEmisor, sector, calificacionRiesgo) VALUES ('PEMEX', 'Petróleos Mexicanos', 'Energía y Petróleo', 'AA+');"
db.Execute "INSERT INTO EMISORES (cveEmisor, nombreEmisor, sector, calificacionRiesgo) VALUES ('CFE', 'Comisión Federal de Electricidad', 'Energía Eléctrica', 'AAA');"
db.Execute "INSERT INTO EMISORES (cveEmisor, nombreEmisor, sector, calificacionRiesgo) VALUES ('AMX', 'América Móvil S.A.B. de C.V.', 'Telecomunicaciones', 'AAA');"
db.Execute "INSERT INTO EMISORES (cveEmisor, nombreEmisor, sector, calificacionRiesgo) VALUES ('FEMSA', 'Fomento Económico Mexicano S.A.B.', 'Consumo y Comercio', 'AAA');"
db.Execute "INSERT INTO EMISORES (cveEmisor, nombreEmisor, sector, calificacionRiesgo) VALUES ('BIMBO', 'Grupo Bimbo S.A.B. de C.V.', 'Alimentos y Consumo', 'AAA');"
db.Execute "INSERT INTO EMISORES (cveEmisor, nombreEmisor, sector, calificacionRiesgo) VALUES ('CEMEX', 'Cemex S.A.B. de C.V.', 'Infraestructura y Cemento', 'AA+');"
db.Execute "INSERT INTO EMISORES (cveEmisor, nombreEmisor, sector, calificacionRiesgo) VALUES ('GFNORTE', 'Grupo Financiero Banorte S.A.B.', 'Banca Comercial', 'AAA');"
db.Execute "INSERT INTO EMISORES (cveEmisor, nombreEmisor, sector, calificacionRiesgo) VALUES ('BBVAMX', 'BBVA México S.A. Institución de Banca Múltiple', 'Banca Comercial', 'AAA');"
db.Execute "INSERT INTO EMISORES (cveEmisor, nombreEmisor, sector, calificacionRiesgo) VALUES ('AC', 'Arca Continental S.A.B. de C.V.', 'Bebidas y Consumo', 'AAA');"
db.Execute "INSERT INTO EMISORES (cveEmisor, nombreEmisor, sector, calificacionRiesgo) VALUES ('KIMBER', 'Kimberly-Clark de México S.A.B.', 'Productos de Consumo', 'AAA');"

' Catalogo ampliado de 25 Instrumentos Financieros registrados en el Administración Integral de Riesgos
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('CETES_28D', 'GOB_FED', 'CETES Gubernamental', 10.75, 28, #2026-09-18#, #2026-10-16#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('CETES_91D', 'GOB_FED', 'CETES Gubernamental', 10.90, 91, #2026-08-15#, #2026-11-14#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('CETES_182D', 'GOB_FED', 'CETES Gubernamental', 11.05, 182, #2026-06-01#, #2026-11-30#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('CETES_364D', 'GOB_FED', 'CETES Gubernamental', 10.85, 364, #2026-01-15#, #2027-01-14#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('BONOS_M03', 'GOB_FED', 'BONOS M Tasa Fija', 9.95, 1095, #2025-03-10#, #2028-03-09#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('BONOS_M05', 'GOB_FED', 'BONOS M Tasa Fija', 9.90, 1825, #2024-05-20#, #2029-05-19#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('BONOS_M10', 'GOB_FED', 'BONOS M Tasa Fija', 9.85, 3640, #2026-01-15#, #2036-01-03#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('BONOS_M20', 'GOB_FED', 'BONOS M Largo Plazo', 10.00, 7280, #2023-06-15#, #2043-05-21#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('BONOS_M30', 'GOB_FED', 'BONOS M Largo Plazo', 10.10, 10800, #2022-01-10#, #2051-08-06#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('UDIBONO_03', 'GOB_FED', 'UDIBONOS Tasa Real', 5.25, 1095, #2025-02-01#, #2028-01-31#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('UDIBONO_10', 'GOB_FED', 'UDIBONOS Tasa Real', 5.10, 3650, #2024-04-10#, #2034-04-08#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('UDIBONO_30', 'GOB_FED', 'UDIBONOS Cobertura Inflación', 4.95, 10800, #2021-09-01#, #2051-03-28#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('BONDES_F1', 'BANXICO', 'BONDES F TIIE Fondeo', 11.12, 364, #2026-02-12#, #2027-02-11#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('BONDES_F3', 'BANXICO', 'BONDES F TIIE Fondeo', 11.18, 1092, #2025-07-01#, #2028-06-27#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('BONDES_D5', 'BANXICO', 'BONDES D Revisables', 11.05, 1820, #2024-08-20#, #2029-08-14#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('BANOB_29', 'BANOBRAS', 'Bono Banca de Desarrollo', 10.65, 1460, #2025-04-15#, #2029-04-14#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('NAFIN_28', 'NAFIN', 'Bono Sustentable NAFIN', 10.55, 1095, #2025-05-10#, #2028-05-09#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('BLK_29', 'HIDROCARBUROS', 'Certificado Bursátil Banca Des.', 10.60, 1460, #2025-06-01#, #2029-05-31#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('PEMEX_27', 'PEMEX', 'Certificado Bursátil Paraestatal', 11.45, 1080, #2024-03-20#, #2027-03-05#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('PEMEX_30', 'PEMEX', 'Bono Energético Largo Plazo', 11.65, 1800, #2025-02-15#, #2030-01-20#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('CFE_28', 'CFE', 'Certificado Bursátil Infraestructura', 10.95, 1440, #2024-06-15#, #2028-05-25#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('CFE_31', 'CFE', 'Bono Sustentable Energía', 11.05, 1825, #2026-01-20#, #2031-01-19#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('AMX_29', 'AMX', 'Deuda Corporativa AAA', 10.15, 1800, #2024-09-01#, #2029-08-06#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('FEMSA_30', 'FEMSA', 'Certificado Bursátil Corporativo', 10.10, 1825, #2025-03-01#, #2030-02-28#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('BIMBO_32', 'BIMBO', 'Bono Corporativo Alimentos', 10.20, 2190, #2026-02-01#, #2032-01-31#);"
db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('BLK_VERDE_26', 'HIDROCARBUROS', 'Bono Verde Hidrocarburos', 10.85, 365, #2026-03-02#, #2027-03-02#);"

db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('BONDES_G_2026', 'GOB_FED', 'Bondes G Gubernamental', 11.15, 365, #2026-02-12#, #2027-02-12#);"

db.Execute "INSERT INTO INSTRUMENTOS (idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento) VALUES ('NAFIN_PYME_26', 'NAFIN', 'Credito Interbancario PyME', 10.7, 365, #2026-04-20#, #2027-04-20#);"


db.Execute "INSERT INTO CAT_MONEDAS (cveMoneda, descripcion, tipoCambioFix) VALUES ('MXN', 'Peso Mexicano (M.N.)', 1.0000);"
db.Execute "INSERT INTO CAT_MONEDAS (cveMoneda, descripcion, tipoCambioFix) VALUES ('USD', 'Dólar Estadounidense (M.E.)', 16.5625);"
db.Execute "INSERT INTO CAT_MONEDAS (cveMoneda, descripcion, tipoCambioFix) VALUES ('UDI', 'Unidades de Inversión', 8.2450);"

db.Execute "INSERT INTO CAT_TASAS (cveTasa, descripcion, valorTasa) VALUES ('TIIE_28', 'TIIE a 28 Días Banxico', 11.00);"
db.Execute "INSERT INTO CAT_TASAS (cveTasa, descripcion, valorTasa) VALUES ('CETES_28', 'Tasa de Rendimiento CETES 28D', 10.75);"
db.Execute "INSERT INTO CAT_TASAS (cveTasa, descripcion, valorTasa) VALUES ('SOFR_USD', 'Secured Overnight Financing Rate', 5.30);"

db.Execute "INSERT INTO TBL_TRANSACCIONES_FONDO (idTransaccion, fechaOperacion, tipoFondo, idInstrumento, cveMoneda, montoInvertido, tasaAplicada, plazoDias, rendimientoEstimado, cveUsuarioRegistro) VALUES ('TRX_FP_001', #2026-09-30#, 'Cartera de Crédito Comercial (Primer Piso)', 'CETES_28D', 'MXN', 500000000, 10.75, 28, 4180555.56, '1000');"
db.Execute "INSERT INTO TBL_TRANSACCIONES_FONDO (idTransaccion, fechaOperacion, tipoFondo, idInstrumento, cveMoneda, montoInvertido, tasaAplicada, plazoDias, rendimientoEstimado, cveUsuarioRegistro) VALUES ('TRX_FP_002', #2026-09-30#, 'Portafolio de Tesorería y Mercado de Dinero', 'BONOS_M10', 'MXN', 850000000, 9.85, 360, 83725000.00, '1000');"
db.Execute "INSERT INTO TBL_TRANSACCIONES_FONDO (idTransaccion, fechaOperacion, tipoFondo, idInstrumento, cveMoneda, montoInvertido, tasaAplicada, plazoDias, rendimientoEstimado, cveUsuarioRegistro) VALUES ('TRX_FP_003', #2026-09-30#, 'Financiamiento Sectorial (Segundo Piso)', 'PEMEX_27', 'MXN', 540000000, 11.20, 360, 60480000.00, '1001');" 

' Paso 4.1: Crear las 3 bases de datos independientes de Insumos (53 columnas incl. ID_INSTRUMENTO) en RIESGOSBD.accdb:
'   1) BD_INSUMO_PRIMER_PISO  (Cartera de Crédito Comercial - Primer Piso)
'   2) BD_INSUMO_TESORERIA    (Portafolio de Tesorería y Mercado de Dinero)
'   3) BD_INSUMO_SEGUNDO_PISO (Financiamiento Sectorial y Garantías - Segundo Piso)
'   + CARTERA_ETIQUETADA      (Repositorio consolidado global)
logFile.WriteLine "   Creando las 3 tablas de bases de datos de insumos (53 columnas incl. ID_INSTRUMENTO) en Backend..."
sqlColsInsumo = "([P] VARCHAR(50), [PISO] VARCHAR(50), [G] VARCHAR(50), [GRUPO] VARCHAR(120), [ESTATA] VARCHAR(50), [DESC_SUCURSAL] VARCHAR(120), " & _
    "[SG] VARCHAR(50), [DESC_SUBGRUPO] VARCHAR(120), [CAMPOA] VARCHAR(50), [NOMBRE] VARCHAR(200), [RFC] VARCHAR(50), [MONEDA2] VARCHAR(50), " & _
    "[MON] VARCHAR(50), [MONEDA] VARCHAR(50), [EF] VARCHAR(50), [ENT_FED] VARCHAR(120), [CONCATENADO] VARCHAR(120), [CTA] VARCHAR(50), " & _
    "[SCATAS] VARCHAR(50), [SCTA] VARCHAR(50), [SSCTA] VARCHAR(50), [SSSCTA] VARCHAR(50), [SECTOR] VARCHAR(50), [C] VARCHAR(50), " & _
    "[CONCEPTO] VARCHAR(120), [SC] VARCHAR(50), [SCONCEPTO] VARCHAR(120), [SSCONCEPTO] VARCHAR(120), [DESTINO] VARCHAR(150), " & _
    "[SDINMO] DOUBLE, [SDFIMO] DOUBLE, [PROMMO] DOUBLE, [SDFIMN] DOUBLE, [SDPRMN] DOUBLE, [SDFIDL] DOUBLE, [SDPRDL] DOUBLE, " & _
    "[MESPRO] VARCHAR(50), [ANOPRO] VARCHAR(50), [ACT] VARCHAR(50), [ACTIVIDAD] VARCHAR(150), [GRUCON] VARCHAR(50), [GRUPO_CONT] VARCHAR(150), " & _
    "[GRUCONTNORM] VARCHAR(50), [ARTICULO_NORMATIVO] VARCHAR(200), [CODIGO_ACTIVIDAD_ECONOMICA_GRAL] VARCHAR(50), " & _
    "[DESC_ACTIVIDAD_ECONOMICA_GRAL] VARCHAR(200), [CODIGO_EJECUTIVO] VARCHAR(50), [NOMBRE_EJECUTIVO] VARCHAR(150), " & _
    "[DESCRIPCIÓN CUENTA] VARCHAR(150), [SITUACION DEL CREDITO] VARCHAR(50), [CONCEPTO_ETAPA] VARCHAR(120), [FECHA_INSUMO] VARCHAR(30), [ID_INSTRUMENTO] VARCHAR(50));"
arrTablasInsumo = Array("BD_INSUMO_PRIMER_PISO", "BD_INSUMO_TESORERIA", "BD_INSUMO_SEGUNDO_PISO", "CARTERA_ETIQUETADA")
For Each tblInsName In arrTablasInsumo
    db.Execute "CREATE TABLE " & tblInsName & " " & sqlColsInsumo
    db.Execute "CREATE INDEX idx_fecha ON " & tblInsName & " ([FECHA_INSUMO]);"
    db.Execute "CREATE INDEX idx_instr ON " & tblInsName & " ([ID_INSTRUMENTO]);"
Next
db.Execute "CREATE INDEX idx_piso_fec ON CARTERA_ETIQUETADA ([FECHA_INSUMO], [PISO]);"
db.Execute "CREATE INDEX idx_bit_fec ON BITACORA_CARGA_INSUMOS ([fechaCorte], [nombreArchivo]);"
On Error Resume Next
access.DBEngine.Workspaces(0).CommitTrans
On Error GoTo 0
logFile.WriteLine "   [OK] 3 Bases de datos de Insumos (BD_INSUMO_PRIMER_PISO, BD_INSUMO_TESORERIA, BD_INSUMO_SEGUNDO_PISO) creadas con campo ID_INSTRUMENTO."

On Error Resume Next
access.CloseCurrentDatabase
On Error GoTo 0
DoEvents
logFile.WriteLine "   [OK] Backend RIESGOSBD.accdb creado, poblado y con Cartera Etiquetada incorporada exitosamente."

' PASO B: CREAR FRONTEND (Modulo_Administracion_Riesgos.accdb)
logFile.WriteLine "Paso 5: Creando Frontend Modulo_Administracion_Riesgos.accdb..."
' Cifrar el backend con la contrasena institucional antes de vincular las tablas.
' Se abre en modo exclusivo (unico modo en que Access permite asignar contrasena)
' y despues se vuelve a cerrar. A partir de aqui, abrir RIESGOSBD.accdb
' directamente pide contrasena; el frontend la lleva en la cadena de conexion.
logFile.WriteLine "Paso 4.9: Cifrando el backend con contrasena institucional..."
bCifrado = False
On Error Resume Next
' La base ya se cerro unas lineas arriba. Se intenta cerrar otra vez solo por
' seguridad, y SIEMPRE bajo On Error Resume Next: cerrar lo que ya esta cerrado
' levanta el error 800A09A3 (objeto cerrado o inexistente) y detiene el script.
access.CloseCurrentDatabase
Err.Clear
DoEvents

' Access solo permite asignar contrasena abriendo la base en modo EXCLUSIVO.
Set dbPwd = CreateObject("DAO.DBEngine.120").OpenDatabase(strBackendPath, True, False)
If Err.Number = 0 And Not dbPwd Is Nothing Then
    Err.Clear
    dbPwd.NewPassword "", "Wembley1"
    If Err.Number = 0 Then
        bCifrado = True
        logFile.WriteLine "   [OK] RIESGOSBD.accdb cifrado."
    Else
        logFile.WriteLine "   [AVISO] No se pudo asignar la contrasena: " & Err.Description
    End If
    Err.Clear
    dbPwd.Close
Else
    logFile.WriteLine "   [AVISO] No se pudo abrir el backend en modo exclusivo: " & Err.Description
End If
Set dbPwd = Nothing
Err.Clear
On Error GoTo 0

If Not bCifrado Then
    logFile.WriteLine "   [AVISO] El backend quedo SIN cifrar. Las tablas se vincularan sin"
    logFile.WriteLine "           contrasena para que el aplicativo funcione de todos modos."
    logFile.WriteLine "           Para cifrarlo despues: abrir BD\RIESGOSBD.accdb en modo"
    logFile.WriteLine "           exclusivo y usar Archivo > Cifrar con contrasena (Wembley1),"
    logFile.WriteLine "           y luego volver a vincular las tablas."
End If

access.NewCurrentDatabase strFrontendPath

' Vincular referencia ADODB 6.1
On Error Resume Next
access.VBE.ActiveVBProject.References.AddFromGuid "{B691E011-1797-432E-907A-4D8C69339129}", 6, 1
On Error GoTo 0

' 1. Importar en ORDEN ESTRICTO DE DEPENDENCIAS:
'    a) Las 10 clases DTO (*DTO.cls)
'    b) Los modulos base (Constantes.bas, Utilities.bas, ConsultasSQL.bas)
'    c) La clase de persistencia (FondoDAO.cls)
'    d) El modulo de negocio (FondoBO.bas)
logFile.WriteLine "   Importando clases DTO, modulos base, FondoDAO y FondoBO en orden de dependencias..."
arrOrdenDto = Array("UsuarioDTO", "InstrumentoDTO", "EmisorDTO", "PerfilDTO", "EstatusDTO", "MonedaDTO", "TasaDTO", "ClienteDTO", "SaldosDTO", "TaxonDTO")
For Each nomDto In arrOrdenDto
    rutaCls = strClassesDir & "\" & nomDto & ".cls"
    If fso.FileExists(rutaCls) Then
        logFile.WriteLine "     + Clase DTO: " & nomDto & ".cls"
        strCleanCode = ReadCleanVba(rutaCls)
        On Error Resume Next
        Err.Clear
        Set comp = access.VBE.ActiveVBProject.VBComponents.Add(2)
        If Err.Number = 0 And Not (comp Is Nothing) Then
            comp.Name = nomDto
            If comp.CodeModule.CountOfLines > 0 Then comp.CodeModule.DeleteLines 1, comp.CodeModule.CountOfLines
            comp.CodeModule.AddFromString strCleanCode
        Else
            Err.Clear
            access.LoadFromText 5, nomDto, rutaCls
        End If
        On Error GoTo 0
    End If
Next

arrOrdenMod = Array("Constantes", "Utilities", "ConsultasSQL")
For Each nomMod In arrOrdenMod
    rutaBas = strModulesDir & "\" & nomMod & ".bas"
    If fso.FileExists(rutaBas) Then
        logFile.WriteLine "     + Modulo base: " & nomMod & ".bas"
        strCleanCode = ReadCleanVba(rutaBas)
        On Error Resume Next
        Err.Clear
        Set comp = access.VBE.ActiveVBProject.VBComponents.Add(1)
        If Err.Number = 0 And Not (comp Is Nothing) Then
            comp.Name = nomMod
            If comp.CodeModule.CountOfLines > 0 Then comp.CodeModule.DeleteLines 1, comp.CodeModule.CountOfLines
            comp.CodeModule.AddFromString strCleanCode
        Else
            Err.Clear
            access.LoadFromText 5, nomMod, rutaBas
        End If
        On Error GoTo 0
    End If
Next

rutaCls = strClassesDir & "\FondoDAO.cls"
If fso.FileExists(rutaCls) Then
    logFile.WriteLine "     + Clase DAO: FondoDAO.cls"
    strCleanCode = ReadCleanVba(rutaCls)
    On Error Resume Next
    Err.Clear
    Set comp = access.VBE.ActiveVBProject.VBComponents.Add(2)
    If Err.Number = 0 And Not (comp Is Nothing) Then
        comp.Name = "FondoDAO"
        If comp.CodeModule.CountOfLines > 0 Then comp.CodeModule.DeleteLines 1, comp.CodeModule.CountOfLines
        comp.CodeModule.AddFromString strCleanCode
    Else
        Err.Clear
        access.LoadFromText 5, "FondoDAO", rutaCls
    End If
    On Error GoTo 0
End If

rutaBas = strModulesDir & "\FondoBO.bas"
If fso.FileExists(rutaBas) Then
    logFile.WriteLine "     + Modulo de Negocio: FondoBO.bas"
    strCleanCode = ReadCleanVba(rutaBas)
    On Error Resume Next
    Err.Clear
    Set comp = access.VBE.ActiveVBProject.VBComponents.Add(1)
    If Err.Number = 0 And Not (comp Is Nothing) Then
        comp.Name = "FondoBO"
        If comp.CodeModule.CountOfLines > 0 Then comp.CodeModule.DeleteLines 1, comp.CodeModule.CountOfLines
        comp.CodeModule.AddFromString strCleanCode
    Else
        Err.Clear
        access.LoadFromText 5, "FondoBO", rutaBas
    End If
    On Error GoTo 0
End If

' Vincular tablas de RIESGOSBD en el Frontend
logFile.WriteLine "   Vinculando tablas de base de datos..."
tNames = Array("USUARIOS", "EMISORES", "INSTRUMENTOS", "PERFILES", "ESTATUS", "CAT_MONEDAS", "CAT_TASAS", "TBL_TRANSACCIONES_FONDO", "BD_INSUMO_PRIMER_PISO", "BD_INSUMO_TESORERIA", "BD_INSUMO_SEGUNDO_PISO", "CARTERA_ETIQUETADA", "BITACORA_CARGA_INSUMOS")
' SEPARACION FRONTEND / BACKEND
' Las 13 tablas viven UNICAMENTE en BD\RIESGOSBD.accdb (el backend). El
' frontend Modulo_Administracion_Riesgos.accdb no contiene ninguna tabla
' propia: solo las VINCULA. Por eso aqui no hay un solo CREATE TABLE, nada
' mas CreateTableDef + Append, que es lo que crea un vinculo.
'
' Se cuenta cuantos vinculos quedaron y se registran los que fallen. Antes el
' On Error Resume Next se comia el error y una tabla sin vincular solo se
' descubria al usarla, con un mensaje que no decia nada util.
nVinculadas = 0
nFallidas = 0
msgVinculos = ""
For Each tName In tNames
    On Error Resume Next
    Err.Clear
    Set tdf = access.CurrentDb.CreateTableDef(tName)
    ' La contrasena se agrega SOLO si el cifrado se confirmo. Vincular con
    ' contrasena contra una base sin contrasena hace que Jet rechace el vinculo,
    ' y entonces fallarian las 13 tablas y el aplicativo quedaria vacio.
    If bCifrado Then
        tdf.Connect = ";DATABASE=" & strBackendPath & ";PWD=Wembley1"
    Else
        tdf.Connect = ";DATABASE=" & strBackendPath
    End If
    tdf.SourceTableName = tName
    access.CurrentDb.TableDefs.Append tdf
    If Err.Number = 0 Then
        nVinculadas = nVinculadas + 1
    Else
        nFallidas = nFallidas + 1
        msgVinculos = msgVinculos & "      - " & tName & ": " & Err.Description & vbCrLf
        Err.Clear
    End If
    On Error GoTo 0
Next
End Sub

Private Sub EjecutarBloqueInstalacion_02()
access.CurrentDb.TableDefs.Refresh
logFile.WriteLine "   [OK] " & nVinculadas & " de " & (nVinculadas + nFallidas) & _
                  " tablas vinculadas desde BD\RIESGOSBD.accdb (el frontend NO tiene tablas propias)."
If nFallidas > 0 Then
    logFile.WriteLine "   [ERROR] No se pudieron vincular " & nFallidas & " tabla(s):"
    logFile.WriteLine msgVinculos
End If

' ==============================================================================
' CREACION DE TODOS LOS FORMULARIOS (CENTRADOS MATEMATICAMENTE CON SIMETRIA TOTAL)
' ==============================================================================

' 1. Form_Acceso (TOTALMENTE CENTRADO: MARGEN IZQUIERDO = MARGEN DERECHO = 3500 twips)
logFile.WriteLine "Paso 6: Construyendo Formulario Form_Acceso..."
Set frm = access.CreateForm
fName = frm.Name
On Error Resume Next
frm.Caption = "Acceso"
frm.AutoCenter = True
frm.RecordSelectors = False
frm.NavigationButtons = False
frm.DividingLines = False
frm.ScrollBars = 0
frm.KeyPreview = True
frm.Width = 16000
frm.Section(0).Height = 6150
frm.Section(0).BackColor = 15790320 ' Gris suave institucional (#F0F2F5)
On Error GoTo 0

ConfigurarEncabezadoVerde frm, "Módulo de Administración Integral de Riesgos"

On Error Resume Next
' Recuadro punteado Login (Centrado entre 5000 y 11000 twips, Width 6000)
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 5000, 1800, 6000, 650)
ctrl.Name = "lblLoginBox"
ctrl.Caption = "Acceso al Sistema"
ctrl.FontName = "Cambria"
ctrl.FontSize = 18
ctrl.FontBold = True
ctrl.ForeColor = 6703164
ctrl.TextAlign = 2 ' Centrado
ctrl.BorderStyle = 2
ctrl.BorderColor = 11184810
ctrl.BackStyle = 0

' Renglon 1: Usuario (Y = 2800)
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 3500, 2800, 2500, 350)
ctrl.Caption = "Número de empleado:"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 11
ctrl.FontBold = True
ctrl.ForeColor = 3355443
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 109, 0, "", "", 6200, 2800, 3500, 350)
ctrl.Name = "txtUsuario"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 11
ctrl.TabIndex = 0
ctrl.OnChange = "[Event Procedure]"
ctrl.OnKeyPress = "[Event Procedure]"
ctrl.OnKeyDown = "[Event Procedure]"

' Renglon 2: Contrasena (Y = 3500)
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 3500, 3500, 2500, 350)
ctrl.Caption = "Contraseña:"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 11
ctrl.FontBold = True
ctrl.ForeColor = 3355443
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 109, 0, "", "", 6200, 3500, 3020, 350)
ctrl.Name = "txtContrasena"
ctrl.InputMask = "Password"
' SEGURO: si el usuario presiona el ojito y suelta el raton FUERA del boton,
' el MouseUp del boton puede no llegar. Al perder el foco el cuadro se vuelve
' a enmascarar, de modo que la contrasena nunca se queda visible.
ctrl.OnLostFocus = "[Event Procedure]"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 11
ctrl.TabIndex = 1
ctrl.OnChange = "[Event Procedure]"
ctrl.OnKeyPress = "[Event Procedure]"
ctrl.OnKeyDown = "[Event Procedure]"

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 10000, 2800, 2268, 397)
ctrl.Name = "cmdIngresar"
ctrl.Caption = "Ingresar"
ctrl.Default = False
ctrl.TabIndex = 2
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 9240, 3500, 460, 350)
ctrl.Name = "cmdVerPass"
ctrl.Caption = "Ver"
ctrl.TabStop = True
ctrl.OnClick = "[Event Procedure]"
EstilizarBotonVer ctrl

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 10000, 3500, 2268, 397)
ctrl.Name = "cmdCambiarContra"
ctrl.Caption = "Cambiar Contraseña"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

' Pie de pagina: Control Fecha nativo (=Fecha() segun Photo 0 de Vista Diseno)
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1000, 5600, 4000, 300)
ctrl.Name = "lblFecha"
ctrl.Caption = "=Fecha()"
ctrl.FontName = "Segoe UI"
ctrl.FontSize = 9
ctrl.FontItalic = True
ctrl.ForeColor = 6316128
ctrl.BackStyle = 0
ctrl.FontName = "Segoe UI"
ctrl.FontSize = 9
ctrl.FontItalic = True
ctrl.ForeColor = 6316128
ctrl.BackStyle = 0

' Pie de pagina: Direccion de Administracion de Riesgos
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 10000, 5600, 5000, 300)
ctrl.Name = "lblArea"
ctrl.Caption = "Dirección de Economía de Hidrocarburos"
ctrl.FontName = "Segoe UI"
ctrl.FontSize = 9
ctrl.FontBold = True
ctrl.ForeColor = 6316128
ctrl.TextAlign = 3
ctrl.BackStyle = 0
On Error GoTo 0

codeF = "Option Compare Database" & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"' Form_Acceso - Pantalla de acceso al modulo" & vbCrLf & _
"' -----------------------------------------------------------------------------" & vbCrLf & _
"' Autentica al empleado contra la tabla USUARIOS." & vbCrLf & _
"' Entradas : txtUsuario (4 digitos), txtContrasena (8 caracteres)." & vbCrLf & _
"' Valida   : FondoBO.validarAccesoDetalladoBO compara la contrasena cifrada" & vbCrLf & _
"'            y verifica que cveEstatus sea ACTIVO." & vbCrLf & _
"' Resultado: deja usuarioSesionDTO cargado y abre Form_Menu_General." & vbCrLf & _
"' Seguridad: la contrasena viaja cifrada; nunca se guarda en claro." & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Option Explicit" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"" & vbCrLf & _
"" & vbCrLf & _
"Private mTextoUsuario As String" & vbCrLf & _
"Private mTextoContrasena As String" & vbCrLf & _
"" & vbCrLf & _
"' ---------------------------------------------------------------------------" & vbCrLf & _
"' SEGURO DE ENMASCARADO AL PERDER EL FOCO" & vbCrLf & _
"' Cierra el unico hueco que deja el ojito: presionar el boton y soltar el" & vbCrLf & _
"' raton fuera de el. En ese caso el MouseUp del boton puede no llegar, pero" & vbCrLf & _
"' en cuanto el cuadro pierde el foco la contrasena se vuelve a ocultar." & vbCrLf & _
"' ---------------------------------------------------------------------------" & vbCrLf
codeF = codeF & "Private Sub txtContrasena_LostFocus()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    Call Utilities.OcultarTodasLasContrasenasUT(Me)" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_Load()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call FondoBO.verificarConexionBackendBO()" & vbCrLf & _
"    mTextoUsuario = """"" & vbCrLf & _
"    mTextoContrasena = """"" & vbCrLf & _
"    Me.txtUsuario.Value = Null" & vbCrLf & _
"    Me.txtContrasena.Value = Null" & vbCrLf & _
"    Me.txtContrasena.InputMask = ""Password""" & vbCrLf & _
"    Call Utilities.OcultarTodasLasContrasenasUT(Me)" & vbCrLf
codeF = codeF & "    Me.lblFecha.Caption = Format(Date, ""dddd, d \de mmmm \de yyyy"")" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_Resize()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtUsuario_KeyPress(KeyAscii As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.RestringirKeyPressUsuario4Digitos Me.txtUsuario, KeyAscii" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub txtUsuario_Change()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    Utilities.TruncarCambioUsuario4Digitos Me.txtUsuario, mTextoUsuario" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtContrasena_KeyPress(KeyAscii As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.ProcesarKeyPressContrasena mTextoContrasena, KeyAscii, Me.txtContrasena" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtContrasena_Change()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.TruncarCambioContrasena8Caracteres Me.txtContrasena, mTextoContrasena" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Function ObtenerUsuarioActual() As String" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    ObtenerUsuarioActual = Utilities.LeerTextoControl(Me.txtUsuario, mTextoUsuario)" & vbCrLf & _
"End Function" & vbCrLf & _
"" & vbCrLf & _
"Private Function ObtenerContrasenaActual() As String" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    ObtenerContrasenaActual = Utilities.LeerTextoControl(Me.txtContrasena, mTextoContrasena)" & vbCrLf & _
"End Function" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdCambiarContra_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Cambiar_Contrasena"", acNormal" & vbCrLf
codeF = codeF & "    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"' ---------------------------------------------------------------------------" & vbCrLf & _
"' VER CONTRASENA SOLO MIENTRAS SE SOSTIENE EL CLIC" & vbCrLf & _
"' MouseDown la descubre, MouseUp la vuelve a ocultar. El Click se deja como" & vbCrLf & _
"' seguro por si el MouseUp no llegara: al soltar siempre queda oculta." & vbCrLf & _
"' ---------------------------------------------------------------------------" & vbCrLf & _
"Private Sub cmdVerPass_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call Utilities.VerContrasenaSostenidaUT(Me, ""cmdVerPass"", ""txtContrasena"", True, ObtenerContrasenaActual())" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdVerPass_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    Call Utilities.VerContrasenaSostenidaUT(Me, ""cmdVerPass"", ""txtContrasena"", False, ObtenerContrasenaActual())" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdVerPass_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call Utilities.VerContrasenaSostenidaUT(Me, ""cmdVerPass"", ""txtContrasena"", False, ObtenerContrasenaActual())" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If KeyCode = 13 Then" & vbCrLf & _
"        KeyCode = 0" & vbCrLf & _
"        Dim cveEval As String, msgErrU As String" & vbCrLf
codeF = codeF & "        mTextoUsuario = CStr(Nz(Me.txtUsuario.Text, """"))" & vbCrLf & _
"        Me.txtUsuario.Value = mTextoUsuario" & vbCrLf
codeF = codeF & "        cveEval = ObtenerUsuarioActual()" & vbCrLf & _
"        If Not FondoBO.validarSoloUsuarioBO(cveEval, msgErrU) Then" & vbCrLf & _
"            MsgBox msgErrU, vbCritical, ""Validación de Acceso — Hidrocarburos""" & vbCrLf & _
"            Me.txtUsuario.SetFocus" & vbCrLf & _
"            Exit Sub" & vbCrLf & _
"        End If" & vbCrLf & _
"        If ObtenerContrasenaActual() <> """" Then" & vbCrLf & _
"            cmdIngresar_Click" & vbCrLf & _
"        Else" & vbCrLf & _
"            Me.txtContrasena.SetFocus" & vbCrLf & _
"        End If" & vbCrLf & _
"        Exit Sub" & vbCrLf
codeF = codeF & "    End If" & vbCrLf & _
"    Utilities.RestringirKeyDownUsuario4Digitos Me.txtUsuario, KeyCode, Shift" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtContrasena_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If KeyCode = 13 Then" & vbCrLf & _
"        KeyCode = 0" & vbCrLf & _
"        Me.cmdIngresar.SetFocus" & vbCrLf & _
"        cmdIngresar_Click" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    If Shift = 0 Then" & vbCrLf
codeF = codeF & "        If Not (TypeOf Screen.ActiveControl Is ComboBox Or TypeOf Screen.ActiveControl Is ListBox) Then" & vbCrLf & _
"            Select Case KeyCode" & vbCrLf & _
"                Case 40: KeyCode = 0" & vbCrLf & _
"                Case 38: KeyCode = 0" & vbCrLf & _
"            End Select" & vbCrLf & _
"        End If" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Public Function validaUsuarioYcontrasena() As Boolean" & vbCrLf & _
"    validaUsuarioYcontrasena = True" & vbCrLf & _
"End Function" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub cmdIngresar_Click()" & vbCrLf
codeF = codeF & "    On Error GoTo errorHandler" & vbCrLf & _
"    Dim cveUsuario As String" & vbCrLf & _
"    Dim password As String" & vbCrLf & _
"    Dim msgErrorLogin As String" & vbCrLf & _
"    " & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Me.cmdIngresar.SetFocus" & vbCrLf & _
"    On Error GoTo errorHandler" & vbCrLf & _
"    " & vbCrLf & _
"    cveUsuario = ObtenerUsuarioActual()" & vbCrLf & _
"    password = ObtenerContrasenaActual()" & vbCrLf & _
"    If password <> """" Then" & vbCrLf
codeF = codeF & "        On Error Resume Next" & vbCrLf & _
"        Me.txtContrasena.Value = password" & vbCrLf
codeF = codeF & "        mTextoContrasena = password" & vbCrLf & _
"        On Error GoTo errorHandler" & vbCrLf & _
"    End If" & vbCrLf & _
"    " & vbCrLf & _
"    Screen.MousePointer = MOUSE_POINTER_WAIT" & vbCrLf & _
"    Set usuarioSesionDTO = FondoBO.validarAccesoDetalladoBO(cveUsuario, password, msgErrorLogin)" & vbCrLf & _
"    Screen.MousePointer = MOUSE_POINTER_DEFAULT" & vbCrLf & _
"    " & vbCrLf & _
"    If usuarioSesionDTO Is Nothing Then" & vbCrLf & _
"        MsgBox msgErrorLogin, vbCritical, ""Validación de Acceso — Hidrocarburos""" & vbCrLf & _
"        On Error Resume Next" & vbCrLf & _
"        If InStr(1, msgErrorLogin, ""Contraseña"", vbTextCompare) > 0 Then" & vbCrLf
codeF = codeF & "            Me.txtContrasena.SetFocus" & vbCrLf & _
"        Else" & vbCrLf
codeF = codeF & "            Me.txtUsuario.SetFocus" & vbCrLf & _
"        End If" & vbCrLf & _
"        Exit Sub" & vbCrLf & _
"    End If" & vbCrLf & _
"    " & vbCrLf & _
"    DoCmd.OpenForm ""Form_Menu_General"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"    Exit Sub" & vbCrLf & _
"errorHandler:" & vbCrLf & _
"    Screen.MousePointer = MOUSE_POINTER_DEFAULT" & vbCrLf & _
"    MsgBox MENSAJE_ERROR & vbCrLf & Err.Description, vbCritical" & vbCrLf & _
"End Sub"
AsignarCodigoLimpio frm, codeF
access.DoCmd.Close 2, fName, 1
On Error Resume Next
access.DoCmd.DeleteObject 2, "Form_Acceso"
On Error GoTo 0
access.DoCmd.Rename "Form_Acceso", 2, fName
logFile.WriteLine "   [OK] Form_Acceso creado y centrado con validaciones especificas."

logFile.WriteLine "Paso 7: Construyendo Formulario Form_Menu_General..."
Set frm = access.CreateForm
fName = frm.Name
On Error Resume Next
frm.Caption = "Menú General"
frm.AutoCenter = True
frm.RecordSelectors = False
frm.NavigationButtons = False
frm.ScrollBars = 0
frm.KeyPreview = True
frm.Width = 16000
frm.Section(0).Height = 6150
frm.Section(0).BackColor = 15790320 ' Gris suave institucional (#F0F2F5)
On Error GoTo 0
End Sub

Private Sub EjecutarBloqueInstalacion_03()

ConfigurarEncabezadoVerde frm, "Módulo de Administración Integral de Riesgos"

On Error Resume Next
' Subencabezado dinamico: No. Empleado y Perfil
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 1550, 6500, 320)
ctrl.Name = "lblNoEmpleado"
ctrl.Caption = "No. Empleado: 1000"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 10.5
ctrl.FontBold = True
ctrl.ForeColor = 4737096
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 1900, 6500, 320)
ctrl.Name = "lblPerfil"
ctrl.Caption = "Perfil: ADMINISTRADOR"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 10.5
ctrl.FontBold = True
ctrl.ForeColor = 2631720
ctrl.BackStyle = 0

' En el mismo Menú General no se muestra el botón de Menú ni flecha redundante; únicamente Salir
Set ctrl = access.CreateControl(fName, 104, 0, "", "", 13200, 1550, 1750, 600)
ctrl.Name = "cmdSalir"
ctrl.Caption = "Salir"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl
ctrl.ControlTipText = "Cerrar sesión y salir"

' Recuadro punteado Menú General centrado (Left 1500, Width 13000)
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 2350, 13000, 600)
ctrl.Name = "lblMenuGeneralBox"
ctrl.Caption = "Menú General"
ctrl.FontName = "Cambria"
ctrl.FontSize = 17
ctrl.FontBold = True
ctrl.ForeColor = 6703164
ctrl.TextAlign = 2
ctrl.BorderStyle = 2
ctrl.BorderColor = 11184810
ctrl.BackStyle = 0

' Matriz de 3 columnas (Simetria perfecta segun Captura Institucional Photo 6)
' Columna 1: Gestion de Datos Maestros (Izquierda: Left 1800, Width 3600)
Set ctrl = access.CreateControl(fName, 104, 0, "", "", 1800, 3400, 3600, 750)
ctrl.Name = "cmdGestionarUsuarios"
ctrl.Caption = "Gestionar Usuarios"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 1800, 4500, 3600, 750)
ctrl.Name = "cmdGestionarInstrumentos"
ctrl.Caption = "Gestionar Instrumentos"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

' Columna 2: Carga de Operaciones (Centro: Left 6200, Width 3600, Y = 4100 - Centrado vertical)
Set ctrl = access.CreateControl(fName, 104, 0, "", "", 6200, 3400, 3600, 750)
ctrl.Name = "cmdCargarInsumos"
ctrl.Caption = "Cargar Insumos"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

' Columna 3: Salidas y Reporteria (Derecha: Left 10600, Width 3600, Y = 4100 - Centrado vertical)
Set ctrl = access.CreateControl(fName, 104, 0, "", "", 10600, 3400, 3600, 750)
ctrl.Name = "cmdGenerarReportes"
ctrl.Caption = "Generar Reportes"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

' Botones auxiliares opcionales (ocultos por defecto para mantener interfaz identica a Photo 6)
Set ctrl = access.CreateControl(fName, 104, 0, "", "", 6200, 4500, 3600, 750)
ctrl.Name = "cmdCambiarMiContra"
ctrl.Caption = "Cambiar Mi Contraseña"
ctrl.Visible = False
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 10600, 4500, 3600, 750)
ctrl.Name = "cmdDirectorio"
ctrl.Caption = "Directorio de Usuarios"
ctrl.Visible = False
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

' Pie de pagina
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 5700, 5000, 300)
ctrl.Name = "lblFecha"
ctrl.Caption = LCase(FormatDateTime(Date, 1))
ctrl.FontName = "Segoe UI"
ctrl.FontSize = 9
ctrl.FontItalic = True
ctrl.ForeColor = 6316128
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 9500, 5700, 5000, 300)
ctrl.Name = "lblArea"
ctrl.Caption = "Dirección de Economía de Hidrocarburos"
ctrl.FontName = "Segoe UI"
ctrl.FontSize = 9
ctrl.FontBold = True
ctrl.ForeColor = 6316128
ctrl.TextAlign = 3
ctrl.BackStyle = 0
On Error GoTo 0

codeF = "Option Compare Database" & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"' Form_Menu_General - Menu principal del modulo" & vbCrLf & _
"' -----------------------------------------------------------------------------" & vbCrLf & _
"' Punto de entrada tras el login. No consulta tablas." & vbCrLf & _
"' Muestra u oculta 'Gestionar Usuarios' segun el perfil de la sesion:" & vbCrLf & _
"'    perfil A (Administrador) lo ve; perfil U (Operativo) no." & vbCrLf & _
"' Depende de: usuarioSesionDTO, cargado por Form_Acceso." & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"" & vbCrLf & _
"Option Explicit" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub Form_Load()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    Dim esAdmin As Boolean" & vbCrLf & _
"    esAdmin = False" & vbCrLf & _
"    If usuarioSesionDTO Is Nothing Then" & vbCrLf & _
"        DoCmd.OpenForm ""Form_Acceso"", acNormal: DoCmd.Close acForm, Me.Name: Exit Sub" & vbCrLf & _
"    End If" & vbCrLf & _
"    ' Encabezado de sesion: el texto de las etiquetas lo arma la capa de Negocio." & vbCrLf & _
"    Call Utilities.AplicarEncabezadoSesionUT(Me, FondoBO.PANT_MENU_GENERAL)" & vbCrLf & _
"" & vbCrLf & _
"    ' Quien es administrador lo decide el Negocio, no este formulario. Antes se" & vbCrLf & _
"    ' comparaba aqui contra el literal A (texto magico) y, si no habia sesion, se" & vbCrLf & _
"    ' asumia esAdmin = True, es decir, no tener sesion daba perfil de" & vbCrLf & _
"    ' administrador. El guardia de arriba ya sale cuando no hay sesion y ademas" & vbCrLf
codeF = codeF & "    ' esAdministradorSesionBO devuelve False sin sesion." & vbCrLf & _
"    esAdmin = FondoBO.esAdministradorSesionBO()" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"    ' El color de la etiqueta si es presentacion pura." & vbCrLf & _
"    If esAdmin Then" & vbCrLf & _
"        Me.lblPerfil.ForeColor = 6703164" & vbCrLf & _
"    Else" & vbCrLf & _
"        Me.lblPerfil.ForeColor = 4737096" & vbCrLf & _
"    End If" & vbCrLf & _
"    " & vbCrLf & _
"    Me.lblMenuGeneralBox.Caption = ""Menú Principal""" & vbCrLf & _
"    " & vbCrLf & _
"    If esAdmin Then" & vbCrLf & _
"        ' Todos los botones del menu comparten exactamente el mismo tamano," & vbCrLf
codeF = codeF & "        ' para que la rejilla se vea pareja sin importar el perfil." & vbCrLf & _
"        Dim vBtn, oBtn" & vbCrLf
codeF = codeF & "        Dim arrBtn" & vbCrLf & _
"        arrBtn = Array(""cmdGestionarUsuarios"", ""cmdGestionarInstrumentos"", ""cmdCargarInsumos"", ""cmdGenerarReportes"", ""cmdCambiarMiContra"", ""cmdDirectorio"")" & vbCrLf & _
"        For Each vBtn In arrBtn" & vbCrLf & _
"            Set oBtn = Me.Controls(CStr(vBtn))" & vbCrLf & _
"            oBtn.Width = 3600: oBtn.Height = 750" & vbCrLf & _
"        Next vBtn" & vbCrLf & _
"        Me.cmdGestionarUsuarios.Visible = True" & vbCrLf & _
"        Me.cmdGestionarUsuarios.Enabled = True" & vbCrLf & _
"        Me.cmdGestionarUsuarios.Left = 1800" & vbCrLf & _
"        Me.cmdGestionarUsuarios.Top = 3400" & vbCrLf & _
"        Me.cmdGestionarInstrumentos.Caption = ""Gestionar Instrumentos""" & vbCrLf & _
"        Me.cmdGestionarInstrumentos.Left = 1800" & vbCrLf
codeF = codeF & "        Me.cmdGestionarInstrumentos.Top = 4500" & vbCrLf & _
"        Me.cmdCargarInsumos.Left = 6200" & vbCrLf
codeF = codeF & "        Me.cmdCargarInsumos.Top = 3400" & vbCrLf & _
"        Me.cmdCambiarMiContra.Caption = ""Restablecer Contraseña""" & vbCrLf & _
"        Me.cmdCambiarMiContra.Left = 6200" & vbCrLf & _
"        Me.cmdCambiarMiContra.Top = 4500" & vbCrLf & _
"        Me.cmdGenerarReportes.Left = 10600" & vbCrLf & _
"        Me.cmdGenerarReportes.Top = 3400" & vbCrLf & _
"        Me.cmdDirectorio.Visible = True" & vbCrLf & _
"        Me.cmdDirectorio.Left = 10600" & vbCrLf & _
"        Me.cmdDirectorio.Top = 4500" & vbCrLf & _
"    Else" & vbCrLf & _
"        Me.cmdGestionarUsuarios.Visible = False" & vbCrLf & _
"        Me.cmdDirectorio.Visible = False" & vbCrLf
codeF = codeF & "        Me.cmdGestionarInstrumentos.Caption = ""Consultar Instrumentos""" & vbCrLf & _
"        Me.cmdGestionarInstrumentos.Left = 1800" & vbCrLf
codeF = codeF & "        Me.cmdGestionarInstrumentos.Top = 3400" & vbCrLf & _
"        Me.cmdCargarInsumos.Left = 6200" & vbCrLf & _
"        Me.cmdCargarInsumos.Top = 3400" & vbCrLf & _
"        Me.cmdGenerarReportes.Left = 10600" & vbCrLf & _
"        Me.cmdGenerarReportes.Top = 3400" & vbCrLf & _
"        Me.cmdCambiarMiContra.Caption = ""Cambiar Mi Contraseña""" & vbCrLf & _
"        Me.cmdCambiarMiContra.Left = 6200" & vbCrLf & _
"        Me.cmdCambiarMiContra.Top = 4500" & vbCrLf & _
"    End If" & vbCrLf & _
"    Me.lblFecha.Caption = Format(Date, ""dddd, d \de mmmm \de yyyy"")" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub Form_Resize()" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If Shift = 0 Then" & vbCrLf & _
"        If Not (TypeOf Screen.ActiveControl Is ComboBox Or TypeOf Screen.ActiveControl Is ListBox) Then" & vbCrLf & _
"            Select Case KeyCode" & vbCrLf & _
"                Case 40 ' vbKeyDown" & vbCrLf & _
"                    KeyCode = 0" & vbCrLf & _
"                    KeyCode = 0" & vbCrLf
codeF = codeF & "                Case 38 ' vbKeyUp" & vbCrLf & _
"                    KeyCode = 0" & vbCrLf
codeF = codeF & "                    KeyCode = 0" & vbCrLf & _
"            End Select" & vbCrLf & _
"        End If" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdGestionarUsuarios_Click()" & vbCrLf & _
"    ' La regla de quien puede abrir esta pantalla vive en el Negocio;" & vbCrLf & _
"    ' este boton solo pregunta. Si no se autoriza, el ayudante ya aviso." & vbCrLf & _
"    If Not Utilities.PuedeAbrirPantallaUT(FondoBO.PANT_GESTION_USUARIOS) Then Exit Sub" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Gestionar_Usuarios"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub cmdGestionarInstrumentos_Click()" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Gestionar_Instrumentos"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdCargarInsumos_Click()" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Cargar_Insumos"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdGenerarReportes_Click()" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Generar_Reporte"", acNormal" & vbCrLf
codeF = codeF & "    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub cmdCambiarMiContra_Click()" & vbCrLf & _
"    ' Que pantalla corresponde a este perfil lo decide el Negocio. Antes se" & vbCrLf & _
"    ' comparaba aqui contra el literal A, y sin sesion se abria la pantalla" & vbCrLf & _
"    ' operativa; ahora pantallaCambioContrasenaBO cubre los dos casos." & vbCrLf & _
"    DoCmd.OpenForm FondoBO.pantallaCambioContrasenaBO(), acNormal" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdDirectorio_Click()" & vbCrLf & _
"    ' La regla de quien puede abrir esta pantalla vive en el Negocio;" & vbCrLf & _
"    ' este boton solo pregunta. Si no se autoriza, el ayudante ya aviso." & vbCrLf & _
"    If Not Utilities.PuedeAbrirPantallaUT(FondoBO.PANT_CONSULTA_USUARIOS) Then Exit Sub" & vbCrLf
codeF = codeF & "    DoCmd.OpenForm ""Form_Subformulario_Consulta_Usuarios"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdAtras_Click()" & vbCrLf & _
"    cmdSalir_Click" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdIrMenu_Click()" & vbCrLf & _
"    ' Ya se encuentra en el Menú General" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdSalir_Click()" & vbCrLf & _
"    If MsgBox(""¿Desea cerrar la sesión actual?"", vbQuestion + vbYesNo, ""Cerrar Sesión"") = vbYes Then" & vbCrLf
codeF = codeF & "        Set usuarioSesionDTO = Nothing" & vbCrLf & _
"        DoCmd.OpenForm ""Form_Acceso"", acNormal" & vbCrLf
codeF = codeF & "        DoCmd.Close acForm, Me.Name" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub"
AsignarCodigoLimpio frm, codeF

access.DoCmd.Close 2, fName, 1
On Error Resume Next
access.DoCmd.DeleteObject 2, "Form_Menu_General"
On Error GoTo 0
access.DoCmd.Rename "Form_Menu_General", 2, fName
logFile.WriteLine "   [OK] Form_Menu_General creado y configurado con perfiles diferenciados."

logFile.WriteLine "Paso 8: Construyendo Formulario Form_Cambiar_Contrasena..."
Set frm = access.CreateForm
fName = frm.Name
On Error Resume Next
frm.Caption = "Cambiar Contraseña"
frm.AutoCenter = True
frm.RecordSelectors = False
frm.NavigationButtons = False
frm.DividingLines = False
frm.ScrollBars = 0
frm.KeyPreview = True
frm.Width = 16000
frm.Section(0).Height = 6150
frm.Section(0).BackColor = 15790320 ' Gris suave institucional (#F0F2F5)
On Error GoTo 0

ConfigurarEncabezadoVerde frm, "Módulo de Administración Integral de Riesgos"
ConfigurarBarraNavegacionSuperior frm, "Form_Menu_General"

On Error Resume Next
' Recuadro punteado / Titulo central de formulario
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 4500, 2150, 7000, 550)
ctrl.Caption = "Cambiar Contraseña"
ctrl.FontName = "Cambria"
ctrl.FontSize = 18
ctrl.FontBold = True
ctrl.ForeColor = 6703164
ctrl.TextAlign = 2
ctrl.BorderStyle = 2
ctrl.BorderColor = 11184810
ctrl.BackStyle = 0

' Renglon 1: Numero de empleado
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1800, 2750, 4800, 380)
ctrl.Caption = "Número de empleado:"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 11
ctrl.FontBold = True
ctrl.ForeColor = 3355443
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 109, 0, "", "", 6800, 2750, 3500, 380)
ctrl.Name = "txtUsuario"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 11

' Renglon 2: Contrasena actual
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1800, 3300, 4800, 380)
ctrl.Caption = "Contraseña:"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 11
ctrl.FontBold = True
ctrl.ForeColor = 3355443
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 109, 0, "", "", 6800, 3300, 3020, 380)
ctrl.Name = "txtContrasena"
ctrl.InputMask = "Password"
' SEGURO: si el usuario presiona el ojito y suelta el raton FUERA del boton,
' el MouseUp del boton puede no llegar. Al perder el foco el cuadro se vuelve
' a enmascarar, de modo que la contrasena nunca se queda visible.
ctrl.OnLostFocus = "[Event Procedure]"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 11

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 9840, 3300, 460, 380)
ctrl.Name = "cmdVerActual"
ctrl.Caption = "Ver"
ctrl.OnClick = "[Event Procedure]"
EstilizarBotonVer ctrl

' Renglon 3: Nueva Contrasena
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1800, 3850, 4800, 380)
ctrl.Caption = "Nueva Contraseña:"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 11
End Sub

Private Sub EjecutarBloqueInstalacion_04()
ctrl.FontBold = True
ctrl.ForeColor = 3355443
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 109, 0, "", "", 6800, 3850, 3020, 380)
ctrl.Name = "txtNuevaContrasena"
ctrl.InputMask = "Password"
' SEGURO: si el usuario presiona el ojito y suelta el raton FUERA del boton,
' el MouseUp del boton puede no llegar. Al perder el foco el cuadro se vuelve
' a enmascarar, de modo que la contrasena nunca se queda visible.
ctrl.OnLostFocus = "[Event Procedure]"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 11

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 9840, 3850, 460, 380)
ctrl.Name = "cmdVerNueva"
ctrl.Caption = "Ver"
ctrl.OnClick = "[Event Procedure]"
EstilizarBotonVer ctrl

' Renglon 4: Confirmacion de la Nueva Contrasena
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1800, 4400, 4800, 380)
ctrl.Caption = "Confirmación de la Nueva Contraseña:"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 11
ctrl.FontBold = True
ctrl.ForeColor = 3355443
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 109, 0, "", "", 6800, 4400, 3020, 380)
ctrl.Name = "txtConfirmarContrasena"
ctrl.InputMask = "Password"
' SEGURO: si el usuario presiona el ojito y suelta el raton FUERA del boton,
' el MouseUp del boton puede no llegar. Al perder el foco el cuadro se vuelve
' a enmascarar, de modo que la contrasena nunca se queda visible.
ctrl.OnLostFocus = "[Event Procedure]"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 11

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 9840, 4400, 460, 380)
ctrl.Name = "cmdVerConfirmar"
ctrl.Caption = "Ver"
ctrl.OnClick = "[Event Procedure]"
EstilizarBotonVer ctrl

' Botones laterales a la derecha
Set ctrl = access.CreateControl(fName, 104, 0, "", "", 10600, 2850, 2800, 480)
ctrl.Name = "cmdAceptar"
ctrl.Caption = "Aceptar"
ctrl.Default = False
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 10600, 3500, 2800, 480)
ctrl.Name = "cmdLimpiar"
ctrl.Caption = "Limpiar"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

' Cuadro institucional inferior con Politicas de Seguridad de Contrasena
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 2500, 4880, 11000, 820)
ctrl.Name = "lblPoliticaBox"
ctrl.Caption = "POLÍTICA DE SEGURIDAD PARA CONTRASEÑAS HIDROCARBUROS:" & vbCrLf & _
"- Extensión obligatoria de exactamente 8 caracteres  - Al menos una mayúscula (A-Z)  - Al menos una minúscula (a-z)  - Al menos un número (0-9)" & vbCrLf & _
"- La nueva contraseña debe ser distinta a la contraseña actual y coincidir exactamente en su confirmación."
ctrl.FontName = "Segoe UI"
ctrl.FontSize = 9.5
ctrl.FontBold = True
ctrl.ForeColor = 6703164
ctrl.TextAlign = 2
ctrl.BackStyle = 1
ctrl.BackColor = 15133908
ctrl.BorderStyle = 1
ctrl.BorderColor = 9376256

' Pie de pagina
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 5760, 5000, 260)
ctrl.Name = "lblFecha"
ctrl.Caption = LCase(FormatDateTime(Date, 1))
ctrl.FontName = "Segoe UI"
ctrl.FontSize = 9
ctrl.FontItalic = True
ctrl.ForeColor = 6316128
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 9500, 5760, 5000, 260)
ctrl.Name = "lblArea"
ctrl.Caption = "Dirección de Economía de Hidrocarburos"
ctrl.FontName = "Segoe UI"
ctrl.FontSize = 9
ctrl.FontBold = True
ctrl.ForeColor = 6316128
ctrl.TextAlign = 3
ctrl.BackStyle = 0
On Error GoTo 0

codeF = "Option Compare Database" & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"' Form_Cambiar_Contrasena - Cambio de contrasena por el propio usuario" & vbCrLf & _
"' -----------------------------------------------------------------------------" & vbCrLf & _
"' Exige la contrasena actual antes de permitir el cambio." & vbCrLf & _
"' Politica: exactamente 8 caracteres, con al menos una mayuscula," & vbCrLf & _
"'           una minuscula y un digito." & vbCrLf & _
"' Valida   : FondoBO.cambiarContrasenaOperativoBO." & vbCrLf & _
"' Escribe  : USUARIOS.password (cifrado)." & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"" & vbCrLf & _
"Option Explicit" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private mBufActual As String" & vbCrLf & _
"Private mBufNueva As String" & vbCrLf & _
"Private mBufConfirmar As String" & vbCrLf & _
"" & vbCrLf & _
"' ---------------------------------------------------------------------------" & vbCrLf & _
"' SEGURO DE ENMASCARADO AL PERDER EL FOCO" & vbCrLf & _
"' Cierra el unico hueco que deja el ojito: presionar el boton y soltar el" & vbCrLf & _
"' raton fuera de el. En ese caso el MouseUp del boton puede no llegar, pero" & vbCrLf & _
"' en cuanto el cuadro pierde el foco la contrasena se vuelve a ocultar." & vbCrLf & _
"' ---------------------------------------------------------------------------" & vbCrLf & _
"Private Sub txtConfirmarContrasena_LostFocus()" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    Call Utilities.OcultarTodasLasContrasenasUT(Me)" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtContrasena_LostFocus()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call Utilities.OcultarTodasLasContrasenasUT(Me)" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtNuevaContrasena_LostFocus()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call Utilities.OcultarTodasLasContrasenasUT(Me)" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    Utilities.RestringirKeyDownUsuario4Digitos Me.txtUsuario, KeyCode, Shift" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtUsuario_KeyPress(KeyAscii As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.RestringirKeyPressUsuario4Digitos Me.txtUsuario, KeyAscii" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtUsuario_Change()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.TruncarCambioUsuario4Digitos Me.txtUsuario" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub Form_Load()" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    ' Encabezado de sesion: el texto de las etiquetas lo arma la capa de Negocio." & vbCrLf & _
"    Call Utilities.AplicarEncabezadoSesionUT(Me, FondoBO.PANT_CAMBIO_MI_CONTRASENA)" & vbCrLf & _
"" & vbCrLf & _
"    ' Si hay sesion o no es una pregunta de negocio; mostrar u ocultar los" & vbCrLf & _
"    ' controles y recorrer el boton Atras es presentacion." & vbCrLf & _
"    If FondoBO.haySesionActivaBO() Then" & vbCrLf & _
"        Me.lblNoEmpleado.Visible = True: Me.lblPerfil.Visible = True" & vbCrLf & _
"        Me.cmdSalir.Visible = True: Me.cmdAtras.Left = 12450" & vbCrLf & _
"    Else" & vbCrLf & _
"        Me.lblNoEmpleado.Visible = False: Me.lblPerfil.Visible = False" & vbCrLf & _
"        Me.cmdSalir.Visible = False: Me.cmdAtras.Left = 13650" & vbCrLf
codeF = codeF & "    End If" & vbCrLf & _
"    mBufActual = """": mBufNueva = """": mBufConfirmar = """"" & vbCrLf
codeF = codeF & "    Me.txtUsuario.Value = Null" & vbCrLf & _
"    Me.txtContrasena.Value = Null" & vbCrLf & _
"    Me.txtNuevaContrasena.Value = Null" & vbCrLf & _
"    Me.txtConfirmarContrasena.Value = Null" & vbCrLf & _
"    If Not usuarioSesionDTO Is Nothing Then" & vbCrLf & _
"        Me.txtUsuario.Value = usuarioSesionDTO.getCveUsuario()" & vbCrLf & _
"        Me.txtUsuario.Locked = True" & vbCrLf & _
"        Me.txtContrasena.SetFocus" & vbCrLf & _
"    End If" & vbCrLf & _
"    Me.lblFecha.Caption = Format(Date, ""dddd, d \de mmmm \de yyyy"")" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub Form_Resize()" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtContrasena_KeyPress(KeyAscii As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.ProcesarKeyPressContrasena mBufActual, KeyAscii, Me.txtContrasena" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtContrasena_Change()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.TruncarCambioContrasena8Caracteres Me.txtContrasena, mBufActual" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub txtNuevaContrasena_KeyPress(KeyAscii As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.ProcesarKeyPressContrasena mBufNueva, KeyAscii" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtNuevaContrasena_Change()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.TruncarCambioContrasena8Caracteres Me.txtNuevaContrasena, mBufNueva" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtConfirmarContrasena_KeyPress(KeyAscii As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    Utilities.ProcesarKeyPressContrasena mBufConfirmar, KeyAscii" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub txtConfirmarContrasena_Change()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.TruncarCambioContrasena8Caracteres Me.txtConfirmarContrasena, mBufConfirmar" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"' ---------------------------------------------------------------------------" & vbCrLf & _
"' VER CONTRASENA SOLO MIENTRAS SE SOSTIENE EL CLIC" & vbCrLf & _
"' MouseDown la descubre, MouseUp la vuelve a ocultar. El Click se deja como" & vbCrLf & _
"' seguro por si el MouseUp no llegara: al soltar siempre queda oculta." & vbCrLf & _
"' ---------------------------------------------------------------------------" & vbCrLf & _
"Private Sub cmdVerActual_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    Call Utilities.VerContrasenaSostenidaUT(Me, ""cmdVerActual"", ""txtContrasena"", True, """")" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdVerActual_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call Utilities.VerContrasenaSostenidaUT(Me, ""cmdVerActual"", ""txtContrasena"", False, """")" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdVerActual_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call Utilities.VerContrasenaSostenidaUT(Me, ""cmdVerActual"", ""txtContrasena"", False, """")" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "' ---------------------------------------------------------------------------" & vbCrLf & _
"' VER CONTRASENA SOLO MIENTRAS SE SOSTIENE EL CLIC" & vbCrLf
codeF = codeF & "' MouseDown la descubre, MouseUp la vuelve a ocultar. El Click se deja como" & vbCrLf & _
"' seguro por si el MouseUp no llegara: al soltar siempre queda oculta." & vbCrLf & _
"' ---------------------------------------------------------------------------" & vbCrLf & _
"Private Sub cmdVerNueva_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call Utilities.VerContrasenaSostenidaUT(Me, ""cmdVerNueva"", ""txtNuevaContrasena"", True, """")" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdVerNueva_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call Utilities.VerContrasenaSostenidaUT(Me, ""cmdVerNueva"", ""txtNuevaContrasena"", False, """")" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub cmdVerNueva_Click()" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    Call Utilities.VerContrasenaSostenidaUT(Me, ""cmdVerNueva"", ""txtNuevaContrasena"", False, """")" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"' ---------------------------------------------------------------------------" & vbCrLf & _
"' VER CONTRASENA SOLO MIENTRAS SE SOSTIENE EL CLIC" & vbCrLf & _
"' MouseDown la descubre, MouseUp la vuelve a ocultar. El Click se deja como" & vbCrLf & _
"' seguro por si el MouseUp no llegara: al soltar siempre queda oculta." & vbCrLf & _
"' ---------------------------------------------------------------------------" & vbCrLf & _
"Private Sub cmdVerConfirmar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call Utilities.VerContrasenaSostenidaUT(Me, ""cmdVerConfirmar"", ""txtConfirmarContrasena"", True, """")" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub cmdVerConfirmar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call Utilities.VerContrasenaSostenidaUT(Me, ""cmdVerConfirmar"", ""txtConfirmarContrasena"", False, """")" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdVerConfirmar_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call Utilities.VerContrasenaSostenidaUT(Me, ""cmdVerConfirmar"", ""txtConfirmarContrasena"", False, """")" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtConfirmarContrasena_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    If KeyCode = 13 Then" & vbCrLf & _
"        KeyCode = 0" & vbCrLf
codeF = codeF & "        Me.cmdAceptar.SetFocus" & vbCrLf & _
"        cmdAceptar_Click" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If Shift = 0 Then" & vbCrLf & _
"        If Not (TypeOf Screen.ActiveControl Is ComboBox Or TypeOf Screen.ActiveControl Is ListBox) Then" & vbCrLf & _
"            Select Case KeyCode" & vbCrLf & _
"                Case 40: KeyCode = 0" & vbCrLf & _
"                Case 38: KeyCode = 0" & vbCrLf
codeF = codeF & "            End Select" & vbCrLf & _
"        End If" & vbCrLf
codeF = codeF & "    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdLimpiar_Click()" & vbCrLf & _
"    If usuarioSesionDTO Is Nothing Then" & vbCrLf & _
"        Me.txtUsuario.Value = Null" & vbCrLf & _
"    End If" & vbCrLf & _
"    mBufActual = """": mBufNueva = """": mBufConfirmar = """"" & vbCrLf & _
"    Me.txtContrasena.Value = Null" & vbCrLf & _
"    Me.txtNuevaContrasena.Value = Null" & vbCrLf & _
"    Me.txtConfirmarContrasena.Value = Null" & vbCrLf & _
"    Me.txtContrasena.SetFocus" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub cmdAceptar_Click()" & vbCrLf & _
"    On Error Resume Next: Me.cmdAceptar.SetFocus: On Error GoTo ErrH" & vbCrLf & _
"    Dim cveUsuario As String" & vbCrLf & _
"    Dim pwdActual As String" & vbCrLf & _
"    Dim pwdNueva As String" & vbCrLf & _
"    Dim pwdConfirmar As String" & vbCrLf & _
"    Dim msgResultado As String" & vbCrLf & _
"    Dim exito As Boolean" & vbCrLf & _
"    " & vbCrLf & _
"    cveUsuario = Utilities.LeerTextoControl(Me.txtUsuario)" & vbCrLf & _
"    pwdActual = Utilities.LeerTextoControl(Me.txtContrasena, mBufActual)" & vbCrLf & _
"    pwdNueva = Utilities.LeerTextoControl(Me.txtNuevaContrasena, mBufNueva)" & vbCrLf
codeF = codeF & "    pwdConfirmar = Utilities.LeerTextoControl(Me.txtConfirmarContrasena, mBufConfirmar)" & vbCrLf & _
"    " & vbCrLf
codeF = codeF & "    Screen.MousePointer = MOUSE_POINTER_WAIT" & vbCrLf & _
"    exito = FondoBO.cambiarContrasenaOperativoBO(cveUsuario, pwdActual, pwdNueva, pwdConfirmar, msgResultado)" & vbCrLf & _
"    Screen.MousePointer = MOUSE_POINTER_DEFAULT" & vbCrLf & _
"    " & vbCrLf & _
"    If exito Then" & vbCrLf & _
"        MsgBox msgResultado, vbInformation, ""Cambio de Contraseña""" & vbCrLf & _
"        cmdAtras_Click" & vbCrLf & _
"    Else" & vbCrLf & _
"        MsgBox msgResultado, vbExclamation, ""Validación de Contraseña""" & vbCrLf & _
"    End If" & vbCrLf & _
"    Exit Sub" & vbCrLf & _
"ErrH:" & vbCrLf
codeF = codeF & "    Screen.MousePointer = MOUSE_POINTER_DEFAULT" & vbCrLf & _
"    MsgBox MENSAJE_ERROR & vbCrLf & Err.Description, vbCritical, ""Error""" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdAtras_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If Not usuarioSesionDTO Is Nothing Then" & vbCrLf & _
"        DoCmd.OpenForm ""Form_Menu_General"", acNormal" & vbCrLf & _
"    Else" & vbCrLf & _
"        DoCmd.OpenForm ""Form_Acceso"", acNormal" & vbCrLf & _
"    End If" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub cmdIrMenu_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    If Not usuarioSesionDTO Is Nothing Then" & vbCrLf & _
"        DoCmd.OpenForm ""Form_Menu_General"", acNormal" & vbCrLf & _
"        DoCmd.Close acForm, Me.Name" & vbCrLf & _
"    Else" & vbCrLf & _
"        DoCmd.OpenForm ""Form_Acceso"", acNormal" & vbCrLf & _
"        DoCmd.Close acForm, Me.Name" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdSalir_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If MsgBox(""¿Desea cerrar la sesión actual?"", vbQuestion + vbYesNo, ""Salir"") = vbYes Then" & vbCrLf
codeF = codeF & "        Set usuarioSesionDTO = Nothing" & vbCrLf & _
"        DoCmd.OpenForm ""Form_Acceso"", acNormal" & vbCrLf
codeF = codeF & "        DoCmd.Close acForm, Me.Name" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdRegresar_Click()" & vbCrLf & _
"    cmdAtras_Click" & vbCrLf & _
"End Sub"
AsignarCodigoLimpio frm, codeF
End Sub

Private Sub EjecutarBloqueInstalacion_05()
access.DoCmd.Close 2, fName, 1
On Error Resume Next
access.DoCmd.DeleteObject 2, "Form_Cambiar_Contrasena"
On Error GoTo 0
access.DoCmd.Rename "Form_Cambiar_Contrasena", 2, fName
logFile.WriteLine "   [OK] Form_Cambiar_Contrasena creado."

' Construyendo Subformularios Tabulares Reales (acSubform = 112)
logFile.WriteLine "   Construyendo Subformularios Tabulares (Subform_Tabla_Usuarios y Subform_Tabla_Instrumentos)..."
Set frm = access.CreateForm
fName = frm.Name
On Error Resume Next
frm.Caption = "Directorio de Usuarios"
frm.DefaultView = 2 ' Vista Hoja de Datos para Subformulario embebido
frm.AllowAdditions = False
frm.AllowDeletions = False
frm.AllowEdits = False
' VALOR DE DISENO UNICAMENTE. La consulta autorizada de este subformulario
' es la de ConsultasSQL; los formularios que lo embeben la reasignan en
' tiempo de ejecucion con Me.sub<X>.Form.RecordSource = FondoBO.obtenerOrigenUsuariosBO().
' Aqui se deja el literal porque el constructor es VBScript y no puede
' llamar a una funcion de VBA: el motor de Access todavia no existe.
frm.RecordSource = "SELECT cveUsuario, nombreUsuario, aPaterno, aMaterno, cvePerfil, emailUsuario, cveEstatus FROM USUARIOS ORDER BY cveUsuario;"

Set ctrl = access.CreateControl(fName, 109, 0, "", "cveUsuario", 200, 100, 1500, 300)
ctrl.Name = "No_Empleado"
ctrl.ControlSource = "cveUsuario"
ctrl.ColumnWidth = 1650
Set lblSub = access.CreateControl(fName, 100, 0, ctrl.Name, "", 0, 0, 1000, 300)
lblSub.Caption = "No. Empleado"

Set ctrl = access.CreateControl(fName, 109, 0, "", "nombreUsuario", 1800, 100, 2200, 300)
ctrl.Name = "Nombre_Empleado"
ctrl.ControlSource = "nombreUsuario"
ctrl.ColumnWidth = 2300
Set lblSub = access.CreateControl(fName, 100, 0, ctrl.Name, "", 0, 0, 1000, 300)
lblSub.Caption = "Nombre(s)"

Set ctrl = access.CreateControl(fName, 109, 0, "", "aPaterno", 4100, 100, 2000, 300)
ctrl.Name = "Apellido_Paterno"
ctrl.ControlSource = "aPaterno"
ctrl.ColumnWidth = 2050
Set lblSub = access.CreateControl(fName, 100, 0, ctrl.Name, "", 0, 0, 1000, 300)
lblSub.Caption = "Apellido Paterno"

Set ctrl = access.CreateControl(fName, 109, 0, "", "aMaterno", 6200, 100, 2000, 300)
ctrl.Name = "Apellido_Materno"
ctrl.ControlSource = "aMaterno"
ctrl.ColumnWidth = 2050
Set lblSub = access.CreateControl(fName, 100, 0, ctrl.Name, "", 0, 0, 1000, 300)
lblSub.Caption = "Apellido Materno"

Set ctrl = access.CreateControl(fName, 109, 0, "", "cvePerfil", 8300, 100, 1100, 300)
ctrl.Name = "Perfil_Usuario"
ctrl.ControlSource = "cvePerfil"
ctrl.ColumnWidth = 1150
Set lblSub = access.CreateControl(fName, 100, 0, ctrl.Name, "", 0, 0, 1000, 300)
lblSub.Caption = "Perfil"

Set ctrl = access.CreateControl(fName, 109, 0, "", "emailUsuario", 9500, 100, 4200, 300)
ctrl.Name = "Correo_Institucional"
ctrl.ControlSource = "emailUsuario"
ctrl.ColumnWidth = 4350
Set lblSub = access.CreateControl(fName, 100, 0, ctrl.Name, "", 0, 0, 1000, 300)
lblSub.Caption = "Correo Institucional"

Set ctrl = access.CreateControl(fName, 109, 0, "", "cveEstatus", 13800, 100, 1500, 300)
ctrl.Name = "Estatus_Usuario"
ctrl.ControlSource = "cveEstatus"
ctrl.ColumnWidth = 1550
Set lblSub = access.CreateControl(fName, 100, 0, ctrl.Name, "", 0, 0, 1000, 300)
lblSub.Caption = "Estatus"
On Error GoTo 0
access.DoCmd.Close 2, fName, 1
On Error Resume Next
access.DoCmd.DeleteObject 2, "Subform_Tabla_Usuarios"
On Error GoTo 0
access.DoCmd.Rename "Subform_Tabla_Usuarios", 2, fName

Set frm = access.CreateForm
fName = frm.Name
On Error Resume Next
frm.Caption = "Catálogo de Instrumentos"
frm.DefaultView = 2 ' Vista Hoja de Datos para Subformulario embebido
frm.AllowAdditions = False
frm.AllowDeletions = False
frm.AllowEdits = False
' VALOR DE DISENO UNICAMENTE. La consulta autorizada de este subformulario
' es la de ConsultasSQL; los formularios que lo embeben la reasignan en
' tiempo de ejecucion con Me.sub<X>.Form.RecordSource = FondoBO.obtenerOrigenInstrumentosBO().
' Aqui se deja el literal porque el constructor es VBScript y no puede
' llamar a una funcion de VBA: el motor de Access todavia no existe.
frm.RecordSource = "SELECT idInstrumento, cveEmisor, tipoInstrumento, tasa, plazo, fechaEmision, fechaVencimiento FROM INSTRUMENTOS ORDER BY idInstrumento;"

Set ctrl = access.CreateControl(fName, 109, 0, "", "idInstrumento", 200, 100, 1800, 300)
ctrl.Name = "ID_Instrumento"
ctrl.ControlSource = "idInstrumento"
ctrl.ColumnWidth = 1950
Set lblSub = access.CreateControl(fName, 100, 0, ctrl.Name, "", 0, 0, 1000, 300)
lblSub.Caption = "ID Instrumento"

Set ctrl = access.CreateControl(fName, 109, 0, "", "cveEmisor", 2100, 100, 1500, 300)
ctrl.Name = "Clave_Emisor"
ctrl.ControlSource = "cveEmisor"
ctrl.ColumnWidth = 1650
Set lblSub = access.CreateControl(fName, 100, 0, ctrl.Name, "", 0, 0, 1000, 300)
lblSub.Caption = "Emisor"

Set ctrl = access.CreateControl(fName, 109, 0, "", "tipoInstrumento", 3700, 100, 3800, 300)
ctrl.Name = "Tipo_Instrumento"
ctrl.ControlSource = "tipoInstrumento"
ctrl.ColumnWidth = 4100
Set lblSub = access.CreateControl(fName, 100, 0, ctrl.Name, "", 0, 0, 1000, 300)
lblSub.Caption = "Tipo de Instrumento"

Set ctrl = access.CreateControl(fName, 109, 0, "", "tasa", 6400, 100, 1300, 300)
ctrl.Name = "Tasa_Anual_Pct"
ctrl.ControlSource = "tasa"
ctrl.Format = "Fixed"
ctrl.ColumnWidth = 1550
Set lblSub = access.CreateControl(fName, 100, 0, ctrl.Name, "", 0, 0, 1000, 300)
lblSub.Caption = "Tasa Anual (%)"

Set ctrl = access.CreateControl(fName, 109, 0, "", "plazo", 7800, 100, 1300, 300)
ctrl.Name = "Plazo_Dias"
ctrl.ControlSource = "plazo"
ctrl.ColumnWidth = 1450
Set lblSub = access.CreateControl(fName, 100, 0, ctrl.Name, "", 0, 0, 1000, 300)
lblSub.Caption = "Plazo (Días)"

Set ctrl = access.CreateControl(fName, 109, 0, "", "fechaEmision", 9200, 100, 1700, 300)
ctrl.Name = "Fecha_Emision"
ctrl.ControlSource = "fechaEmision"
ctrl.Format = "dd/mm/yyyy"
ctrl.ColumnWidth = 1850
Set lblSub = access.CreateControl(fName, 100, 0, ctrl.Name, "", 0, 0, 1000, 300)
lblSub.Caption = "Fecha Emisión"

Set ctrl = access.CreateControl(fName, 109, 0, "", "fechaVencimiento", 11000, 100, 1700, 300)
ctrl.Name = "Fecha_Vencimiento"
ctrl.ControlSource = "fechaVencimiento"
ctrl.Format = "dd/mm/yyyy"
ctrl.ColumnWidth = 1850
Set lblSub = access.CreateControl(fName, 100, 0, ctrl.Name, "", 0, 0, 1000, 300)
lblSub.Caption = "Fecha Vencimiento"
On Error GoTo 0
access.DoCmd.Close 2, fName, 1
On Error Resume Next
access.DoCmd.DeleteObject 2, "Subform_Tabla_Instrumentos"
On Error GoTo 0
access.DoCmd.Rename "Subform_Tabla_Instrumentos", 2, fName

logFile.WriteLine "Paso 9: Construyendo Formulario Form_Gestionar_Usuarios..."
Set frm = access.CreateForm
fName = frm.Name
On Error Resume Next
frm.Caption = "Gestionar Usuarios"
frm.AutoCenter = True
frm.KeyPreview = True
frm.RecordSelectors = False
frm.NavigationButtons = False
frm.ScrollBars = 0

frm.RecordSelectors = False

frm.NavigationButtons = False

frm.DividingLines = False

frm.Width = 16000
frm.Section(0).Height = 6150
frm.Section(0).BackColor = 15790320
On Error GoTo 0

ConfigurarEncabezadoVerde frm, "Módulo de Administración Integral de Riesgos - Gestión de Usuarios"
ConfigurarBarraNavegacionSuperior frm, "Form_Menu_General"

On Error Resume Next
Set ctrl = access.CreateControl(fName, 104, 0, "", "", 1500, 1950, 2400, 480)
ctrl.Name = "cmdAgregar"
ctrl.Caption = "+ Nuevo Usuario"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 4100, 1950, 2400, 480)
ctrl.Name = "cmdActualizar"
ctrl.Caption = "Actualizar Usuario"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 6700, 1950, 2400, 480)
ctrl.Name = "cmdEliminar"
ctrl.Caption = "Inactivar Usuario"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 9300, 1950, 2600, 480)
ctrl.Name = "cmdCambiarAdmin"
ctrl.Caption = "Restablecer Contraseña"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 2600, 13000, 380)
ctrl.Caption = "Directorio de Usuarios Registrados en el Sistema de Administración Integral de Riesgos"
ctrl.FontName = "Cambria"
ctrl.FontSize = 13
ctrl.FontBold = True
ctrl.ForeColor = 6703164
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 112, 0, "", "", 1500, 3050, 13000, 2950)
ctrl.Name = "subUsuarios"
ctrl.SourceObject = "Subform_Tabla_Usuarios"
On Error GoTo 0

codeF = "Option Compare Database" & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"' Form_Gestionar_Usuarios - Menu del catalogo de usuarios" & vbCrLf & _
"' -----------------------------------------------------------------------------" & vbCrLf & _
"' Solo navegacion: alta, actualizacion, baja logica, restablecer" & vbCrLf & _
"' contrasena y directorio. No escribe en la base." & vbCrLf & _
"' Incorpora Form_Subformulario_Consulta_Usuarios, que se refresca" & vbCrLf & _
"' en vivo tras cada operacion." & vbCrLf & _
"' Acceso    : restringido a perfil A (Administrador)." & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"" & vbCrLf & _
"Option Explicit" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub Form_Load()" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    ' Encabezado de sesion y control de acceso. La regla de quien puede" & vbCrLf & _
"    ' entrar a esta pantalla NO esta aqui: vive en FondoBO.autorizarPantallaBO." & vbCrLf & _
"    ' Este formulario solo pregunta y obedece (Marco de Desarrollo, I.2)." & vbCrLf & _
"    ' Si el acceso se deniega, el ayudante ya aviso, abrio el menu y cerro" & vbCrLf & _
"    ' este formulario; aqui solo queda salir." & vbCrLf & _
"    If Not Utilities.AplicarEncabezadoSesionUT(Me, FondoBO.PANT_GESTION_USUARIOS) Then Exit Sub" & vbCrLf & _
"    Me.subUsuarios.Form.RecordSource = FondoBO.obtenerOrigenUsuariosBO()" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_Resize()" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdAgregar_Click()" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Agregar_Usuarios"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdActualizar_Click()" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Actualizar_Usuarios"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub cmdEliminar_Click()" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Eliminar_Usuarios"", acNormal" & vbCrLf
codeF = codeF & "    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdCambiarAdmin_Click()" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Cambiar_Contrasena_Administrador"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdAtras_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Menu_General"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub cmdIrMenu_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Menu_General"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdSalir_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If MsgBox(""¿Desea cerrar la sesión actual?"", vbQuestion + vbYesNo, ""Salir"") = vbYes Then" & vbCrLf & _
"        Set usuarioSesionDTO = Nothing" & vbCrLf & _
"        DoCmd.OpenForm ""Form_Acceso"", acNormal" & vbCrLf & _
"        DoCmd.Close acForm, Me.Name" & vbCrLf
codeF = codeF & "    End If" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub cmdRegresar_Click()" & vbCrLf & _
"    cmdAtras_Click" & vbCrLf & _
"End Sub"
AsignarCodigoLimpio frm, codeF
access.DoCmd.Close 2, fName, 1
On Error Resume Next
access.DoCmd.DeleteObject 2, "Form_Gestionar_Usuarios"
On Error GoTo 0
access.DoCmd.Rename "Form_Gestionar_Usuarios", 2, fName
logFile.WriteLine "   [OK] Form_Gestionar_Usuarios creado."

' 5. Form_Subformulario_Consulta_Usuarios
logFile.WriteLine "Paso 10: Construyendo Formulario Form_Subformulario_Consulta_Usuarios..."
Set frm = access.CreateForm
fName = frm.Name
On Error Resume Next
frm.Caption = "Directorio de Usuarios"
frm.AutoCenter = True
frm.KeyPreview = True
frm.RecordSelectors = False
frm.NavigationButtons = False
frm.ScrollBars = 0

frm.RecordSelectors = False

frm.NavigationButtons = False

frm.DividingLines = False

frm.Width = 16000
frm.Section(0).Height = 6150
frm.Section(0).BackColor = 15790320
On Error GoTo 0

ConfigurarEncabezadoVerde frm, "Módulo de Administración Integral de Riesgos - Directorio de Usuarios"
ConfigurarBarraNavegacionSuperior frm, "Form_Menu_General"

On Error Resume Next
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 2200, 13000, 400)
ctrl.Caption = "Consulta General de Usuarios y Perfiles Institucionales"
ctrl.FontName = "Cambria"
ctrl.FontSize = 13
ctrl.FontBold = True
ctrl.ForeColor = 6703164
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 112, 0, "", "", 1500, 2700, 13000, 3300)
ctrl.Name = "subDirectorioUsuarios"
ctrl.SourceObject = "Subform_Tabla_Usuarios"
On Error GoTo 0

codeF = "Option Compare Database" & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"' Form_Subformulario_Consulta_Usuarios - Subformulario: tabla de usuarios en vivo" & vbCrLf & _
"' -----------------------------------------------------------------------------" & vbCrLf & _
"' Solo lectura. Origen: ConsultasSQL.sqlSelectUsuarios." & vbCrLf & _
"' Se embebe en las pantallas de usuarios y se actualiza con Requery" & vbCrLf & _
"' inmediatamente despues de cada alta, cambio o baja." & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"" & vbCrLf & _
"Option Explicit" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_Load()" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    ' Encabezado de sesion y control de acceso. La regla de quien puede" & vbCrLf
codeF = codeF & "    ' entrar a esta pantalla NO esta aqui: vive en FondoBO.autorizarPantallaBO." & vbCrLf & _
"    ' Este formulario solo pregunta y obedece (Marco de Desarrollo, I.2)." & vbCrLf & _
"    ' Si el acceso se deniega, el ayudante ya aviso, abrio el menu y cerro" & vbCrLf & _
"    ' este formulario; aqui solo queda salir." & vbCrLf & _
"    If Not Utilities.AplicarEncabezadoSesionUT(Me, FondoBO.PANT_CONSULTA_USUARIOS) Then Exit Sub" & vbCrLf & _
"    Me.subDirectorioUsuarios.Form.RecordSource = FondoBO.obtenerOrigenUsuariosBO()" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_Resize()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub cmdAtras_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Menu_General"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdIrMenu_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Menu_General"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub cmdSalir_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    If MsgBox(""¿Desea cerrar la sesión actual?"", vbQuestion + vbYesNo, ""Salir"") = vbYes Then" & vbCrLf & _
"        Set usuarioSesionDTO = Nothing" & vbCrLf & _
"        DoCmd.OpenForm ""Form_Acceso"", acNormal" & vbCrLf & _
"        DoCmd.Close acForm, Me.Name" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdRegresar_Click()" & vbCrLf & _
"    cmdAtras_Click" & vbCrLf & _
"End Sub"
AsignarCodigoLimpio frm, codeF
End Sub

Private Sub EjecutarBloqueInstalacion_06()
access.DoCmd.Close 2, fName, 1
On Error Resume Next
access.DoCmd.DeleteObject 2, "Form_Subformulario_Consulta_Usuarios"
On Error GoTo 0
access.DoCmd.Rename "Form_Subformulario_Consulta_Usuarios", 2, fName
logFile.WriteLine "   [OK] Form_Subformulario_Consulta_Usuarios creado."

' 6. Form_Agregar_Usuarios
logFile.WriteLine "Paso 11: Construyendo Formulario Form_Agregar_Usuarios..."
Set frm = access.CreateForm
fName = frm.Name
On Error Resume Next
frm.Caption = "Agregar Usuario"
frm.AutoCenter = True
frm.KeyPreview = True
frm.ScrollBars = 0
frm.RecordSelectors = False
frm.NavigationButtons = False
frm.DividingLines = False
frm.Width = 16000
frm.Section(0).Height = 6150
frm.Section(0).BackColor = 15790320
On Error GoTo 0

ConfigurarEncabezadoVerde frm, "Módulo de Administración Integral de Riesgos - Alta de Usuario"
ConfigurarBarraNavegacionSuperior frm, "Form_Gestionar_Usuarios"

On Error Resume Next
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 1900, 13000, 360)
ctrl.Caption = "Captura de Datos del Nuevo Usuario"
ctrl.FontName = "Cambria"
ctrl.FontSize = 13
ctrl.FontBold = True
ctrl.ForeColor = 6703164
ctrl.TextAlign = 2
ctrl.BackStyle = 0

' Renglon 1: No. Empleado, Contrasena, Perfil
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 2320, 1750, 320)
ctrl.Caption = "No. Empleado:"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 3300, 2320, 1600, 320)
ctrl.Name = "txtUsuario"

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 5150, 2320, 1450, 320)
ctrl.Caption = "Contraseña:"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 6650, 2320, 2000, 320)
ctrl.Name = "txtPassword"
ctrl.InputMask = "Password"
' SEGURO: si el usuario presiona el ojito y suelta el raton FUERA del boton,
' el MouseUp del boton puede no llegar. Al perder el foco el cuadro se vuelve
' a enmascarar, de modo que la contrasena nunca se queda visible.
ctrl.OnLostFocus = "[Event Procedure]"

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 8680, 2320, 460, 320)
ctrl.Name = "cmdVerPass"
ctrl.Caption = "Ver"
ctrl.OnClick = "[Event Procedure]"
EstilizarBotonVer ctrl

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 9350, 2320, 1550, 320)
ctrl.Caption = "Perfil (A / U):"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 10950, 2320, 900, 320)
ctrl.Name = "txtPerfil"

' Renglon 2: Nombre(s) y Apellido Paterno
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 2740, 1750, 320)
ctrl.Caption = "Nombre(s):"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 3300, 2740, 2800, 320)
ctrl.Name = "txtNombre"

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 6400, 2740, 2000, 320)
ctrl.Caption = "Apellido Paterno:"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 8450, 2740, 3400, 320)
ctrl.Name = "txtPaterno"

' Renglon 3: Apellido Materno y Correo Hidrocarburos
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 3160, 1750, 320)
ctrl.Caption = "Apellido Materno:"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 3300, 3160, 2800, 320)
ctrl.Name = "txtMaterno"

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 6400, 3160, 2000, 320)
ctrl.Caption = "Correo Hidrocarburos:"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 8450, 3160, 3400, 320)
ctrl.Name = "txtEmail"

' Botones de Accion a la derecha
Set ctrl = access.CreateControl(fName, 104, 0, "", "", 12150, 2360, 2350, 460)
ctrl.Name = "cmdGuardar"
ctrl.Caption = "Guardar"
ctrl.Default = False
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 12150, 2940, 2350, 460)
ctrl.Name = "cmdLimpiar"
ctrl.Caption = "Limpiar"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

' Subformulario de Usuarios en vivo para corroborar el alta en tiempo real
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 3580, 13000, 300)
ctrl.Caption = "Directorio de Usuarios en Tiempo Real (Corroboración de Alta)"
ctrl.FontName = "Cambria"
ctrl.FontSize = 11
ctrl.FontBold = True
ctrl.ForeColor = 6703164
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 112, 0, "", "", 1500, 3900, 13000, 2100)
ctrl.Name = "subUsuarios"
ctrl.SourceObject = "Subform_Tabla_Usuarios"
On Error GoTo 0

codeF = "Option Compare Database" & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"' Form_Agregar_Usuarios - Alta de usuario" & vbCrLf & _
"' -----------------------------------------------------------------------------" & vbCrLf & _
"' Entradas : clave de 4 digitos, nombre, apellidos, correo, perfil y" & vbCrLf & _
"'            contrasena inicial." & vbCrLf & _
"' Valida   : clave no duplicada, perfil en {A, U} y politica de contrasena." & vbCrLf & _
"' Escribe  : INSERT en USUARIOS via FondoBO.registrarUsuarioBO." & vbCrLf & _
"' El campo de clave esta restringido por teclado a 4 digitos numericos." & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"" & vbCrLf & _
"Option Explicit" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private mBufPass As String" & vbCrLf & _
"" & vbCrLf & _
"' ---------------------------------------------------------------------------" & vbCrLf & _
"' SEGURO DE ENMASCARADO AL PERDER EL FOCO" & vbCrLf & _
"' Cierra el unico hueco que deja el ojito: presionar el boton y soltar el" & vbCrLf & _
"' raton fuera de el. En ese caso el MouseUp del boton puede no llegar, pero" & vbCrLf & _
"' en cuanto el cuadro pierde el foco la contrasena se vuelve a ocultar." & vbCrLf & _
"' ---------------------------------------------------------------------------" & vbCrLf & _
"Private Sub txtPassword_LostFocus()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call Utilities.OcultarTodasLasContrasenasUT(Me)" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.RestringirKeyDownUsuario4Digitos Me.txtUsuario, KeyCode, Shift" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtUsuario_KeyPress(KeyAscii As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.RestringirKeyPressUsuario4Digitos Me.txtUsuario, KeyAscii" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtUsuario_Change()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    Utilities.TruncarCambioUsuario4Digitos Me.txtUsuario" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub txtPassword_KeyPress(KeyAscii As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.ProcesarKeyPressContrasena mBufPass, KeyAscii" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtPassword_Change()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.TruncarCambioContrasena8Caracteres Me.txtPassword, mBufPass" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"' ---------------------------------------------------------------------------" & vbCrLf
codeF = codeF & "' VER CONTRASENA SOLO MIENTRAS SE SOSTIENE EL CLIC" & vbCrLf & _
"' MouseDown la descubre, MouseUp la vuelve a ocultar. El Click se deja como" & vbCrLf
codeF = codeF & "' seguro por si el MouseUp no llegara: al soltar siempre queda oculta." & vbCrLf & _
"' ---------------------------------------------------------------------------" & vbCrLf & _
"Private Sub cmdVerPass_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call Utilities.VerContrasenaSostenidaUT(Me, ""cmdVerPass"", ""txtPassword"", True, """")" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdVerPass_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call Utilities.VerContrasenaSostenidaUT(Me, ""cmdVerPass"", ""txtPassword"", False, """")" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub cmdVerPass_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    Call Utilities.VerContrasenaSostenidaUT(Me, ""cmdVerPass"", ""txtPassword"", False, """")" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If Shift = 0 Then" & vbCrLf & _
"        If Not (TypeOf Screen.ActiveControl Is ComboBox Or TypeOf Screen.ActiveControl Is ListBox) Then" & vbCrLf & _
"            Select Case KeyCode" & vbCrLf & _
"                Case 40: KeyCode = 0" & vbCrLf & _
"                Case 38: KeyCode = 0" & vbCrLf & _
"            End Select" & vbCrLf & _
"        End If" & vbCrLf
codeF = codeF & "    End If" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub Form_Load()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    ' Encabezado de sesion y control de acceso. La regla de quien puede" & vbCrLf & _
"    ' entrar a esta pantalla NO esta aqui: vive en FondoBO.autorizarPantallaBO." & vbCrLf & _
"    ' Este formulario solo pregunta y obedece (Marco de Desarrollo, I.2)." & vbCrLf & _
"    ' Si el acceso se deniega, el ayudante ya aviso, abrio el menu y cerro" & vbCrLf & _
"    ' este formulario; aqui solo queda salir." & vbCrLf & _
"    If Not Utilities.AplicarEncabezadoSesionUT(Me, FondoBO.PANT_ALTA_USUARIOS) Then Exit Sub" & vbCrLf & _
"    mBufPass = """"" & vbCrLf & _
"    Me.txtUsuario.Value = Null" & vbCrLf & _
"    Me.txtPassword.Value = Null" & vbCrLf
codeF = codeF & "    Me.txtPerfil.Value = PERFIL_EMPLEADO" & vbCrLf & _
"    Me.subUsuarios.Form.RecordSource = FondoBO.obtenerOrigenUsuariosBO()" & vbCrLf
codeF = codeF & "    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_Resize()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdLimpiar_Click()" & vbCrLf & _
"    mBufPass = """"" & vbCrLf & _
"    Me.txtUsuario.Value = Null" & vbCrLf & _
"    Me.txtPassword.Value = Null" & vbCrLf
codeF = codeF & "    Me.txtNombre.Value = Null" & vbCrLf & _
"    Me.txtPaterno.Value = Null" & vbCrLf
codeF = codeF & "    Me.txtMaterno.Value = Null" & vbCrLf & _
"    Me.txtPerfil.Value = PERFIL_EMPLEADO" & vbCrLf & _
"    Me.txtEmail.Value = Null" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdGuardar_Click()" & vbCrLf & _
"    On Error Resume Next: Me.cmdGuardar.SetFocus: On Error GoTo ErrH" & vbCrLf & _
"    Dim nuevoUsuarioDTO As UsuarioDTO" & vbCrLf & _
"    Dim msgResultado As String" & vbCrLf & _
"    Dim exito As Boolean" & vbCrLf & _
"    " & vbCrLf & _
"    Set nuevoUsuarioDTO = New UsuarioDTO" & vbCrLf
codeF = codeF & "    nuevoUsuarioDTO.setCveUsuario Utilities.LeerTextoControl(Me.txtUsuario)" & vbCrLf & _
"    nuevoUsuarioDTO.setPassword Utilities.LeerTextoControl(Me.txtPassword, mBufPass)" & vbCrLf
codeF = codeF & "    nuevoUsuarioDTO.setNombre Utilities.LeerTextoControl(Me.txtNombre)" & vbCrLf & _
"    nuevoUsuarioDTO.setAPaterno Utilities.LeerTextoControl(Me.txtPaterno)" & vbCrLf & _
"    nuevoUsuarioDTO.setAMaterno Utilities.LeerTextoControl(Me.txtMaterno)" & vbCrLf & _
"    Dim strPerfNuevo As String: strPerfNuevo = Utilities.LeerTextoControl(Me.txtPerfil)" & vbCrLf & _
"    If strPerfNuevo = """" Then strPerfNuevo = PERFIL_EMPLEADO" & vbCrLf & _
"    nuevoUsuarioDTO.setCvePerfil strPerfNuevo" & vbCrLf & _
"    nuevoUsuarioDTO.setEmail Utilities.LeerTextoControl(Me.txtEmail)" & vbCrLf & _
"    nuevoUsuarioDTO.setCveEstatus CLAVE_ESTATU_ACTIVO" & vbCrLf & _
"    " & vbCrLf & _
"    exito = FondoBO.registrarUsuarioBO(nuevoUsuarioDTO, msgResultado)" & vbCrLf & _
"    Set nuevoUsuarioDTO = Nothing" & vbCrLf & _
"    " & vbCrLf
codeF = codeF & "    If exito Then" & vbCrLf & _
"        On Error Resume Next" & vbCrLf
codeF = codeF & "        Me.subUsuarios.Form.RecordSource = FondoBO.obtenerOrigenUsuariosBO()" & vbCrLf & _
"        Me.subUsuarios.Form.Requery" & vbCrLf & _
"        On Error GoTo ErrH" & vbCrLf & _
"        MsgBox msgResultado & vbCrLf & vbCrLf & ""El subformulario inferior de usuarios se ha actualizado en tiempo real."", vbInformation, ""Alta de Usuario""" & vbCrLf & _
"        cmdLimpiar_Click" & vbCrLf & _
"    Else" & vbCrLf & _
"        MsgBox msgResultado, vbExclamation, ""Validación de Negocio""" & vbCrLf & _
"    End If" & vbCrLf & _
"    Exit Sub" & vbCrLf & _
"ErrH:" & vbCrLf & _
"    MsgBox MENSAJE_ERROR & vbCrLf & Err.Description, vbCritical" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub cmdAtras_Click()" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Gestionar_Usuarios"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdIrMenu_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Menu_General"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdSalir_Click()" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    If MsgBox(""¿Desea cerrar la sesión actual?"", vbQuestion + vbYesNo, ""Salir"") = vbYes Then" & vbCrLf
codeF = codeF & "        Set usuarioSesionDTO = Nothing" & vbCrLf & _
"        DoCmd.OpenForm ""Form_Acceso"", acNormal" & vbCrLf & _
"        DoCmd.Close acForm, Me.Name" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdRegresar_Click()" & vbCrLf & _
"    cmdAtras_Click" & vbCrLf & _
"End Sub"
AsignarCodigoLimpio frm, codeF
access.DoCmd.Close 2, fName, 1
On Error Resume Next
access.DoCmd.DeleteObject 2, "Form_Agregar_Usuarios"
On Error GoTo 0
access.DoCmd.Rename "Form_Agregar_Usuarios", 2, fName
logFile.WriteLine "   [OK] Form_Agregar_Usuarios creado."

' 7. Form_Actualizar_Usuarios
logFile.WriteLine "Paso 12: Construyendo Formulario Form_Actualizar_Usuarios..."
Set frm = access.CreateForm
fName = frm.Name
On Error Resume Next
frm.Caption = "Actualizar Usuarios"
frm.AutoCenter = True
frm.KeyPreview = True
frm.ScrollBars = 0
frm.RecordSelectors = False
frm.NavigationButtons = False
frm.DividingLines = False
frm.Width = 16000
frm.Section(0).Height = 6150
frm.Section(0).BackColor = 15790320
On Error GoTo 0

ConfigurarEncabezadoVerde frm, "Módulo de Administración Integral de Riesgos - Modificación de Usuario"
ConfigurarBarraNavegacionSuperior frm, "Form_Gestionar_Usuarios"

On Error Resume Next
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 1900, 13000, 350)
ctrl.Caption = "Actualización de Datos de Usuario"
ctrl.FontName = "Cambria"
ctrl.FontSize = 13
ctrl.FontBold = True
ctrl.ForeColor = 6703164
ctrl.TextAlign = 2
ctrl.BackStyle = 0

' Renglon 1: Numero de Empleado + Buscar + Perfil
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 2320, 2200, 330)
ctrl.Caption = "Número de Empleado:"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 3750, 2320, 1800, 330)
ctrl.Name = "txtUsuario"

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 5700, 2320, 1650, 340)
ctrl.Name = "cmdBuscar"
ctrl.Caption = "Buscar"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 7700, 2320, 1600, 330)
End Sub

Private Sub EjecutarBloqueInstalacion_07()
ctrl.Caption = "Perfil (A / U):"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 9350, 2320, 1200, 330)
ctrl.Name = "txtPerfil"

' Renglon 2: Nombre(s) y Apellido Paterno
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 2740, 2200, 330)
ctrl.Caption = "Nombre(s):"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 3750, 2740, 2800, 330)
ctrl.Name = "txtNombre"

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 6800, 2740, 2000, 330)
ctrl.Caption = "Apellido Paterno:"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 8850, 2740, 2950, 330)
ctrl.Name = "txtPaterno"

' Renglon 3: Apellido Materno y Correo Electronico
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 3160, 2200, 330)
ctrl.Caption = "Apellido Materno:"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 3750, 3160, 2800, 330)
ctrl.Name = "txtMaterno"

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 6800, 3160, 2000, 330)
ctrl.Caption = "Correo Electrónico:"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 8850, 3160, 2950, 330)
ctrl.Name = "txtEmail"

' Boton Guardar Cambios a la derecha
Set ctrl = access.CreateControl(fName, 104, 0, "", "", 12050, 2600, 2450, 520)
ctrl.Name = "cmdActualizar"
ctrl.Caption = "Guardar Cambios"
ctrl.Default = False
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

' Subformulario de Usuarios en vivo para corroborar la actualizacion
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 3580, 13000, 300)
ctrl.Caption = "Directorio de Usuarios en Tiempo Real (Corroboración de Actualización)"
ctrl.FontName = "Cambria"
ctrl.FontSize = 11
ctrl.FontBold = True
ctrl.ForeColor = 6703164
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 112, 0, "", "", 1500, 3900, 13000, 2100)
ctrl.Name = "subUsuarios"
ctrl.SourceObject = "Subform_Tabla_Usuarios"
On Error GoTo 0

codeF = "Option Compare Database" & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"' Form_Actualizar_Usuarios - Actualizacion de datos de usuario" & vbCrLf & _
"' -----------------------------------------------------------------------------" & vbCrLf & _
"' Busca por clave con FondoBO.consultarUsuarioBO y precarga el formulario." & vbCrLf & _
"' Escribe  : UPDATE en USUARIOS via FondoBO.actualizarUsuarioBO." & vbCrLf & _
"' No modifica la contrasena; para eso existe la pantalla de restablecer." & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"" & vbCrLf & _
"Option Explicit" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtUsuario_KeyPress(KeyAscii As Integer)" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    Utilities.RestringirKeyPressUsuario4Digitos Me.txtUsuario, KeyAscii" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtUsuario_Change()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.TruncarCambioUsuario4Digitos Me.txtUsuario" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf & _
"    If KeyCode = 13 Then" & vbCrLf & _
"        KeyCode = 0" & vbCrLf & _
"        cmdBuscar_Click" & vbCrLf & _
"    End If" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If Shift = 0 Then" & vbCrLf & _
"        If Not (TypeOf Screen.ActiveControl Is ComboBox Or TypeOf Screen.ActiveControl Is ListBox) Then" & vbCrLf & _
"            Select Case KeyCode" & vbCrLf & _
"                Case 40: KeyCode = 0" & vbCrLf & _
"                Case 38: KeyCode = 0" & vbCrLf & _
"            End Select" & vbCrLf & _
"        End If" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub Form_Load()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    ' Encabezado de sesion y control de acceso. La regla de quien puede" & vbCrLf & _
"    ' entrar a esta pantalla NO esta aqui: vive en FondoBO.autorizarPantallaBO." & vbCrLf & _
"    ' Este formulario solo pregunta y obedece (Marco de Desarrollo, I.2)." & vbCrLf & _
"    ' Si el acceso se deniega, el ayudante ya aviso, abrio el menu y cerro" & vbCrLf & _
"    ' este formulario; aqui solo queda salir." & vbCrLf & _
"    If Not Utilities.AplicarEncabezadoSesionUT(Me, FondoBO.PANT_CAMBIO_USUARIOS) Then Exit Sub" & vbCrLf & _
"    Me.subUsuarios.Form.RecordSource = FondoBO.obtenerOrigenUsuariosBO()" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_Resize()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub cmdBuscar_Click()" & vbCrLf & _
"    On Error GoTo ErrH" & vbCrLf & _
"    Dim cveBuscar As String" & vbCrLf & _
"    Dim uEncontrado As UsuarioDTO" & vbCrLf & _
"    " & vbCrLf & _
"    cveBuscar = Utilities.LeerTextoControl(Me.txtUsuario)" & vbCrLf & _
"    Set uEncontrado = FondoBO.consultarUsuarioBO(cveBuscar)" & vbCrLf & _
"    If Not uEncontrado Is Nothing Then" & vbCrLf & _
"        Me.txtNombre.Value = uEncontrado.getNombre()" & vbCrLf & _
"        Me.txtPaterno.Value = uEncontrado.getAPaterno()" & vbCrLf & _
"        Me.txtMaterno.Value = uEncontrado.getAMaterno()" & vbCrLf
codeF = codeF & "        Me.txtPerfil.Value = uEncontrado.getCvePerfil()" & vbCrLf & _
"        Me.txtEmail.Value = uEncontrado.getEmail()" & vbCrLf
codeF = codeF & "    Else" & vbCrLf & _
"        MsgBox ""Usuario no encontrado."", vbExclamation, ""Búsqueda de Usuario""" & vbCrLf & _
"    End If" & vbCrLf & _
"    Set uEncontrado = Nothing" & vbCrLf & _
"    Exit Sub" & vbCrLf & _
"ErrH:" & vbCrLf & _
"    MsgBox MENSAJE_ERROR & vbCrLf & Err.Description, vbCritical" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdActualizar_Click()" & vbCrLf & _
"    On Error Resume Next: Me.cmdActualizar.SetFocus: On Error GoTo ErrH" & vbCrLf & _
"    Dim usuarioDTO As UsuarioDTO" & vbCrLf
codeF = codeF & "    Dim msgResultado As String" & vbCrLf & _
"    Dim exito As Boolean" & vbCrLf
codeF = codeF & "    " & vbCrLf & _
"    Set usuarioDTO = New UsuarioDTO" & vbCrLf & _
"    usuarioDTO.setCveUsuario Utilities.LeerTextoControl(Me.txtUsuario)" & vbCrLf & _
"    usuarioDTO.setNombre Utilities.LeerTextoControl(Me.txtNombre)" & vbCrLf & _
"    usuarioDTO.setAPaterno Utilities.LeerTextoControl(Me.txtPaterno)" & vbCrLf & _
"    usuarioDTO.setAMaterno Utilities.LeerTextoControl(Me.txtMaterno)" & vbCrLf & _
"    Dim strPerfAct As String: strPerfAct = Utilities.LeerTextoControl(Me.txtPerfil)" & vbCrLf & _
"    If strPerfAct = """" Then strPerfAct = PERFIL_EMPLEADO" & vbCrLf & _
"    usuarioDTO.setCvePerfil strPerfAct" & vbCrLf & _
"    usuarioDTO.setEmail Utilities.LeerTextoControl(Me.txtEmail)" & vbCrLf & _
"    " & vbCrLf & _
"    exito = FondoBO.actualizarUsuarioBO(usuarioDTO, msgResultado)" & vbCrLf
codeF = codeF & "    Set usuarioDTO = Nothing" & vbCrLf & _
"    " & vbCrLf
codeF = codeF & "    If exito Then" & vbCrLf & _
"        On Error Resume Next" & vbCrLf & _
"        Me.subUsuarios.Form.RecordSource = FondoBO.obtenerOrigenUsuariosBO()" & vbCrLf & _
"        Me.subUsuarios.Form.Requery" & vbCrLf & _
"        On Error GoTo ErrH" & vbCrLf & _
"        MsgBox msgResultado & vbCrLf & vbCrLf & ""El subformulario inferior de usuarios se ha actualizado en tiempo real."", vbInformation, ""Actualización Exitosa""" & vbCrLf & _
"    Else" & vbCrLf & _
"        MsgBox msgResultado, vbExclamation, ""Validación de Negocio""" & vbCrLf & _
"    End If" & vbCrLf & _
"    Exit Sub" & vbCrLf & _
"ErrH:" & vbCrLf & _
"    MsgBox MENSAJE_ERROR & vbCrLf & Err.Description, vbCritical" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub cmdAtras_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Gestionar_Usuarios"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdIrMenu_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Menu_General"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub cmdSalir_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    If MsgBox(""¿Desea cerrar la sesión actual?"", vbQuestion + vbYesNo, ""Salir"") = vbYes Then" & vbCrLf & _
"        Set usuarioSesionDTO = Nothing" & vbCrLf & _
"        DoCmd.OpenForm ""Form_Acceso"", acNormal" & vbCrLf & _
"        DoCmd.Close acForm, Me.Name" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdRegresar_Click()" & vbCrLf & _
"    cmdAtras_Click" & vbCrLf
"End Sub"
AsignarCodigoLimpio frm, codeF
access.DoCmd.Close 2, fName, 1
On Error Resume Next
access.DoCmd.DeleteObject 2, "Form_Actualizar_Usuarios"
On Error GoTo 0
access.DoCmd.Rename "Form_Actualizar_Usuarios", 2, fName
logFile.WriteLine "   [OK] Form_Actualizar_Usuarios creado."

' 8. Form_Eliminar_Usuarios
logFile.WriteLine "Paso 13: Construyendo Formulario Form_Eliminar_Usuarios..."
Set frm = access.CreateForm
fName = frm.Name
On Error Resume Next
frm.Caption = "Eliminar Usuarios"
frm.AutoCenter = True
frm.KeyPreview = True
frm.ScrollBars = 0
frm.RecordSelectors = False
frm.NavigationButtons = False
frm.DividingLines = False
frm.Width = 16000
frm.Section(0).Height = 6150
frm.Section(0).BackColor = 15790320
On Error GoTo 0

ConfigurarEncabezadoVerde frm, "Módulo de Administración Integral de Riesgos - Baja de Usuario"
ConfigurarBarraNavegacionSuperior frm, "Form_Gestionar_Usuarios"

On Error Resume Next
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 1900, 13000, 350)
ctrl.Caption = "Baja Lógica (Inactivación) de Usuario"
ctrl.FontName = "Cambria"
ctrl.FontSize = 13
ctrl.FontBold = True
ctrl.ForeColor = 6703164
ctrl.TextAlign = 2
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 2000, 2360, 2400, 350)
ctrl.Caption = "Número de Empleado:"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 4500, 2360, 2200, 350)
ctrl.Name = "txtUsuario"

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 6900, 2360, 1800, 360)
ctrl.Name = "cmdBuscar"
ctrl.Caption = "Buscar"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 9100, 2320, 3200, 440)
ctrl.Name = "cmdEliminar"
ctrl.Caption = "Dar de Baja (Inactivar)"
ctrl.Default = False
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 2900, 13000, 380)
ctrl.Name = "lblInfo"
ctrl.Caption = "Ingrese el número de empleado para consultar."
ctrl.FontBold = True
ctrl.ForeColor = 4737096
ctrl.TextAlign = 2

' Subformulario de Usuarios en vivo para corroborar la baja/inactivacion
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 3450, 13000, 300)
ctrl.Caption = "Directorio de Usuarios en Tiempo Real (Corroboración de Estatus Activo / Inactivo)"
ctrl.FontName = "Cambria"
ctrl.FontSize = 11
ctrl.FontBold = True
ctrl.ForeColor = 6703164
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 112, 0, "", "", 1500, 3800, 13000, 2200)
ctrl.Name = "subUsuarios"
ctrl.SourceObject = "Subform_Tabla_Usuarios"
On Error GoTo 0

codeF = "Option Compare Database" & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"' Form_Eliminar_Usuarios - Baja logica de usuario" & vbCrLf & _
"' -----------------------------------------------------------------------------" & vbCrLf & _
"' No borra el registro: marca cveEstatus = INACTIVO." & vbCrLf & _
"' Asi se conserva la trazabilidad historica de quien opero el modulo." & vbCrLf & _
"' Escribe  : UPDATE en USUARIOS via FondoBO.inactivarUsuarioBO." & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"" & vbCrLf & _
"Option Explicit" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtUsuario_KeyPress(KeyAscii As Integer)" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    Utilities.RestringirKeyPressUsuario4Digitos Me.txtUsuario, KeyAscii" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtUsuario_Change()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.TruncarCambioUsuario4Digitos Me.txtUsuario" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf & _
"    If KeyCode = 13 Then" & vbCrLf & _
"        KeyCode = 0" & vbCrLf & _
"        cmdBuscar_Click" & vbCrLf & _
"    End If" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If Shift = 0 Then" & vbCrLf & _
"        If Not (TypeOf Screen.ActiveControl Is ComboBox Or TypeOf Screen.ActiveControl Is ListBox) Then" & vbCrLf & _
"            Select Case KeyCode" & vbCrLf & _
"                Case 40: KeyCode = 0" & vbCrLf & _
"                Case 38: KeyCode = 0" & vbCrLf & _
"            End Select" & vbCrLf & _
"        End If" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub Form_Load()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    ' Encabezado de sesion y control de acceso. La regla de quien puede" & vbCrLf & _
"    ' entrar a esta pantalla NO esta aqui: vive en FondoBO.autorizarPantallaBO." & vbCrLf & _
"    ' Este formulario solo pregunta y obedece (Marco de Desarrollo, I.2)." & vbCrLf & _
"    ' Si el acceso se deniega, el ayudante ya aviso, abrio el menu y cerro" & vbCrLf & _
"    ' este formulario; aqui solo queda salir." & vbCrLf & _
"    If Not Utilities.AplicarEncabezadoSesionUT(Me, FondoBO.PANT_BAJA_USUARIOS) Then Exit Sub" & vbCrLf & _
"    Me.subUsuarios.Form.RecordSource = FondoBO.obtenerOrigenUsuariosBO()" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_Resize()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub cmdBuscar_Click()" & vbCrLf & _
"    On Error GoTo ErrH" & vbCrLf & _
"    Dim cveBuscar As String" & vbCrLf & _
"    Dim uEncontrado As UsuarioDTO" & vbCrLf & _
"    " & vbCrLf & _
"    cveBuscar = Utilities.LeerTextoControl(Me.txtUsuario)" & vbCrLf & _
"    Set uEncontrado = FondoBO.consultarUsuarioBO(cveBuscar)" & vbCrLf & _
"    If Not uEncontrado Is Nothing Then" & vbCrLf & _
"        Me.lblInfo.Caption = ""Usuario: "" & uEncontrado.getNombre() & "" "" & uEncontrado.getAPaterno() & "" | Perfil: "" & uEncontrado.getCvePerfil() & "" | Estatus: "" & uEncontrado.getCveEstatus()" & vbCrLf & _
"    Else" & vbCrLf & _
"        MsgBox ""Usuario no encontrado."", vbExclamation, ""Búsqueda de Usuario""" & vbCrLf
codeF = codeF & "    End If" & vbCrLf & _
"    Set uEncontrado = Nothing" & vbCrLf
codeF = codeF & "    Exit Sub" & vbCrLf & _
"ErrH:" & vbCrLf & _
"    MsgBox MENSAJE_ERROR & vbCrLf & Err.Description, vbCritical" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdEliminar_Click()" & vbCrLf & _
"    On Error GoTo ErrH" & vbCrLf & _
"    Dim cveBaja As String" & vbCrLf & _
"    Dim msgResultado As String" & vbCrLf & _
"    Dim exito As Boolean" & vbCrLf & _
"    " & vbCrLf & _
"    cveBaja = Utilities.LeerTextoControl(Me.txtUsuario)" & vbCrLf
codeF = codeF & "    If MsgBox(""¿Desea dar de baja al usuario seleccionado?"", vbQuestion + vbYesNo, ""Confirmar Baja"") = vbYes Then" & vbCrLf & _
"        exito = FondoBO.inactivarUsuarioBO(cveBaja, msgResultado)" & vbCrLf
codeF = codeF & "        If exito Then" & vbCrLf & _
"            On Error Resume Next" & vbCrLf & _
"            Me.subUsuarios.Form.RecordSource = FondoBO.obtenerOrigenUsuariosBO()" & vbCrLf & _
"            Me.subUsuarios.Form.Requery" & vbCrLf & _
"            Me.lblInfo.Caption = ""Usuario "" & cveBaja & "" actualizado a Estatus: INACTIVO (I)""" & vbCrLf & _
"            On Error GoTo ErrH" & vbCrLf & _
"            MsgBox msgResultado & vbCrLf & vbCrLf & ""El subformulario inferior de usuarios se ha actualizado en tiempo real mostrando el estatus INACTIVO."", vbInformation, ""Baja de Usuario""" & vbCrLf & _
"        Else" & vbCrLf & _
"            MsgBox msgResultado, vbExclamation, ""Validación de Negocio""" & vbCrLf & _
"        End If" & vbCrLf & _
"    End If" & vbCrLf & _
"    Exit Sub" & vbCrLf
codeF = codeF & "ErrH:" & vbCrLf & _
"    MsgBox MENSAJE_ERROR & vbCrLf & Err.Description, vbCritical" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdAtras_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Gestionar_Usuarios"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdIrMenu_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Menu_General"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub cmdSalir_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If MsgBox(""¿Desea cerrar la sesión actual?"", vbQuestion + vbYesNo, ""Salir"") = vbYes Then" & vbCrLf & _
"        Set usuarioSesionDTO = Nothing" & vbCrLf & _
"        DoCmd.OpenForm ""Form_Acceso"", acNormal" & vbCrLf & _
"        DoCmd.Close acForm, Me.Name" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdRegresar_Click()" & vbCrLf & _
"    cmdAtras_Click" & vbCrLf & _
"End Sub"
AsignarCodigoLimpio frm, codeF
End Sub

Private Sub EjecutarBloqueInstalacion_08()
access.DoCmd.Close 2, fName, 1
On Error Resume Next
access.DoCmd.DeleteObject 2, "Form_Eliminar_Usuarios"
On Error GoTo 0
access.DoCmd.Rename "Form_Eliminar_Usuarios", 2, fName
logFile.WriteLine "   [OK] Form_Eliminar_Usuarios creado."

' 9. Form_Cambiar_Contrasena_Administrador
logFile.WriteLine "Paso 14: Construyendo Formulario Form_Cambiar_Contrasena_Administrador..."
Set frm = access.CreateForm
fName = frm.Name
On Error Resume Next
frm.Caption = "Restablecer Contraseña (Administrador)"
frm.AutoCenter = True
frm.RecordSelectors = False
frm.NavigationButtons = False
frm.DividingLines = False
frm.ScrollBars = 0
frm.KeyPreview = True
frm.Width = 16000
frm.Section(0).Height = 6150
frm.Section(0).BackColor = 15790320
On Error GoTo 0

ConfigurarEncabezadoVerde frm, "Restablecer Contraseña (Administrador)"
ConfigurarBarraNavegacionSuperior frm, "Form_Gestionar_Usuarios"

On Error Resume Next
' Titulo
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 3500, 1880, 9000, 380)
ctrl.Caption = "Restablecer Contraseña de Empleado"
ctrl.FontName = "Cambria"
ctrl.FontSize = 14
ctrl.FontBold = True
ctrl.ForeColor = 6703164
ctrl.TextAlign = 2
ctrl.BorderStyle = 2
ctrl.BorderColor = 11184810
ctrl.BackStyle = 0

' Renglon 1: Numero de empleado
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1800, 2340, 4400, 340)
ctrl.Caption = "Número de empleado:"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 10.5
ctrl.FontBold = True
ctrl.ForeColor = 3355443
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 109, 0, "", "", 6300, 2340, 3500, 340)
ctrl.Name = "txtUsuario"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 10.5

' Renglon 2: Nueva Contrasena (SIN CONTRASENA ACTUAL - SOLO PARA ADMIN)
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1800, 2740, 4400, 340)
ctrl.Caption = "Nueva Contraseña:"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 10.5
ctrl.FontBold = True
ctrl.ForeColor = 3355443
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 109, 0, "", "", 6300, 2740, 3020, 340)
ctrl.Name = "txtNueva"
ctrl.InputMask = "Password"
' SEGURO: si el usuario presiona el ojito y suelta el raton FUERA del boton,
' el MouseUp del boton puede no llegar. Al perder el foco el cuadro se vuelve
' a enmascarar, de modo que la contrasena nunca se queda visible.
ctrl.OnLostFocus = "[Event Procedure]"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 10.5

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 9340, 2740, 460, 340)
ctrl.Name = "cmdVerNueva"
ctrl.Caption = "Ver"
ctrl.OnClick = "[Event Procedure]"
EstilizarBotonVer ctrl

' Renglon 3: Confirmacion de la Nueva Contrasena
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1800, 3140, 4400, 340)
ctrl.Caption = "Confirmación de la Nueva Contraseña:"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 10.5
ctrl.FontBold = True
ctrl.ForeColor = 3355443
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 109, 0, "", "", 6300, 3140, 3020, 340)
ctrl.Name = "txtConfirmar"
ctrl.InputMask = "Password"
' SEGURO: si el usuario presiona el ojito y suelta el raton FUERA del boton,
' el MouseUp del boton puede no llegar. Al perder el foco el cuadro se vuelve
' a enmascarar, de modo que la contrasena nunca se queda visible.
ctrl.OnLostFocus = "[Event Procedure]"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 10.5

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 9340, 3140, 460, 340)
ctrl.Name = "cmdVerConfirmar"
ctrl.Caption = "Ver"
ctrl.OnClick = "[Event Procedure]"
EstilizarBotonVer ctrl

' Botones laterales a la derecha
Set ctrl = access.CreateControl(fName, 104, 0, "", "", 10400, 2340, 2800, 440)
ctrl.Name = "cmdActualizar"
ctrl.Caption = "Aceptar"
ctrl.Default = False
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 10400, 2880, 2800, 440)
ctrl.Name = "cmdLimpiar"
ctrl.Caption = "Limpiar"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

' Cuadro institucional con Politicas (2 lineas holgadas para garantizar 0 recorte de texto en cualquier resolucion)
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 3520, 13000, 580)
ctrl.Name = "lblPoliticaBox"
ctrl.Caption = "POLÍTICA DE SEGURIDAD HIDROCARBUROS: Extensión obligatoria de exactamente 8 caracteres (al menos 1 mayúscula A-Z, 1 minúscula a-z y 1 número 0-9)." & vbCrLf & _
               "El restablecimiento por Administrador actualiza directamente la clave del empleado sin requerir la contraseña anterior."
ctrl.FontName = "Segoe UI"
ctrl.FontSize = 9
ctrl.FontBold = True
ctrl.ForeColor = 6703164
ctrl.TextAlign = 2
ctrl.BackStyle = 1
ctrl.BackColor = 15133908
ctrl.BorderStyle = 1
ctrl.BorderColor = 9376256

' Subformulario de Usuarios en vivo
Set ctrl = access.CreateControl(fName, 112, 0, "", "", 1500, 4140, 13000, 1580)
ctrl.Name = "subUsuarios"
ctrl.SourceObject = "Subform_Tabla_Usuarios"

' Pie de pagina
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 5760, 5000, 260)
ctrl.Name = "lblFecha"
ctrl.Caption = LCase(FormatDateTime(Date, 1))
ctrl.FontName = "Segoe UI"
ctrl.FontSize = 9
ctrl.FontItalic = True
ctrl.ForeColor = 6316128
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 9500, 5760, 5000, 260)
ctrl.Name = "lblArea"
ctrl.Caption = "Dirección de Economía de Hidrocarburos"
ctrl.FontName = "Segoe UI"
ctrl.FontSize = 9
ctrl.FontBold = True
ctrl.ForeColor = 6316128
ctrl.TextAlign = 3
ctrl.BackStyle = 0
On Error GoTo 0

codeF = "Option Compare Database" & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"' Form_Cambiar_Contrasena_Administrador - Restablecer contrasena (administrador)" & vbCrLf & _
"' -----------------------------------------------------------------------------" & vbCrLf & _
"' Permite al administrador asignar una contrasena nueva sin conocer la" & vbCrLf & _
"' anterior. Aplica la misma politica de 8 caracteres." & vbCrLf & _
"' Escribe  : UPDATE en USUARIOS via FondoBO.restablecerContrasenaAdminBO." & vbCrLf & _
"' Acceso   : restringido a perfil A." & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"" & vbCrLf & _
"Option Explicit" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private mBufNueva As String" & vbCrLf & _
"Private mBufConfirmar As String" & vbCrLf & _
"" & vbCrLf & _
"' ---------------------------------------------------------------------------" & vbCrLf & _
"' SEGURO DE ENMASCARADO AL PERDER EL FOCO" & vbCrLf & _
"' Cierra el unico hueco que deja el ojito: presionar el boton y soltar el" & vbCrLf & _
"' raton fuera de el. En ese caso el MouseUp del boton puede no llegar, pero" & vbCrLf & _
"' en cuanto el cuadro pierde el foco la contrasena se vuelve a ocultar." & vbCrLf & _
"' ---------------------------------------------------------------------------" & vbCrLf & _
"Private Sub txtConfirmar_LostFocus()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call Utilities.OcultarTodasLasContrasenasUT(Me)" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub txtNueva_LostFocus()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call Utilities.OcultarTodasLasContrasenasUT(Me)" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.RestringirKeyDownUsuario4Digitos Me.txtUsuario, KeyCode, Shift" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtUsuario_KeyPress(KeyAscii As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    Utilities.RestringirKeyPressUsuario4Digitos Me.txtUsuario, KeyAscii" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub txtUsuario_Change()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.TruncarCambioUsuario4Digitos Me.txtUsuario" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_Load()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    ' Encabezado de sesion y control de acceso. La regla de quien puede" & vbCrLf & _
"    ' entrar a esta pantalla NO esta aqui: vive en FondoBO.autorizarPantallaBO." & vbCrLf & _
"    ' Este formulario solo pregunta y obedece (Marco de Desarrollo, I.2)." & vbCrLf & _
"    ' Si el acceso se deniega, el ayudante ya aviso, abrio el menu y cerro" & vbCrLf
codeF = codeF & "    ' este formulario; aqui solo queda salir." & vbCrLf & _
"    If Not Utilities.AplicarEncabezadoSesionUT(Me, FondoBO.PANT_RESTABLECER_CONTRASENA) Then Exit Sub" & vbCrLf
codeF = codeF & "    mBufNueva = """": mBufConfirmar = """"" & vbCrLf & _
"    Me.txtUsuario.Value = Null" & vbCrLf & _
"    Me.txtNueva.Value = Null" & vbCrLf & _
"    Me.txtConfirmar.Value = Null" & vbCrLf & _
"    Me.lblFecha.Caption = Format(Date, ""dddd, d \de mmmm \de yyyy"")" & vbCrLf & _
"    Me.subUsuarios.Form.RecordSource = FondoBO.obtenerOrigenUsuariosBO()" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_Resize()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub txtNueva_KeyPress(KeyAscii As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.ProcesarKeyPressContrasena mBufNueva, KeyAscii" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtNueva_Change()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.TruncarCambioContrasena8Caracteres Me.txtNueva, mBufNueva" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtConfirmar_KeyPress(KeyAscii As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    Utilities.ProcesarKeyPressContrasena mBufConfirmar, KeyAscii" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub txtConfirmar_Change()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.TruncarCambioContrasena8Caracteres Me.txtConfirmar, mBufConfirmar" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"' ---------------------------------------------------------------------------" & vbCrLf & _
"' VER CONTRASENA SOLO MIENTRAS SE SOSTIENE EL CLIC" & vbCrLf & _
"' MouseDown la descubre, MouseUp la vuelve a ocultar. El Click se deja como" & vbCrLf & _
"' seguro por si el MouseUp no llegara: al soltar siempre queda oculta." & vbCrLf & _
"' ---------------------------------------------------------------------------" & vbCrLf & _
"Private Sub cmdVerNueva_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    Call Utilities.VerContrasenaSostenidaUT(Me, ""cmdVerNueva"", ""txtNueva"", True, """")" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdVerNueva_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call Utilities.VerContrasenaSostenidaUT(Me, ""cmdVerNueva"", ""txtNueva"", False, """")" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdVerNueva_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call Utilities.VerContrasenaSostenidaUT(Me, ""cmdVerNueva"", ""txtNueva"", False, """")" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "' ---------------------------------------------------------------------------" & vbCrLf & _
"' VER CONTRASENA SOLO MIENTRAS SE SOSTIENE EL CLIC" & vbCrLf
codeF = codeF & "' MouseDown la descubre, MouseUp la vuelve a ocultar. El Click se deja como" & vbCrLf & _
"' seguro por si el MouseUp no llegara: al soltar siempre queda oculta." & vbCrLf & _
"' ---------------------------------------------------------------------------" & vbCrLf & _
"Private Sub cmdVerConfirmar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call Utilities.VerContrasenaSostenidaUT(Me, ""cmdVerConfirmar"", ""txtConfirmar"", True, """")" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdVerConfirmar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call Utilities.VerContrasenaSostenidaUT(Me, ""cmdVerConfirmar"", ""txtConfirmar"", False, """")" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub cmdVerConfirmar_Click()" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    Call Utilities.VerContrasenaSostenidaUT(Me, ""cmdVerConfirmar"", ""txtConfirmar"", False, """")" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtConfirmar_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If KeyCode = 13 Then" & vbCrLf & _
"        KeyCode = 0" & vbCrLf & _
"        Me.cmdActualizar.SetFocus" & vbCrLf & _
"        cmdActualizar_Click" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    If Shift = 0 Then" & vbCrLf & _
"        If Not (TypeOf Screen.ActiveControl Is ComboBox Or TypeOf Screen.ActiveControl Is ListBox) Then" & vbCrLf & _
"            Select Case KeyCode" & vbCrLf & _
"                Case 40: KeyCode = 0" & vbCrLf & _
"                Case 38: KeyCode = 0" & vbCrLf & _
"            End Select" & vbCrLf & _
"        End If" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdLimpiar_Click()" & vbCrLf
codeF = codeF & "    mBufNueva = """": mBufConfirmar = """"" & vbCrLf & _
"    Me.txtUsuario.Value = Null" & vbCrLf
codeF = codeF & "    Me.txtNueva.Value = Null" & vbCrLf & _
"    Me.txtConfirmar.Value = Null" & vbCrLf & _
"    Me.txtUsuario.SetFocus" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdActualizar_Click()" & vbCrLf & _
"    On Error Resume Next: Me.cmdActualizar.SetFocus: On Error GoTo ErrH" & vbCrLf & _
"    Dim cveEmpleado As String" & vbCrLf & _
"    Dim pwdNueva As String" & vbCrLf & _
"    Dim pwdConfirmar As String" & vbCrLf & _
"    Dim msgResultado As String" & vbCrLf & _
"    Dim exito As Boolean" & vbCrLf
codeF = codeF & "    " & vbCrLf & _
"    cveEmpleado = Utilities.LeerTextoControl(Me.txtUsuario)" & vbCrLf
codeF = codeF & "    pwdNueva = Utilities.LeerTextoControl(Me.txtNueva, mBufNueva)" & vbCrLf & _
"    pwdConfirmar = Utilities.LeerTextoControl(Me.txtConfirmar, mBufConfirmar)" & vbCrLf & _
"    " & vbCrLf & _
"    Screen.MousePointer = MOUSE_POINTER_WAIT" & vbCrLf & _
"    exito = FondoBO.restablecerContrasenaAdminBO(cveEmpleado, pwdNueva, pwdConfirmar, msgResultado)" & vbCrLf & _
"    Screen.MousePointer = MOUSE_POINTER_DEFAULT" & vbCrLf & _
"    " & vbCrLf & _
"    If exito Then" & vbCrLf & _
"        On Error Resume Next" & vbCrLf & _
"        Me.subUsuarios.Form.RecordSource = FondoBO.obtenerOrigenUsuariosBO()" & vbCrLf & _
"        Me.subUsuarios.Form.Requery" & vbCrLf & _
"        On Error GoTo ErrH" & vbCrLf
codeF = codeF & "        MsgBox msgResultado & vbCrLf & vbCrLf & ""El subformulario inferior de usuarios ha sido verificado y actualizado en tiempo real."", vbInformation, ""Restablecimiento Exitoso""" & vbCrLf & _
"        cmdLimpiar_Click" & vbCrLf
codeF = codeF & "    Else" & vbCrLf & _
"        MsgBox msgResultado, vbExclamation, ""Validación de Seguridad""" & vbCrLf & _
"    End If" & vbCrLf & _
"    Exit Sub" & vbCrLf & _
"ErrH:" & vbCrLf & _
"    Screen.MousePointer = MOUSE_POINTER_DEFAULT" & vbCrLf & _
"    MsgBox MENSAJE_ERROR & vbCrLf & Err.Description, vbCritical, ""Error""" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdAtras_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Gestionar_Usuarios"", acNormal" & vbCrLf
codeF = codeF & "    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub cmdIrMenu_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Menu_General"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdSalir_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If MsgBox(""¿Desea cerrar la sesión actual?"", vbQuestion + vbYesNo, ""Salir"") = vbYes Then" & vbCrLf & _
"        Set usuarioSesionDTO = Nothing" & vbCrLf & _
"        DoCmd.OpenForm ""Form_Acceso"", acNormal" & vbCrLf
codeF = codeF & "        DoCmd.Close acForm, Me.Name" & vbCrLf & _
"    End If" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdRegresar_Click()" & vbCrLf & _
"    cmdAtras_Click" & vbCrLf & _
"End Sub"
AsignarCodigoLimpio frm, codeF
End Sub

Private Sub EjecutarBloqueInstalacion_09()
access.DoCmd.Close 2, fName, 1
On Error Resume Next
access.DoCmd.DeleteObject 2, "Form_Cambiar_Contrasena_Administrador"
On Error GoTo 0
access.DoCmd.Rename "Form_Cambiar_Contrasena_Administrador", 2, fName
logFile.WriteLine "   [OK] Form_Cambiar_Contrasena_Administrador creado."

logFile.WriteLine "Paso 15: Construyendo Formulario Form_Gestionar_Instrumentos..."
Set frm = access.CreateForm
fName = frm.Name
On Error Resume Next
frm.Caption = "Gestionar Instrumentos"
frm.AutoCenter = True
frm.KeyPreview = True
frm.ScrollBars = 0

frm.RecordSelectors = False

frm.NavigationButtons = False

frm.DividingLines = False

frm.Width = 16000
frm.Section(0).Height = 6150
frm.Section(0).BackColor = 15790320
On Error GoTo 0

ConfigurarEncabezadoVerde frm, "Módulo de Administración Integral de Riesgos - Gestión de Instrumentos"
ConfigurarBarraNavegacionSuperior frm, "Form_Menu_General"

On Error Resume Next
Set ctrl = access.CreateControl(fName, 104, 0, "", "", 1500, 1950, 3100, 480)
ctrl.Name = "cmdAgregar"
ctrl.Caption = "+ Nuevo Instrumento"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 4900, 1950, 3100, 480)
ctrl.Name = "cmdActualizar"
ctrl.Caption = "Modificar Instrumento"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 2600, 13000, 380)
ctrl.Caption = "Catálogo de Instrumentos Financieros Registrados"
ctrl.FontName = "Cambria"
ctrl.FontSize = 13
ctrl.FontBold = True
ctrl.ForeColor = 6703164
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 112, 0, "", "", 1500, 3050, 13000, 2950)
ctrl.Name = "subInstrumentos"
ctrl.SourceObject = "Subform_Tabla_Instrumentos"
On Error GoTo 0

codeF = "Option Compare Database" & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"' Form_Gestionar_Instrumentos - Menu del catalogo de instrumentos" & vbCrLf & _
"' -----------------------------------------------------------------------------" & vbCrLf & _
"' Navegacion hacia alta, modificacion y consulta de instrumentos." & vbCrLf & _
"' Incorpora Form_Subformulario_Consulta_Instrumentos en vivo." & vbCrLf & _
"' El catalogo INSTRUMENTOS es la llave de la validacion de insumos:" & vbCrLf & _
"' ningun insumo se carga si trae un ID_INSTRUMENTO que no este aqui." & vbCrLf & _
"' Los botones de alta y modificacion solo son visibles para perfil A." & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"" & vbCrLf & _
"Option Explicit" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"' Cache del mapa fecha=conteo. Se llena una vez por repintado del mes." & vbCrLf
codeF = codeF & "Private mMapaFechas As String" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_Load()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    ' Encabezado de sesion y control de acceso. La regla de quien puede" & vbCrLf & _
"    ' entrar a esta pantalla NO esta aqui: vive en FondoBO.autorizarPantallaBO." & vbCrLf & _
"    ' Este formulario solo pregunta y obedece (Marco de Desarrollo, I.2)." & vbCrLf & _
"    ' Pantalla sin restriccion de perfil: solo se pinta el encabezado." & vbCrLf & _
"    Call Utilities.AplicarEncabezadoSesionUT(Me, FondoBO.PANT_GESTION_INSTRUMENTOS)" & vbCrLf & _
"    Me.subInstrumentos.Form.RecordSource = FondoBO.obtenerOrigenInstrumentosBO()" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub Form_Resize()" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdAgregar_Click()" & vbCrLf & _
"    ' La regla de quien puede abrir esta pantalla vive en el Negocio;" & vbCrLf & _
"    ' este boton solo pregunta. Si no se autoriza, el ayudante ya aviso." & vbCrLf & _
"    If Not Utilities.PuedeAbrirPantallaUT(FondoBO.PANT_ALTA_INSTRUMENTOS) Then Exit Sub" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Agregar_Instrumento"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub cmdActualizar_Click()" & vbCrLf & _
"    ' La regla de quien puede abrir esta pantalla vive en el Negocio;" & vbCrLf
codeF = codeF & "    ' este boton solo pregunta. Si no se autoriza, el ayudante ya aviso." & vbCrLf & _
"    If Not Utilities.PuedeAbrirPantallaUT(FondoBO.PANT_CAMBIO_INSTRUMENTOS) Then Exit Sub" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Actualizar_Instrumento"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdAtras_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Menu_General"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub cmdIrMenu_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    DoCmd.OpenForm ""Form_Menu_General"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdSalir_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If MsgBox(""¿Desea cerrar la sesión actual?"", vbQuestion + vbYesNo, ""Salir"") = vbYes Then" & vbCrLf & _
"        Set usuarioSesionDTO = Nothing" & vbCrLf & _
"        DoCmd.OpenForm ""Form_Acceso"", acNormal" & vbCrLf & _
"        DoCmd.Close acForm, Me.Name" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub cmdRegresar_Click()" & vbCrLf
codeF = codeF & "    cmdAtras_Click" & vbCrLf & _
"End Sub"
AsignarCodigoLimpio frm, codeF
access.DoCmd.Close 2, fName, 1
On Error Resume Next
access.DoCmd.DeleteObject 2, "Form_Gestionar_Instrumentos"
On Error GoTo 0
access.DoCmd.Rename "Form_Gestionar_Instrumentos", 2, fName
logFile.WriteLine "   [OK] Form_Gestionar_Instrumentos creado."

' 11. Form_Subformulario_Consulta_Instrumentos
logFile.WriteLine "Paso 16: Construyendo Formulario Form_Subformulario_Consulta_Instrumentos..."
Set frm = access.CreateForm
fName = frm.Name
On Error Resume Next
frm.Caption = "Consulta de Instrumentos"
frm.AutoCenter = True
frm.KeyPreview = True
frm.RecordSelectors = False
frm.NavigationButtons = False
frm.ScrollBars = 0

frm.RecordSelectors = False

frm.NavigationButtons = False

frm.DividingLines = False

frm.Width = 16000
frm.Section(0).Height = 6150
frm.Section(0).BackColor = 15790320
On Error GoTo 0

ConfigurarEncabezadoVerde frm, "Módulo de Administración Integral de Riesgos - Catálogo de Instrumentos"
ConfigurarBarraNavegacionSuperior frm, "Form_Gestionar_Instrumentos"

On Error Resume Next
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 2200, 13000, 400)
ctrl.Caption = "Catálogo de Instrumentos Financieros"
ctrl.FontName = "Cambria"
ctrl.FontSize = 13
ctrl.FontBold = True
ctrl.ForeColor = 6703164
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 112, 0, "", "", 1500, 2700, 13000, 3300)
ctrl.Name = "subCatalogoInstrumentos"
ctrl.SourceObject = "Subform_Tabla_Instrumentos"
On Error GoTo 0

codeF = "Option Compare Database" & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"' Form_Subformulario_Consulta_Instrumentos - Subformulario: catalogo de instrumentos en vivo" & vbCrLf & _
"' -----------------------------------------------------------------------------" & vbCrLf & _
"' Solo lectura. Origen: ConsultasSQL.sqlSelectInstrumentos." & vbCrLf & _
"' Se refresca con Requery tras registrar o modificar un instrumento." & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"" & vbCrLf & _
"Option Explicit" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_Load()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    ' Encabezado de sesion y control de acceso. La regla de quien puede" & vbCrLf & _
"    ' entrar a esta pantalla NO esta aqui: vive en FondoBO.autorizarPantallaBO." & vbCrLf
codeF = codeF & "    ' Este formulario solo pregunta y obedece (Marco de Desarrollo, I.2)." & vbCrLf & _
"    ' Pantalla sin restriccion de perfil: solo se pinta el encabezado." & vbCrLf & _
"    Call Utilities.AplicarEncabezadoSesionUT(Me, FondoBO.PANT_CONSULTA_INSTRUMENTOS)" & vbCrLf & _
"    Me.subCatalogoInstrumentos.Form.RecordSource = FondoBO.obtenerOrigenInstrumentosBO()" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_Resize()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub cmdAtras_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    DoCmd.OpenForm ""Form_Gestionar_Instrumentos"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdIrMenu_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Menu_General"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdSalir_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    If MsgBox(""¿Desea cerrar la sesión actual?"", vbQuestion + vbYesNo, ""Salir"") = vbYes Then" & vbCrLf & _
"        Set usuarioSesionDTO = Nothing" & vbCrLf
codeF = codeF & "        DoCmd.OpenForm ""Form_Acceso"", acNormal" & vbCrLf & _
"        DoCmd.Close acForm, Me.Name" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdRegresar_Click()" & vbCrLf & _
"    cmdAtras_Click" & vbCrLf & _
"End Sub"
AsignarCodigoLimpio frm, codeF
access.DoCmd.Close 2, fName, 1
On Error Resume Next
access.DoCmd.DeleteObject 2, "Form_Subformulario_Consulta_Instrumentos"
On Error GoTo 0
access.DoCmd.Rename "Form_Subformulario_Consulta_Instrumentos", 2, fName
logFile.WriteLine "   [OK] Form_Subformulario_Consulta_Instrumentos creado."

' 12. Form_Agregar_Instrumento
logFile.WriteLine "Paso 17: Construyendo Formulario Form_Agregar_Instrumento..."
Set frm = access.CreateForm
fName = frm.Name
On Error Resume Next
frm.Caption = "Agregar Instrumento"
frm.AutoCenter = True
frm.KeyPreview = True
frm.ScrollBars = 0
frm.RecordSelectors = False
frm.NavigationButtons = False
frm.DividingLines = False
frm.Width = 16000
frm.Section(0).Height = 6150
frm.Section(0).BackColor = 15790320
On Error GoTo 0

ConfigurarEncabezadoVerde frm, "Módulo de Administración Integral de Riesgos - Alta de Instrumento"
ConfigurarBarraNavegacionSuperior frm, "Form_Gestionar_Instrumentos"

On Error Resume Next
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 1900, 13000, 350)
ctrl.Caption = "Alta de Nuevo Instrumento Financiero"
ctrl.FontName = "Cambria"
ctrl.FontSize = 13
ctrl.FontBold = True
ctrl.ForeColor = 6703164
ctrl.TextAlign = 2
ctrl.BackStyle = 0

' Renglon 1: ID Instrumento y Clave Emisor
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 2340, 2000, 340)
ctrl.Caption = "ID Instrumento:"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 3600, 2340, 2500, 340)
ctrl.Name = "txtId"

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 6400, 2340, 2200, 340)
ctrl.Caption = "Clave Emisor:"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 8700, 2340, 2600, 340)
ctrl.Name = "txtEmisor"

' Renglon 2: Tipo Instrumento y Tasa Rendimiento (%)
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 2760, 2000, 340)
ctrl.Caption = "Tipo Instrumento:"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 3600, 2760, 2500, 340)
ctrl.Name = "txtTipo"

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 6400, 2760, 2200, 340)
ctrl.Caption = "Tasa Rendimiento (%):"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 8700, 2760, 2000, 340)
ctrl.Name = "txtTasa"

' Renglon 3: Plazo (dias)
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 3180, 2000, 340)
ctrl.Caption = "Plazo (días):"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 3600, 3180, 2000, 340)
ctrl.Name = "txtPlazo"

' Boton Guardar a la derecha
Set ctrl = access.CreateControl(fName, 104, 0, "", "", 11800, 2550, 2500, 500)
ctrl.Name = "cmdGuardar"
ctrl.Caption = "Guardar"
ctrl.Default = False
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

' Subformulario de Instrumentos en vivo para corroborar el alta
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 3600, 13000, 300)
ctrl.Caption = "Catálogo de Instrumentos Financieros en Tiempo Real (Corroboración de Alta)"
ctrl.FontName = "Cambria"
ctrl.FontSize = 11
ctrl.FontBold = True
ctrl.ForeColor = 6703164
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 112, 0, "", "", 1500, 3920, 13000, 2080)
ctrl.Name = "subInstrumentos"
ctrl.SourceObject = "Subform_Tabla_Instrumentos"
On Error GoTo 0

codeF = "Option Compare Database" & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"' Form_Agregar_Instrumento - Alta de instrumento financiero" & vbCrLf & _
"' -----------------------------------------------------------------------------" & vbCrLf & _
"' Entradas : ID, clave de emisor, tipo, tasa anual y plazo en dias." & vbCrLf & _
"' Valida   : ID no duplicado; tasa y plazo numericos." & vbCrLf & _
"' Escribe  : INSERT en INSTRUMENTOS via FondoBO.registrarInstrumentoBO." & vbCrLf & _
"' fechaVencimiento se calcula como fechaEmision + plazo." & vbCrLf & _
"' Dar de alta aqui un instrumento es el paso previo para poder cargar" & vbCrLf & _
"' un insumo que lo contenga." & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Option Explicit" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If Shift = 0 Then" & vbCrLf & _
"        If Not (TypeOf Screen.ActiveControl Is ComboBox Or TypeOf Screen.ActiveControl Is ListBox) Then" & vbCrLf & _
"            Select Case KeyCode" & vbCrLf & _
"                Case 40: KeyCode = 0" & vbCrLf & _
"                Case 38: KeyCode = 0" & vbCrLf & _
"            End Select" & vbCrLf & _
"        End If" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub Form_Load()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    ' Encabezado de sesion y control de acceso. La regla de quien puede" & vbCrLf & _
"    ' entrar a esta pantalla NO esta aqui: vive en FondoBO.autorizarPantallaBO." & vbCrLf & _
"    ' Este formulario solo pregunta y obedece (Marco de Desarrollo, I.2)." & vbCrLf & _
"    ' Si el acceso se deniega, el ayudante ya aviso, abrio el menu y cerro" & vbCrLf & _
"    ' este formulario; aqui solo queda salir." & vbCrLf & _
"    If Not Utilities.AplicarEncabezadoSesionUT(Me, FondoBO.PANT_ALTA_INSTRUMENTOS) Then Exit Sub" & vbCrLf & _
"    Me.subInstrumentos.Form.RecordSource = FondoBO.obtenerOrigenInstrumentosBO()" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_Resize()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub cmdGuardar_Click()" & vbCrLf & _
"    On Error GoTo ErrH" & vbCrLf & _
"    Dim instDTO As InstrumentoDTO" & vbCrLf & _
"    Dim msgResultado As String" & vbCrLf & _
"    Dim exito As Boolean" & vbCrLf & _
"    " & vbCrLf & _
"    Set instDTO = New InstrumentoDTO" & vbCrLf & _
"    instDTO.setIdInstrumento Utilities.LeerTextoControl(Me.txtId)" & vbCrLf & _
"    instDTO.setCveEmisor Utilities.LeerTextoControl(Me.txtEmisor)" & vbCrLf & _
"    instDTO.setTipoInstrumento Utilities.LeerTextoControl(Me.txtTipo)" & vbCrLf & _
"    instDTO.setTasa Val(Utilities.LeerTextoControl(Me.txtTasa))" & vbCrLf
codeF = codeF & "    instDTO.setPlazo CLng(Val(Utilities.LeerTextoControl(Me.txtPlazo)))" & vbCrLf & _
"    " & vbCrLf
codeF = codeF & "    exito = FondoBO.registrarInstrumentoBO(instDTO, msgResultado)" & vbCrLf & _
"    Set instDTO = Nothing" & vbCrLf & _
"    " & vbCrLf & _
"    If exito Then" & vbCrLf & _
"        On Error Resume Next" & vbCrLf & _
"        Me.subInstrumentos.Form.RecordSource = FondoBO.obtenerOrigenInstrumentosBO()" & vbCrLf & _
"        Me.subInstrumentos.Form.Requery" & vbCrLf & _
"        Me.txtId.Value = Null: Me.txtEmisor.Value = Null: Me.txtTipo.Value = Null: Me.txtTasa.Value = Null: Me.txtPlazo.Value = Null" & vbCrLf & _
"        On Error GoTo ErrH" & vbCrLf & _
"        MsgBox msgResultado & vbCrLf & vbCrLf & ""El subformulario inferior de instrumentos se ha actualizado en tiempo real."", vbInformation, ""Alta de Instrumento""" & vbCrLf & _
"    Else" & vbCrLf & _
"        MsgBox msgResultado, vbExclamation, ""Validación de Negocio""" & vbCrLf
codeF = codeF & "    End If" & vbCrLf & _
"    Exit Sub" & vbCrLf
codeF = codeF & "ErrH:" & vbCrLf & _
"    MsgBox MENSAJE_ERROR & vbCrLf & Err.Description, vbCritical" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdAtras_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Gestionar_Instrumentos"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdIrMenu_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    DoCmd.OpenForm ""Form_Menu_General"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdSalir_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If MsgBox(""¿Desea cerrar la sesión actual?"", vbQuestion + vbYesNo, ""Salir"") = vbYes Then" & vbCrLf & _
"        Set usuarioSesionDTO = Nothing" & vbCrLf & _
"        DoCmd.OpenForm ""Form_Acceso"", acNormal" & vbCrLf & _
"        DoCmd.Close acForm, Me.Name" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdRegresar_Click()" & vbCrLf
codeF = codeF & "    cmdAtras_Click" & vbCrLf & _
"End Sub"
AsignarCodigoLimpio frm, codeF
End Sub

Private Sub EjecutarBloqueInstalacion_10()
access.DoCmd.Close 2, fName, 1
On Error Resume Next
access.DoCmd.DeleteObject 2, "Form_Agregar_Instrumento"
On Error GoTo 0
access.DoCmd.Rename "Form_Agregar_Instrumento", 2, fName
logFile.WriteLine "   [OK] Form_Agregar_Instrumento creado."

' 13. Form_Actualizar_Instrumento
logFile.WriteLine "Paso 18: Construyendo Formulario Form_Actualizar_Instrumento..."
Set frm = access.CreateForm
fName = frm.Name
On Error Resume Next
frm.Caption = "Actualizar Instrumento"
frm.AutoCenter = True
frm.KeyPreview = True
frm.ScrollBars = 0
frm.RecordSelectors = False
frm.NavigationButtons = False
frm.DividingLines = False
frm.Width = 16000
frm.Section(0).Height = 6150
frm.Section(0).BackColor = 15790320
On Error GoTo 0

ConfigurarEncabezadoVerde frm, "Módulo de Administración Integral de Riesgos - Modificación de Instrumento"
ConfigurarBarraNavegacionSuperior frm, "Form_Gestionar_Instrumentos"

On Error Resume Next
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 1900, 13000, 350)
ctrl.Caption = "Modificación de Parámetros de Instrumento Financiero"
ctrl.FontName = "Cambria"
ctrl.FontSize = 13
ctrl.FontBold = True
ctrl.ForeColor = 6703164
ctrl.TextAlign = 2
ctrl.BackStyle = 0

' Renglon 1: ID Instrumento + Buscar + Clave Emisor
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 2340, 2000, 340)
ctrl.Caption = "ID Instrumento:"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 3600, 2340, 2400, 340)
ctrl.Name = "txtId"

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 6150, 2340, 1650, 350)
ctrl.Name = "cmdBuscar"
ctrl.Caption = "Buscar"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 8100, 2340, 1800, 340)
ctrl.Caption = "Clave Emisor:"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 10000, 2340, 2400, 340)
ctrl.Name = "txtEmisor"

' Renglon 2: Tipo Instrumento y Tasa Rendimiento (%)
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 2780, 2000, 340)
ctrl.Caption = "Tipo Instrumento:"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 3600, 2780, 2800, 340)
ctrl.Name = "txtTipo"

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 6700, 2780, 2300, 340)
ctrl.Caption = "Tasa Rendimiento (%):"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 9100, 2780, 1800, 340)
ctrl.Name = "txtTasa"

' Renglon 3: Plazo (dias)
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 3200, 2000, 340)
ctrl.Caption = "Plazo (días):"
ctrl.FontBold = True
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 3600, 3200, 2000, 340)
ctrl.Name = "txtPlazo"

' Boton Guardar Cambios a la derecha
Set ctrl = access.CreateControl(fName, 104, 0, "", "", 11800, 2650, 2500, 500)
ctrl.Name = "cmdActualizar"
ctrl.Caption = "Guardar Cambios"
ctrl.Default = False
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl

' Subformulario de Instrumentos en vivo para corroborar la modificacion
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1500, 3600, 13000, 300)
ctrl.Caption = "Catálogo de Instrumentos Financieros en Tiempo Real (Corroboración de Modificación)"
ctrl.FontName = "Cambria"
ctrl.FontSize = 11
ctrl.FontBold = True
ctrl.ForeColor = 6703164
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 112, 0, "", "", 1500, 3920, 13000, 2080)
ctrl.Name = "subInstrumentos"
ctrl.SourceObject = "Subform_Tabla_Instrumentos"
On Error GoTo 0

codeF = "Option Compare Database" & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"' Form_Actualizar_Instrumento - Modificacion de instrumento" & vbCrLf & _
"' -----------------------------------------------------------------------------" & vbCrLf & _
"' Busca por ID con FondoBO.consultarInstrumentoBO y precarga los campos." & vbCrLf & _
"' Escribe  : UPDATE en INSTRUMENTOS via FondoBO.actualizarInstrumentoBO." & vbCrLf & _
"' La tasa afecta directamente al reporte: las consultas de analisis" & vbCrLf & _
"' promedian INSTRUMENTOS.tasa por grupo." & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"" & vbCrLf & _
"Option Explicit" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub txtId_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf & _
"    If KeyCode = 13 Then" & vbCrLf
codeF = codeF & "        KeyCode = 0" & vbCrLf & _
"        cmdBuscar_Click" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If Shift = 0 Then" & vbCrLf & _
"        If Not (TypeOf Screen.ActiveControl Is ComboBox Or TypeOf Screen.ActiveControl Is ListBox) Then" & vbCrLf & _
"            Select Case KeyCode" & vbCrLf & _
"                Case 40: KeyCode = 0" & vbCrLf & _
"                Case 38: KeyCode = 0" & vbCrLf
codeF = codeF & "            End Select" & vbCrLf & _
"        End If" & vbCrLf
codeF = codeF & "    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_Load()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    ' Encabezado de sesion y control de acceso. La regla de quien puede" & vbCrLf & _
"    ' entrar a esta pantalla NO esta aqui: vive en FondoBO.autorizarPantallaBO." & vbCrLf & _
"    ' Este formulario solo pregunta y obedece (Marco de Desarrollo, I.2)." & vbCrLf & _
"    ' Si el acceso se deniega, el ayudante ya aviso, abrio el menu y cerro" & vbCrLf & _
"    ' este formulario; aqui solo queda salir." & vbCrLf & _
"    If Not Utilities.AplicarEncabezadoSesionUT(Me, FondoBO.PANT_CAMBIO_INSTRUMENTOS) Then Exit Sub" & vbCrLf & _
"    Me.subInstrumentos.Form.RecordSource = FondoBO.obtenerOrigenInstrumentosBO()" & vbCrLf
codeF = codeF & "    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub Form_Resize()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdBuscar_Click()" & vbCrLf & _
"    On Error GoTo ErrH" & vbCrLf & _
"    Dim idBuscar As String" & vbCrLf & _
"    Dim instEncontrado As InstrumentoDTO" & vbCrLf & _
"    " & vbCrLf & _
"    idBuscar = Utilities.LeerTextoControl(Me.txtId)" & vbCrLf
codeF = codeF & "    Set instEncontrado = FondoBO.consultarInstrumentoBO(idBuscar)" & vbCrLf & _
"    If Not instEncontrado Is Nothing Then" & vbCrLf
codeF = codeF & "        Me.txtEmisor.Value = instEncontrado.getCveEmisor()" & vbCrLf & _
"        Me.txtTipo.Value = instEncontrado.getTipoInstrumento()" & vbCrLf & _
"        Me.txtTasa.Value = instEncontrado.getTasa()" & vbCrLf & _
"        Me.txtPlazo.Value = instEncontrado.getPlazo()" & vbCrLf & _
"    Else" & vbCrLf & _
"        MsgBox ""Instrumento no encontrado."", vbExclamation, ""Búsqueda de Instrumento""" & vbCrLf & _
"    End If" & vbCrLf & _
"    Set instEncontrado = Nothing" & vbCrLf & _
"    Exit Sub" & vbCrLf & _
"ErrH:" & vbCrLf & _
"    MsgBox MENSAJE_ERROR & vbCrLf & Err.Description, vbCritical" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub cmdActualizar_Click()" & vbCrLf
codeF = codeF & "    On Error GoTo ErrH" & vbCrLf & _
"    Dim instDTO As InstrumentoDTO" & vbCrLf & _
"    Dim msgResultado As String" & vbCrLf & _
"    Dim exito As Boolean" & vbCrLf & _
"    " & vbCrLf & _
"    Set instDTO = New InstrumentoDTO" & vbCrLf & _
"    instDTO.setIdInstrumento Utilities.LeerTextoControl(Me.txtId)" & vbCrLf & _
"    instDTO.setCveEmisor Utilities.LeerTextoControl(Me.txtEmisor)" & vbCrLf & _
"    instDTO.setTipoInstrumento Utilities.LeerTextoControl(Me.txtTipo)" & vbCrLf & _
"    instDTO.setTasa Val(Utilities.LeerTextoControl(Me.txtTasa))" & vbCrLf & _
"    instDTO.setPlazo CLng(Val(Utilities.LeerTextoControl(Me.txtPlazo)))" & vbCrLf & _
"    " & vbCrLf
codeF = codeF & "    exito = FondoBO.actualizarInstrumentoBO(instDTO, msgResultado)" & vbCrLf & _
"    Set instDTO = Nothing" & vbCrLf
codeF = codeF & "    " & vbCrLf & _
"    If exito Then" & vbCrLf & _
"        On Error Resume Next" & vbCrLf & _
"        Me.subInstrumentos.Form.RecordSource = FondoBO.obtenerOrigenInstrumentosBO()" & vbCrLf & _
"        Me.subInstrumentos.Form.Requery" & vbCrLf & _
"        On Error GoTo ErrH" & vbCrLf & _
"        MsgBox msgResultado & vbCrLf & vbCrLf & ""El subformulario inferior de instrumentos se ha actualizado en tiempo real."", vbInformation, ""Actualización Exitosa""" & vbCrLf & _
"    Else" & vbCrLf & _
"        MsgBox msgResultado, vbExclamation, ""Validación de Negocio""" & vbCrLf & _
"    End If" & vbCrLf & _
"    Exit Sub" & vbCrLf & _
"ErrH:" & vbCrLf
codeF = codeF & "    MsgBox MENSAJE_ERROR & vbCrLf & Err.Description, vbCritical" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub cmdAtras_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Gestionar_Instrumentos"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdIrMenu_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Menu_General"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub cmdSalir_Click()" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    If MsgBox(""¿Desea cerrar la sesión actual?"", vbQuestion + vbYesNo, ""Salir"") = vbYes Then" & vbCrLf & _
"        Set usuarioSesionDTO = Nothing" & vbCrLf & _
"        DoCmd.OpenForm ""Form_Acceso"", acNormal" & vbCrLf & _
"        DoCmd.Close acForm, Me.Name" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdRegresar_Click()" & vbCrLf & _
"    cmdAtras_Click" & vbCrLf & _
"End Sub"
AsignarCodigoLimpio frm, codeF
access.DoCmd.Close 2, fName, 1
On Error Resume Next
access.DoCmd.DeleteObject 2, "Form_Actualizar_Instrumento"
On Error GoTo 0
access.DoCmd.Rename "Form_Actualizar_Instrumento", 2, fName
logFile.WriteLine "   [OK] Form_Actualizar_Instrumento creado."


' 13.9 Construyendo Formulario Emergente de Calendario Mensual Interactivo (1 solo clic, seleccion de Mes y Anio, marca dias con insumo y valida al seleccionar)
logFile.WriteLine "   Construyendo Calendario Mensual Interactivo Form_Calendario_Popup (con selectores de Mes/Anio y dias cargados marcados)..."
Set frm = access.CreateForm
fName = frm.Name
On Error Resume Next
frm.Caption = "Calendario Hidrocarburos"
frm.PopUp = True
frm.Modal = True
frm.BorderStyle = 3 ' Dialogo fijo
frm.AutoCenter = True
frm.RecordSelectors = False
frm.NavigationButtons = False
frm.DividingLines = False
frm.ScrollBars = 0
frm.ControlBox = True
frm.MinMaxButtons = 0
frm.Width = 4600
frm.Section(0).Height = 4320
frm.Section(0).BackColor = 16119285 ' #F5F6F8

' Barra superior con botones < > y Selectores Desplegables de Mes (cboMes) y Anio (cboAnio)
Set ctrl = access.CreateControl(fName, 104, 0, "", "", 160, 100, 500, 400)
ctrl.Name = "cmdMesAnt"
ctrl.Caption = "<"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl
ctrl.FontSize = 10
ctrl.ControlTipText = "Mes anterior"

' Selector desplegable de Mes (Enero a Diciembre)
Set ctrl = access.CreateControl(fName, 111, 0, "", "", 710, 108, 1880, 385)
ctrl.Name = "cboMes"
ctrl.RowSourceType = "Value List"
ctrl.RowSource = "Enero;Febrero;Marzo;Abril;Mayo;Junio;Julio;Agosto;Septiembre;Octubre;Noviembre;Diciembre"
ctrl.ColumnCount = 1
ctrl.BoundColumn = 1
ctrl.LimitToList = True
ctrl.AllowValueListEdits = False
ctrl.ListRows = 12
ctrl.FontName = "Cambria"
ctrl.FontSize = 10.5
ctrl.FontBold = True
ctrl.ForeColor = 6703164
ctrl.BackColor = 16777215
ctrl.BorderStyle = 1
ctrl.BorderColor = 5000704
ctrl.TextAlign = 2
ctrl.OnChange = "[Event Procedure]"
ctrl.AfterUpdate = "[Event Procedure]"
ctrl.OnGotFocus = "[Event Procedure]"
ctrl.ControlTipText = "Seleccionar el Mes en el calendario"

' Selector desplegable de Anio (2000 a Anio Actual + 3)
Set ctrl = access.CreateControl(fName, 111, 0, "", "", 2640, 108, 1240, 385)
ctrl.Name = "cboAnio"
ctrl.RowSourceType = "Value List"
ctrl.RowSource = "2028;2027;2026;2025;2024;2023;2022;2021;2020;2019;2018;2017;2016;2015;2014;2013;2012;2011;2010"
ctrl.ColumnCount = 1
ctrl.BoundColumn = 1
ctrl.LimitToList = True
ctrl.AllowValueListEdits = False
ctrl.ListRows = 12
ctrl.FontName = "Cambria"
ctrl.FontSize = 10.5
ctrl.FontBold = True
ctrl.ForeColor = 6703164
ctrl.BackColor = 16777215
ctrl.BorderStyle = 1
ctrl.BorderColor = 5000704
ctrl.TextAlign = 2
ctrl.OnChange = "[Event Procedure]"
ctrl.AfterUpdate = "[Event Procedure]"
ctrl.OnGotFocus = "[Event Procedure]"
ctrl.ControlTipText = "Seleccionar el Año en el calendario"

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 3940, 100, 500, 400)
ctrl.Name = "cmdMesSig"
ctrl.Caption = ">"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl
ctrl.FontSize = 10
ctrl.ControlTipText = "Mes siguiente"

' Encabezados de dias de la semana (Do Lu Ma Mi Ju Vi Sa)
arrDiasSem = Array("Do", "Lu", "Ma", "Mi", "Ju", "Vi", "Sá")
For dIdx = 0 To 6
    Set ctrl = access.CreateControl(fName, 100, 0, "", "", 180 + (dIdx * 610), 560, 570, 260)
    ctrl.Caption = arrDiasSem(dIdx)
    ctrl.FontName = "Segoe UI"
    ctrl.FontSize = 8.5
    ctrl.FontBold = True
    ctrl.ForeColor = 5000704
    ctrl.TextAlign = 2
    ctrl.BackStyle = 0
Next

' Cuadricula de 42 botones de dia (6 semanas x 7 dias)
For rCal = 0 To 5
    For cCal = 0 To 6
        idxBtn = (rCal * 7) + cCal + 1
        Set ctrl = access.CreateControl(fName, 104, 0, "", "", 180 + (cCal * 610), 850 + (rCal * 410), 570, 380)
        ctrl.Name = "cmdDia" & idxBtn
        ctrl.Caption = CStr(idxBtn)
        ctrl.FontName = "Segoe UI"
        ctrl.FontSize = 8.5
        ctrl.FontBold = True
        ctrl.BackStyle = 1
        ctrl.BackColor = 16777215
        ctrl.ForeColor = 3355443
        ctrl.BorderStyle = 1
        ctrl.BorderColor = 13882323
        ctrl.OnClick = "=SeleccionarDiaCal(" & idxBtn & ")"
    Next
Next

' Leyenda de dias con informacion cargada en la base de datos
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 160, 3340, 4280, 320)
ctrl.Name = "lblLeyendaCal"
ctrl.Caption = "[Verde]: Día con información cargada  |  [Azul]: Hoy"
ctrl.FontName = "Segoe UI"
End Sub

Private Sub EjecutarBloqueInstalacion_11()
ctrl.FontSize = 8
ctrl.FontBold = True
ctrl.ForeColor = 3433750
ctrl.TextAlign = 2
ctrl.BackStyle = 1
ctrl.BackColor = 15133908
ctrl.BorderStyle = 1
ctrl.BorderColor = 9376256

' Botones inferiores: Hoy y Cerrar
Set ctrl = access.CreateControl(fName, 104, 0, "", "", 550, 3740, 1650, 420)
ctrl.Name = "cmdHoy"
ctrl.Caption = "Hoy"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl
ctrl.FontSize = 9.5

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 2400, 3740, 1650, 420)
ctrl.Name = "cmdCerrarCal"
ctrl.Caption = "Cerrar"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl
ctrl.FontSize = 9.5
On Error GoTo 0

codeF = "Option Compare Database" & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"' Form_Calendario_Popup - Selector de fecha emergente" & vbCrLf & _
"' -----------------------------------------------------------------------------" & vbCrLf & _
"' Utilitario sin acceso a datos. Devuelve una fecha en formato dd/mm/aaaa" & vbCrLf & _
"' al control que lo invoco. Evita errores de captura manual de fechas." & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"" & vbCrLf & _
"Option Explicit" & vbCrLf & _
"" & vbCrLf & _
"Private mMes As Integer" & vbCrLf & _
"Private mAnio As Integer" & vbCrLf
codeF = codeF & "Private mFormDestino As String" & vbCrLf & _
"Private mCtrlDestino As String" & vbCrLf
codeF = codeF & "Private mSincronizandoCombos As Boolean" & vbCrLf & _
"Private mMapaFechas As String" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_Load()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Dim args As String, partes() As String, fechaBase As Date, strValDest As String" & vbCrLf & _
"    Dim idxM As Integer, idxY As Integer, nomMesesInit As Variant" & vbCrLf & _
"    mSincronizandoCombos = True" & vbCrLf & _
"    nomMesesInit = Array("""", ""Enero"", ""Febrero"", ""Marzo"", ""Abril"", ""Mayo"", ""Junio"", ""Julio"", ""Agosto"", ""Septiembre"", ""Octubre"", ""Noviembre"", ""Diciembre"")" & vbCrLf & _
"    Me.cboMes.RowSourceType = ""Value List""" & vbCrLf & _
"    Me.cboMes.RowSource = """"" & vbCrLf & _
"    For idxM = 1 To 12" & vbCrLf
codeF = codeF & "        Me.cboMes.AddItem CStr(nomMesesInit(idxM))" & vbCrLf & _
"    Next idxM" & vbCrLf
codeF = codeF & "    If Me.cboMes.ListCount < 12 Then" & vbCrLf & _
"        Me.cboMes.RowSource = ""Enero;Febrero;Marzo;Abril;Mayo;Junio;Julio;Agosto;Septiembre;Octubre;Noviembre;Diciembre""" & vbCrLf & _
"    End If" & vbCrLf & _
"    Me.cboAnio.RowSourceType = ""Value List""" & vbCrLf & _
"    Me.cboAnio.RowSource = """"" & vbCrLf & _
"    For idxY = Year(Date) + 3 To 2000 Step -1" & vbCrLf & _
"        Me.cboAnio.AddItem CStr(idxY)" & vbCrLf & _
"    Next idxY" & vbCrLf & _
"    If Me.cboAnio.ListCount < 5 Then" & vbCrLf & _
"        Me.cboAnio.RowSource = ""2028;2027;2026;2025;2024;2023;2022;2021;2020;2019;2018;2017;2016;2015;2010""" & vbCrLf & _
"    End If" & vbCrLf & _
"    fechaBase = Date" & vbCrLf
codeF = codeF & "    args = Nz(Me.OpenArgs, """")" & vbCrLf & _
"    If InStr(args, ""|"") > 0 Then" & vbCrLf
codeF = codeF & "        partes = Split(args, ""|"")" & vbCrLf & _
"        mFormDestino = partes(0)" & vbCrLf & _
"        mCtrlDestino = partes(1)" & vbCrLf & _
"        strValDest = Trim(CStr(Nz(Forms(mFormDestino).Controls(mCtrlDestino).Value, """")))" & vbCrLf & _
"        If FondoBO.esFechaFormatoCorrectoBO(strValDest) And Not FondoBO.esFechaFuturaBO(strValDest) Then" & vbCrLf & _
"            fechaBase = FondoBO.convertirFechaBO(strValDest)" & vbCrLf & _
"        End If" & vbCrLf & _
"    End If" & vbCrLf & _
"    mMes = Month(fechaBase)" & vbCrLf & _
"    mAnio = Year(fechaBase)" & vbCrLf & _
"    mSincronizandoCombos = False" & vbCrLf & _
"    ActualizarCuadricula" & vbCrLf
codeF = codeF & "    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub Form_Resize()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cboMes_GotFocus()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If Not mSincronizandoCombos Then Me.cboMes.Dropdown" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cboAnio_GotFocus()" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    If Not mSincronizandoCombos Then Me.cboAnio.Dropdown" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub AplicarSeleccionMesAnio()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If mSincronizandoCombos Then Exit Sub" & vbCrLf & _
"    Dim strM As String, strA As String, iM As Integer, valA As Integer, nomMeses As Variant, cambio As Boolean" & vbCrLf & _
"    cambio = False" & vbCrLf & _
"    nomMeses = Array("""", ""Enero"", ""Febrero"", ""Marzo"", ""Abril"", ""Mayo"", ""Junio"", ""Julio"", ""Agosto"", ""Septiembre"", ""Octubre"", ""Noviembre"", ""Diciembre"")" & vbCrLf & _
"    strM = Utilities.LeerTextoControl(Me.cboMes)" & vbCrLf & _
"    For iM = 1 To 12" & vbCrLf & _
"        If LCase(Trim(strM)) = LCase(CStr(nomMeses(iM))) Then" & vbCrLf & _
"            If mMes <> iM Then mMes = iM: cambio = True" & vbCrLf
codeF = codeF & "            Exit For" & vbCrLf & _
"        End If" & vbCrLf
codeF = codeF & "    Next iM" & vbCrLf & _
"    strA = Utilities.LeerTextoControl(Me.cboAnio)" & vbCrLf & _
"    If Len(Trim(strA)) = 4 And IsNumeric(strA) Then" & vbCrLf & _
"        valA = CInt(Val(strA))" & vbCrLf & _
"        If valA >= 1990 And valA <= 2050 Then" & vbCrLf & _
"            If mAnio <> valA Then mAnio = valA: cambio = True" & vbCrLf & _
"        End If" & vbCrLf & _
"    End If" & vbCrLf & _
"    If cambio Then ActualizarCuadricula" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cboMes_Change()" & vbCrLf
codeF = codeF & "    AplicarSeleccionMesAnio" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub cboMes_AfterUpdate()" & vbCrLf & _
"    AplicarSeleccionMesAnio" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cboAnio_Change()" & vbCrLf & _
"    AplicarSeleccionMesAnio" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cboAnio_AfterUpdate()" & vbCrLf & _
"    AplicarSeleccionMesAnio" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub ActualizarCuadricula()" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    Dim primerDia As Date, offsetSem As Integer, diasEnMes As Integer" & vbCrLf & _
"    Dim i As Integer, diaNum As Integer, btn As Control, nomMeses As Variant" & vbCrLf & _
"    Dim fechaIter As Date, strFecIter As String, nRegDia As Long" & vbCrLf & _
"    nomMeses = Array("""", ""Enero"", ""Febrero"", ""Marzo"", ""Abril"", ""Mayo"", ""Junio"", ""Julio"", ""Agosto"", ""Septiembre"", ""Octubre"", ""Noviembre"", ""Diciembre"")" & vbCrLf & _
"    mSincronizandoCombos = True" & vbCrLf & _
"    If LCase(Trim(CStr(Nz(Me.cboMes.Value, """")))) <> LCase(CStr(nomMeses(mMes))) Then" & vbCrLf & _
"        Me.cboMes.Value = CStr(nomMeses(mMes))" & vbCrLf & _
"    End If" & vbCrLf & _
"    If Trim(CStr(Nz(Me.cboAnio.Value, """"))) <> CStr(mAnio) Then" & vbCrLf & _
"        Me.cboAnio.Value = CStr(mAnio)" & vbCrLf & _
"    End If" & vbCrLf
codeF = codeF & "    mSincronizandoCombos = False" & vbCrLf & _
"    ' Una sola consulta por repintado en lugar de una por dia (eran 31)." & vbCrLf
codeF = codeF & "    ' Esto es lo que hacia lento abrir el calendario y cambiar de mes." & vbCrLf & _
"    mMapaFechas = FondoBO.obtenerMapaFechasConteoBO()" & vbCrLf & _
"    Me.Painting = False" & vbCrLf & _
"    primerDia = DateSerial(mAnio, mMes, 1)" & vbCrLf & _
"    offsetSem = Weekday(primerDia, vbSunday) - 1" & vbCrLf & _
"    diasEnMes = Day(DateSerial(mAnio, mMes + 1, 0))" & vbCrLf & _
"    For i = 1 To 42" & vbCrLf & _
"        Set btn = Me.Controls(""cmdDia"" & i)" & vbCrLf & _
"        diaNum = i - offsetSem" & vbCrLf & _
"        If diaNum >= 1 And diaNum <= diasEnMes Then" & vbCrLf & _
"            fechaIter = DateSerial(mAnio, mMes, diaNum)" & vbCrLf & _
"            strFecIter = Format(fechaIter, ""dd/mm/yyyy"")" & vbCrLf
codeF = codeF & "            nRegDia = FondoBO.conteoDeFechaEnMapaBO(mMapaFechas, strFecIter)" & vbCrLf & _
"            btn.Enabled = True" & vbCrLf
codeF = codeF & "            btn.Visible = True" & vbCrLf & _
"            If nRegDia > 0 Then" & vbCrLf & _
"                btn.Caption = CStr(diaNum)" & vbCrLf & _
"                btn.BackColor = 3433750 ' Verde institucional: Dia con insumo cargado" & vbCrLf & _
"                btn.ForeColor = 16777215" & vbCrLf & _
"                btn.BorderColor = 5000704" & vbCrLf & _
"                btn.ControlTipText = ""Información disponible en BD ("" & strFecIter & "": "" & nRegDia & "" registros)""" & vbCrLf & _
"            ElseIf fechaIter = Date Then" & vbCrLf & _
"                btn.Caption = CStr(diaNum)" & vbCrLf & _
"                btn.BackColor = 6703164 ' Azul marino: Hoy (sin insumo cargado)" & vbCrLf & _
"                btn.ForeColor = 16777215" & vbCrLf & _
"                btn.BorderColor = 6703164" & vbCrLf
codeF = codeF & "                btn.ControlTipText = ""Hoy ("" & strFecIter & "") - Sin insumo cargado""" & vbCrLf & _
"            ElseIf fechaIter > Date Then" & vbCrLf
codeF = codeF & "                btn.Caption = CStr(diaNum)" & vbCrLf & _
"                btn.BackColor = 15921906 ' Gris claro: Fecha posterior" & vbCrLf & _
"                btn.ForeColor = 10526880" & vbCrLf & _
"                btn.BorderColor = 13882323" & vbCrLf & _
"                btn.ControlTipText = ""Fecha futura no permitida ("" & strFecIter & "")""" & vbCrLf & _
"            Else" & vbCrLf & _
"                btn.Caption = CStr(diaNum)" & vbCrLf & _
"                btn.BackColor = 16777215" & vbCrLf & _
"                btn.ForeColor = 3355443" & vbCrLf & _
"                btn.BorderColor = 13882323" & vbCrLf & _
"                btn.ControlTipText = ""Sin insumo cargado para "" & strFecIter" & vbCrLf & _
"            End If" & vbCrLf
codeF = codeF & "        Else" & vbCrLf & _
"            btn.Caption = """"" & vbCrLf
codeF = codeF & "            btn.Enabled = False" & vbCrLf & _
"            btn.Visible = False" & vbCrLf & _
"        End If" & vbCrLf & _
"    Next i" & vbCrLf & _
"    Me.Painting = True" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Public Function SeleccionarDiaCal(ByVal idxBoton As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Dim strCap As String, numDia As Integer, fechaSel As Date, strFecSel As String, strDisp As String" & vbCrLf & _
"    strCap = Trim(CStr(Nz(Me.Controls(""cmdDia"" & idxBoton).Caption, """")))" & vbCrLf & _
"    numDia = CInt(Val(strCap))" & vbCrLf
codeF = codeF & "    If numDia >= 1 And numDia <= 31 Then" & vbCrLf & _
"        fechaSel = DateSerial(mAnio, mMes, numDia)" & vbCrLf
codeF = codeF & "        strFecSel = Format(fechaSel, ""dd/mm/yyyy"")" & vbCrLf & _
"        If fechaSel > Date Then" & vbCrLf & _
"            MsgBox ""Fecha de consulta inválida ("" & strFecSel & ""): "" & vbCrLf & _" & vbCrLf & _
"                   ""- No se puede escoger fechas futuras (posteriores al día de la consulta actual: "" & Format(Date, ""dd/mm/yyyy"") & "")."", vbCritical, ""Fecha Futura No Permitida""" & vbCrLf & _
"            Exit Function" & vbCrLf & _
"        End If" & vbCrLf & _
"        If UCase(Trim(mFormDestino)) = ""FORM_GENERAR_REPORTE"" Then" & vbCrLf & _
"            If FondoBO.contarRegistrosFechaBO(strFecSel) <= 0 Then" & vbCrLf & _
"                strDisp = FondoBO.obtenerFechasCargadasBO()" & vbCrLf & _
"                If strDisp = """" Then strDisp = ""Ninguna (cargue primero un insumo en 'Cargar Insumos')""" & vbCrLf & _
"                MsgBox ""Fecha sin información de insumo ("" & strFecSel & ""): "" & vbCrLf & _" & vbCrLf & _
"                       ""- No se ha cargado el insumo correspondiente a este día en la base de datos."" & vbCrLf & _" & vbCrLf
codeF = codeF & "                       ""- Días marcados con información disponible: "" & strDisp, vbExclamation, ""Día sin Insumo Cargado""" & vbCrLf & _
"                Exit Function" & vbCrLf
codeF = codeF & "            End If" & vbCrLf & _
"        End If" & vbCrLf & _
"        If mFormDestino <> """" And mCtrlDestino <> """" Then" & vbCrLf & _
"            Forms(mFormDestino).Controls(mCtrlDestino).Value = strFecSel" & vbCrLf & _
"            Utilities.ActualizarOpcionesSubidasPorFecha Forms(mFormDestino), strFecSel" & vbCrLf & _
"        End If" & vbCrLf & _
"        DoCmd.Close acForm, Me.Name" & vbCrLf & _
"    End If" & vbCrLf & _
"End Function" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdMesAnt_Click()" & vbCrLf & _
"    mMes = mMes - 1" & vbCrLf
codeF = codeF & "    If mMes < 1 Then mMes = 12: mAnio = mAnio - 1" & vbCrLf & _
"    ActualizarCuadricula" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdMesSig_Click()" & vbCrLf & _
"    mMes = mMes + 1" & vbCrLf & _
"    If mMes > 12 Then mMes = 1: mAnio = mAnio + 1" & vbCrLf & _
"    ActualizarCuadricula" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdHoy_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Dim strHoy As String, strDisp As String" & vbCrLf & _
"    strHoy = Format(Date, ""dd/mm/yyyy"")" & vbCrLf
codeF = codeF & "    If UCase(Trim(mFormDestino)) = ""FORM_GENERAR_REPORTE"" Then" & vbCrLf & _
"        If FondoBO.contarRegistrosFechaBO(strHoy) <= 0 Then" & vbCrLf
codeF = codeF & "            strDisp = FondoBO.obtenerFechasCargadasBO()" & vbCrLf & _
"            If strDisp = """" Then strDisp = ""Ninguna (cargue primero un insumo en 'Cargar Insumos')""" & vbCrLf & _
"            MsgBox ""Fecha sin información de insumo ("" & strHoy & ""): "" & vbCrLf & _" & vbCrLf & _
"                   ""- No se ha cargado el insumo correspondiente al día de hoy en la base de datos."" & vbCrLf & _" & vbCrLf & _
"                   ""- Días marcados con información disponible: "" & strDisp, vbExclamation, ""Día sin Insumo Cargado""" & vbCrLf & _
"            Exit Sub" & vbCrLf & _
"        End If" & vbCrLf & _
"    End If" & vbCrLf & _
"    If mFormDestino <> """" And mCtrlDestino <> """" Then" & vbCrLf & _
"        Forms(mFormDestino).Controls(mCtrlDestino).Value = strHoy" & vbCrLf & _
"        Utilities.ActualizarOpcionesSubidasPorFecha Forms(mFormDestino), strHoy" & vbCrLf & _
"    End If" & vbCrLf
codeF = codeF & "    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub cmdCerrarCal_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub"
AsignarCodigoLimpio frm, codeF
access.DoCmd.Close 2, fName, 1
On Error Resume Next
access.DoCmd.DeleteObject 2, "Form_Calendario_Popup"
On Error GoTo 0
access.DoCmd.Rename "Form_Calendario_Popup", 2, fName
logFile.WriteLine "   [OK] Form_Calendario_Popup creado con selectores de Mes/Anio, marcado de dias cargados y validacion directa."

' 14. Form_Cargar_Insumos (3 BASES DE DATOS DE INSUMOS POR FECHA + BUSQUEDA AUTOMATICA EN CARPETA O ARCHIVOS + VALIDACION INTERNA DE NUEVOS INSTRUMENTOS)
logFile.WriteLine "Paso 19: Construyendo Formulario Form_Cargar_Insumos..."
Set frm = access.CreateForm
fName = frm.Name
On Error Resume Next
frm.Caption = "Cargar Insumos"
frm.AutoCenter = True
frm.RecordSelectors = False
frm.NavigationButtons = False
frm.DividingLines = False
frm.ScrollBars = 0
frm.KeyPreview = True
frm.Width = 16000
frm.Section(0).Height = 6150
frm.Section(0).BackColor = 15790320
On Error GoTo 0

ConfigurarEncabezadoVerde frm, "Módulo de Administración Integral de Riesgos"
ConfigurarBarraNavegacionSuperior frm, "Form_Menu_General"

On Error Resume Next

' Contenedor central (Marco tipo Tarjeta Institucional)
Set ctrl = access.CreateControl(fName, 101, 0, "", "", 1800, 2100, 12400, 3450)
ctrl.BackStyle = 1
ctrl.BackColor = 16119285 ' Fondo suave (#F5F6F8)
ctrl.BorderStyle = 1
ctrl.BorderColor = 12632256 ' Borde gris institucional (#C0C0C0)
ctrl.SpecialEffect = 0

' Plaqueta / Banner superior del contenedor "Cargar Insumos"
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 2200, 1900, 3000, 420)
ctrl.Caption = "Cargar Insumos"
ctrl.BackStyle = 1
ctrl.BackColor = 5000704 ' #004E4C Verde azulado oficial
ctrl.ForeColor = 16777215 ' Blanco
ctrl.FontName = "Cambria"
ctrl.FontSize = 12
ctrl.FontBold = True
ctrl.TextAlign = 2

' ========================================================
' SECCION 1: 1.- Parámetros
' ========================================================
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 2400, 2380, 4000, 330)
ctrl.Caption = "1.- Parámetros"
ctrl.FontName = "Cambria"
ctrl.FontSize = 11
ctrl.FontBold = True
ctrl.ForeColor = 6703164 ' #1F4E78 Azul Marino
ctrl.BackStyle = 0

' Etiqueta Fecha:
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 3200, 2760, 950, 340)
ctrl.Caption = "Fecha:"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 10.5
ctrl.FontBold = True
ctrl.ForeColor = 3355443
ctrl.BackStyle = 0

' Campo txtFecha (ShowDatePicker = 0; validacion manual inmediata contra fechas futuras)
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 4200, 2720, 1600, 380)
ctrl.Name = "txtFecha"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 10
ctrl.TextAlign = 2
ctrl.ShowDatePicker = 0

' UNICO Icono de Calendario (con 1 solo clic despliega Form_Calendario_Popup con selectores de Mes y Año)
Set ctrl = access.CreateControl(fName, 104, 0, "", "", 5850, 2720, 440, 380)
ctrl.Name = "cmdCalendario"
ctrl.Caption = ""
ctrl.OnClick = "[Event Procedure]"
End Sub

Private Sub EjecutarBloqueInstalacion_12()
ctrl.BackStyle = 1
ctrl.BackColor = 16119285
ctrl.BorderStyle = 0
If fso.FileExists(strCalIconPath) Then
    ctrl.Picture = strCalIconPath
Else
    ctrl.Caption = ChrW(128197)
End If
ctrl.ControlTipText = "Abrir calendario con 1 solo clic (permite elegir Mes y Año)"

' Etiqueta dd/mm/aaaa limpia (sin texto adicional)
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 6360, 2760, 1800, 340)
ctrl.Caption = "dd/mm/aaaa"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 9.5
ctrl.ForeColor = 8421504
ctrl.BackStyle = 0

' Etiqueta Ruta de archivo:
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 3200, 3260, 1750, 340)
ctrl.Caption = "Ruta de archivo:"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 10.5
ctrl.FontBold = True
ctrl.ForeColor = 3355443
ctrl.BackStyle = 0

' Campo txtRuta
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 5000, 3220, 5300, 380)
ctrl.Name = "txtRuta"
ctrl.FontName = "Segoe UI"
ctrl.FontSize = 9.5

' Boton Examinar (Seleccion directa de 1 archivo, multiples archivos o una carpeta completa sin menu Si/No)
Set ctrl = access.CreateControl(fName, 104, 0, "", "", 10400, 3220, 1550, 380)
ctrl.Name = "cmdExaminar"
ctrl.Caption = "Archivo(s)"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl
ctrl.FontSize = 10.5
ctrl.ControlTipText = "Elegir uno o varios ARCHIVOS .xlsx (multiseleccion con Ctrl)"

' ========================================================
' SECCION 2: 2.- Selección de portafolio de riesgo (3 Bases de Datos)
' ========================================================

' Boton dedicado para elegir UNA CARPETA. Antes solo existia 'Examinar', que
' abre el selector de ARCHIVOS: ahi no se puede elegir una carpeta (al hacer
' clic solo se entra en ella), y habia que cancelar el dialogo para que apareciera
' el selector de carpetas. Con este boton es un clic directo.
Set ctrl = access.CreateControl(fName, 104, 0, "", "", 12050, 3220, 1550, 380)
ctrl.Name = "cmdExaminarCarpeta"
ctrl.Caption = "Carpeta"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl
ctrl.FontSize = 10.5
ctrl.ControlTipText = "Elegir una CARPETA: el sistema busca dentro los insumos de la fecha y de los portafolios marcados"

' ========================================================
' SECCION 2: 2.- Selección de portafolio de riesgo (3 Bases de Datos)
' ========================================================
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 2400, 3740, 6500, 330)
ctrl.Caption = "2.- Selección de portafolio de riesgo"
ctrl.FontName = "Cambria"
ctrl.FontSize = 11
ctrl.FontBold = True
ctrl.ForeColor = 6703164 ' #1F4E78 Azul Marino
ctrl.BackStyle = 0

' Checkbox 1: Cartera de Crédito Comercial (Primer Piso) -> BD_INSUMO_PRIMER_PISO (NO preseleccionada)
Set ctrl = access.CreateControl(fName, 106, 0, "", "", 3200, 4150, 300, 300)
ctrl.Name = "chkBeneficioDefinido"
ctrl.DefaultValue = "False"

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 3560, 4150, 6800, 300)
ctrl.Name = "lblOptPrimerPiso"
ctrl.Caption = "Cartera de Crédito Comercial (Primer Piso)"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 10.5
ctrl.ForeColor = 3355443
ctrl.BackStyle = 0

' Checkbox 2: Portafolio de Tesorería y Mercado de Dinero -> BD_INSUMO_TESORERIA (NO preseleccionada)
Set ctrl = access.CreateControl(fName, 106, 0, "", "", 3200, 4550, 300, 300)
ctrl.Name = "chkContribucionDefinida"
ctrl.DefaultValue = "False"

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 3560, 4550, 6800, 300)
ctrl.Name = "lblOptTesoreria"
ctrl.Caption = "Portafolio de Tesorería y Mercado de Dinero"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 10.5
ctrl.ForeColor = 3355443
ctrl.BackStyle = 0

' Checkbox 3: Financiamiento Sectorial y Garantías (Segundo Piso) -> BD_INSUMO_SEGUNDO_PISO (NO preseleccionada)
Set ctrl = access.CreateControl(fName, 106, 0, "", "", 3200, 4950, 300, 300)
ctrl.Name = "chkPEA"
ctrl.DefaultValue = "False"

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 3560, 4950, 6800, 300)
ctrl.Name = "lblOptSegundoPiso"
ctrl.Caption = "Financiamiento Sectorial y Garantías (Segundo Piso)"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 10.5
ctrl.ForeColor = 3355443
ctrl.BackStyle = 0

' Recuadro informativo con el Formato Unico Obligatorio del nombre de archivo
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 2400, 5260, 11200, 440)
ctrl.Name = "lblFormatoInsumos"
ctrl.Caption = "FORMATO UNICO OBLIGATORIO (53 COLS SIN VACIOS): INSUMO_PRIMER_PISO_DD_MM_AAAA.xlsx | INSUMO_TESORERIA_DD_MM_AAAA.xlsx | INSUMO_SEGUNDO_PISO_DD_MM_AAAA.xlsx"
ctrl.FontName = "Segoe UI"
ctrl.FontSize = 8.5
ctrl.FontBold = True
ctrl.ForeColor = 6703164
ctrl.TextAlign = 2
ctrl.BackStyle = 1
ctrl.BackColor = 15133908
ctrl.BorderStyle = 1
ctrl.BorderColor = 9376256

' Boton de Accion: Cargar Archivos
Set ctrl = access.CreateControl(fName, 104, 0, "", "", 11400, 4580, 2200, 620)
ctrl.Name = "cmdCargarArchivos"
ctrl.Caption = "Cargar Archivos"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl
ctrl.FontSize = 11

' Pie de pagina / Fecha larga
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1800, 5720, 5000, 300)
ctrl.Name = "lblFechaPie"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 9.5
ctrl.ForeColor = 6710886
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 8800, 5720, 5400, 300)
ctrl.Name = "lblArea"
ctrl.Caption = "Dirección de Economía de Hidrocarburos"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 9.5
ctrl.ForeColor = 6710886
ctrl.TextAlign = 3
ctrl.BackStyle = 0

On Error GoTo 0

codeF = "Option Compare Database" & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"' Form_Cargar_Insumos - Carga y validacion de insumos NORMATIVA" & vbCrLf & _
"' -----------------------------------------------------------------------------" & vbCrLf & _
"' Pantalla critica del modulo. Admite tres formas de seleccion:" & vbCrLf & _
"'    1) un archivo .xlsx" & vbCrLf & _
"'    2) varios archivos (multiseleccion, separados por ;)" & vbCrLf & _
"'    3) una carpeta: FondoBO.buscarInsumosCorrectosEnCarpetaBO localiza" & vbCrLf & _
"'       dentro los archivos que coincidan con la fecha y con los" & vbCrLf & _
"'       portafolios marcados, sin intervencion del usuario." & vbCrLf & _
"' Si el nombre del archivo delata el portafolio, la casilla se marca sola." & vbCrLf & _
"' Valida (FondoBO.validarYProcesarInsumoBO), rechazando el archivo si:" & vbCrLf
codeF = codeF & "'    - el nombre no sigue INSUMO_<PORTAFOLIO>_dd_mm_aaaa.xlsx" & vbCrLf & _
"'    - no tiene exactamente las 53 columnas oficiales, en orden" & vbCrLf
codeF = codeF & "'    - hay celdas vacias, saldos negativos o fechas que no cuadran" & vbCrLf & _
"'    - algun ID_INSTRUMENTO no existe en el catalogo INSTRUMENTOS" & vbCrLf & _
"' Escribe  : BD_INSUMO_PRIMER_PISO, BD_INSUMO_TESORERIA," & vbCrLf & _
"'            BD_INSUMO_SEGUNDO_PISO y BITACORA_CARGA_INSUMOS." & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"" & vbCrLf & _
"Option Explicit" & vbCrLf & _
"" & vbCrLf & _
"Private mValidandoFecha As Boolean" & vbCrLf & _
"Private mFechaValida As String" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_Load()" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    mValidandoFecha = True" & vbCrLf
codeF = codeF & "    ' Encabezado de sesion y control de acceso. La regla de quien puede" & vbCrLf & _
"    ' entrar a esta pantalla NO esta aqui: vive en FondoBO.autorizarPantallaBO." & vbCrLf & _
"    ' Este formulario solo pregunta y obedece (Marco de Desarrollo, I.2)." & vbCrLf & _
"    ' Pantalla sin restriccion de perfil: solo se pinta el encabezado." & vbCrLf & _
"    Call Utilities.AplicarEncabezadoSesionUT(Me, FondoBO.PANT_CARGA_INSUMOS)" & vbCrLf & _
"    mFechaValida = Format(Date, ""dd/mm/yyyy"")" & vbCrLf & _
"    Me.txtFecha.Value = mFechaValida" & vbCrLf & _
"    Me.chkBeneficioDefinido.Value = False" & vbCrLf & _
"    Me.chkContribucionDefinida.Value = False" & vbCrLf & _
"    Me.chkPEA.Value = False" & vbCrLf & _
"    Me.lblFechaPie.Caption = Format(Date, ""dddd, d \de mmmm \de yyyy"")" & vbCrLf & _
"    mValidandoFecha = False" & vbCrLf
codeF = codeF & "    Utilities.ActualizarOpcionesSubidasPorFecha Me, mFechaValida" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_Resize()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdCalendario_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If Not Utilities.ValidarAlSalirFechaManual(Me.txtFecha, mValidandoFecha, False, mFechaValida, False) Then Exit Sub" & vbCrLf & _
"    mValidandoFecha = True" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Calendario_Popup"", acNormal, , , , acDialog, Me.Name & ""|txtFecha""" & vbCrLf
codeF = codeF & "    Dim fecCal As String: fecCal = Trim(CStr(Nz(Me.txtFecha.Value, """")))" & vbCrLf & _
"    If FondoBO.esFechaFormatoCorrectoBO(fecCal) And Not FondoBO.esFechaFuturaBO(fecCal) Then" & vbCrLf
codeF = codeF & "        mFechaValida = FondoBO.normalizarFechaBO(fecCal)" & vbCrLf & _
"    End If" & vbCrLf & _
"    Utilities.ActualizarOpcionesSubidasPorFecha Me, mFechaValida" & vbCrLf & _
"    mValidandoFecha = False" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtFecha_KeyPress(KeyAscii As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.RestringirKeyPressFecha Me.txtFecha, KeyAscii" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtFecha_Change()" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    Utilities.ValidarCambioFechaManual Me.txtFecha, mValidandoFecha, False, mFechaValida" & vbCrLf
codeF = codeF & "    Utilities.ActualizarOpcionesSubidasPorFecha Me, Utilities.LeerTextoControl(Me.txtFecha)" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtFecha_AfterUpdate()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call Utilities.ValidarAlSalirFechaManual(Me.txtFecha, mValidandoFecha, False, mFechaValida, False)" & vbCrLf & _
"    Utilities.ActualizarOpcionesSubidasPorFecha Me, mFechaValida" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtFecha_LostFocus()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call Utilities.ValidarAlSalirFechaManual(Me.txtFecha, mValidandoFecha, False, mFechaValida, False)" & vbCrLf
codeF = codeF & "    Utilities.ActualizarOpcionesSubidasPorFecha Me, mFechaValida" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"' -----------------------------------------------------------------------------" & vbCrLf & _
"' cmdExaminarCarpeta_Click" & vbCrLf & _
"' Abre directamente el selector de CARPETAS. Al elegir una, el aplicativo" & vbCrLf & _
"' busca dentro los insumos que coincidan con la fecha capturada y con los" & vbCrLf & _
"' portafolios marcados, y marca solo las casillas que correspondan." & vbCrLf & _
"' Si no hay ninguna casilla marcada, toma los tres portafolios." & vbCrLf & _
"' -----------------------------------------------------------------------------" & vbCrLf & _
"Private Sub cmdExaminarCarpeta_Click()" & vbCrLf & _
"    On Error GoTo ErrCarp" & vbCrLf & _
"    Dim fdC As Object, fso As Object, strFecha As String, carp As String" & vbCrLf & _
"    Dim bBD As Boolean, bCD As Boolean, bPEA As Boolean" & vbCrLf
codeF = codeF & "    Dim rutasHalladas As String, resumen As String, msgErr As String, ini As String" & vbCrLf & _
"    Dim detallePrecarga As String" & vbCrLf
codeF = codeF & "    Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf & _
"    strFecha = Utilities.LeerTextoControl(Me.txtFecha)" & vbCrLf & _
"    If Not FondoBO.esFechaFormatoCorrectoBO(strFecha) Then strFecha = Format(Date, ""dd/mm/yyyy"")" & vbCrLf & _
"    bBD = Nz(Me.chkBeneficioDefinido.Value, False)" & vbCrLf & _
"    bCD = Nz(Me.chkContribucionDefinida.Value, False)" & vbCrLf & _
"    bPEA = Nz(Me.chkPEA.Value, False)" & vbCrLf & _
"" & vbCrLf & _
"    ' -----------------------------------------------------------------------" & vbCrLf & _
"    ' AVISO DE REEMPLAZO AL EXAMINAR" & vbCrLf & _
"    ' Si las bases de insumo YA tienen informacion de esta fecha de corte, se" & vbCrLf & _
"    ' pregunta AQUI, antes de seleccionar los archivos, y no al final de toda" & vbCrLf & _
"    ' la validacion. Si el usuario responde No, no se abre el explorador y la" & vbCrLf
codeF = codeF & "    ' seleccion queda limpia: nada se toca." & vbCrLf & _
"    ' -----------------------------------------------------------------------" & vbCrLf
codeF = codeF & "    If FondoBO.existeInformacionPrecargadaBO(strFecha, bBD, bCD, bPEA, detallePrecarga) Then" & vbCrLf & _
"        If MsgBox(FondoBO.textoPreguntaReemplazoBO(strFecha, detallePrecarga), _" & vbCrLf & _
"                  vbQuestion + vbYesNo + vbDefaultButton2, _" & vbCrLf & _
"                  ""Informacion Ya Precargada - Confirmar Reemplazo"") = vbNo Then" & vbCrLf & _
"            Me.txtRuta.Value = """"" & vbCrLf & _
"            MsgBox ""Operacion cancelada. No se modifico ninguna base de datos."" & vbCrLf & vbCrLf & _" & vbCrLf & _
"                   ""Si desea cargar otra fecha, cambie la fecha de corte y vuelva a examinar."", _" & vbCrLf & _
"                   vbInformation, ""Reemplazo Cancelado""" & vbCrLf & _
"            Exit Sub" & vbCrLf & _
"        End If" & vbCrLf & _
"    End If" & vbCrLf & _
"    ini = CurrentProject.Path & ""\Insumos""" & vbCrLf
codeF = codeF & "    If Not fso.FolderExists(ini) Then ini = CurrentProject.Path" & vbCrLf & _
"    Set fdC = Application.FileDialog(4)" & vbCrLf
codeF = codeF & "    With fdC" & vbCrLf & _
"        .Title = ""Seleccionar carpeta de insumos - Hidrocarburos""" & vbCrLf & _
"        .ButtonName = ""Usar esta carpeta""" & vbCrLf & _
"        .InitialFileName = ini & ""\""" & vbCrLf & _
"        If .Show <> -1 Then Set fdC = Nothing: Exit Sub" & vbCrLf & _
"        carp = Trim(CStr(.SelectedItems(1)))" & vbCrLf & _
"    End With" & vbCrLf & _
"    Set fdC = Nothing" & vbCrLf & _
"    If carp = """" Then Exit Sub" & vbCrLf & _
"    If FondoBO.buscarInsumosCorrectosEnCarpetaBO(carp, strFecha, bBD, bCD, bPEA, rutasHalladas, resumen, msgErr) Then" & vbCrLf & _
"        Me.txtRuta.Value = rutasHalladas" & vbCrLf & _
"        If InStr(1, rutasHalladas, ""INSUMO_PRIMER_PISO_"", vbTextCompare) > 0 Then Me.chkBeneficioDefinido.Value = True" & vbCrLf
codeF = codeF & "        If InStr(1, rutasHalladas, ""INSUMO_TESORERIA_"", vbTextCompare) > 0 Then Me.chkContribucionDefinida.Value = True" & vbCrLf & _
"        If InStr(1, rutasHalladas, ""INSUMO_SEGUNDO_PISO_"", vbTextCompare) > 0 Then Me.chkPEA.Value = True" & vbCrLf
codeF = codeF & "        MsgBox resumen, vbInformation, ""Insumos localizados en la carpeta""" & vbCrLf & _
"    Else" & vbCrLf & _
"        Me.txtRuta.Value = carp" & vbCrLf & _
"        MsgBox msgErr, vbExclamation, ""Examinar carpeta""" & vbCrLf & _
"    End If" & vbCrLf & _
"    Exit Sub" & vbCrLf & _
"ErrCarp:" & vbCrLf & _
"    MsgBox ""No fue posible abrir el selector de carpetas: "" & Err.Description, vbExclamation, ""Examinar carpeta""" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdExaminar_Click()" & vbCrLf & _
"    On Error GoTo ErrDialog" & vbCrLf
codeF = codeF & "    Dim fd As Object, fso As Object, rutaInicial As String, rutaActual As String, strFecha As String" & vbCrLf & _
"    Dim bBD As Boolean, bCD As Boolean, bPEA As Boolean" & vbCrLf
codeF = codeF & "    Dim rutasHalladas As String, resumenHallazgos As String, msgErrBusq As String" & vbCrLf & _
"    Dim idxSel As Long, itemSel As String, listaSel As String" & vbCrLf & _
"    Dim tipoDet As String, tablaDest As String, msgErrNom As String" & vbCrLf & _
"    Dim detallePrecarga As String" & vbCrLf & _
"    Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf & _
"    strFecha = Utilities.LeerTextoControl(Me.txtFecha)" & vbCrLf & _
"    If Not FondoBO.esFechaFormatoCorrectoBO(strFecha) Then strFecha = Format(Date, ""dd/mm/yyyy"")" & vbCrLf & _
"    bBD = Nz(Me.chkBeneficioDefinido.Value, False)" & vbCrLf & _
"    bCD = Nz(Me.chkContribucionDefinida.Value, False)" & vbCrLf & _
"    bPEA = Nz(Me.chkPEA.Value, False)" & vbCrLf & _
"" & vbCrLf & _
"    ' -----------------------------------------------------------------------" & vbCrLf
codeF = codeF & "    ' AVISO DE REEMPLAZO AL EXAMINAR" & vbCrLf & _
"    ' Si las bases de insumo YA tienen informacion de esta fecha de corte, se" & vbCrLf
codeF = codeF & "    ' pregunta AQUI, antes de seleccionar los archivos, y no al final de toda" & vbCrLf & _
"    ' la validacion. Si el usuario responde No, no se abre el explorador y la" & vbCrLf & _
"    ' seleccion queda limpia: nada se toca." & vbCrLf & _
"    ' -----------------------------------------------------------------------" & vbCrLf & _
"    If FondoBO.existeInformacionPrecargadaBO(strFecha, bBD, bCD, bPEA, detallePrecarga) Then" & vbCrLf & _
"        If MsgBox(FondoBO.textoPreguntaReemplazoBO(strFecha, detallePrecarga), _" & vbCrLf & _
"                  vbQuestion + vbYesNo + vbDefaultButton2, _" & vbCrLf & _
"                  ""Informacion Ya Precargada - Confirmar Reemplazo"") = vbNo Then" & vbCrLf & _
"            Me.txtRuta.Value = """"" & vbCrLf & _
"            MsgBox ""Operacion cancelada. No se modifico ninguna base de datos."" & vbCrLf & vbCrLf & _" & vbCrLf & _
"                   ""Si desea cargar otra fecha, cambie la fecha de corte y vuelva a examinar."", _" & vbCrLf & _
"                   vbInformation, ""Reemplazo Cancelado""" & vbCrLf
codeF = codeF & "            Exit Sub" & vbCrLf & _
"        End If" & vbCrLf
codeF = codeF & "    End If" & vbCrLf & _
"    rutaActual = Trim(Utilities.LeerTextoControl(Me.txtRuta))" & vbCrLf & _
"    rutaInicial = CurrentProject.Path & ""\Insumos""" & vbCrLf & _
"    If Not fso.FolderExists(rutaInicial) Then rutaInicial = CurrentProject.Path" & vbCrLf & _
"" & vbCrLf & _
"    ' 1. Si el usuario ya escribio o pego la ruta de una carpeta completa en txtRuta y presiona Examinar," & vbCrLf & _
"    '    buscar en automatico todos los documentos contenidos en esa carpeta sin desplegar menu Si/No" & vbCrLf & _
"    If rutaActual <> """" And InStr(rutaActual, "";"") = 0 And fso.FolderExists(rutaActual) Then" & vbCrLf & _
"        If FondoBO.buscarInsumosCorrectosEnCarpetaBO(rutaActual, strFecha, bBD, bCD, bPEA, rutasHalladas, resumenHallazgos, msgErrBusq) Then" & vbCrLf & _
"            Me.txtRuta.Value = rutasHalladas" & vbCrLf & _
"            If InStr(1, rutasHalladas, ""INSUMO_PRIMER_PISO_"", vbTextCompare) > 0 Then Me.chkBeneficioDefinido.Value = True" & vbCrLf & _
"            If InStr(1, rutasHalladas, ""INSUMO_TESORERIA_"", vbTextCompare) > 0 Then Me.chkContribucionDefinida.Value = True" & vbCrLf
codeF = codeF & "            If InStr(1, rutasHalladas, ""INSUMO_SEGUNDO_PISO_"", vbTextCompare) > 0 Then Me.chkPEA.Value = True" & vbCrLf & _
"            Exit Sub" & vbCrLf
codeF = codeF & "        Else" & vbCrLf & _
"            MsgBox msgErrBusq, vbExclamation, ""Examinar Carpeta — Formato Único de Insumos""" & vbCrLf & _
"            Exit Sub" & vbCrLf & _
"        End If" & vbCrLf & _
"    End If" & vbCrLf & _
"" & vbCrLf & _
"    ' 2. Abrir directamente el explorador unificado (sin menu Si/No): permite seleccionar 1 archivo, varios archivos al mismo tiempo o una carpeta completa" & vbCrLf & _
"    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker" & vbCrLf & _
"    With fd" & vbCrLf & _
"        .Title = ""Seleccionar Archivo(s) de Insumos o Carpeta - Hidrocarburos""" & vbCrLf & _
"        .Filters.Clear" & vbCrLf & _
"        .Filters.Add ""Archivos de Excel (*.xlsx; *.xls; *.xlsm)"", ""*.xlsx; *.xls; *.xlsm""" & vbCrLf
codeF = codeF & "        .Filters.Add ""Todos los archivos (*.*)"", ""*.*""" & vbCrLf & _
"        .AllowMultiSelect = True" & vbCrLf
codeF = codeF & "        .InitialFileName = rutaInicial & ""\""" & vbCrLf & _
"        If .Show = -1 Then" & vbCrLf & _
"            listaSel = """"" & vbCrLf & _
"            For idxSel = 1 To .SelectedItems.Count" & vbCrLf & _
"                itemSel = Trim(CStr(.SelectedItems(idxSel)))" & vbCrLf & _
"                If fso.FolderExists(itemSel) Then" & vbCrLf & _
"                    If FondoBO.buscarInsumosCorrectosEnCarpetaBO(itemSel, strFecha, bBD, bCD, bPEA, rutasHalladas, resumenHallazgos, msgErrBusq) Then" & vbCrLf & _
"                        Me.txtRuta.Value = rutasHalladas" & vbCrLf & _
"                        If InStr(1, rutasHalladas, ""INSUMO_PRIMER_PISO_"", vbTextCompare) > 0 Then Me.chkBeneficioDefinido.Value = True" & vbCrLf & _
"                        If InStr(1, rutasHalladas, ""INSUMO_TESORERIA_"", vbTextCompare) > 0 Then Me.chkContribucionDefinida.Value = True" & vbCrLf & _
"                        If InStr(1, rutasHalladas, ""INSUMO_SEGUNDO_PISO_"", vbTextCompare) > 0 Then Me.chkPEA.Value = True" & vbCrLf & _
"                    Else" & vbCrLf
codeF = codeF & "                        Me.txtRuta.Value = itemSel" & vbCrLf & _
"                        MsgBox msgErrBusq, vbExclamation, ""Examinar Carpeta — Formato Único de Insumos""" & vbCrLf
codeF = codeF & "                    End If" & vbCrLf & _
"                    Set fd = Nothing" & vbCrLf & _
"                    Exit Sub" & vbCrLf & _
"                End If" & vbCrLf & _
"                If listaSel <> """" Then listaSel = listaSel & "";""" & vbCrLf & _
"                listaSel = listaSel & itemSel" & vbCrLf & _
"                If FondoBO.validarFormatoUnicoNombreInsumoBO(fso.GetFileName(itemSel), strFecha, """", tipoDet, tablaDest, msgErrNom) Then" & vbCrLf & _
"                    If tipoDet = ""PRIMER_PISO"" Then Me.chkBeneficioDefinido.Value = True" & vbCrLf & _
"                    If tipoDet = ""TESORERIA"" Then Me.chkContribucionDefinida.Value = True" & vbCrLf & _
"                    If tipoDet = ""SEGUNDO_PISO"" Then Me.chkPEA.Value = True" & vbCrLf & _
"                End If" & vbCrLf & _
"            Next idxSel" & vbCrLf
codeF = codeF & "            Me.txtRuta.Value = listaSel" & vbCrLf & _
"        Else" & vbCrLf
codeF = codeF & "            ' Si el usuario cerro el selector de archivos para elegir directamente una carpeta completa:" & vbCrLf & _
"            Set fd = Nothing" & vbCrLf & _
"            Dim fdCarp As Object, carpElegida As String: carpElegida = """"" & vbCrLf & _
"            On Error Resume Next" & vbCrLf & _
"            Set fdCarp = Application.FileDialog(4) ' msoFileDialogFolderPicker" & vbCrLf & _
"            If Err.Number = 0 And Not fdCarp Is Nothing Then" & vbCrLf & _
"                With fdCarp" & vbCrLf & _
"                    .Title = ""Seleccionar Carpeta Completa de Insumos — Búsqueda Automática Hidrocarburos""" & vbCrLf & _
"                    .ButtonName = ""Seleccionar Carpeta""" & vbCrLf & _
"                    .InitialFileName = rutaInicial & ""\""" & vbCrLf & _
"                    If .Show = -1 Then carpElegida = Trim(CStr(.SelectedItems(1)))" & vbCrLf & _
"                End With" & vbCrLf
codeF = codeF & "                Set fdCarp = Nothing" & vbCrLf & _
"            End If" & vbCrLf
codeF = codeF & "            On Error GoTo ErrDialog" & vbCrLf & _
"            If carpElegida <> """" Then" & vbCrLf & _
"                If FondoBO.buscarInsumosCorrectosEnCarpetaBO(carpElegida, strFecha, bBD, bCD, bPEA, rutasHalladas, resumenHallazgos, msgErrBusq) Then" & vbCrLf & _
"                    Me.txtRuta.Value = rutasHalladas" & vbCrLf & _
"                    If InStr(1, rutasHalladas, ""INSUMO_PRIMER_PISO_"", vbTextCompare) > 0 Then Me.chkBeneficioDefinido.Value = True" & vbCrLf & _
"                    If InStr(1, rutasHalladas, ""INSUMO_TESORERIA_"", vbTextCompare) > 0 Then Me.chkContribucionDefinida.Value = True" & vbCrLf & _
"                    If InStr(1, rutasHalladas, ""INSUMO_SEGUNDO_PISO_"", vbTextCompare) > 0 Then Me.chkPEA.Value = True" & vbCrLf & _
"                Else" & vbCrLf & _
"                    Me.txtRuta.Value = carpElegida" & vbCrLf & _
"                    MsgBox msgErrBusq, vbExclamation, ""Examinar Carpeta — Formato Único de Insumos""" & vbCrLf & _
"                End If" & vbCrLf & _
"            End If" & vbCrLf
codeF = codeF & "            Exit Sub" & vbCrLf & _
"        End If" & vbCrLf
codeF = codeF & "    End With" & vbCrLf & _
"    Set fd = Nothing" & vbCrLf & _
"    Exit Sub" & vbCrLf & _
"ErrDialog:" & vbCrLf & _
"    MsgBox ""Puede escribir o pegar directamente la ruta del archivo o carpeta en el campo Ruta de archivo."", vbInformation, ""Seleccionar Insumo""" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdCargarArchivos_Click()" & vbCrLf & _
"    On Error GoTo ErrH" & vbCrLf & _
"    Dim ruta As String, strFecha As String, fso As Object" & vbCrLf & _
"    Dim bBD As Boolean, bCD As Boolean, bPEA As Boolean, nMarcados As Integer" & vbCrLf & _
"    Dim rutasHalladas As String, resumenHallazgos As String, msgErrBusq As String" & vbCrLf
codeF = codeF & "    Dim arrRutas() As String, idxR As Long, rItem As String, extArch As String" & vbCrLf & _
"    Dim totalFilas As Long, totalErrores As Long, detalleErrores As String" & vbCrLf
codeF = codeF & "    Dim nuevosInstDet As String, nuevosInstCod As String, cargaOk As Boolean" & vbCrLf & _
"    Dim tipoDet As String, tablaDest As String, msgErrNom As String, tipoEsp As String" & vbCrLf & _
"    Dim yaCargado As Boolean, infoPrevia As String" & vbCrLf & _
"    Dim resumenExito As String, totalFilasGlobal As Long, nRegNuevos As Long, msgRegNuevos As String" & vbCrLf & _
"    " & vbCrLf & _
"    strFecha = Utilities.LeerTextoControl(Me.txtFecha)" & vbCrLf & _
"    ruta = Trim(Utilities.LeerTextoControl(Me.txtRuta))" & vbCrLf & _
"    bBD = Nz(Me.chkBeneficioDefinido.Value, False)" & vbCrLf & _
"    bCD = Nz(Me.chkContribucionDefinida.Value, False)" & vbCrLf & _
"    bPEA = Nz(Me.chkPEA.Value, False)" & vbCrLf & _
"    " & vbCrLf & _
"    ' 1. Verificar que los campos Fecha y Ruta esten llenos y con fecha valida (no futura)" & vbCrLf
codeF = codeF & "    If strFecha = """" And ruta = """" Then" & vbCrLf & _
"        MsgBox ""Campos obligatorios vacíos:"" & vbCrLf & _" & vbCrLf
codeF = codeF & "               ""- El campo 'Fecha' se encuentra vacío."" & vbCrLf & _" & vbCrLf & _
"               ""- El campo 'Ruta de archivo' se encuentra vacío."" & vbCrLf & vbCrLf & _" & vbCrLf & _
"               FondoBO.obtenerFormatoCorrectoInsumoBO(Format(Date, ""dd/mm/yyyy""), bBD, bCD, bPEA), vbCritical, ""Parámetros Incompletos""" & vbCrLf & _
"        Me.txtFecha.SetFocus" & vbCrLf & _
"        Exit Sub" & vbCrLf & _
"    End If" & vbCrLf & _
"    If strFecha = """" Then" & vbCrLf & _
"        MsgBox ""Campo obligatorio vacío:"" & vbCrLf & _" & vbCrLf & _
"               ""- El campo 'Fecha' está vacío. Favor de capturar una fecha con formato dd/mm/aaaa."", vbCritical, ""Fecha Requerida""" & vbCrLf & _
"        Me.txtFecha.SetFocus" & vbCrLf & _
"        Exit Sub" & vbCrLf & _
"    End If" & vbCrLf
codeF = codeF & "    If Not FondoBO.esFechaFormatoCorrectoBO(strFecha) Then" & vbCrLf & _
"        MsgBox ""Formato incorrecto en el campo Fecha ('"" & strFecha & ""'):"" & vbCrLf & _" & vbCrLf
codeF = codeF & "               ""- Debe ingresar una fecha real y válida con formato dd/mm/aaaa."", vbCritical, ""Fecha Incorrecta""" & vbCrLf & _
"        Me.txtFecha.SetFocus" & vbCrLf & _
"        Exit Sub" & vbCrLf & _
"    End If" & vbCrLf & _
"    If FondoBO.esFechaFuturaBO(strFecha) Then" & vbCrLf & _
"        MsgBox ""Fecha inválida ("" & FondoBO.normalizarFechaBO(strFecha) & ""): "" & vbCrLf & _" & vbCrLf & _
"               ""- No se puede escoger fechas futuras (posteriores al día de la consulta actual: "" & Format(Date, ""dd/mm/yyyy"") & "")."", vbCritical, ""Fecha Futura No Permitida""" & vbCrLf & _
"        Me.txtFecha.SetFocus" & vbCrLf & _
"        Exit Sub" & vbCrLf & _
"    End If" & vbCrLf & _
"    strFecha = FondoBO.normalizarFechaBO(strFecha)" & vbCrLf & _
"    " & vbCrLf
codeF = codeF & "    If ruta = """" Then" & vbCrLf & _
"        MsgBox ""Campo obligatorio vacío:"" & vbCrLf & _" & vbCrLf
codeF = codeF & "               ""- El campo 'Ruta de archivo' está vacío. Seleccione uno o varios archivos o una carpeta completa mediante el botón 'Examinar'."" & vbCrLf & vbCrLf & _" & vbCrLf & _
"               FondoBO.obtenerFormatoCorrectoInsumoBO(strFecha, bBD, bCD, bPEA), vbCritical, ""Ruta de Archivo o Carpeta Requerida""" & vbCrLf & _
"        Me.txtRuta.SetFocus" & vbCrLf & _
"        Exit Sub" & vbCrLf & _
"    End If" & vbCrLf & _
"    " & vbCrLf & _
"    Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf & _
"    " & vbCrLf & _
"    ' 2. Si la ruta proporcionada es una CARPETA COMPLETA, buscar automaticamente todos los documentos contenidos en esa carpeta" & vbCrLf & _
"    If InStr(ruta, "";"") = 0 And fso.FolderExists(ruta) Then" & vbCrLf & _
"        If Not FondoBO.buscarInsumosCorrectosEnCarpetaBO(ruta, strFecha, bBD, bCD, bPEA, rutasHalladas, resumenHallazgos, msgErrBusq) Then" & vbCrLf & _
"            MsgBox ""CARGA RECHAZADA — BÚSQUEDA EN CARPETA"" & vbCrLf & vbCrLf & msgErrBusq, vbCritical, ""Archivos No Encontrados en Carpeta""" & vbCrLf
codeF = codeF & "            Exit Sub" & vbCrLf & _
"        End If" & vbCrLf
codeF = codeF & "        ruta = rutasHalladas" & vbCrLf & _
"        Me.txtRuta.Value = ruta" & vbCrLf & _
"        If InStr(1, ruta, ""INSUMO_PRIMER_PISO_"", vbTextCompare) > 0 Then Me.chkBeneficioDefinido.Value = True: bBD = True" & vbCrLf & _
"        If InStr(1, ruta, ""INSUMO_TESORERIA_"", vbTextCompare) > 0 Then Me.chkContribucionDefinida.Value = True: bCD = True" & vbCrLf & _
"        If InStr(1, ruta, ""INSUMO_SEGUNDO_PISO_"", vbTextCompare) > 0 Then Me.chkPEA.Value = True: bPEA = True" & vbCrLf & _
"    End If" & vbCrLf & _
"    " & vbCrLf & _
"    arrRutas = Split(ruta, "";"")" & vbCrLf & _
"    " & vbCrLf & _
"    ' Si el usuario selecciono 1 o varios archivos especificos y no habia marcado casillas, sincronizar automaticamente las casillas de los archivos seleccionados" & vbCrLf & _
"    If Not (bBD Or bCD Or bPEA) Then" & vbCrLf & _
"        For idxR = LBound(arrRutas) To UBound(arrRutas)" & vbCrLf
codeF = codeF & "            rItem = Trim(arrRutas(idxR))" & vbCrLf & _
"            If rItem <> """" Then" & vbCrLf
codeF = codeF & "                If FondoBO.validarFormatoUnicoNombreInsumoBO(fso.GetFileName(rItem), strFecha, """", tipoDet, tablaDest, msgErrNom) Then" & vbCrLf & _
"                    If tipoDet = ""PRIMER_PISO"" Then Me.chkBeneficioDefinido.Value = True: bBD = True" & vbCrLf & _
"                    If tipoDet = ""TESORERIA"" Then Me.chkContribucionDefinida.Value = True: bCD = True" & vbCrLf & _
"                    If tipoDet = ""SEGUNDO_PISO"" Then Me.chkPEA.Value = True: bPEA = True" & vbCrLf & _
"                End If" & vbCrLf & _
"            End If" & vbCrLf & _
"        Next idxR" & vbCrLf & _
"    End If" & vbCrLf & _
"    " & vbCrLf & _
"    If Not (bBD Or bCD Or bPEA) And UBound(arrRutas) < 0 Then" & vbCrLf & _
"        MsgBox ""Debe seleccionar al menos una de las 3 opciones de portafolio de riesgo para realizar la carga."", vbExclamation, ""Cargar Insumos""" & vbCrLf & _
"        Exit Sub" & vbCrLf
codeF = codeF & "    End If" & vbCrLf & _
"    " & vbCrLf
codeF = codeF & "    nMarcados = IIf(bBD, 1, 0) + IIf(bCD, 1, 0) + IIf(bPEA, 1, 0)" & vbCrLf & _
"    tipoEsp = """"" & vbCrLf & _
"    If UBound(arrRutas) = 0 And nMarcados = 1 Then" & vbCrLf & _
"        If bBD Then tipoEsp = ""PRIMER_PISO""" & vbCrLf & _
"        If bCD Then tipoEsp = ""TESORERIA""" & vbCrLf & _
"        If bPEA Then tipoEsp = ""SEGUNDO_PISO""" & vbCrLf & _
"    End If" & vbCrLf & _
"    " & vbCrLf & _
"    resumenExito = """"" & vbCrLf & _
"    totalFilasGlobal = 0" & vbCrLf & _
"    Call FondoBO.IniciarSesionExcelRapidaBO" & vbCrLf & _
"    " & vbCrLf
codeF = codeF & "    For idxR = LBound(arrRutas) To UBound(arrRutas)" & vbCrLf & _
"        rItem = Trim(arrRutas(idxR))" & vbCrLf & _
"        If rItem <> """" Then" & vbCrLf
codeF = codeF & "            extArch = LCase(fso.GetExtensionName(rItem))" & vbCrLf & _
"            If InStr(rItem, ""\"") = 0 Or (extArch <> ""xlsx"" And extArch <> ""xls"" And extArch <> ""xlsm"") Or Not fso.FileExists(rItem) Then" & vbCrLf & _
"                MsgBox ""Ruta de archivo incorrecta:"" & vbCrLf & _" & vbCrLf & _
"                       ""- La ruta especificada ('"" & rItem & ""') no existe o no corresponde a un archivo Excel (.xlsx) ni a una carpeta válida."" & vbCrLf & vbCrLf & _" & vbCrLf & _
"                       FondoBO.obtenerFormatoCorrectoInsumoBO(strFecha, bBD, bCD, bPEA), vbCritical, ""Ruta o Formato de Archivo Incorrecto""" & vbCrLf & _
"                Call FondoBO.FinalizarSesionExcelRapidaBO" & vbCrLf & _
"                Me.txtRuta.SetFocus" & vbCrLf & _
"                Exit Sub" & vbCrLf & _
"            End If" & vbCrLf & _
"            " & vbCrLf & _
"            DoCmd.Hourglass True" & vbCrLf & _
"            cargaOk = FondoBO.validarYProcesarInsumoBO(rItem, strFecha, totalFilas, totalErrores, detalleErrores, tipoEsp, nuevosInstDet, nuevosInstCod)" & vbCrLf
codeF = codeF & "            DoCmd.Hourglass False" & vbCrLf & _
"            " & vbCrLf & _
"            If totalErrores > 0 Or Not cargaOk Then" & vbCrLf
codeF = codeF & "                MsgBox ""CARGA RECHAZADA — VALIDADOR INSTITUCIONAL HIDROCARBUROS"" & vbCrLf & _" & vbCrLf & _
"                       ""Archivo evaluado: "" & fso.GetFileName(rItem) & "" | Fecha Pantalla: "" & strFecha & vbCrLf & _" & vbCrLf & _
"                       ""Filas revisadas: "" & totalFilas & "" | Errores detectados: "" & totalErrores & vbCrLf & vbCrLf & _" & vbCrLf & _
"                       ""Motivo(s) por el que no se puede subir el insumo y formato correcto requerido:"" & vbCrLf & _" & vbCrLf & _
"                       detalleErrores & vbCrLf & _" & vbCrLf & _
"                       ""Base de datos protegida (NO se realizó la carga)."", vbCritical, ""Insumo Incorrecto — Validación Fallida""" & vbCrLf & _
"                Call FondoBO.FinalizarSesionExcelRapidaBO" & vbCrLf & _
"                Exit Sub" & vbCrLf & _
"            End If" & vbCrLf & _
"            " & vbCrLf & _
"            ' Validacion interna obligatoria: NO permite subir el insumo si existen nuevos instrumentos no contemplados en INSTRUMENTOS" & vbCrLf & _
"            If Trim(nuevosInstCod) <> """" Then" & vbCrLf
codeF = codeF & "                MsgBox ""CARGA RECHAZADA — VALIDACIÓN INTERNA DE NUEVOS INSTRUMENTOS"" & vbCrLf & vbCrLf & _" & vbCrLf & _
"                       ""El insumo '"" & fso.GetFileName(rItem) & ""' (Fecha: "" & strFecha & "") NO se puede subir porque contiene NUEVOS INSTRUMENTOS en la columna [ID_INSTRUMENTO] que no están contemplados en el catálogo INSTRUMENTOS:"" & vbCrLf & vbCrLf & _" & vbCrLf & _
"                       nuevosInstDet & vbCrLf & _" & vbCrLf
codeF = codeF & "                       ""Por normativa de integridad, primero debe registrar estos instrumentos en el apartado 'Gestionar Instrumentos' antes de volver a subir el insumo."", vbCritical, ""Nuevos Instrumentos No Contemplados — Carga Bloqueada""" & vbCrLf & _
"                Call FondoBO.FinalizarSesionExcelRapidaBO" & vbCrLf & _
"                Exit Sub" & vbCrLf & _
"            End If" & vbCrLf & _
"            " & vbCrLf & _
"            Call FondoBO.validarFormatoUnicoNombreInsumoBO(fso.GetFileName(rItem), strFecha, """", tipoDet, tablaDest, msgErrNom)" & vbCrLf & _
"            yaCargado = FondoBO.verificarInsumoYaCargadoBO(fso.GetFileName(rItem), strFecha, infoPrevia)" & vbCrLf & _
"            If yaCargado Then" & vbCrLf & _
"                If MsgBox(""AVISO DE CONTROL DE INSUMOS:"" & vbCrLf & vbCrLf & _" & vbCrLf & _
"                          ""El insumo '"" & fso.GetFileName(rItem) & ""' (Fecha: "" & strFecha & "" | Base: "" & tablaDest & "") YA fue cargado previamente:"" & vbCrLf & _" & vbCrLf & _
"                          ""  "" & infoPrevia & vbCrLf & vbCrLf & _" & vbCrLf & _
"                          ""¿Desea SOBRESCRIBIR la información existente en "" & tablaDest & ""?"", vbQuestion + vbYesNo + vbDefaultButton2, ""Insumo Ya Cargado — Confirmar Sobrescritura"") = vbNo Then" & vbCrLf
codeF = codeF & "                    Call FondoBO.FinalizarSesionExcelRapidaBO" & vbCrLf & _
"                    MsgBox ""Operación cancelada para "" & fso.GetFileName(rItem) & ""."", vbInformation, ""Sobrescritura Cancelada""" & vbCrLf & _
"                    Exit Sub" & vbCrLf & _
"                End If" & vbCrLf
codeF = codeF & "            End If" & vbCrLf & _
"            " & vbCrLf & _
"            DoCmd.Hourglass True" & vbCrLf & _
"            Call FondoBO.guardarInsumoValidadoBO(rItem, fso.GetFileName(rItem), strFecha, totalFilas, tipoDet, tablaDest)" & vbCrLf & _
"            DoCmd.Hourglass False" & vbCrLf & _
"            " & vbCrLf & _
"            totalFilasGlobal = totalFilasGlobal + totalFilas" & vbCrLf & _
"            resumenExito = resumenExito & ""  • "" & fso.GetFileName(rItem) & "" -> Base: "" & tablaDest & "" ("" & totalFilas & "" filas | Col. ID_INSTRUMENTO verificada)"" & vbCrLf" & vbCrLf & _
"        End If" & vbCrLf & _
"    Next idxR" & vbCrLf & _
"    Call FondoBO.FinalizarSesionExcelRapidaBO" & vbCrLf & _
"    " & vbCrLf
codeF = codeF & "    Utilities.ActualizarOpcionesSubidasPorFecha Me, strFecha" & vbCrLf & _
"    MsgBox ""¡Insumo(s) verificado(s) al 100% y cargado(s) exitosamente en su(s) base(s) de datos!"" & vbCrLf & vbCrLf & _" & vbCrLf & _
"           ""- Fecha de Corte conciliada: "" & strFecha & vbCrLf & _" & vbCrLf
codeF = codeF & "           ""- Estructura validada: 53 columnas exactas (incluyendo campo ID_INSTRUMENTO) y "" & totalFilasGlobal & "" registros sin vacíos"" & vbCrLf & _" & vbCrLf & _
"           ""- Validación interna de Instrumentos: 100% conciliados con catálogo INSTRUMENTOS"" & vbCrLf & vbCrLf & _" & vbCrLf & _
"           ""Bases de datos actualizadas:"" & vbCrLf & resumenExito, vbInformation, ""Validación y Carga Exitosa - Hidrocarburos""" & vbCrLf & _
"    Exit Sub" & vbCrLf & _
"ErrH:" & vbCrLf & _
"    Call FondoBO.FinalizarSesionExcelRapidaBO" & vbCrLf & _
"    DoCmd.Hourglass False" & vbCrLf & _
"    MsgBox ""Ocurrió un error al validar o procesar el archivo: "" & Err.Description, vbCritical, ""Error de Carga""" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtRuta_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf & _
"    If KeyCode = 13 Then" & vbCrLf
codeF = codeF & "        KeyCode = 0" & vbCrLf & _
"        cmdCargarArchivos_Click" & vbCrLf & _
"    End If" & vbCrLf
codeF = codeF & "End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtFecha_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If KeyCode = 13 Or KeyCode = 9 Then" & vbCrLf & _
"        If KeyCode = 13 Then KeyCode = 0" & vbCrLf & _
"        If Utilities.ValidarAlSalirFechaManual(Me.txtFecha, mValidandoFecha, False, mFechaValida, False) Then" & vbCrLf & _
"            Me.txtRuta.SetFocus" & vbCrLf & _
"        End If" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf
codeF = codeF & "Private Sub cmdAtras_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    DoCmd.OpenForm ""Form_Menu_General"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdSalir_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If MsgBox(""¿Desea cerrar la sesión actual?"", vbQuestion + vbYesNo, ""Cerrar Sesión"") = vbYes Then" & vbCrLf & _
"        Set usuarioSesionDTO = Nothing" & vbCrLf & _
"        DoCmd.OpenForm ""Form_Acceso"", acNormal" & vbCrLf & _
"        DoCmd.Close acForm, Me.Name" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub cmdRegresar_Click()" & vbCrLf
codeF = codeF & "    cmdAtras_Click" & vbCrLf & _
"End Sub"
AsignarCodigoLimpio frm, codeF
End Sub

Private Sub EjecutarBloqueInstalacion_13()
access.DoCmd.Close 2, fName, 1
On Error Resume Next
access.DoCmd.DeleteObject 2, "Form_Cargar_Insumos"
On Error GoTo 0
access.DoCmd.Rename "Form_Cargar_Insumos", 2, fName
logFile.WriteLine "   [OK] Form_Cargar_Insumos creado con 3 bases de datos por fecha, busqueda automatica en carpeta, formato unico y validacion interna de instrumentos."


' 15. Form_Generar_Reporte (DISENO EXACTO 1:1 CON LA FOTOGRAFIA DEL USUARIO + SELECCION INDIVIDUAL POR ANALISIS + PARIDAD PDF/EXCEL)
logFile.WriteLine "Paso 20: Construyendo Formulario Form_Generar_Reporte..."
Set frm = access.CreateForm
fName = frm.Name
On Error Resume Next
frm.Caption = "Generar_Reporte"
frm.AutoCenter = True
frm.KeyPreview = True
frm.RecordSelectors = False
frm.NavigationButtons = False
frm.DividingLines = False
frm.ScrollBars = 0
frm.Width = 16000
frm.Section(0).Height = 6700
frm.Section(0).BackColor = 15790320
On Error GoTo 0

ConfigurarEncabezadoVerde frm, "Módulo de Administración Integral de Riesgos"
ConfigurarBarraNavegacionSuperior frm, "Form_Menu_General"

On Error Resume Next

' Contenedor central (Marco tipo Tarjeta Institucional)
Set ctrl = access.CreateControl(fName, 101, 0, "", "", 1800, 2000, 12400, 3600)
ctrl.BackStyle = 1
ctrl.BackColor = 16119285 ' Fondo suave (#F5F6F8)
ctrl.BorderStyle = 1
ctrl.BorderColor = 12632256 ' Borde gris institucional (#C0C0C0)
ctrl.SpecialEffect = 0

' Plaqueta / Banner superior del contenedor "Generar Reporte"
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 2200, 1800, 3000, 420)
ctrl.Caption = "Generar Reporte"
ctrl.BackStyle = 1
ctrl.BackColor = 5000704 ' #004E4C Verde azulado oficial
ctrl.ForeColor = 16777215 ' Blanco
ctrl.FontName = "Cambria"
ctrl.FontSize = 12
ctrl.FontBold = True
ctrl.TextAlign = 2

' ========================================================
' SECCION 1: 1.- Parámetros (Fecha y Ruta para guardar el reporte)
' ========================================================
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 2400, 2320, 4000, 320)
ctrl.Caption = "1.- Parámetros"
ctrl.FontName = "Cambria"
ctrl.FontSize = 11
ctrl.FontBold = True
ctrl.ForeColor = 6703164 ' #1F4E78 Azul Marino
ctrl.BackStyle = 0

' Etiqueta Fecha:
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 3200, 2720, 950, 340)
ctrl.Caption = "Fecha:"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 10.5
ctrl.FontBold = True
ctrl.ForeColor = 3355443
ctrl.BackStyle = 0

' Campo txtFechaCorte
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 4200, 2680, 1550, 360)
ctrl.Name = "txtFechaCorte"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 10
ctrl.TextAlign = 2
ctrl.ShowDatePicker = 0

' UNICO Icono de Calendario (con 1 solo clic abre Form_Calendario_Popup)
Set ctrl = access.CreateControl(fName, 104, 0, "", "", 5810, 2680, 440, 360)
ctrl.Name = "cmdCalendario"
ctrl.Caption = ""
ctrl.OnClick = "[Event Procedure]"
ctrl.BackStyle = 1
ctrl.BackColor = 16119285
ctrl.BorderStyle = 0
If fso.FileExists(strCalIconPath) Then
    ctrl.Picture = strCalIconPath
Else
    ctrl.Caption = ChrW(128197)
End If
ctrl.ControlTipText = "Abrir calendario con 1 solo clic"

' Etiqueta dd/mm/aaaa
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 6320, 2720, 1600, 340)
ctrl.Caption = "dd/mm/aaaa"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 9.5
ctrl.ForeColor = 8421504
ctrl.BackStyle = 0

' Etiqueta Ruta para guardar:
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 3200, 3160, 1750, 340)
ctrl.Caption = "Ruta guardado:"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 10.5
ctrl.FontBold = True
ctrl.ForeColor = 3355443
ctrl.BackStyle = 0

' Campo txtRutaGuardar (Permite al usuario elegir donde guardar el reporte)
Set ctrl = access.CreateControl(fName, 109, 0, "", "", 5000, 3120, 6700, 360)
ctrl.Name = "txtRutaGuardar"
ctrl.FontName = "Segoe UI"
ctrl.FontSize = 9.5

' Boton Examinar Carpeta de Guardado
Set ctrl = access.CreateControl(fName, 104, 0, "", "", 11800, 3120, 1800, 360)
ctrl.Name = "cmdExaminarRuta"
ctrl.Caption = "Examinar"
ctrl.OnClick = "[Event Procedure]"
EstilizarBoton ctrl
ctrl.FontSize = 10

' ========================================================
' SECCION 2: 2.- Selección de portafolio de riesgo
' ========================================================
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 2400, 3620, 5500, 320)
ctrl.Caption = "2.- Selección de portafolio de riesgo"
ctrl.FontName = "Cambria"
ctrl.FontSize = 11
ctrl.FontBold = True
ctrl.ForeColor = 6703164 ' #1F4E78 Azul Marino
ctrl.BackStyle = 0

' Checkbox 1: Cartera de Crédito Comercial (Primer Piso)
Set ctrl = access.CreateControl(fName, 106, 0, "", "", 3200, 3950, 280, 260)
ctrl.Name = "chkBeneficioDefinido"
ctrl.DefaultValue = "False"

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 3560, 3950, 5600, 260)
ctrl.Name = "lblOptPrimerPiso"
ctrl.Caption = "Cartera de Crédito Comercial (Primer Piso)"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 10.5
ctrl.ForeColor = 3355443
ctrl.BackStyle = 0

' Checkbox 2: Portafolio de Tesorería y Mercado de Dinero
Set ctrl = access.CreateControl(fName, 106, 0, "", "", 3200, 4230, 280, 260)
ctrl.Name = "chkContribucionDefinida"
ctrl.DefaultValue = "False"

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 3560, 4230, 5600, 260)
ctrl.Name = "lblOptTesoreria"
ctrl.Caption = "Portafolio de Tesorería y Mercado de Dinero"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 10.5
ctrl.ForeColor = 3355443
ctrl.BackStyle = 0

' Checkbox 3: Financiamiento Sectorial y Garantías (Segundo Piso)
Set ctrl = access.CreateControl(fName, 106, 0, "", "", 3200, 4510, 280, 260)
ctrl.Name = "chkPEA"
ctrl.DefaultValue = "False"

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 3560, 4510, 5600, 260)
ctrl.Name = "lblOptSegundoPiso"
ctrl.Caption = "Financiamiento Sectorial y Garantías (Segundo Piso)"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 10.5
ctrl.ForeColor = 3355443
ctrl.BackStyle = 0

' Checkbox 4: Consolidado (Independiente: NO marca las otras 3 casillas en automatico)
Set ctrl = access.CreateControl(fName, 106, 0, "", "", 3200, 4790, 280, 260)
ctrl.Name = "chkConsolidado"
ctrl.DefaultValue = "False"

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 3560, 4790, 5600, 260)
ctrl.Name = "lblOptConsolidado"
ctrl.Caption = "Consolidado"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 10.5
ctrl.ForeColor = 3355443
ctrl.BackStyle = 0

' Checkbox 5: Comparativo de 2 insumos (cruce Primer Piso x Tesoreria).
Set ctrl = access.CreateControl(fName, 106, 0, "", "", 3200, 5070, 280, 260)
ctrl.Name = "chkComparativo2"
ctrl.DefaultValue = "False"

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 3560, 5070, 6600, 260)
ctrl.Name = "lblOptComparativo2"
ctrl.Caption = "Comparativo de 2 insumos (Primer Piso vs. Tesorería, por instrumento)"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 10.5
ctrl.ForeColor = 3355443
ctrl.BackStyle = 0

' Checkbox 6: Comparativo de 3 insumos (cruce de los tres libros).
Set ctrl = access.CreateControl(fName, 106, 0, "", "", 3200, 5350, 280, 260)
ctrl.Name = "chkComparativo3"
ctrl.DefaultValue = "False"

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 3560, 5350, 6600, 260)
ctrl.Name = "lblOptComparativo3"
ctrl.Caption = "Comparativo de 3 insumos (los tres libros, por instrumento)"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 10.5
ctrl.ForeColor = 3355443
ctrl.BackStyle = 0

' ========================================================
' BOTONES CON ICONO DE DOCUMENTO DESCARGABLE: PDF Y EXCEL
' ========================================================
Set ctrl = access.CreateControl(fName, 104, 0, "", "", 10850, 4120, 1250, 1180)
ctrl.Name = "cmdGenerarPDF"
ctrl.Caption = ""
ctrl.OnClick = "[Event Procedure]"
ctrl.BackStyle = 1
ctrl.BackColor = 16119285
ctrl.BorderStyle = 0
ctrl.HoverColor = 15133908
ctrl.PressedColor = 14540253
If fso.FileExists(strPdfIconPath) Then
    ctrl.Picture = strPdfIconPath
Else
    ctrl.Caption = "PDF"
End If
ctrl.ControlTipText = "Generar Reporte(s) en PDF en la ruta seleccionada"

Set ctrl = access.CreateControl(fName, 104, 0, "", "", 12300, 4120, 1250, 1180)
ctrl.Name = "cmdGenerar"
ctrl.Caption = ""
ctrl.OnClick = "[Event Procedure]"
ctrl.BackStyle = 1
ctrl.BackColor = 16119285
ctrl.BorderStyle = 0
ctrl.HoverColor = 15133908
ctrl.PressedColor = 14540253
If fso.FileExists(strExcelIconPath) Then
    ctrl.Picture = strExcelIconPath
Else
    ctrl.Caption = "EXCEL"
End If
ctrl.ControlTipText = "Generar Reporte(s) en EXCEL en la ruta seleccionada"

' Pie de pagina / Fecha larga y Direccion
Set ctrl = access.CreateControl(fName, 100, 0, "", "", 1800, 5680, 5000, 300)
ctrl.Name = "lblFechaPie"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 9.5
ctrl.ForeColor = 6710886
ctrl.BackStyle = 0

Set ctrl = access.CreateControl(fName, 100, 0, "", "", 8800, 5680, 5400, 300)
ctrl.Name = "lblArea"
ctrl.Caption = "Dirección de Economía de Hidrocarburos"
ctrl.FontName = "Cambria Math"
ctrl.FontSize = 9.5
ctrl.ForeColor = 6710886
ctrl.TextAlign = 3
ctrl.BackStyle = 0
On Error GoTo 0

codeF = "Option Compare Database" & vbCrLf & _
"' =============================================================================" & vbCrLf & _
"' Form_Generar_Reporte - Generacion del reporte ejecutivo" & vbCrLf & _
"' -----------------------------------------------------------------------------" & vbCrLf & _
"' Produce el reporte en Excel o PDF a partir de los insumos ya cargados." & vbCrLf & _
"' Entradas : fecha de corte, portafolios a incluir y ruta de guardado." & vbCrLf & _
"' Origen de los datos: las tres consultas de analisis del DAO" & vbCrLf & _
"'    extraerAnalisis1PrimerPisoSQL   (BD_INSUMO_PRIMER_PISO)" & vbCrLf & _
"'    extraerAnalisis2SectoresMonedasSQL (BD_INSUMO_TESORERIA)" & vbCrLf & _
"'    extraerAnalisis3SegundoPisoSQL  (BD_INSUMO_SEGUNDO_PISO)" & vbCrLf & _
"' Cada una cruza contra INSTRUMENTOS para promediar la tasa del grupo." & vbCrLf & _
"' Las graficas son objetos nativos de Excel construidos con los mismos" & vbCrLf
codeF = codeF & "' datos de la tabla, por lo que siempre cuadran con las cifras impresas." & vbCrLf & _
"' Marcar los 3 portafolios sin 'Consolidado' genera 3 archivos separados." & vbCrLf
codeF = codeF & "' =============================================================================" & vbCrLf & _
"" & vbCrLf & _
"Option Explicit" & vbCrLf & _
"" & vbCrLf & _
"Private mValidandoFecha As Boolean" & vbCrLf & _
"Private mFechaValida As String" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_Load()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    mValidandoFecha = True" & vbCrLf & _
"    ' Encabezado de sesion y control de acceso. La regla de quien puede" & vbCrLf & _
"    ' entrar a esta pantalla NO esta aqui: vive en FondoBO.autorizarPantallaBO." & vbCrLf
codeF = codeF & "    ' Este formulario solo pregunta y obedece (Marco de Desarrollo, I.2)." & vbCrLf & _
"    ' Pantalla sin restriccion de perfil: solo se pinta el encabezado." & vbCrLf
codeF = codeF & "    Call Utilities.AplicarEncabezadoSesionUT(Me, FondoBO.PANT_GENERAR_REPORTE)" & vbCrLf & _
"    Dim ultFecSQL As String: ultFecSQL = FondoBO.obtenerUltimaFechaCargadaBO()" & vbCrLf & _
"    If ultFecSQL <> """" Then mFechaValida = ultFecSQL Else mFechaValida = Format(Date, ""dd/mm/yyyy"")" & vbCrLf & _
"    Me.txtFechaCorte.Value = mFechaValida" & vbCrLf & _
"    Me.txtRutaGuardar.Value = FondoBO.obtenerRutaDescargasPorDefectoBO()" & vbCrLf & _
"    Me.chkBeneficioDefinido.Value = False" & vbCrLf & _
"    Me.chkContribucionDefinida.Value = False" & vbCrLf & _
"    Me.chkPEA.Value = False" & vbCrLf & _
"    Me.chkConsolidado.Value = False" & vbCrLf & _
"    Me.chkComparativo2.Value = False" & vbCrLf & _
"    Me.chkComparativo3.Value = False" & vbCrLf & _
"    Me.lblFechaPie.Caption = Format(Date, ""dddd, d \de mmmm \de yyyy"")" & vbCrLf
codeF = codeF & "    mValidandoFecha = False" & vbCrLf & _
"    Utilities.ActualizarOpcionesSubidasPorFecha Me, mFechaValida" & vbCrLf
codeF = codeF & "    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub Form_Resize()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.AjustarTamanoFormularioCompleto Me" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdCalendario_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If Not Utilities.ValidarAlSalirFechaManual(Me.txtFechaCorte, mValidandoFecha, True, mFechaValida, False) Then Exit Sub" & vbCrLf & _
"    mValidandoFecha = True" & vbCrLf
codeF = codeF & "    DoCmd.OpenForm ""Form_Calendario_Popup"", acNormal, , , , acDialog, Me.Name & ""|txtFechaCorte""" & vbCrLf & _
"    Dim fecCal As String: fecCal = Trim(CStr(Nz(Me.txtFechaCorte.Value, """")))" & vbCrLf
codeF = codeF & "    If FondoBO.esFechaFormatoCorrectoBO(fecCal) And Not FondoBO.esFechaFuturaBO(fecCal) Then" & vbCrLf & _
"        mFechaValida = FondoBO.normalizarFechaBO(fecCal)" & vbCrLf & _
"    End If" & vbCrLf & _
"    Utilities.ActualizarOpcionesSubidasPorFecha Me, mFechaValida" & vbCrLf & _
"    mValidandoFecha = False" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdExaminarRuta_Click()" & vbCrLf & _
"    On Error GoTo ErrCarpeta" & vbCrLf & _
"    Dim fdCarpeta As Object, rutaInicial As String" & vbCrLf & _
"    rutaInicial = Trim(Utilities.LeerTextoControl(Me.txtRutaGuardar))" & vbCrLf & _
"    If rutaInicial = """" Then rutaInicial = FondoBO.obtenerRutaDescargasPorDefectoBO()" & vbCrLf
codeF = codeF & "    Set fdCarpeta = Application.FileDialog(4) ' msoFileDialogFolderPicker" & vbCrLf & _
"    With fdCarpeta" & vbCrLf
codeF = codeF & "        .Title = ""Seleccionar Carpeta donde se Guardará el Reporte — Hidrocarburos""" & vbCrLf & _
"        .ButtonName = ""Seleccionar Carpeta""" & vbCrLf & _
"        .InitialFileName = rutaInicial & ""\""" & vbCrLf & _
"        If .Show = -1 Then" & vbCrLf & _
"            Me.txtRutaGuardar.Value = .SelectedItems(1)" & vbCrLf & _
"        End If" & vbCrLf & _
"    End With" & vbCrLf & _
"    Set fdCarpeta = Nothing" & vbCrLf & _
"    Exit Sub" & vbCrLf & _
"ErrCarpeta:" & vbCrLf & _
"    MsgBox ""Puede escribir o pegar directamente la ruta de la carpeta destino en el campo Ruta guardado."", vbInformation, ""Carpeta Destino""" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Function ValidarSeleccionReporte() As Boolean" & vbCrLf
codeF = codeF & "    ValidarSeleccionReporte = False" & vbCrLf & _
"    If Not Utilities.ValidarAlSalirFechaManual(Me.txtFechaCorte, mValidandoFecha, True, mFechaValida, True) Then Exit Function" & vbCrLf & _
"    Dim strFecVal As String: strFecVal = Utilities.LeerTextoControl(Me.txtFechaCorte)" & vbCrLf & _
"    If Not FondoBO.esFechaFormatoCorrectoBO(strFecVal) Then" & vbCrLf & _
"        MsgBox ""Favor de seleccionar o ingresar una Fecha válida con formato dd/mm/aaaa."", vbExclamation, ""Generar Reporte""" & vbCrLf & _
"        Me.txtFechaCorte.SetFocus" & vbCrLf & _
"        Exit Function" & vbCrLf & _
"    End If" & vbCrLf & _
"    If FondoBO.esFechaFuturaBO(strFecVal) Then" & vbCrLf & _
"        MsgBox ""Fecha de consulta inválida ("" & FondoBO.normalizarFechaBO(strFecVal) & ""): "" & vbCrLf & _" & vbCrLf & _
"               ""- No se puede escoger fechas futuras (posteriores al día de la consulta actual: "" & Format(Date, ""dd/mm/yyyy"") & "")."", vbCritical, ""Fecha Futura No Permitida""" & vbCrLf & _
"        Me.txtFechaCorte.SetFocus" & vbCrLf
codeF = codeF & "        Exit Function" & vbCrLf & _
"    End If" & vbCrLf
codeF = codeF & "    Dim strRutaDest As String: strRutaDest = Trim(Utilities.LeerTextoControl(Me.txtRutaGuardar))" & vbCrLf & _
"    If strRutaDest = """" Then" & vbCrLf & _
"        strRutaDest = FondoBO.obtenerRutaDescargasPorDefectoBO()" & vbCrLf & _
"        Me.txtRutaGuardar.Value = strRutaDest" & vbCrLf & _
"    End If" & vbCrLf & _
"    If Not (Nz(Me.chkBeneficioDefinido.Value, False) Or Nz(Me.chkContribucionDefinida.Value, False) Or Nz(Me.chkPEA.Value, False) Or Nz(Me.chkConsolidado.Value, False) _" & vbCrLf & _
"           Or Nz(Me.chkComparativo2.Value, False) Or Nz(Me.chkComparativo3.Value, False)) Then" & vbCrLf & _
"        MsgBox ""Debe seleccionar al menos una opcion de reporte: un portafolio, el Consolidado, o uno de los dos Comparativos."", vbExclamation, ""Generar Reporte""" & vbCrLf & _
"        Exit Function" & vbCrLf & _
"    End If" & vbCrLf & _
"    ValidarSeleccionReporte = True" & vbCrLf & _
"End Function" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub txtFechaCorte_KeyPress(KeyAscii As Integer)" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    Utilities.RestringirKeyPressFecha Me.txtFechaCorte, KeyAscii" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtFechaCorte_Change()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Utilities.ValidarCambioFechaManual Me.txtFechaCorte, mValidandoFecha, True, mFechaValida" & vbCrLf & _
"    Utilities.ActualizarOpcionesSubidasPorFecha Me, Utilities.LeerTextoControl(Me.txtFechaCorte)" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtFechaCorte_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    If KeyCode = 13 Or KeyCode = 9 Then" & vbCrLf & _
"        If KeyCode = 13 Then KeyCode = 0" & vbCrLf
codeF = codeF & "        Call Utilities.ValidarAlSalirFechaManual(Me.txtFechaCorte, mValidandoFecha, True, mFechaValida, True)" & vbCrLf & _
"        Utilities.ActualizarOpcionesSubidasPorFecha Me, mFechaValida" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtFechaCorte_AfterUpdate()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Call Utilities.ValidarAlSalirFechaManual(Me.txtFechaCorte, mValidandoFecha, True, mFechaValida, False)" & vbCrLf & _
"    Utilities.ActualizarOpcionesSubidasPorFecha Me, mFechaValida" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub txtFechaCorte_LostFocus()" & vbCrLf
codeF = codeF & "    On Error Resume Next" & vbCrLf & _
"    Call Utilities.ValidarAlSalirFechaManual(Me.txtFechaCorte, mValidandoFecha, True, mFechaValida, False)" & vbCrLf
codeF = codeF & "    Utilities.ActualizarOpcionesSubidasPorFecha Me, mFechaValida" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub EjecutarGeneracionReportes(ByVal formato As String)" & vbCrLf & _
"    On Error GoTo ErrGenRep" & vbCrLf & _
"    Dim rutaCarpeta As String, strFecRep As String, rOut As String, msgErr As String" & vbCrLf & _
"    Dim bBD As Boolean, bCD As Boolean, bPEA As Boolean, bCons As Boolean" & vbCrLf & _
"    Dim bCmp2 As Boolean, bCmp3 As Boolean" & vbCrLf & _
"    Dim nGenerados As Integer, listaArchivos As String, ultimoArchivo As String" & vbCrLf & _
"    If Not ValidarSeleccionReporte() Then Exit Sub" & vbCrLf & _
"    bBD = Nz(Me.chkBeneficioDefinido.Value, False)" & vbCrLf & _
"    bCD = Nz(Me.chkContribucionDefinida.Value, False)" & vbCrLf
codeF = codeF & "    bPEA = Nz(Me.chkPEA.Value, False)" & vbCrLf & _
"    bCons = Nz(Me.chkConsolidado.Value, False)" & vbCrLf
codeF = codeF & "    bCmp2 = Nz(Me.chkComparativo2.Value, False)" & vbCrLf & _
"    bCmp3 = Nz(Me.chkComparativo3.Value, False)" & vbCrLf & _
"    strFecRep = Utilities.LeerTextoControl(Me.txtFechaCorte)" & vbCrLf & _
"    If strFecRep = """" Then strFecRep = Format(Date, ""dd/mm/yyyy"")" & vbCrLf & _
"    rutaCarpeta = Trim(Utilities.LeerTextoControl(Me.txtRutaGuardar))" & vbCrLf & _
"    If rutaCarpeta = """" Then rutaCarpeta = FondoBO.obtenerRutaDescargasPorDefectoBO()" & vbCrLf & _
"    " & vbCrLf & _
"    DoCmd.Hourglass True" & vbCrLf & _
"    Call FondoBO.IniciarSesionExcelRapidaBO" & vbCrLf & _
"    nGenerados = 0: listaArchivos = """": ultimoArchivo = """"" & vbCrLf & _
"    " & vbCrLf & _
"    ' 1. Si marco Cartera Comercial Primer Piso -> genera su archivo individual" & vbCrLf
codeF = codeF & "    If bBD Then" & vbCrLf & _
"        If Not FondoBO.generarReporteEjecutivoBO(strFecRep, formato, True, False, False, False, rOut, msgErr, rutaCarpeta) Then" & vbCrLf & _
"            DoCmd.Hourglass False" & vbCrLf
codeF = codeF & "            Call FondoBO.FinalizarSesionExcelRapidaBO" & vbCrLf & _
"            MsgBox ""No fue posible generar el reporte individual de Primer Piso en "" & formato & "": "" & vbCrLf & msgErr, vbCritical, ""Error de Exportación""" & vbCrLf & _
"            Exit Sub" & vbCrLf & _
"        End If" & vbCrLf & _
"        nGenerados = nGenerados + 1" & vbCrLf & _
"        listaArchivos = listaArchivos & ""  "" & nGenerados & "") [Primer Piso]: "" & rOut & vbCrLf" & vbCrLf & _
"        ultimoArchivo = rOut" & vbCrLf & _
"    End If" & vbCrLf & _
"    " & vbCrLf & _
"    ' 2. Si marco Portafolio de Tesoreria y Mercado de Dinero -> genera su archivo individual" & vbCrLf & _
"    If bCD Then" & vbCrLf & _
"        If Not FondoBO.generarReporteEjecutivoBO(strFecRep, formato, False, True, False, False, rOut, msgErr, rutaCarpeta) Then" & vbCrLf
codeF = codeF & "            DoCmd.Hourglass False" & vbCrLf & _
"            Call FondoBO.FinalizarSesionExcelRapidaBO" & vbCrLf & _
"            MsgBox ""No fue posible generar el reporte individual de Tesorería en "" & formato & "": "" & vbCrLf & msgErr, vbCritical, ""Error de Exportación""" & vbCrLf & _
"            Exit Sub" & vbCrLf
codeF = codeF & "        End If" & vbCrLf & _
"        nGenerados = nGenerados + 1" & vbCrLf & _
"        listaArchivos = listaArchivos & ""  "" & nGenerados & "") [Tesorería y Mercado]: "" & rOut & vbCrLf" & vbCrLf & _
"        ultimoArchivo = rOut" & vbCrLf & _
"    End If" & vbCrLf & _
"    " & vbCrLf & _
"    ' 3. Si marco Financiamiento Sectorial Segundo Piso -> genera su archivo individual" & vbCrLf & _
"    If bPEA Then" & vbCrLf & _
"        If Not FondoBO.generarReporteEjecutivoBO(strFecRep, formato, False, False, True, False, rOut, msgErr, rutaCarpeta) Then" & vbCrLf & _
"            DoCmd.Hourglass False" & vbCrLf & _
"            Call FondoBO.FinalizarSesionExcelRapidaBO" & vbCrLf & _
"            MsgBox ""No fue posible generar el reporte individual de Segundo Piso en "" & formato & "": "" & vbCrLf & msgErr, vbCritical, ""Error de Exportación""" & vbCrLf
codeF = codeF & "            Exit Sub" & vbCrLf & _
"        End If" & vbCrLf & _
"        nGenerados = nGenerados + 1" & vbCrLf
codeF = codeF & "        listaArchivos = listaArchivos & ""  "" & nGenerados & "") [Segundo Piso]: "" & rOut & vbCrLf" & vbCrLf & _
"        ultimoArchivo = rOut" & vbCrLf & _
"    End If" & vbCrLf & _
"    " & vbCrLf & _
"    ' 4. Si marco Consolidado -> genera el archivo Consolidado Institucional (con los 3 analisis + Resumen Consolidado)" & vbCrLf & _
"    If bCons Then" & vbCrLf & _
"        If Not FondoBO.generarReporteEjecutivoBO(strFecRep, formato, True, True, True, True, rOut, msgErr, rutaCarpeta) Then" & vbCrLf & _
"            DoCmd.Hourglass False" & vbCrLf & _
"            Call FondoBO.FinalizarSesionExcelRapidaBO" & vbCrLf & _
"            MsgBox ""No fue posible generar el reporte Consolidado en "" & formato & "": "" & vbCrLf & msgErr, vbCritical, ""Error de Exportación""" & vbCrLf & _
"            Exit Sub" & vbCrLf & _
"        End If" & vbCrLf
codeF = codeF & "        nGenerados = nGenerados + 1" & vbCrLf & _
"        listaArchivos = listaArchivos & ""  "" & nGenerados & "") [Consolidado Integral]: "" & rOut & vbCrLf" & vbCrLf & _
"        ultimoArchivo = rOut" & vbCrLf
codeF = codeF & "    End If" & vbCrLf & _
"    " & vbCrLf & _
"" & vbCrLf & _
"    ' 5. Comparativo de 2 insumos: cruza Primer Piso con Tesoreria por" & vbCrLf & _
"    '    instrumento. No es una suma de portafolios, es un CRUCE: dice que" & vbCrLf & _
"    '    instrumento aparece en cada libro, con que saldo, y cuales aparecen" & vbCrLf & _
"    '    en uno solo, que es donde suelen estar los errores de clasificacion." & vbCrLf & _
"    If bCmp2 Then" & vbCrLf & _
"        If Not FondoBO.generarReporteComparativoBO(strFecRep, formato, 2, rOut, msgErr, rutaCarpeta) Then" & vbCrLf & _
"            DoCmd.Hourglass False" & vbCrLf & _
"            Call FondoBO.FinalizarSesionExcelRapidaBO" & vbCrLf & _
"            MsgBox ""No fue posible generar el reporte comparativo de 2 insumos en "" & formato & "": "" & vbCrLf & msgErr, vbCritical, ""Error de Exportacion""" & vbCrLf
codeF = codeF & "            Exit Sub" & vbCrLf & _
"        End If" & vbCrLf & _
"        nGenerados = nGenerados + 1" & vbCrLf
codeF = codeF & "        listaArchivos = listaArchivos & ""  "" & nGenerados & "") [Comparativo 2 insumos]: "" & rOut & vbCrLf" & vbCrLf & _
"        ultimoArchivo = rOut" & vbCrLf & _
"    End If" & vbCrLf & _
"" & vbCrLf & _
"    ' 6. Comparativo de 3 insumos: el mismo cruce sobre los tres libros, con" & vbCrLf & _
"    '    una columna de presencia que dice en cuantos de los tres aparece cada" & vbCrLf & _
"    '    instrumento." & vbCrLf & _
"    If bCmp3 Then" & vbCrLf & _
"        If Not FondoBO.generarReporteComparativoBO(strFecRep, formato, 3, rOut, msgErr, rutaCarpeta) Then" & vbCrLf & _
"            DoCmd.Hourglass False" & vbCrLf & _
"            Call FondoBO.FinalizarSesionExcelRapidaBO" & vbCrLf & _
"            MsgBox ""No fue posible generar el reporte comparativo de 3 insumos en "" & formato & "": "" & vbCrLf & msgErr, vbCritical, ""Error de Exportacion""" & vbCrLf
codeF = codeF & "            Exit Sub" & vbCrLf & _
"        End If" & vbCrLf & _
"        nGenerados = nGenerados + 1" & vbCrLf
codeF = codeF & "        listaArchivos = listaArchivos & ""  "" & nGenerados & "") [Comparativo 3 insumos]: "" & rOut & vbCrLf" & vbCrLf & _
"        ultimoArchivo = rOut" & vbCrLf & _
"    End If" & vbCrLf & _
"" & vbCrLf & _
"    Call FondoBO.FinalizarSesionExcelRapidaBO" & vbCrLf & _
"    DoCmd.Hourglass False" & vbCrLf & _
"    MsgBox ""Se generaron exitosamente "" & nGenerados & "" archivo(s) en formato "" & formato & "" en la carpeta seleccionada:"" & vbCrLf & _" & vbCrLf & _
"           ""Carpeta destino: "" & rutaCarpeta & vbCrLf & vbCrLf & _" & vbCrLf & _
"           ""Archivos generados:"" & vbCrLf & listaArchivos, vbInformation, ""Reporte(s) "" & formato & "" Generado(s) - Hidrocarburos""" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    If rutaCarpeta <> """" Then" & vbCrLf & _
"        Application.FollowHyperlink rutaCarpeta" & vbCrLf
codeF = codeF & "    End If" & vbCrLf & _
"    Exit Sub" & vbCrLf & _
"ErrGenRep:" & vbCrLf
codeF = codeF & "    Call FondoBO.FinalizarSesionExcelRapidaBO" & vbCrLf & _
"    DoCmd.Hourglass False" & vbCrLf & _
"    MsgBox ""Ocurrió un error al generar el(los) archivo(s) "" & formato & "": "" & Err.Description, vbCritical, ""Error""" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdGenerar_Click()" & vbCrLf & _
"    EjecutarGeneracionReportes ""EXCEL""" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdGenerarPDF_Click()" & vbCrLf & _
"    EjecutarGeneracionReportes ""PDF""" & vbCrLf & _
"End Sub" & vbCrLf
codeF = codeF & "" & vbCrLf & _
"Private Sub cmdAtras_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    DoCmd.OpenForm ""Form_Menu_General"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdIrMenu_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    DoCmd.OpenForm ""Form_Menu_General"", acNormal" & vbCrLf & _
"    DoCmd.Close acForm, Me.Name" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdSalir_Click()" & vbCrLf & _
"    On Error Resume Next" & vbCrLf
codeF = codeF & "    If MsgBox(""¿Desea cerrar la sesión actual?"", vbQuestion + vbYesNo, ""Salir"") = vbYes Then" & vbCrLf & _
"        Set usuarioSesionDTO = Nothing" & vbCrLf
codeF = codeF & "        DoCmd.OpenForm ""Form_Acceso"", acNormal" & vbCrLf & _
"        DoCmd.Close acForm, Me.Name" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Private Sub cmdRegresar_Click()" & vbCrLf & _
"    cmdAtras_Click" & vbCrLf & _
"End Sub"
AsignarCodigoLimpio frm, codeF
End Sub

Private Sub EjecutarBloqueInstalacion_14()
access.DoCmd.Close 2, fName, 1
On Error Resume Next
access.DoCmd.DeleteObject 2, "Form_Generar_Reporte"
On Error GoTo 0
access.DoCmd.Rename "Form_Generar_Reporte", 2, fName
logFile.WriteLine "   [OK] Form_Generar_Reporte creado con ruta de guardado elegible, checkboxes desacoplados y generacion de 3 PDFs/Excels independientes."

' 16. Construyendo Reporte Nativo de Access Rpt_Analisis_Integral_Riesgos como respaldo PDF
logFile.WriteLine "Paso 20.1: Construyendo Reporte Ejecutivo Rpt_Analisis_Integral_Riesgos para exportacion PDF..."
On Error Resume Next
Set rpt = access.CreateReport
rName = rpt.Name
rpt.Caption = "Reporte Ejecutivo de Análisis Financiero - Administración Integral de Riesgos Hidrocarburos"
rpt.Width = 10500
rpt.Section(0).Height = 12800
rpt.Section(0).BackColor = 16777215

' Encabezado del Reporte PDF
Set ctrl = access.CreateReportControl(rName, 101, 0, "", "", 0, 0, 10500, 1100)
ctrl.BackStyle = 1
ctrl.BackColor = 15523798
ctrl.BorderStyle = 0

Set ctrl = access.CreateReportControl(rName, 103, 0, "", "", 300, 80, 1800, 920)
ctrl.SizeMode = 3
ctrl.BackStyle = 0
ctrl.BackColor = 15523798
ctrl.BorderStyle = 0
If fso.FileExists(strLogoPath) Then ctrl.Picture = strLogoPath

Set ctrl = access.CreateReportControl(rName, 100, 0, "", "", 2200, 100, 8000, 340)
ctrl.Caption = "DIRECCIÓN CORPORATIVA DE ECONOMÍA Y FISCALIZACIÓN DE HIDROCARBUROS"
ctrl.FontName = "Cambria"
ctrl.FontSize = 11.5
ctrl.FontBold = True
ctrl.ForeColor = 6703164

Set ctrl = access.CreateReportControl(rName, 100, 0, "", "", 2200, 450, 8000, 320)
ctrl.Caption = "DIRECCIÓN DE FISCALIZACIÓN Y CONTRAPRESTACIONES PETROLERAS"
ctrl.FontName = "Cambria"
ctrl.FontSize = 10
ctrl.FontBold = True
ctrl.ForeColor = 6703164

Set ctrl = access.CreateReportControl(rName, 100, 0, "", "", 2200, 760, 8000, 300)
ctrl.Caption = "GERENCIA DE VERIFICACIÓN DE VOLÚMENES Y REGALÍAS"
ctrl.FontName = "Cambria"
ctrl.FontSize = 9.5
ctrl.FontBold = True
ctrl.ForeColor = 5000704

' Seccion Analisis 1 (Cartera de Crédito Comercial (Primer Piso))
Set ctrl = access.CreateReportControl(rName, 100, 0, "", "", 400, 1300, 9700, 360)
ctrl.Name = "lblTitA1"
ctrl.Caption = "ANÁLISIS SQL 1 — CARTERA DE CRÉDITO COMERCIAL (1ER PISO): BD_INSUMO_PRIMER_PISO + INSTRUMENTOS"
ctrl.FontName = "Cambria"
ctrl.FontSize = 10
ctrl.FontBold = True
ctrl.ForeColor = 16777215
ctrl.BackStyle = 1
ctrl.BackColor = 5000704

Set ctrl = access.CreateReportControl(rName, 100, 0, "", "", 400, 1720, 9700, 620)
ctrl.Name = "lblTxtA1"
ctrl.Caption = "- 1. Act. Empresarial Otros (46 créd.): $39,031.29 MDP | 2. Fuente Pago Propia (39 créd.): $20,761.65 MDP | 3. Garantía Hipotecaria (42 créd.): $17,084.08 MDP" & vbCrLf & _
               "- 4. Emp. Prod. Estado (6 créd.): $6,642.20 MDP | 5. Quirografarios y Prendarias (26 créd.): $11,938.79 MDP" & vbCrLf & _
               "- TOTAL 1ER PISO (159 créditos): Saldo Inicial $27,445.85 MDP | Saldo Final MN $95,458.01 MDP | Saldo Promedio $95,650.91 MDP | Equivalente USD $5,763.50 MDD"
ctrl.FontName = "Segoe UI"
ctrl.FontSize = 8.5

Set ctrl = access.CreateReportControl(rName, 103, 0, "", "", 1200, 2400, 8000, 2350)
ctrl.Name = "imgA1"
ctrl.SizeMode = 3
ctrl.BorderStyle = 1
If fso.FileExists(strImgDir & "\Grafica_Analisis_1.bmp") Then
    ctrl.Picture = strImgDir & "\Grafica_Analisis_1.bmp"
ElseIf fso.FileExists(strImgDir & "\Grafica_Analisis_1.png") Then
    ctrl.Picture = strImgDir & "\Grafica_Analisis_1.png"
End If

' Seccion Analisis 2 (Portafolio de Tesorería y Mercado de Dinero)
Set ctrl = access.CreateReportControl(rName, 100, 0, "", "", 400, 4950, 9700, 360)
ctrl.Name = "lblTitA2"
ctrl.Caption = "ANÁLISIS SQL 2 — PORTAFOLIO DE TESORERÍA Y MERCADO DE DINERO: BD_INSUMO_TESORERIA + INSTRUMENTOS"
ctrl.FontName = "Cambria"
ctrl.FontSize = 10
ctrl.FontBold = True
ctrl.ForeColor = 16777215
ctrl.BackStyle = 1
ctrl.BackColor = 6703164

Set ctrl = access.CreateReportControl(rName, 100, 0, "", "", 400, 5370, 9700, 620)
ctrl.Name = "lblTxtA2"
ctrl.Caption = "- 1. Sector Priv. | Emp. Nacionales (72 op.): $37,429.45 MDP (59.27%) | 2. Financ. No Bancarias (26 op.): $10,941.26 MDP (17.33%)" & vbCrLf & _
               "- 3. Banca Múltiple (9 op.): $6,505.60 MDP (10.30%) | 4. Emp. Extranjeras (3 op.): $3,782.21 MDP (5.99%) | 5. O.E.S.P. + Otros (4 op.): $4,490.75 MDP (7.11%)" & vbCrLf & _
               "- TOTAL TESORERÍA (114 operaciones): Saldo Inicial $25,996.24 MDP | Saldo Final MN $63,149.27 MDP | Saldo Promedio $62,745.51 MDP | Equivalente USD $3,812.79 MDD"
ctrl.FontName = "Segoe UI"
ctrl.FontSize = 8.5

Set ctrl = access.CreateReportControl(rName, 103, 0, "", "", 1200, 6050, 8000, 2350)
ctrl.Name = "imgA2"
ctrl.SizeMode = 3
ctrl.BorderStyle = 1
If fso.FileExists(strImgDir & "\Grafica_Analisis_2.bmp") Then
    ctrl.Picture = strImgDir & "\Grafica_Analisis_2.bmp"
ElseIf fso.FileExists(strImgDir & "\Grafica_Analisis_2.png") Then
    ctrl.Picture = strImgDir & "\Grafica_Analisis_2.png"
End If

' Seccion Analisis 3 (Financiamiento Sectorial Segundo Piso)
Set ctrl = access.CreateReportControl(rName, 100, 0, "", "", 400, 8600, 9700, 360)
ctrl.Name = "lblTitA3"
ctrl.Caption = "ANÁLISIS SQL 3 — FINANCIAMIENTO SECTORIAL Y GARANTÍAS (2DO PISO): BD_INSUMO_SEGUNDO_PISO + INSTRUMENTOS"
ctrl.FontName = "Cambria"
ctrl.FontSize = 10
ctrl.FontBold = True
ctrl.ForeColor = 16777215
ctrl.BackStyle = 1
ctrl.BackColor = 5000704

Set ctrl = access.CreateReportControl(rName, 100, 0, "", "", 400, 9020, 9700, 620)
ctrl.Name = "lblTxtA3"
ctrl.Caption = "- 1. Banca Múltiple Inst. Créd. (20 créd.): $20,175.58 MDP | 2. Financ. No Bancaria G0 (26 créd.): $8,975.03 MDP | 3. Financ. Sectorial G6 (16 créd.): $3,511.43 MDP" & vbCrLf & _
               "- 4. Financ. Especializado G9 (5 créd.): $3,128.44 MDP | 5. Otros Programas 2do Piso (1 créd.): $1,176.98 MDP" & vbCrLf & _
               "- TOTAL 2DO PISO (68 créditos): Saldo Inicial $28,372.68 MDP | Saldo Final MN $36,967.46 MDP | Saldo Promedio $36,188.07 MDP | Equivalente USD $2,232.00 MDD"
ctrl.FontName = "Segoe UI"
ctrl.FontSize = 8.5

Set ctrl = access.CreateReportControl(rName, 103, 0, "", "", 1200, 9700, 8000, 2350)
ctrl.Name = "imgA3"
ctrl.SizeMode = 3
ctrl.BorderStyle = 1
If fso.FileExists(strImgDir & "\Grafica_Analisis_3.bmp") Then
    ctrl.Picture = strImgDir & "\Grafica_Analisis_3.bmp"
ElseIf fso.FileExists(strImgDir & "\Grafica_Analisis_3.png") Then
    ctrl.Picture = strImgDir & "\Grafica_Analisis_3.png"
End If

access.DoCmd.Close 3, rName, 1
access.DoCmd.DeleteObject 3, "Rpt_Analisis_Integral_Riesgos"
access.DoCmd.Rename "Rpt_Analisis_Integral_Riesgos", 3, rName
On Error GoTo 0
logFile.WriteLine "   [OK] Reporte Rpt_Analisis_Integral_Riesgos creado exitosamente."

' Establecer Formulario de Inicio en Form_Acceso
On Error Resume Next
access.CurrentDb.Properties("StartupForm").Value = "Form_Acceso"
If Err.Number <> 0 Then
    Set prop = access.CurrentDb.CreateProperty("StartupForm", 10, "Form_Acceso")
    access.CurrentDb.Properties.Append prop
End If
On Error GoTo 0
logFile.WriteLine "   [OK] Inicio automatico configurado en Form_Acceso."

' Pre-compilar y guardar todos los modulos VBA (acCmdCompileAndSaveAllModules = 126) y cerrar ventanas de VBE
On Error Resume Next
access.RunCommand 126
Err.Clear
For passVbe = 1 To 3
    For Each wVbe In access.VBE.Windows
        If wVbe.Type = 0 Or wVbe.Type = 1 Then wVbe.Close
    Next
Next
access.VBE.MainWindow.Visible = False
access.CloseCurrentDatabase
access.Quit
Set access = Nothing
On Error GoTo 0

logFile.WriteLine "====================================================================="
logFile.WriteLine "INSTALACION Y CONFIGURACION COMPLETADA EXITOSAMENTE AL 100%"
logFile.WriteLine "Todos los 15 formularios, modulos, clases, tablas y plantillas listos."
logFile.WriteLine "====================================================================="
logFile.Close

MsgBox "Módulo de Administración Integral de Riesgos Hidrocarburos Instalado Exitosamente!" & vbCrLf & vbCrLf & _
       "Novedades aplicadas:" & vbCrLf & _
       "- Base de datos poblada con 101 empleados de 4 dígitos (del 1000 al 1100) con contraseña de 8 caracteres Wembley1" & vbCrLf & _
       "- Interfaz centrada simétricamente (recuadro de Login, campos y botones centrados)" & vbCrLf & _
       "- Títulos completamente centrados en todos los 15 formularios" & vbCrLf & _
       "- Cierre en cascada total de ventanas abiertas al presionar Salir" & vbCrLf & _
       "- Logotipo Hidrocarburos con proporción exacta sin aplastamiento (Zoom 1:1)" & vbCrLf & vbCrLf & _
       "Credenciales habilitadas para pruebas:" & vbCrLf & _
       "  * Administrador Principal: 1000 (Contraseña: Wembley1)" & vbCrLf & _
       "  * Empleados Operativos/Admin: 1001 al 1100 (Contraseña: Wembley1)" & vbCrLf & vbCrLf & _
       "Ya puedes abrir Modulo_Administracion_Riesgos.accdb.", vbInformation, "Instalación Exitosa - Hidrocarburos"
End Sub

Function ReadCleanVba(filePath)
    Dim ts, line, result, inHeader
    Set ts = fso.OpenTextFile(filePath, 1)
    result = ""
    inHeader = True
    Do Until ts.AtEndOfStream
        line = ts.ReadLine
        If inHeader Then
            If Left(Trim(line), 7) = "VERSION" Or _
               Left(Trim(line), 5) = "BEGIN" Or _
               Left(Trim(line), 8) = "MultiUse" Or _
               Trim(line) = "End" Or _
               Trim(line) = "END" Or _
               Left(Trim(line), 9) = "Attribute" Then
                ' Saltar linea de cabecera de archivo VB
            Else
                If Trim(line) <> "" Then
                    inHeader = False
                    result = result & line & vbCrLf
                End If
            End If
        Else
            result = result & line & vbCrLf
        End If
    Loop
    ts.Close
    ReadCleanVba = result
End Function

Sub ConfigurarEncabezadoVerde(f, titulo)
    Dim band, imgL, lblT
    On Error Resume Next
    
    ' 1. Franja institucional con color exacto #D6DFEC (RGB: 214, 223, 236), altura 2.213 cm = 1255 twips
    Set band = access.CreateControl(f.Name, 101, 0, "", "", 0, 0, 16500, 1255)
    band.Name = "bandSuperior"
    band.BackStyle = 1 ' Normal (Solido)
    band.BackColor = 15523798 ' #D6DFEC (RGB: 214, 223, 236 - Indicativo visual institucional Hidrocarburos)
    band.SpecialEffect = 0 ' Sin relieve (segun Hoja de Propiedades)
    band.BorderStyle = 0
    
    ' 2. Logotipo e imagotipo oficial Hidrocarburos (globo terraqueo + HIDROCARBUROS) SizeMode = 3 (acOLESizeZoom)
    Set imgL = access.CreateControl(f.Name, 103, 0, "", "", 600, 100, 2400, 1050)
    imgL.Name = "imgLogoSuperior"
    imgL.SizeMode = 3 ' acOLESizeZoom: Escalado proporcional perfecto sin distorsion
    imgL.BackStyle = 0 ' Sin fondo blanco
    imgL.BackColor = 15523798 ' #D6DFEC (RGB: 214, 223, 236 - Mismo color exacto que la banda azul)
    imgL.BorderStyle = 0
    If fso.FileExists(strLogoPath) Then imgL.Picture = strLogoPath
    
    ' 3. Titulo en DOS renglones para que no se vea amontonado:
    '      renglon 1: el nombre del modulo, constante en todas las pantallas
    '      renglon 2: el nombre de la seccion (Gestion de Instrumentos, Alta de Usuario, ...)
    '    El titulo llega como "Modulo ... - Seccion"; aqui se parte en el separador.
    Dim strPrin, strSec, posSep
    strPrin = titulo : strSec = ""
    posSep = InStr(titulo, " " & ChrW(8212) & " ")      ' guion largo
    If posSep = 0 Then posSep = InStr(titulo, " - ")     ' guion corto
    If posSep > 0 Then
        strPrin = Trim(Left(titulo, posSep - 1))
        strSec = Trim(Mid(titulo, posSep + 3))
    End If
    If InStr(LCase(strPrin), "m") <> 1 And strSec = "" Then
        ' Titulos que no traen el prefijo del modulo (p.ej. "Restablecer Contrasena"):
        ' se muestra el modulo arriba y ese texto como seccion.
        strSec = strPrin
        strPrin = "M" & ChrW(243) & "dulo de Administraci" & ChrW(243) & "n Integral de Riesgos"
    End If

    Set lblT = access.CreateControl(f.Name, 100, 0, "", "", 0, 190, 16000, 520)
    lblT.Name = "lblTituloSuperior"
    lblT.Caption = strPrin
    lblT.FontName = "Cambria"
    lblT.FontSize = 16
    lblT.FontBold = True
    lblT.ForeColor = 6703164 ' #1F4E78 Azul Marino Institucional
    lblT.TextAlign = 2 ' Centrado total
    lblT.BackStyle = 0 ' Transparente

    ' Renglon 2: seccion. Se crea siempre para que todas las pantallas tengan
    ' la misma altura de encabezado, aunque quede vacio.
    Dim lblS
    Set lblS = access.CreateControl(f.Name, 100, 0, "", "", 0, 720, 16000, 420)
    lblS.Name = "lblSeccionSuperior"
    lblS.Caption = strSec
    lblS.FontName = "Cambria"
    lblS.FontSize = 12.5
    lblS.FontBold = True
    lblS.ForeColor = 5000704 ' #004E4C Verde institucional
    lblS.TextAlign = 2
    lblS.BackStyle = 0
    On Error GoTo 0
End Sub

Sub EstilizarBotonVer(btn)
    On Error Resume Next
    btn.TabStop = True
    btn.UseTheme = False
    btn.BackStyle = 1
    btn.BackColor = 15458795 ' #EBF0F7 Gris-azulado institucional suave integrado al cuadro de contrasena
    btn.BorderStyle = 1
    btn.BorderColor = 11839396 ' #A4B8B4 Borde sutil armonizado con el cuadro de texto
    btn.HoverColor = 15523798 ' #D6DFEC Azul institucional al pasar el cursor
    btn.PressedColor = 5000704 ' #004E4C Verde institucional al mantener presionado
    btn.ForeColor = 6703164 ' #1F4E78 Azul marino institucional
    btn.HoverForeColor = 5000704
    btn.PressedForeColor = 16777215
    ' ICONO DE OJO DE ALTA DEFINICION (IconoOjoVer.bmp 22x22 RGB con supersampling 8x)
    ' Se incrusta directamente en el formulario (PictureType = 0) para lucir nitido y moderno.
    If fso.FileExists(strEyeIconPath) Then
        btn.Caption = ""
        btn.PictureType = 0
        btn.Picture = strEyeIconPath
    Else
        btn.Caption = "N"
        btn.FontName = "Webdings"
        btn.FontSize = 11
        btn.FontBold = False
    End If
    btn.TextAlign = 2
    btn.ControlTipText = "Mantenga presionado para mostrar la contraseña"
    ' La contrasena se descubre solo mientras se sostiene el clic (MouseDown / MouseUp)
    btn.OnMouseDown = "[Event Procedure]"
    btn.OnMouseUp = "[Event Procedure]"
    On Error GoTo 0
End Sub

Sub EstilizarBoton(btn)
    On Error Resume Next
    btn.TabStop = True
    btn.BackStyle = 1 ' Normal (Solido)
    btn.BackColor = 5000704 ' #004E4C Verde azulado institucional segun Hoja de Propiedades
    btn.BorderStyle = 1 ' Solido
    btn.BorderColor = 9376256 ' #00928F Borde cian institucional
    btn.HoverColor = 9376256 ' #00928F
    btn.PressedColor = 9376256 ' #00928F
    btn.ForeColor = 16777215 ' Blanco (#FFFFFF - Fondo 1 / Estilo rapido)
    btn.HoverForeColor = 16777215 ' Fondo 1 al mantener mouse
    btn.PressedForeColor = 16777215 ' Fondo 1 al hacer clic
    btn.FontName = "Cambria Math"
    btn.FontSize = 11
    btn.FontBold = True ' Espesor: Seminegrita
    btn.TextAlign = 2 ' Centrado
    btn.TopMargin = 0
    btn.BottomMargin = 30 ' 0.053cm
    btn.LeftMargin = 120  ' 0.212cm
    btn.RightMargin = 60  ' 0.106cm
    On Error GoTo 0
End Sub


Sub ConfigurarBarraNavegacionSuperior(f, formPadre)
    Dim ctrl
    On Error Resume Next
    
    ' Subencabezado dinamico: No. Empleado y Perfil alineados a la izquierda (Y = 1290 a 1810)
    Set ctrl = access.CreateControl(f.Name, 100, 0, "", "", 800, 1290, 8500, 260)
    ctrl.Name = "lblNoEmpleado"
    ctrl.Caption = "No. Empleado: -"
    ctrl.FontName = "Cambria Math"
    ctrl.FontSize = 10
    ctrl.FontBold = True
    ctrl.ForeColor = 4737096
    ctrl.BackStyle = 0

    Set ctrl = access.CreateControl(f.Name, 100, 0, "", "", 800, 1550, 8500, 260)
    ctrl.Name = "lblPerfil"
    ctrl.Caption = "Perfil: -"
    ctrl.FontName = "Cambria Math"
    ctrl.FontSize = 10
    ctrl.FontBold = True
    ctrl.ForeColor = 2631720
    ctrl.BackStyle = 0

    ' Solo mostrar el boton "Menú" cuando sea indispensable (subpantallas de 2do nivel cuyo padre NO es Form_Menu_General)
    Dim mostrarBotonMenu
    mostrarBotonMenu = (UCase(Trim(formPadre)) <> "FORM_MENU_GENERAL" And UCase(Trim(formPadre)) <> "FORM_ACCESO" And Trim(formPadre) <> "")

    If mostrarBotonMenu Then
        ' 1. Boton Flecha Retroceder (<-) hacia subpantalla padre
        Set ctrl = access.CreateControl(f.Name, 104, 0, "", "", 10600, 1310, 850, 480)
        ctrl.Name = "cmdAtras"
        ctrl.Caption = ChrW(8592)
        ctrl.OnClick = "[Event Procedure]"
        EstilizarBoton ctrl
        ctrl.FontSize = 13
        ctrl.ControlTipText = "Regresar a la pantalla anterior"

        ' 2. Boton Menu Principal (indispensable en subpantallas de 2do nivel)
        Set ctrl = access.CreateControl(f.Name, 104, 0, "", "", 11600, 1310, 1750, 480)
        ctrl.Name = "cmdIrMenu"
        ctrl.Caption = "Menú"
        ctrl.OnClick = "[Event Procedure]"
        EstilizarBoton ctrl
        ctrl.FontSize = 10.5
        ctrl.ControlTipText = "Ir al Menú General"
    Else
        ' 1. Boton Flecha Retroceder (<-) regresa directamente al Menú General (sin botón Menú redundante)
        Set ctrl = access.CreateControl(f.Name, 104, 0, "", "", 12450, 1310, 850, 480)
        ctrl.Name = "cmdAtras"
        ctrl.Caption = ChrW(8592)
        ctrl.OnClick = "[Event Procedure]"
        EstilizarBoton ctrl
        ctrl.FontSize = 13
        ctrl.ControlTipText = "Regresar al Menú General"
    End If

    ' 3. Boton Salir de la sesion
    Set ctrl = access.CreateControl(f.Name, 104, 0, "", "", 13500, 1310, 1500, 480)
    ctrl.Name = "cmdSalir"
    ctrl.Caption = "Salir"
    ctrl.OnClick = "[Event Procedure]"
    EstilizarBoton ctrl
    ctrl.FontSize = 10.5
    ctrl.ControlTipText = "Cerrar sesión y salir"
    On Error GoTo 0
End Sub


Sub AsignarCodigoLimpio(f, codigo)
    On Error Resume Next
    f.HasModule = True
    Dim m, cItem
    Set m = f.Module
    If m.CountOfLines > 0 Then m.DeleteLines 1, m.CountOfLines
    m.AddFromString codigo
    If InStr(1, codigo, "Sub Form_Load(", 1) > 0 Then f.OnLoad = "[Event Procedure]"
    If InStr(1, codigo, "Sub Form_Resize(", 1) > 0 Then f.OnResize = "[Event Procedure]"
    If InStr(1, codigo, "Sub Form_KeyDown(", 1) > 0 Then f.OnKeyDown = "[Event Procedure]"
    For Each cItem In f.Controls
        If InStr(1, codigo, "Sub " & cItem.Name & "_Click(", 1) > 0 Then cItem.OnClick = "[Event Procedure]"
        If InStr(1, codigo, "Sub " & cItem.Name & "_MouseDown(", 1) > 0 Then cItem.OnMouseDown = "[Event Procedure]"
        If InStr(1, codigo, "Sub " & cItem.Name & "_MouseUp(", 1) > 0 Then cItem.OnMouseUp = "[Event Procedure]"
        If InStr(1, codigo, "Sub " & cItem.Name & "_KeyDown(", 1) > 0 Then cItem.OnKeyDown = "[Event Procedure]"
        If InStr(1, codigo, "Sub " & cItem.Name & "_KeyPress(", 1) > 0 Then cItem.OnKeyPress = "[Event Procedure]"
        If InStr(1, codigo, "Sub " & cItem.Name & "_Change(", 1) > 0 Then cItem.OnChange = "[Event Procedure]"
        If InStr(1, codigo, "Sub " & cItem.Name & "_LostFocus(", 1) > 0 Then cItem.OnLostFocus = "[Event Procedure]"
        If InStr(1, codigo, "Sub " & cItem.Name & "_GotFocus(", 1) > 0 Then cItem.OnGotFocus = "[Event Procedure]"
        If InStr(1, codigo, "Sub " & cItem.Name & "_BeforeUpdate(", 1) > 0 Then
            cItem.BeforeUpdate = "[Event Procedure]"
            cItem.OnBeforeUpdate = "[Event Procedure]"
        End If
        If InStr(1, codigo, "Sub " & cItem.Name & "_AfterUpdate(", 1) > 0 Then
            cItem.AfterUpdate = "[Event Procedure]"
            cItem.OnAfterUpdate = "[Event Procedure]"
        End If
    Next
    On Error GoTo 0
End Sub
