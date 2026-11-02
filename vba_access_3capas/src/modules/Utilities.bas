Attribute VB_Name = "Utilities"
Option Compare Database
Option Explicit

' =============================================================================
' MÓDULO TRANSVERSAL: Utilities.bas
' -----------------------------------------------------------------------------
' Contiene las funciones utilitarias de apoyo para la Capa de Presentación y la
' Capa de Negocio:
'   1. Validación y restricción de teclado en campos de empleado (4 dígitos),
'      contraseñas (8 caracteres) y fechas (dd/mm/aaaa).
'   2. Cifrado y descifrado simétrico institucional de contraseñas (ENC$...).
'   3. Lectura segura de controles de formulario (LeerTextoControl).
'   4. Adaptación responsiva de formularios a pantalla completa (AjustarTamanoFormularioCompleto).
'   5. Resaltado en verde de portafolios cargados por fecha (ActualizarOpcionesSubidasPorFecha).
'   6. Visualización temporal de contraseña mientras se sostiene el clic (VerContrasenaSostenidaUT).
'   7. Aplicación del encabezado de sesión y control de acceso (AplicarEncabezadoSesionUT).
' =============================================================================

Public Function tieneLetras(ByVal cadena As String) As Boolean
    Dim i As Long, c As String
    For i = 1 To Len(cadena)
        c = Mid(cadena, i, 1)
        If Not (c >= "0" And c <= "9") Then
            tieneLetras = True
            Exit Function
        End If
    Next i
    tieneLetras = False
End Function

' Bloquea en KeyDown cualquier letra, simbolo con Shift o caracter especial en el campo Usuario
Public Sub RestringirKeyDownUsuario4Digitos(ByRef ctrl As Object, ByRef KeyCode As Integer, ByVal Shift As Integer)
    On Error Resume Next
    ' Permitir teclas de control y navegacion: Backspace(8), Tab(9), Enter(13), Esc(27), Fin(35), Inicio(36), Flechas(37-40), Supr(46)
    If KeyCode = 8 Or KeyCode = 9 Or KeyCode = 13 Or KeyCode = 27 Or (KeyCode >= 35 And KeyCode <= 40) Or KeyCode = 46 Then
        Exit Sub
    End If
    ' Si presiona Shift, Ctrl o Alt junto con un numero, bloquear para impedir caracteres especiales (!"#$%&/()=)
    If Shift <> 0 Then
        KeyCode = 0
        Exit Sub
    End If
    ' Permitir unicamente digitos fila superior (48 a 57) o teclado numerico (96 a 105)
    If Not ((KeyCode >= 48 And KeyCode <= 57) Or (KeyCode >= 96 And KeyCode <= 105)) Then
        KeyCode = 0
        Exit Sub
    End If
    Dim txtActual As String, selLen As Long
    txtActual = CStr(Nz(ctrl.Text, ""))
    selLen = CLng(Nz(ctrl.SelLength, 0))
    If (Len(txtActual) - selLen) >= 4 Then
        KeyCode = 0
    End If
    On Error GoTo 0
End Sub

' Impide ingresar letras, caracteres especiales o mas de 4 digitos en cualquier campo de Usuario (evento KeyPress)
Public Sub RestringirKeyPressUsuario4Digitos(ByRef ctrl As Object, ByRef KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 8 Then Exit Sub ' Permitir tecla Backspace
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
        Exit Sub
    End If
    Dim txtActual As String, selLen As Long
    txtActual = CStr(Nz(ctrl.Text, ""))
    selLen = CLng(Nz(ctrl.SelLength, 0))
    If (Len(txtActual) - selLen) >= 4 Then
        KeyAscii = 0
    End If
    On Error GoTo 0
End Sub

' Limpia cualquier letra o caracter especial y trunca a 4 digitos numericos en el campo de Usuario (evento Change)
Public Sub TruncarCambioUsuario4Digitos(ByRef ctrl As Object, Optional ByRef bufferUsuario As String = "")
    On Error Resume Next
    Dim s As String, limpio As String, i As Long, ch As String
    s = CStr(Nz(ctrl.Text, ""))
    limpio = ""
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            If Len(limpio) < 4 Then limpio = limpio & ch
        End If
    Next i
    If s <> limpio Then
        ctrl.Text = limpio
        ctrl.SelStart = Len(limpio)
    End If
    bufferUsuario = limpio
    On Error GoTo 0
End Sub

' Encriptacion institucional de contrasenas usando la llave maestra Wembley1 (8 caracteres)
Public Function cifrarContrasena(ByVal textoPlano As String, Optional ByVal claveSecreta As String = "Wembley1") As String
    Dim i As Long, bTxt As Integer, bKey As Integer, bOut As Integer
    Dim hexStr As String, h As String
    textoPlano = Trim(textoPlano)
    If textoPlano = "" Then
        cifrarContrasena = ""
        Exit Function
    End If
    If Left(textoPlano, 4) = "ENC$" Then
        cifrarContrasena = textoPlano
        Exit Function
    End If
    If claveSecreta = "" Then claveSecreta = "Wembley1"
    hexStr = "ENC$"
    For i = 1 To Len(textoPlano)
        bTxt = Asc(Mid(textoPlano, i, 1))
        bKey = Asc(Mid(claveSecreta, ((i - 1) Mod Len(claveSecreta)) + 1, 1))
        bOut = ((bTxt Xor ((bKey * 3 + i * 29) Mod 256)) + 113) Mod 256
        h = Hex(bOut)
        If Len(h) = 1 Then h = "0" & h
        hexStr = hexStr & h
    Next i
    cifrarContrasena = hexStr
End Function

Public Function descifrarContrasena(ByVal textoCifrado As String, Optional ByVal claveSecreta As String = "Wembley1") As String
    Dim i As Long, pos As Long, bEnc As Integer, bKey As Integer, bOrig As Integer
    Dim hexBody As String, outStr As String
    textoCifrado = Trim(textoCifrado)
    If Left(textoCifrado, 4) <> "ENC$" Then
        descifrarContrasena = textoCifrado
        Exit Function
    End If
    If claveSecreta = "" Then claveSecreta = "Wembley1"
    hexBody = Mid(textoCifrado, 5)
    outStr = ""
    pos = 1
    For i = 1 To Len(hexBody) Step 2
        bEnc = CInt("&H" & Mid(hexBody, i, 2))
        bKey = Asc(Mid(claveSecreta, ((pos - 1) Mod Len(claveSecreta)) + 1, 1))
        bOrig = ((bEnc - 113 + 256) Mod 256) Xor ((bKey * 3 + pos * 29) Mod 256)
        outStr = outStr & Chr(bOrig)
        pos = pos + 1
    Next i
    descifrarContrasena = outStr
End Function

' Registra cada tecla en un campo de Contrasena y restringe fisicamente a maximo 8 caracteres
Public Sub ProcesarKeyPressContrasena(ByRef bufferPwd As String, ByRef KeyAscii As Integer, Optional ByRef ctrl As Object = Nothing)
    On Error Resume Next
    If KeyAscii = 8 Then
        If Len(bufferPwd) > 0 Then
            bufferPwd = Left(bufferPwd, Len(bufferPwd) - 1)
        End If
        Exit Sub
    End If
    If KeyAscii >= 32 And KeyAscii <= 255 Then
        Dim lenVisual As Long, selLen As Long
        lenVisual = Len(bufferPwd)
        selLen = 0
        If Not ctrl Is Nothing Then
            selLen = CLng(Nz(ctrl.SelLength, 0))
            Dim txtCtrl As String
            txtCtrl = CStr(Nz(ctrl.Text, ""))
            If Len(txtCtrl) > lenVisual Then lenVisual = Len(txtCtrl)
            If selLen > 0 And selLen >= lenVisual Then
                bufferPwd = ""
                lenVisual = 0
                selLen = 0
            End If
        End If
        If (lenVisual - selLen) >= 8 Then
            KeyAscii = 0
            Exit Sub
        End If
        If Len(bufferPwd) >= 8 Then
            KeyAscii = 0
            Exit Sub
        End If
        bufferPwd = bufferPwd & Chr(KeyAscii)
    End If
    On Error GoTo 0
End Sub

' Sincroniza y trunca a maximo 8 caracteres cualquier campo de Contrasena en el evento Change
Public Sub TruncarCambioContrasena8Caracteres(ByRef ctrl As Object, ByRef bufferPwd As String)
    On Error Resume Next
    Dim sTxt As String
    sTxt = CStr(Nz(ctrl.Text, ""))
    If Len(sTxt) = 0 Then
        bufferPwd = ""
        Exit Sub
    End If
    If Len(sTxt) > 8 Then
        sTxt = Left(sTxt, 8)
        ctrl.Text = sTxt
        ctrl.SelStart = 8
    End If
    If Len(bufferPwd) > 8 Then
        bufferPwd = Left(bufferPwd, 8)
    End If
    If Replace(sTxt, "*", "") <> "" Then
        bufferPwd = sTxt
    End If
    On Error GoTo 0
End Sub

' Lectura universal segura de cualquier cuadro de texto o combo sin mutar InputMask:
' 1) Si el control tiene el foco (.Text accesible sin error y sin mascara de asteriscos), prioriza .Text (edicion en vivo)
' 2) Si el control no tiene el foco o tiene InputMask="Password", lee .Value confirmado
' 3) Si .Text/.Value devuelven asteriscos por InputMask="Password", utiliza bufferRespaldo capturado en KeyPress
Public Function LeerTextoControl(ByRef ctrl As Object, Optional ByVal bufferRespaldo As String = "") As String
    On Error Resume Next
    Dim valVal As String, valTxt As String
    valVal = ""
    valTxt = ""
    Err.Clear
    valTxt = Trim(CStr(Nz(ctrl.Text, "")))
    If Err.Number = 0 And valTxt <> "" And Replace(valTxt, "*", "") <> "" Then
        If Len(valTxt) > 8 And UCase(Trim(CStr(Nz(ctrl.InputMask, "")))) = "PASSWORD" Then valTxt = Left(valTxt, 8)
        LeerTextoControl = valTxt
        Exit Function
    End If
    Err.Clear
    valVal = Trim(CStr(Nz(ctrl.Value, "")))
    If Err.Number = 0 And valVal <> "" And Replace(valVal, "*", "") <> "" Then
        If Len(valVal) > 8 And UCase(Trim(CStr(Nz(ctrl.InputMask, "")))) = "PASSWORD" Then valVal = Left(valVal, 8)
        LeerTextoControl = valVal
        Exit Function
    End If
    If Trim(bufferRespaldo) <> "" And Replace(Trim(bufferRespaldo), "*", "") <> "" Then
        LeerTextoControl = Left(Trim(bufferRespaldo), 8)
        Exit Function
    End If
    LeerTextoControl = ""
    On Error GoTo 0
End Function

' Restringe fisicamente el campo Fecha a digitos (0-9) y diagonal (/) con maximo 10 caracteres (dd/mm/aaaa)
Public Sub RestringirKeyPressFecha(ByRef ctrl As Object, ByRef KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 8 Then Exit Sub ' Permitir Backspace
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 47) Then
        KeyAscii = 0
        Exit Sub
    End If
    Dim txtActual As String, selLen As Long
    txtActual = CStr(Nz(ctrl.Text, ""))
    selLen = CLng(Nz(ctrl.SelLength, 0))
    If (Len(txtActual) - selLen) >= 10 Then
        KeyAscii = 0
    End If
    On Error GoTo 0
End Sub

' Valida en tiempo real al escribir manualmente una fecha completa dd/mm/aaaa en el campo Fecha:
' Si el usuario escribe una fecha futura, muestra de inmediato el aviso de que no se puede escoger fechas futuras y restaura la fecha valida
Public Sub ValidarCambioFechaManual(ByRef ctrl As Object, ByRef flagValidando As Boolean, ByVal esReporte As Boolean, ByRef fechaUltimaValida As String)
    On Error Resume Next
    If flagValidando Then Exit Sub
    Dim s As String, limpio As String, i As Long, ch As String
    s = CStr(Nz(ctrl.Text, ""))
    limpio = ""
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "/" Then
            If Len(limpio) < 10 Then limpio = limpio & ch
        End If
    Next i
    If s <> limpio Then
        flagValidando = True
        ctrl.Text = limpio
        ctrl.SelStart = Len(limpio)
        flagValidando = False
    End If
    Dim partes() As String, fechaCompleta As Boolean
    fechaCompleta = False
    If InStr(limpio, "/") > 0 Then
        partes = Split(limpio, "/")
        If UBound(partes) = 2 Then
            If Len(partes(0)) = 2 And Len(partes(1)) = 2 And Len(partes(2)) = 4 Then
                fechaCompleta = True
            ElseIf Len(partes(0)) >= 1 And Len(partes(1)) >= 1 And Len(partes(2)) = 4 And CLng(Nz(ctrl.SelStart, 0)) = Len(limpio) Then
                fechaCompleta = True
            End If
        End If
    End If
    If Not fechaCompleta Then Exit Sub

    If fechaUltimaValida = "" Then
        If esReporte Then
            fechaUltimaValida = FondoBO.obtenerUltimaFechaCargadaBO()
        End If
        If fechaUltimaValida = "" Then fechaUltimaValida = Format(Date, "dd/mm/yyyy")
    End If

    If Not FondoBO.esFechaFormatoCorrectoBO(limpio) Then
        flagValidando = True
        MsgBox "Fecha inválida (" & limpio & "):" & vbCrLf & _
               "- Favor de ingresar una fecha real y válida con formato dd/mm/aaaa.", vbExclamation, "Fecha Inválida"
        ctrl.Value = fechaUltimaValida
        ctrl.Text = fechaUltimaValida
        ctrl.SelStart = Len(fechaUltimaValida)
        flagValidando = False
        Exit Sub
    End If

    Dim fecNorm As String
    fecNorm = FondoBO.normalizarFechaBO(limpio)
    If FondoBO.esFechaFuturaBO(fecNorm) Then
        flagValidando = True
        MsgBox "Fecha de consulta inválida (" & fecNorm & "): " & vbCrLf & _
               "- No se puede escoger fechas futuras (posteriores al día de la consulta actual: " & Format(Date, "dd/mm/yyyy") & ").", vbCritical, "Fecha Futura No Permitida"
        ctrl.Value = fechaUltimaValida
        ctrl.Text = fechaUltimaValida
        ctrl.SelStart = Len(fechaUltimaValida)
        flagValidando = False
        Exit Sub
    End If

    If esReporte And fecNorm <> fechaUltimaValida Then
        If FondoBO.contarRegistrosFechaBO(fecNorm) <= 0 Then
            Dim strDisp As String
            strDisp = FondoBO.obtenerFechasCargadasBO()
            If strDisp = "" Then strDisp = "Ninguna (cargue primero un insumo en 'Cargar Insumos')"
            flagValidando = True
            MsgBox "Fecha sin información de insumo (" & fecNorm & "): " & vbCrLf & _
                   "- No se ha cargado el insumo correspondiente a este día en la base de datos." & vbCrLf & _
                   "- Días marcados con información disponible: " & strDisp, vbExclamation, "Día sin Insumo Cargado"
            ctrl.Value = fechaUltimaValida
            ctrl.Text = fechaUltimaValida
            ctrl.SelStart = Len(fechaUltimaValida)
            flagValidando = False
            Exit Sub
        End If
    End If

    fechaUltimaValida = fecNorm
    flagValidando = True
    ctrl.Value = fecNorm
    flagValidando = False
    On Error GoTo 0
End Sub

' Valida al presionar Enter, Tab o salir del campo Fecha (LostFocus / AfterUpdate) que la fecha escrita manualmente no sea futura ni invalida
Public Function ValidarAlSalirFechaManual(ByRef ctrl As Object, ByRef flagValidando As Boolean, ByVal esReporte As Boolean, ByRef fechaUltimaValida As String, Optional ByVal forzarChequeoInsumo As Boolean = False) As Boolean
    On Error Resume Next
    If flagValidando Then
        ValidarAlSalirFechaManual = True
        Exit Function
    End If
    ValidarAlSalirFechaManual = False
    If fechaUltimaValida = "" Then
        If esReporte Then
            fechaUltimaValida = FondoBO.obtenerUltimaFechaCargadaBO()
        End If
        If fechaUltimaValida = "" Then fechaUltimaValida = Format(Date, "dd/mm/yyyy")
    End If

    Dim strFec As String
    strFec = LeerTextoControl(ctrl)
    If strFec = "" Then
        ValidarAlSalirFechaManual = True
        Exit Function
    End If

    If Not FondoBO.esFechaFormatoCorrectoBO(strFec) Then
        flagValidando = True
        MsgBox "Fecha inválida (" & strFec & "):" & vbCrLf & _
               "- Favor de ingresar una fecha real y válida con formato dd/mm/aaaa.", vbExclamation, "Fecha Inválida"
        ctrl.Value = fechaUltimaValida
        ctrl.Text = fechaUltimaValida
        ctrl.SetFocus
        ctrl.SelStart = Len(fechaUltimaValida)
        flagValidando = False
        Exit Function
    End If

    Dim fecNorm As String
    fecNorm = FondoBO.normalizarFechaBO(strFec)
    If FondoBO.esFechaFuturaBO(fecNorm) Then
        flagValidando = True
        MsgBox "Fecha de consulta inválida (" & fecNorm & "): " & vbCrLf & _
               "- No se puede escoger fechas futuras (posteriores al día de la consulta actual: " & Format(Date, "dd/mm/yyyy") & ").", vbCritical, "Fecha Futura No Permitida"
        ctrl.Value = fechaUltimaValida
        ctrl.Text = fechaUltimaValida
        ctrl.SetFocus
        ctrl.SelStart = Len(fechaUltimaValida)
        flagValidando = False
        Exit Function
    End If

    If esReporte And (forzarChequeoInsumo Or fecNorm <> fechaUltimaValida) Then
        If FondoBO.contarRegistrosFechaBO(fecNorm) <= 0 Then
            Dim strDisp As String
            strDisp = FondoBO.obtenerFechasCargadasBO()
            If strDisp = "" Then strDisp = "Ninguna (cargue primero un insumo en 'Cargar Insumos')"
            flagValidando = True
            MsgBox "Fecha sin información de insumo (" & fecNorm & "): " & vbCrLf & _
                   "- No se ha cargado el insumo correspondiente a este día en la base de datos." & vbCrLf & _
                   "- Días marcados con información disponible: " & strDisp, vbExclamation, "Día sin Insumo Cargado"
            ctrl.Value = fechaUltimaValida
            ctrl.Text = fechaUltimaValida
            ctrl.SetFocus
            ctrl.SelStart = Len(fechaUltimaValida)
            flagValidando = False
            Exit Function
        End If
    End If

    fechaUltimaValida = fecNorm
    flagValidando = True
    ctrl.Value = fecNorm
    ctrl.Text = fecNorm
    flagValidando = False
    ValidarAlSalirFechaManual = True
    On Error GoTo 0
End Function

' ==============================================================================================
' ADAPTACION RESPONSIVA DE PANTALLA COMPLETA (FULL-PAGE ACCESS WINDOW ADAPTATION)
' Ajusta dinamicamente cualquier formulario al 100% del ancho y alto de la pagina de Access,
' respetando el limite entero de 16 bits de Access (max 31,000 twips) y preservando las
' coordenadas originales de diseno en c.Tag. Nunca evalua propiedades exclusivas (.Caption)
' en controles que no sean etiquetas para evitar el bug de VBA On Error Resume Next en If/ElseIf.
' ==============================================================================================
Public Sub AjustarTamanoFormularioCompleto(ByRef frm As Object)
    On Error Resume Next
    If frm Is Nothing Then Exit Sub
    If UCase(Trim(CStr(frm.Name))) = "FORM_CALENDARIO_POPUP" Then Exit Sub

    Dim wWin As Long, hWin As Long
    wWin = CLng(Nz(frm.InsideWidth, 0))
    hWin = CLng(Nz(frm.InsideHeight, 0))
    If wWin < 16000 Then wWin = 16000
    If wWin > 30500 Then wWin = 30500
    If hWin < 6150 Then hWin = 6150
    If hWin > 20000 Then hWin = 20000

    ' Expandir el lienzo del formulario exactamente al ancho y alto de la ventana de Access
    If frm.Width < wWin Then frm.Width = wWin
    If frm.Section(0).Height < hWin Then frm.Section(0).Height = hWin

    Dim c As Object, cName As String, cType As Long, tagPartes() As String
    Dim origL As Long, origT As Long, origW As Long, origH As Long
    Dim tieneSubform As Boolean, tieneIrMenu As Boolean
    Dim minLeft As Long, maxRight As Long, shiftX As Long
    Dim margenX As Long, anchoUtil As Long, yPie As Long, subHeight As Long

    tieneSubform = False
    tieneIrMenu = False
    minLeft = 999999
    maxRight = 0

    ' 1. Guardar las coordenadas originales de diseno en c.Tag la primera vez que se abre el formulario
    For Each c In frm.Controls
        Err.Clear
        cName = UCase(Trim(CStr(c.Name)))
        cType = CLng(Nz(c.ControlType, 0))
        If InStr(CStr(Nz(c.Tag, "")), "|") = 0 Then
            c.Tag = CStr(CLng(c.Left)) & "|" & CStr(CLng(c.Top)) & "|" & CStr(CLng(c.Width)) & "|" & CStr(CLng(c.Height))
        End If
        If cType = 112 Then tieneSubform = True
        If cName = "CMDIRMENU" Then
            If c.Visible Then tieneIrMenu = True
        End If
    Next c

    ' 2. Calcular desplazamiento horizontal (shiftX) para centrar el bloque principal en formularios sin subformulario
    If Not tieneSubform Then
        For Each c In frm.Controls
            Err.Clear
            cName = UCase(Trim(CStr(c.Name)))
            tagPartes = Split(CStr(Nz(c.Tag, "")), "|")
            If UBound(tagPartes) = 3 Then
                origL = CLng(Val(tagPartes(0)))
                origT = CLng(Val(tagPartes(1)))
                origW = CLng(Val(tagPartes(2)))
                If origT >= 1800 And origT <= 5650 And cName <> "LBLFECHA" And cName <> "LBLFECHAPIE" And cName <> "LBLAREA" And cName <> "BANDSUPERIOR" And cName <> "LBLTITULOSUPERIOR" And cName <> "LBLSECCIONSUPERIOR" And cName <> "IMGLOGOSUPERIOR" Then
                    If origL < minLeft Then minLeft = origL
                    If (origL + origW) > maxRight Then maxRight = origL + origW
                End If
            End If
        Next c
        If maxRight > minLeft And minLeft < 999999 Then
            shiftX = (wWin \ 2) - ((minLeft + maxRight) \ 2)
            If (minLeft + shiftX) < 600 Then shiftX = 600 - minLeft
        Else
            shiftX = 0
        End If
    Else
        shiftX = 0
    End If

    margenX = 800
    anchoUtil = wWin - (margenX * 2)
    If anchoUtil < 14400 Then anchoUtil = 14400
    If (margenX + anchoUtil) > 30500 Then anchoUtil = 30500 - margenX

    yPie = hWin - 400
    If yPie < 5780 Then yPie = 5780

    ' 3. Aplicar posiciones exactas a cada control a partir de sus coordenadas de diseno originales
    For Each c In frm.Controls
        Err.Clear
        cName = UCase(Trim(CStr(c.Name)))
        cType = CLng(Nz(c.ControlType, 0))
        tagPartes = Split(CStr(Nz(c.Tag, "")), "|")
        If UBound(tagPartes) = 3 Then
            origL = CLng(Val(tagPartes(0)))
            origT = CLng(Val(tagPartes(1)))
            origW = CLng(Val(tagPartes(2)))
            origH = CLng(Val(tagPartes(3)))

            If cName = "BANDSUPERIOR" Or (origT = 0 And origH = 1255 And cType = 101) Then
                c.Left = 0
                c.Top = 0
                c.Width = wWin
                c.Height = 1255
            ElseIf cName = "LBLTITULOSUPERIOR" Or (origT = 300 And origH = 700 And cType = 100) Then
                c.Left = 0
                c.Top = origT
                c.Width = wWin
                c.Height = origH
            ElseIf cName = "LBLSECCIONSUPERIOR" Then
                c.Left = 0
                c.Top = origT
                c.Width = wWin
                c.Height = origH
            ElseIf cName = "IMGLOGOSUPERIOR" Or (origT = 100 And origH = 1050 And cType = 103) Then
                c.Left = 600
                c.Top = 100
                c.Width = 2400
                c.Height = 1050
            ElseIf cName = "CMDSALIR" Then
                c.Top = origT
                c.Width = origW
                c.Height = origH
                c.Left = wWin - origW - 800
            ElseIf cName = "CMDIRMENU" Then
                c.Top = origT
                c.Width = origW
                c.Height = origH
                c.Left = wWin - 1500 - 800 - origW - 200
            ElseIf cName = "CMDATRAS" Then
                c.Top = origT
                c.Width = origW
                c.Height = origH
                If tieneIrMenu Then
                    c.Left = wWin - 1500 - 800 - 1750 - 200 - origW - 150
                Else
                    c.Left = wWin - 1500 - 800 - origW - 200
                End If
            ElseIf cName = "LBLFECHA" Or cName = "LBLFECHAPIE" Then
                c.Top = yPie
                c.Width = origW
                c.Height = origH
                If tieneSubform Then
                    c.Left = margenX
                Else
                    c.Left = origL + shiftX
                End If
            ElseIf cName = "LBLAREA" Then
                c.Top = yPie
                c.Width = origW
                c.Height = origH
                If tieneSubform Then
                    c.Left = wWin - origW - margenX
                Else
                    c.Left = origL + shiftX
                    If (c.Left + origW) < (wWin - 800) Then c.Left = wWin - origW - 800
                End If
            ElseIf tieneSubform Then
                If cType = 112 Then
                    c.Left = margenX
                    c.Top = origT
                    c.Width = anchoUtil
                    subHeight = yPie - origT - 140
                    If subHeight < 1800 Then subHeight = 1800
                    c.Height = subHeight
                ElseIf cName = "LBLPOLITICABOX" Then
                    c.Left = margenX
                    c.Top = origT
                    c.Width = anchoUtil
                    c.Height = origH
                ElseIf cType = 100 And origW >= 11500 And origT >= 1800 And origT <= 3700 Then
                    c.Left = margenX
                    c.Top = origT
                    c.Width = anchoUtil
                    c.Height = origH
                Else
                    c.Left = origL
                    c.Top = origT
                    c.Width = origW
                    c.Height = origH
                End If
            Else
                If origT >= 1800 And origT <= 5650 Then
                    c.Left = origL + shiftX
                    c.Top = origT
                    c.Width = origW
                    c.Height = origH
                End If
            End If
        End If
    Next c
    On Error GoTo 0
End Sub

' ==============================================================================================
' MARCA EN COLOR VERDE UNICAMENTE LAS OPCIONES DE PORTAFOLIO QUE YA SE HAN SUBIDO PARA LA FECHA
' Seleccionada (Primer Piso -> BD_INSUMO_PRIMER_PISO, Tesoreria -> BD_INSUMO_TESORERIA,
' Segundo Piso -> BD_INSUMO_SEGUNDO_PISO; y Consolidado si las 3 ya estan subidas).
' ==============================================================================================
Public Sub ActualizarOpcionesSubidasPorFecha(ByRef frm As Object, ByVal fechaRaw As String)
    On Error Resume Next
    If frm Is Nothing Then Exit Sub
    Dim strFec As String
    Dim nP1 As Long, nP2 As Long, nP3 As Long
    Dim bP1 As Boolean, bP2 As Boolean, bP3 As Boolean, bCons As Boolean
    Dim c As Object, cName As String, cType As Long

    strFec = Trim(fechaRaw)
    bP1 = False: bP2 = False: bP3 = False: bCons = False
    nP1 = 0: nP2 = 0: nP3 = 0

    ' RENDIMIENTO: Solo consultar la BD cuando se tiene una fecha completa (dd/mm/aaaa)
    ' y obtener los 3 conteos en 1 sola llamada cacheada (evitando 12 consultas SQL al teclear o tabular).
    If Len(strFec) >= 8 Then
        If FondoBO.esFechaFormatoCorrectoBO(strFec) And Not FondoBO.esFechaFuturaBO(strFec) Then
            strFec = FondoBO.normalizarFechaBO(strFec)
            Call FondoBO.obtenerConteosTresTablasFechaBO(strFec, nP1, nP2, nP3)
            bP1 = (nP1 > 0)
            bP2 = (nP2 > 0)
            bP3 = (nP3 > 0)
            bCons = (bP1 And bP2 And bP3)
        End If
    End If

    For Each c In frm.Controls
        Err.Clear
        cName = UCase(Trim(CStr(c.Name)))
        cType = CLng(Nz(c.ControlType, 0))
        If cType = 100 Then
            If cName = "LBLOPTPRIMERPISO" Then
                If bP1 Then
                    c.BackStyle = 1
                    c.BackColor = 3433750 ' Verde institucional: opcion ya subida
                    c.ForeColor = 16777215
                    c.FontBold = True
                    c.ControlTipText = "Insumo de Primer Piso YA subido para " & strFec & " (" & nP1 & " registros en BD_INSUMO_PRIMER_PISO)"
                Else
                    c.BackStyle = 0
                    c.BackColor = 16119285
                    c.ForeColor = 3355443
                    c.FontBold = False
                    c.ControlTipText = "Sin insumo de Primer Piso subido para " & strFec
                End If
            ElseIf cName = "LBLOPTTESORERIA" Then
                If bP2 Then
                    c.BackStyle = 1
                    c.BackColor = 3433750 ' Verde institucional: opcion ya subida
                    c.ForeColor = 16777215
                    c.FontBold = True
                    c.ControlTipText = "Insumo de Tesorería YA subido para " & strFec & " (" & nP2 & " registros en BD_INSUMO_TESORERIA)"
                Else
                    c.BackStyle = 0
                    c.BackColor = 16119285
                    c.ForeColor = 3355443
                    c.FontBold = False
                    c.ControlTipText = "Sin insumo de Tesorería subido para " & strFec
                End If
            ElseIf cName = "LBLOPTSEGUNDOPISO" Then
                If bP3 Then
                    c.BackStyle = 1
                    c.BackColor = 3433750 ' Verde institucional: opcion ya subida
                    c.ForeColor = 16777215
                    c.FontBold = True
                    c.ControlTipText = "Insumo de Segundo Piso YA subido para " & strFec & " (" & nP3 & " registros en BD_INSUMO_SEGUNDO_PISO)"
                Else
                    c.BackStyle = 0
                    c.BackColor = 16119285
                    c.ForeColor = 3355443
                    c.FontBold = False
                    c.ControlTipText = "Sin insumo de Segundo Piso subido para " & strFec
                End If
            ElseIf cName = "LBLOPTCONSOLIDADO" Then
                If bCons Then
                    c.BackStyle = 1
                    c.BackColor = 3433750 ' Verde institucional: las 3 bases estan subidas
                    c.ForeColor = 16777215
                    c.FontBold = True
                    c.ControlTipText = "Los 3 portafolios están subidos para " & strFec & " (Consolidado disponible)"
                Else
                    c.BackStyle = 0
                    c.BackColor = 16119285
                    c.ForeColor = 3355443
                    c.FontBold = False
                    c.ControlTipText = "Requiere los 3 portafolios subidos para la fecha " & strFec
                End If
            ElseIf cName = "LBLOPTCOMPARATIVO2" Then
                If (bP1 And bP2) Then
                    c.BackStyle = 1
                    c.BackColor = 3433750 ' Verde institucional: Primer Piso y Tesoreria subidos
                    c.ForeColor = 16777215
                    c.FontBold = True
                    c.ControlTipText = "Primer Piso y Tesorería están subidos para " & strFec & " (Comparativo de 2 insumos disponible)"
                Else
                    c.BackStyle = 0
                    c.BackColor = 16119285
                    c.ForeColor = 3355443
                    c.FontBold = False
                    c.ControlTipText = "Requiere Primer Piso y Tesorería subidos para la fecha " & strFec
                End If
            ElseIf cName = "LBLOPTCOMPARATIVO3" Then
                If bCons Then
                    c.BackStyle = 1
                    c.BackColor = 3433750 ' Verde institucional: los 3 libros subidos
                    c.ForeColor = 16777215
                    c.FontBold = True
                    c.ControlTipText = "Los 3 portafolios están subidos para " & strFec & " (Comparativo de 3 insumos disponible)"
                Else
                    c.BackStyle = 0
                    c.BackColor = 16119285
                    c.ForeColor = 3355443
                    c.FontBold = False
                    c.ControlTipText = "Requiere los 3 portafolios subidos para la fecha " & strFec
                End If
            End If
        End If
    Next c
    On Error GoTo 0
End Sub

' =============================================================================
' VER LA CONTRASENA SOLO MIENTRAS SE SOSTIENE EL CLIC
' -----------------------------------------------------------------------------
' Antes el boton "Ver" alternaba con un clic, de modo que la contrasena se
' quedaba visible en pantalla hasta que alguien volvia a pulsarlo. Ahora se
' descubre en el MouseDown y se vuelve a ocultar en el MouseUp: en cuanto se
' suelta el boton, la contrasena desaparece.
'
' Detalle importante de Access: NO se puede cambiar la propiedad InputMask del
' control que tiene el foco. Por eso la secuencia es siempre
'   1) mover el foco al boton, 2) cambiar la mascara, 3) reponer el valor,
'   4) devolver el foco al cuadro de texto con el cursor al final.
'
' Se usa frm.Controls("nombre") y no Me.nombre porque este procedimiento lo
' comparten varios formularios y los nombres de control no existen en todos;
' Me.nombre se resuelve al compilar y Controls("nombre") al ejecutar.
' =============================================================================
Public Sub VerContrasenaSostenidaUT(ByRef frm As Object, ByVal nombreBoton As String, _
        ByVal nombreTexto As String, ByVal mostrar As Boolean, _
        Optional ByVal textoReal As String = "")
    Dim txt As Object, v As String
    On Error Resume Next
    Set txt = frm.Controls(nombreTexto)
    If txt Is Nothing Then Exit Sub

    v = textoReal
    If v = "" Then v = Trim(CStr(Nz(txt.Value, "")))

    frm.Controls(nombreBoton).SetFocus
    If mostrar Then
        frm.Controls(nombreBoton).BackColor = 15523798 ' #D6DFEC Resaltado activo al sostener clic
        frm.Controls(nombreBoton).BorderColor = 5000704 ' #004E4C Borde verde institucional activo
        txt.InputMask = ""
    Else
        frm.Controls(nombreBoton).BackColor = 15458795 ' #EBF0F7 Fondo suave en reposo
        frm.Controls(nombreBoton).BorderColor = 11839396
        txt.InputMask = "Password"
    End If
    txt.Value = v
    txt.SetFocus
    txt.SelStart = Len(v)
    Err.Clear
End Sub

' Vuelve a ocultar TODAS las contrasenas de un formulario. Se llama desde el
' MouseUp del propio formulario como seguro: si el usuario presiona el ojito y
' arrastra el raton fuera del boton antes de soltar, el MouseUp del boton puede
' no llegar, y sin este seguro la contrasena se quedaria visible.
Public Sub OcultarTodasLasContrasenasUT(ByRef frm As Object)
    Dim nombres As Variant, i As Integer, c As Object, v As String
    On Error Resume Next
    nombres = Split("txtContrasena|txtNuevaContrasena|txtConfirmarContrasena|" & _
                    "txtPassword|txtNueva|txtConfirmar", "|")
    For i = LBound(nombres) To UBound(nombres)
        Set c = Nothing
        Set c = frm.Controls(CStr(nombres(i)))
        If Not c Is Nothing Then
            If c.InputMask <> "Password" Then
                v = Trim(CStr(Nz(c.Value, "")))
                c.InputMask = "Password"
                c.Value = v
            End If
        End If
        Err.Clear
    Next i
    Err.Clear
End Sub

' =============================================================================
' ENCABEZADO DE SESIÓN Y CONTROL DE ACCESO  (CAPA DE PRESENTACIÓN)
' -----------------------------------------------------------------------------
' Este procedimiento NO decide nada: pregunta a la capa de Negocio y obedece.
'   - el texto de las dos etiquetas lo arma FondoBO (etiquetaUsuarioSesionBO,
'     etiquetaPerfilSesionBO);
'   - quién puede entrar lo decide FondoBO.autorizarPantallaBO;
'   - aquí solo se pinta, se avisa y se navega, que es lo que le toca a la
'     Presentación según el Marco de Desarrollo.
'
' Sustituye a 14 copias del mismo bloque repartidas en los Form_Load. Cada
' formulario pasa de 14 líneas a una sola:
'
'     If Not Utilities.AplicarEncabezadoSesionUT(Me, FondoBO.PANT_ALTA_USUARIOS) Then Exit Sub
'
' Se usa frm.Controls("nombre") y no Me.nombre porque lo comparten formularios
' distintos: Me.nombre se resuelve al compilar y Controls("nombre") al ejecutar.
'
' Devuelve False cuando el acceso fue denegado; en ese caso ya mostró el aviso,
' abrió el menú y cerró el formulario, así que el llamador solo tiene que salir.
' =============================================================================
Public Function AplicarEncabezadoSesionUT(ByRef frm As Object, _
        Optional ByVal clavePantalla As String = "") As Boolean
    Dim msgDenegado As String, nombreFrm As String

    On Error Resume Next
    AplicarEncabezadoSesionUT = True
    nombreFrm = frm.Name
    Err.Clear

    ' 1. Pintar el encabezado. Si el formulario no tiene alguna de las dos
    '    etiquetas, la asignación falla en silencio y no pasa nada.
    frm.Controls("lblNoEmpleado").Caption = FondoBO.etiquetaUsuarioSesionBO()
    Err.Clear
    frm.Controls("lblPerfil").Caption = FondoBO.etiquetaPerfilSesionBO()
    Err.Clear

    ' 2. Preguntar si esta pantalla puede abrirse. Sin clave, no hay restricción.
    If Trim(clavePantalla) = "" Then Exit Function
    If FondoBO.autorizarPantallaBO(clavePantalla, msgDenegado) Then Exit Function

    ' 3. Acceso denegado: avisar, regresar al menú y cerrar.
    AplicarEncabezadoSesionUT = False
    MsgBox msgDenegado, vbCritical, "Control de Acceso Hidrocarburos"
    DoCmd.OpenForm "Form_Menu_General", acNormal
    DoCmd.Close acForm, nombreFrm
    Err.Clear
End Function

' Variante para los botones que llevan a una pantalla restringida: avisa y
' devuelve False, pero NO cierra el formulario donde está el botón. La usan los
' botones de Gestionar Instrumentos, que antes evaluaban el perfil ellos mismos.
Public Function PuedeAbrirPantallaUT(ByVal clavePantalla As String) As Boolean
    Dim msgDenegado As String
    On Error Resume Next
    PuedeAbrirPantallaUT = FondoBO.autorizarPantallaBO(clavePantalla, msgDenegado)
    If Not PuedeAbrirPantallaUT Then
        MsgBox msgDenegado, vbCritical, "Control de Acceso Hidrocarburos"
    End If
    Err.Clear
End Function
