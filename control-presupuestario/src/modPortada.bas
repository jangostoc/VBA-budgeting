Attribute VB_Name = "modPortada"
Option Explicit

' =====================================================================
'  modPortada  (libro de Control presupuestario)
'  Crea (o vuelve a crear) la hoja "Portada" al principio del libro con:
'    1. Cabecera con título, periodo y fecha de actualización
'    2. Cifras clave del acumulado (se calculan de Presupuesto y Real)
'    3. Índice con un enlace a cada pestaña y qué contiene
'    4. Instrucciones de uso
'  Además pone un botón "« Portada" en cada hoja para volver.
'
'  Usa:  SumaHasta (modCuadroMando), MesDelInforme (modConclusiones),
'        UMBRAL_VIGILAR (modPresupuesto)
'
'  Macro que hay que ejecutar:  CrearPortada
' =====================================================================

Private Const NOMBRE_PORTADA As String = "Portada"
Private Const FUENTE As String = "Calibri"
Private Const BOTON_VOLVER As String = "btnVolverPortada"

' Colores (Const no admite RGB(), así que se rellenan en IniciarColores)
Private colAzulOscuro As Long
Private colAcento As Long
Private colFondoClaro As Long
Private colTexto As Long
Private colGris As Long
Private colLinea As Long
Private colSubtitulo As Long

' ---------------------------------------------------------------------
'  MACRO PRINCIPAL
' ---------------------------------------------------------------------
Public Sub CrearPortada()
    Dim ws As Worksheet
    Dim fila As Long

    Application.ScreenUpdating = False
    IniciarColores

    ' 1. Si ya hay portada, se borra para crearla limpia
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(NOMBRE_PORTADA).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' 2. Nueva hoja en PRIMERA posición
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1))
    ws.Name = NOMBRE_PORTADA
    ws.Tab.Color = colAzulOscuro

    ' 3. Lienzo: fuente, anchos de columna, sin cuadrícula
    PrepararLienzo ws

    ' 4. Bloques de la portada (cada uno avanza la variable "fila")
    fila = 2
    EscribirCabecera ws, fila
    EscribirCifrasClave ws, fila
    EscribirIndice ws, fila
    EscribirGuia ws, fila
    EscribirPie ws, fila
    AnadirBotonActualizar ws

    ' 5. Botón "« Portada" en el resto de hojas
    AnadirBotonesVolver

    ' 6. Dejar la portada a la vista, arriba del todo
    ws.Activate
    ws.Range("A1").Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Application.ScreenUpdating = True
End Sub

' ---------------------------------------------------------------------
'  PREPARACIÓN
' ---------------------------------------------------------------------
Private Sub IniciarColores()
    colAzulOscuro = RGB(31, 56, 100)
    colAcento = RGB(47, 117, 181)
    colFondoClaro = RGB(242, 246, 252)
    colTexto = RGB(38, 38, 38)
    colGris = RGB(110, 110, 110)
    colLinea = RGB(217, 225, 242)
    colSubtitulo = RGB(189, 215, 238)
End Sub

Private Sub PrepararLienzo(ws As Worksheet)
    With ws.Cells
        .Font.Name = FUENTE
        .Font.Size = 11
        .Font.Color = colTexto
        .VerticalAlignment = xlCenter
    End With

    ' A = margen | B = nº | C = hoja | D:E = descripción | F = macro | G = margen
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 6
    ws.Columns("C:F").ColumnWidth = 26
    ws.Columns("G").ColumnWidth = 2

    ws.Activate                             ' La cuadrícula es de la ventana, no de la hoja
    With ActiveWindow
        .DisplayGridlines = False
        .DisplayHeadings = False
        .Zoom = 100
    End With
End Sub

' ---------------------------------------------------------------------
'  BLOQUE 1: CABECERA
' ---------------------------------------------------------------------
Private Sub EscribirCabecera(ws As Worksheet, ByRef fila As Long)
    Dim periodo As String

    periodo = PeriodoInforme()

    ' Banda azul (5 filas) + línea de acento debajo
    ws.Range(ws.Cells(fila, 2), ws.Cells(fila + 4, 6)).Interior.Color = colAzulOscuro
    ws.Range(ws.Cells(fila + 5, 2), ws.Cells(fila + 5, 6)).Interior.Color = colAcento

    ws.Rows(fila).RowHeight = 10
    ws.Rows(fila + 1).RowHeight = 44
    ws.Rows(fila + 2).RowHeight = 22
    ws.Rows(fila + 3).RowHeight = 20
    ws.Rows(fila + 4).RowHeight = 10
    ws.Rows(fila + 5).RowHeight = 4

    With ws.Cells(fila + 1, 2)
        .Value = "CONTROL PRESUPUESTARIO " & Year(Date)
        .Font.Size = 26
        .Font.Bold = True
        .Font.Color = vbWhite
    End With

    With ws.Cells(fila + 2, 2)
        If periodo = "" Then
            .Value = "Real frente a presupuesto · sin datos todavía (ejecuta CrearDatosPresupuesto)"
        Else
            .Value = "Real frente a presupuesto, cuadro de mando e informe de gestión · acumulado " & periodo
        End If
        .Font.Size = 13
        .Font.Color = colSubtitulo
    End With

    With ws.Cells(fila + 3, 2)
        .Value = "Actualizado el " & Format(Now, "dd/mm/yyyy hh:nn") & _
                 "   ·   " & (ThisWorkbook.Worksheets.Count - 1) & " hojas en el libro"
        .Font.Size = 10
        .Font.Color = colSubtitulo
    End With

    fila = fila + 7
End Sub

' ---------------------------------------------------------------------
'  BLOQUE 2: CIFRAS CLAVE (4 tarjetas en C, D, E y F)
' ---------------------------------------------------------------------
Private Sub EscribirCifrasClave(ws As Worksheet, ByRef fila As Long)
    Dim wsP As Worksheet, wsR As Worksheet
    Dim etiquetas As Variant, valores As Variant, subValores As Variant
    Dim fmtValor As Variant, fmtSub As Variant
    Dim mes As Long, i As Long, col As Long, ultimaFila As Long
    Dim vR As Double, vP As Double, aR As Double, aP As Double
    Dim resR As Double, resP As Double, resAnualP As Double
    Dim mbR As Double, mbP As Double, prevision As Double
    Dim fmtEur As String

    ' Datos: si no existen las hojas, mes se queda en 0 y se muestra "n/d"
    Set wsP = ObtenerHoja("Presupuesto")
    Set wsR = ObtenerHoja("Real")
    If Not wsP Is Nothing And Not wsR Is Nothing Then mes = MesDelInforme(wsP, wsR)

    If mes > 0 Then
        ultimaFila = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row
        vR = SumaHasta(wsR, 2, mes)
        vP = SumaHasta(wsP, 2, mes)
        aR = SumaHasta(wsR, 3, mes)
        aP = SumaHasta(wsP, 3, mes)
        For i = 2 To ultimaFila
            resR = resR + SumaHasta(wsR, i, mes)
            resP = resP + SumaHasta(wsP, i, mes)
            resAnualP = resAnualP + SumaHasta(wsP, i, 12)
        Next i
        mbR = (vR + aR) / vR
        mbP = (vP + aP) / vP
        prevision = resR + (resAnualP - resP)      ' Real acumulado + presupuesto pendiente

        valores = Array(vR, mbR, resR, prevision)
        subValores = Array(vR / vP - 1, (mbR - mbP) * 100, (resR - resP) / Abs(resP), prevision / resAnualP - 1)
    End If

    fmtEur = "#,##0 ""€"";[Red]-#,##0 ""€"""
    etiquetas = Array("Ventas acumuladas", "Margen bruto", "Resultado acumulado", "Previsión de cierre")
    fmtValor = Array(fmtEur, "0.0%", fmtEur, fmtEur)
    fmtSub = Array(FormatoVariacion("vs ppto"), _
                   "[Color10]+0.0 ""p.p. vs ppto"";[Red]-0.0 ""p.p. vs ppto"";0.0 ""p.p. vs ppto""", _
                   FormatoVariacion("vs ppto"), FormatoVariacion("vs ppto anual"))

    If mes > 0 Then
        EscribirSeccion ws, fila, "Cifras clave · acumulado " & PeriodoInforme()
    Else
        EscribirSeccion ws, fila, "Cifras clave"
    End If

    ws.Rows(fila).RowHeight = 22
    ws.Rows(fila + 1).RowHeight = 36
    ws.Rows(fila + 2).RowHeight = 22

    For i = 0 To UBound(etiquetas)
        col = 3 + i                         ' Columnas C, D, E y F

        ' Fondo de la tarjeta, borde izquierdo de color y hueco blanco a la derecha
        With ws.Range(ws.Cells(fila, col), ws.Cells(fila + 2, col))
            .Interior.Color = colFondoClaro
            .HorizontalAlignment = xlLeft
            .IndentLevel = 1
            With .Borders(xlEdgeLeft)
                .Color = colAcento
                .Weight = xlThick
            End With
            With .Borders(xlEdgeRight)
                .Color = vbWhite
                .Weight = xlThick
            End With
        End With

        ' Fila 1: etiqueta
        With ws.Cells(fila, col)
            .Value = UCase(etiquetas(i))
            .Font.Size = 9
            .Font.Bold = True
            .Font.Color = colGris
        End With

        ' Fila 2: importe | Fila 3: comparación con el presupuesto
        With ws.Cells(fila + 1, col)
            .Font.Size = 20
            .Font.Bold = True
            .Font.Color = colAzulOscuro
            If mes = 0 Then
                .Value = "n/d"
            Else
                .Value = valores(i)
                .NumberFormat = fmtValor(i)
            End If
        End With

        With ws.Cells(fila + 2, col)
            .Font.Size = 10
            If mes = 0 Then
                .Value = "Ejecuta CrearDatosPresupuesto"
                .Font.Color = colGris
            Else
                .Value = subValores(i)
                .NumberFormat = fmtSub(i)
            End If
        End With
    Next i

    fila = fila + 4
End Sub

' Formato de número para una variación: verde si sube, rojo si baja, con texto detrás
Private Function FormatoVariacion(texto As String) As String
    FormatoVariacion = "[Color10]+0.0% """ & texto & """;[Red]-0.0% """ & texto & """;0.0% """ & texto & """"
End Function

' ---------------------------------------------------------------------
'  BLOQUE 3: ÍNDICE CON ENLACES
' ---------------------------------------------------------------------
Private Sub EscribirIndice(ws As Worksheet, ByRef fila As Long)
    Dim info As Object                      ' Diccionario: hoja -> Array(descripción, macro)
    Dim orden As Collection
    Dim nombre As Variant
    Dim sh As Worksheet
    Dim n As Long

    Set info = DescripcionesHojas()
    Set orden = New Collection

    ' 1º las hojas conocidas, en el orden lógico de lectura
    For Each nombre In info.Keys
        Set sh = ObtenerHoja(CStr(nombre))
        If Not sh Is Nothing Then
            If sh.Visible = xlSheetVisible Then orden.Add sh.Name
        End If
    Next nombre

    ' 2º cualquier otra hoja visible que no esté en la lista
    For Each sh In ThisWorkbook.Worksheets
        If sh.Name <> NOMBRE_PORTADA And sh.Visible = xlSheetVisible Then
            If Not info.Exists(sh.Name) Then orden.Add sh.Name
        End If
    Next sh

    EscribirSeccion ws, fila, "Contenido del libro"

    ws.Cells(fila, 2).Value = "Nº"
    ws.Cells(fila, 3).Value = "HOJA (clic para ir)"
    ws.Cells(fila, 4).Value = "QUÉ CONTIENE"
    ws.Cells(fila, 6).Value = "MACRO QUE LA GENERA"
    With ws.Range(ws.Cells(fila, 2), ws.Cells(fila, 6))
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = colGris
    End With
    ws.Cells(fila, 2).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(fila, 3), ws.Cells(fila, 6)).IndentLevel = 1
    fila = fila + 1

    For Each nombre In orden
        n = n + 1
        EscribirFilaIndice ws, fila, n, ThisWorkbook.Worksheets(CStr(nombre)), info
        fila = fila + 1
    Next nombre

    fila = fila + 1
End Sub

Private Sub EscribirFilaIndice(ws As Worksheet, ByVal fila As Long, ByVal n As Long, _
                               sh As Worksheet, info As Object)
    Dim descripcion As String, nombreMacro As String
    Dim colorHoja As Long

    If info.Exists(sh.Name) Then
        descripcion = info(sh.Name)(0)
        nombreMacro = info(sh.Name)(1)
    Else
        descripcion = "Hoja de trabajo adicional del libro."
        nombreMacro = "-"
    End If

    If sh.Tab.ColorIndex = xlColorIndexNone Then
        colorHoja = colAcento
    Else
        colorHoja = sh.Tab.Color
    End If

    ws.Rows(fila).RowHeight = 36

    With ws.Range(ws.Cells(fila, 2), ws.Cells(fila, 6))
        If n Mod 2 = 1 Then .Interior.Color = colFondoClaro
        With .Borders(xlEdgeBottom)
            .Color = colLinea
            .Weight = xlThin
        End With
    End With

    With ws.Cells(fila, 2)
        .Value = n
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = colAzulOscuro
        With .Borders(xlEdgeLeft)
            .Color = colorHoja
            .Weight = xlThick
        End With
    End With

    ' Enlace a la hoja (las comillas simples permiten nombres con espacios)
    ws.Hyperlinks.Add Anchor:=ws.Cells(fila, 3), Address:="", _
        SubAddress:="'" & Replace(sh.Name, "'", "''") & "'!A1", _
        ScreenTip:="Ir a la hoja " & sh.Name, TextToDisplay:=sh.Name
    With ws.Cells(fila, 3)
        .IndentLevel = 1
        .Font.Name = FUENTE
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = colAcento
        .Font.Underline = xlUnderlineStyleNone
    End With

    ws.Cells(fila, 4).Value = descripcion
    With ws.Range(ws.Cells(fila, 4), ws.Cells(fila, 5))
        .Merge
        .WrapText = True
        .IndentLevel = 1
        .Font.Size = 10
    End With

    With ws.Cells(fila, 6)
        .Value = nombreMacro
        .IndentLevel = 1
        .Font.Name = "Consolas"
        .Font.Size = 10
        .Font.Color = colGris
    End With
End Sub

' Aquí se explica cada hoja, en el orden en que conviene leerlas.
' Si añades una hoja nueva, añade aquí su línea.
Private Function DescripcionesHojas() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    d.Add "CuadroMando", Array("4 KPI con semáforo y 3 gráficos. Elige el mes en la celda C3 " & _
        "y pulsa Actualizar.", "CrearCuadroMando")
    d.Add "Conclusiones", Array("Informe de gestión redactado automáticamente: resumen, puente " & _
        "del resultado, análisis por partida, previsión de cierre y acciones.", "GenerarConclusionesPpto")
    d.Add "Desviaciones", Array("Real vs presupuesto del mes y del acumulado por epígrafe, " & _
        "con semáforo y gráfico de ventas.", "AnalisisDesviaciones")
    d.Add "Presupuesto", Array("Presupuesto mensual aprobado (datos de partida): epígrafes en " & _
        "filas y meses en columnas.", "CrearDatosPresupuesto")
    d.Add "Real", Array("Importes reales de los meses cerrados, con la misma estructura " & _
        "que el presupuesto.", "CrearDatosPresupuesto")
    d.Add "DatosCuadro", Array("Hoja auxiliar con las series que alimentan los gráficos del " & _
        "cuadro de mando. No se edita a mano.", "ActualizarCuadro")

    Set DescripcionesHojas = d
End Function

' ---------------------------------------------------------------------
'  BLOQUE 4: INSTRUCCIONES DE USO
' ---------------------------------------------------------------------
Private Sub EscribirGuia(ws As Worksheet, ByRef fila As Long)
    Dim lineas As Variant
    Dim i As Long

    lineas = Array( _
        "Haz clic en el nombre de una hoja del índice para ir a ella. En cada hoja, el botón " & _
            "«Portada» te trae de vuelta.", _
        "Primera vez, por este orden:  CrearDatosPresupuesto  >  CrearCuadroMando  >  " & _
            "AnalisisDesviaciones  >  GenerarConclusionesPpto  >  CrearPortada.", _
        "Cierre de cada mes: elige el mes en la celda C3 del CuadroMando y pulsa Actualizar; " & _
            "después ejecuta GenerarConclusionesPpto y pulsa «Actualizar portada».", _
        "Para usar datos propios, pega los importes en Presupuesto y Real con la misma estructura " & _
            "(ingresos en positivo, gastos en negativo, meses sin cerrar vacíos) y no vuelvas a " & _
            "ejecutar CrearDatosPresupuesto, que regenera los datos de ejemplo.", _
        "El semáforo marca en ámbar las desviaciones desfavorables de hasta un " & _
            Format(UMBRAL_VIGILAR, "0%") & " y en rojo las mayores. El umbral se cambia en la " & _
            "constante UMBRAL_VIGILAR del módulo modPresupuesto.", _
        "Las macros se lanzan con Alt + F8 (o Programador > Macros) y el código se ve con Alt + F11.")

    EscribirSeccion ws, fila, "Instrucciones de uso"

    For i = 0 To UBound(lineas)
        With ws.Cells(fila, 2)
            .Value = i + 1
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlTop
            .Font.Bold = True
            .Font.Color = colAcento
        End With

        ws.Cells(fila, 3).Value = lineas(i)
        With ws.Range(ws.Cells(fila, 3), ws.Cells(fila, 6))
            .Merge
            .WrapText = True
            .VerticalAlignment = xlTop
            .Font.Size = 10
        End With

        ' Las celdas combinadas no se autoajustan: calculamos el alto a mano
        ws.Rows(fila).RowHeight = 15 * (Int(Len(lineas(i)) / 115) + 1) + 4
        fila = fila + 1
    Next i

    fila = fila + 1
End Sub

Private Sub EscribirPie(ws As Worksheet, ByRef fila As Long)
    With ws.Range(ws.Cells(fila, 2), ws.Cells(fila, 6)).Borders(xlEdgeTop)
        .Color = colLinea
        .Weight = xlThin
    End With
    With ws.Cells(fila, 2)
        .Value = "Portada generada con la macro CrearPortada (módulo modPortada)."
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = colGris
    End With
    fila = fila + 1
End Sub

' ---------------------------------------------------------------------
'  BOTONES
' ---------------------------------------------------------------------
Private Sub AnadirBotonActualizar(ws As Worksheet)
    Dim shp As Shape
    Dim ref As Range

    Set ref = ws.Cells(3, 6)
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        ref.Left + ref.Width - 130, ref.Top + (ref.Height - 26) / 2, 125, 26)

    FormatearBoton shp, "Actualizar portada", vbWhite, colAzulOscuro
    shp.OnAction = "CrearPortada"
End Sub

Private Sub AnadirBotonesVolver()
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        If sh.Name <> NOMBRE_PORTADA And sh.Visible = xlSheetVisible Then
            PonerBotonVolver sh
        End If
    Next sh
End Sub

Private Sub PonerBotonVolver(sh As Worksheet)
    Dim shp As Shape
    Dim colLibre As Long

    On Error GoTo Salir                     ' Ej.: hoja protegida -> la saltamos

    On Error Resume Next
    sh.Shapes(BOTON_VOLVER).Delete
    On Error GoTo Salir

    ' Una columna a la derecha de los datos, para no tapar nada
    colLibre = sh.UsedRange.Column + sh.UsedRange.Columns.Count + 1
    Set shp = sh.Shapes.AddShape(msoShapeRoundedRectangle, _
        sh.Cells(1, colLibre).Left, sh.Cells(1, 1).Top + 2, 90, 22)
    shp.Name = BOTON_VOLVER

    FormatearBoton shp, "« Portada", colAzulOscuro, vbWhite

    ' Hipervínculo en la forma: funciona aunque las macros estén desactivadas
    sh.Hyperlinks.Add Anchor:=shp, Address:="", _
        SubAddress:="'" & NOMBRE_PORTADA & "'!A1", ScreenTip:="Volver a la portada"
Salir:
End Sub

Private Sub FormatearBoton(shp As Shape, texto As String, colorFondo As Long, colorTexto As Long)
    shp.Placement = xlFreeFloating
    shp.Fill.ForeColor.RGB = colorFondo
    shp.Line.Visible = msoFalse
    With shp.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 4
        .MarginRight = 4
        With .TextRange
            .Text = texto
            .ParagraphFormat.Alignment = msoAlignCenter
            .Font.Name = FUENTE
            .Font.Size = 10
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = colorTexto
        End With
    End With
End Sub

' ---------------------------------------------------------------------
'  UTILIDADES
' ---------------------------------------------------------------------
Private Sub EscribirSeccion(ws As Worksheet, ByRef fila As Long, texto As String)
    With ws.Cells(fila, 2)
        .Value = UCase(texto)
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = colAzulOscuro
    End With
    With ws.Range(ws.Cells(fila, 2), ws.Cells(fila, 6)).Borders(xlEdgeBottom)
        .Color = colAcento
        .Weight = xlMedium
    End With
    ws.Rows(fila).RowHeight = 24
    ws.Rows(fila + 1).RowHeight = 8
    fila = fila + 2
End Sub

' Devuelve la hoja o Nothing si no existe (sin dar error)
Private Function ObtenerHoja(nombre As String) As Worksheet
    On Error Resume Next
    Set ObtenerHoja = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0
End Function

' Texto "Ene - Sep" del periodo analizado, o "" si no hay datos
Private Function PeriodoInforme() As String
    Dim wsP As Worksheet, wsR As Worksheet
    Dim mes As Long

    Set wsP = ObtenerHoja("Presupuesto")
    Set wsR = ObtenerHoja("Real")
    If wsP Is Nothing Or wsR Is Nothing Then Exit Function

    mes = MesDelInforme(wsP, wsR)
    If mes > 0 Then PeriodoInforme = wsP.Cells(1, 2).Value & " - " & wsP.Cells(1, mes + 1).Value
End Function
