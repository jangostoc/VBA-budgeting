Attribute VB_Name = "Mod5Portada"
Option Explicit

' =====================================================================
'  Mod5Portada
'  Crea (o vuelve a crear) la hoja "Portada" al principio del libro con:
'    1. Cabecera con título, subtítulo y fecha de actualización
'    2. Cifras clave del ejercicio (se leen de la hoja Analisis)
'    3. Índice con un enlace a cada pestaña y qué contiene
'    4. Guía rápida de uso
'  Además pone un botón "« Portada" en cada hoja para volver.
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

    Application.ScreenUpdating = False      ' Que no parpadee la pantalla
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
    Dim anio As Long
    anio = AnioActual()

    ' Banda azul (5 filas) + línea de acento debajo
    ws.Range(ws.Cells(fila, 2), ws.Cells(fila + 4, 6)).Interior.Color = colAzulOscuro
    ws.Range(ws.Cells(fila + 5, 2), ws.Cells(fila + 5, 6)).Interior.Color = colAcento

    ws.Rows(fila).RowHeight = 10            ' Relleno superior
    ws.Rows(fila + 1).RowHeight = 44
    ws.Rows(fila + 2).RowHeight = 22
    ws.Rows(fila + 3).RowHeight = 20
    ws.Rows(fila + 4).RowHeight = 10        ' Relleno inferior
    ws.Rows(fila + 5).RowHeight = 4         ' Línea de acento

    ' El texto se escribe en B y "desborda" hacia la derecha (no hace falta combinar)
    With ws.Cells(fila + 1, 2)
        .Value = "CUENTAS ANUALES " & anio
        .Font.Size = 26
        .Font.Bold = True
        .Font.Color = vbWhite
    End With

    With ws.Cells(fila + 2, 2)
        .Value = "Balance, Pérdidas y Ganancias y análisis comparativo " & anio & " - " & (anio - 1)
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
    Dim etiquetas As Variant, buscar As Variant
    Dim wsA As Worksheet, celda As Range
    Dim i As Long, col As Long, anioAnt As Long
    Dim fmtVar As String

    anioAnt = AnioActual() - 1
    fmtVar = "[Color10]+0.0% ""vs " & anioAnt & """;[Red]-0.0% ""vs " & anioAnt & """;0.0% ""vs " & anioAnt & """"

    EscribirSeccion ws, fila, "Cifras clave del ejercicio"

    etiquetas = Array("Cifra de negocios", "Resultado del ejercicio", "Total activo", "Patrimonio neto")
    ' Texto que se busca en la columna A de la hoja Analisis
    buscar = Array("Importe neto de la cifra de negocios", "RESULTADO DEL EJERCICIO", "TOTAL ACTIVO", "A) PATRIMONIO NETO")

    Set wsA = ObtenerHoja("Analisis")

    ws.Rows(fila).RowHeight = 22
    ws.Rows(fila + 1).RowHeight = 36
    ws.Rows(fila + 2).RowHeight = 22

    For i = 0 To UBound(etiquetas)
        col = 3 + i                         ' Columnas C, D, E y F

        ' Fondo de la tarjeta, borde izquierdo de color y hueco blanco a la derecha
        With ws.Range(ws.Cells(fila, col), ws.Cells(fila + 2, col))
            .Interior.Color = colFondoClaro
            .HorizontalAlignment = xlLeft     ' Primero la alineación, luego la sangría
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

        ' Buscamos el epígrafe (MatchCase para no confundir mayúsculas/minúsculas)
        Set celda = Nothing
        If Not wsA Is Nothing Then
            Set celda = wsA.Columns(1).Find(What:=buscar(i), LookIn:=xlValues, _
                                            LookAt:=xlPart, MatchCase:=True)
        End If

        ' Fila 2: importe | Fila 3: variación respecto al año anterior
        With ws.Cells(fila + 1, col)
            .Font.Size = 20
            .Font.Bold = True
            .Font.Color = colAzulOscuro
            If celda Is Nothing Then
                .Value = "n/d"
            Else
                .Value = celda.Offset(0, 1).Value             ' Columna B = año actual
                .NumberFormat = "#,##0 ""€"";[Red]-#,##0 ""€"""
            End If
        End With

        With ws.Cells(fila + 2, col)
            .Font.Size = 10
            If celda Is Nothing Then
                .Value = "Ejecuta AnalisisComparativo"
                .Font.Color = colGris
            Else
                .Value = celda.Offset(0, 4).Value             ' Columna E = Var. %
                .NumberFormat = fmtVar
            End If
        End With
    Next i

    fila = fila + 4
End Sub

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

    ' 1º las hojas conocidas, en el orden lógico de trabajo
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

    ' Cabecera de la tabla
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

    ' Una fila por hoja
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

    ' Color de la pestaña (si no tiene, el azul de acento)
    If sh.Tab.ColorIndex = xlColorIndexNone Then
        colorHoja = colAcento
    Else
        colorHoja = sh.Tab.Color
    End If

    ws.Rows(fila).RowHeight = 36

    With ws.Range(ws.Cells(fila, 2), ws.Cells(fila, 6))
        If n Mod 2 = 1 Then .Interior.Color = colFondoClaro     ' Filas alternas
        With .Borders(xlEdgeBottom)
            .Color = colLinea
            .Weight = xlThin
        End With
    End With

    ' Nº con una barra del color de la pestaña
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
    With ws.Cells(fila, 3)                  ' El hipervínculo cambia el estilo: lo rehacemos
        .IndentLevel = 1
        .Font.Name = FUENTE
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = colAcento
        .Font.Underline = xlUnderlineStyleNone
    End With

    ' Descripción en D:E (primero el texto, luego combinar)
    ws.Cells(fila, 4).Value = descripcion
    With ws.Range(ws.Cells(fila, 4), ws.Cells(fila, 5))
        .Merge
        .WrapText = True
        .IndentLevel = 1
        .Font.Size = 10
    End With

    ' Macro de origen, en letra de "código"
    With ws.Cells(fila, 6)
        .Value = nombreMacro
        .IndentLevel = 1
        .Font.Name = "Consolas"
        .Font.Size = 10
        .Font.Color = colGris
    End With
End Sub

' Aquí se explica cada hoja. El ORDEN es el orden lógico de trabajo.
' Si añades una hoja nueva, añade aquí su línea.
Private Function DescripcionesHojas() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare           ' "analisis" = "Analisis"

    d.Add "SumasSaldos", Array("Balance de sumas y saldos " & AnioActual() & _
        " (datos de partida): cuenta, descripción y saldo.", "CrearDatosEjemplo")
    d.Add "Resumen", Array("Saldos agrupados por los 7 grupos del PGC, con una " & _
        "comprobación de cuadre que debe dar 0.", "ResumenPorGrupos")
    d.Add "Mapeo", Array("Tabla puente: asigna cada cuenta a su masa (ANC, AC, PN, " & _
        "PNC, PC, PyG) y a su epígrafe de las cuentas anuales.", "CrearMapeo")
    d.Add "CuentasAnuales", Array("Balance de situación y Cuenta de pérdidas y " & _
        "ganancias, construidos a partir del mapeo.", "GenerarCuentasAnuales")
    d.Add "SumasSaldos2024", Array("Sumas y saldos del ejercicio anterior, " & _
        "base para la comparación.", "CrearDatos2024")
    d.Add "Analisis", Array("Comparativa entre ejercicios por epígrafe (variación en € y %, " & _
        "peso sobre el total) y ratios de liquidez, solvencia y rentabilidad.", "AnalisisComparativo")
    d.Add "Conclusiones", Array("Interpretación de los resultados: evolución del margen, " & _
        "estructura de gastos y conclusiones finales.", "GenerarConclusiones")

    Set DescripcionesHojas = d
End Function

' ---------------------------------------------------------------------
'  BLOQUE 4: GUÍA DE USO
' ---------------------------------------------------------------------
Private Sub EscribirGuia(ws As Worksheet, ByRef fila As Long)
    Dim lineas As Variant
    Dim flujo As String
    Dim info As Object, clave As Variant
    Dim i As Long

    ' El orden de las macros se construye a partir del diccionario de descripciones
    Set info = DescripcionesHojas()
    For Each clave In info.Keys
        If flujo <> "" Then flujo = flujo & "  >  "
        flujo = flujo & info(clave)(1)
    Next clave

    lineas = Array( _
        "Haz clic en el nombre de una hoja del índice para ir directamente a ella.", _
        "En cada hoja tienes un botón «Portada» arriba a la derecha para volver aquí.", _
        "Orden de ejecución de las macros:  " & flujo & ".", _
        "Las macros se lanzan con Alt + F8 (o en Programador > Macros) y el código se ve con Alt + F11.", _
        "Si regeneras alguna hoja, pulsa «Actualizar portada» (o ejecuta CrearPortada) " & _
        "para refrescar cifras, enlaces y botones.")

    EscribirSeccion ws, fila, "Cómo usar este libro"

    For i = 0 To UBound(lineas)
        With ws.Cells(fila, 2)
            .Value = i + 1
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
            .Font.Color = colAcento
        End With

        ws.Cells(fila, 3).Value = lineas(i)
        With ws.Range(ws.Cells(fila, 3), ws.Cells(fila, 6))
            .Merge
            .WrapText = True
            .Font.Size = 10
        End With

        ' Las líneas largas necesitan dos renglones
        If Len(lineas(i)) > 110 Then
            ws.Rows(fila).RowHeight = 32
        Else
            ws.Rows(fila).RowHeight = 20
        End If
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
        .Value = "Portada generada con la macro CrearPortada (módulo Mod5Portada)."
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = colGris
    End With
    fila = fila + 1
End Sub

' ---------------------------------------------------------------------
'  BOTONES (formas con hipervínculo o macro asignada)
' ---------------------------------------------------------------------
Private Sub AnadirBotonActualizar(ws As Worksheet)
    Dim shp As Shape
    Dim ref As Range

    Set ref = ws.Cells(3, 6)                ' Fila del título, última columna
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        ref.Left + ref.Width - 130, ref.Top + (ref.Height - 26) / 2, 125, 26)

    FormatearBoton shp, "Actualizar portada", vbWhite, colAzulOscuro
    shp.OnAction = "CrearPortada"           ' Al hacer clic se ejecuta la macro
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

    ' Quitar el botón anterior, si lo hubiera
    On Error Resume Next
    sh.Shapes(BOTON_VOLVER).Delete
    On Error GoTo Salir

    ' Se coloca una columna a la derecha de los datos, para no tapar nada
    colLibre = sh.UsedRange.Column + sh.UsedRange.Columns.Count + 1
    Set shp = sh.Shapes.AddShape(msoShapeRoundedRectangle, _
        sh.Cells(1, colLibre).Left, sh.Cells(1, 1).Top + 2, 90, 22)
    shp.Name = BOTON_VOLVER

    FormatearBoton shp, "« Portada", colAzulOscuro, vbWhite

    ' Un hipervínculo en la forma: funciona aunque las macros estén desactivadas
    sh.Hyperlinks.Add Anchor:=shp, Address:="", _
        SubAddress:="'" & NOMBRE_PORTADA & "'!A1", ScreenTip:="Volver a la portada"
Salir:
End Sub

Private Sub FormatearBoton(shp As Shape, texto As String, colorFondo As Long, colorTexto As Long)
    shp.Placement = xlFreeFloating          ' No se deforma al cambiar anchos de columna
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
    ws.Rows(fila + 1).RowHeight = 8         ' Pequeño respiro bajo el título
    fila = fila + 2
End Sub

' Devuelve la hoja o Nothing si no existe (sin dar error)
Private Function ObtenerHoja(nombre As String) As Worksheet
    On Error Resume Next
    Set ObtenerHoja = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0
End Function

' Año del ejercicio: se lee de la cabecera de Analisis (B3); si no, año pasado
Private Function AnioActual() As Long
    Dim wsA As Worksheet
    AnioActual = Year(Date) - 1
    Set wsA = ObtenerHoja("Analisis")
    If Not wsA Is Nothing Then
        If Not IsEmpty(wsA.Range("B3").Value) And IsNumeric(wsA.Range("B3").Value) Then
            AnioActual = CLng(wsA.Range("B3").Value)
        End If
    End If
End Function
