Attribute VB_Name = "Mod3CuadroMando"
Option Explicit

' ---- Devuelve el color de fondo según la desviación ----
Function ColorSemaforo(ByVal desvPct As Double) As Long
    If desvPct >= 0 Then
        ColorSemaforo = RGB(198, 239, 206)          ' Verde
    ElseIf desvPct >= -UMBRAL_VIGILAR Then
        ColorSemaforo = RGB(255, 235, 156)          ' Ámbar
    Else
        ColorSemaforo = RGB(255, 199, 206)          ' Rojo
    End If
End Function

' ---- Suma una fila desde enero hasta el mes indicado ----
Function SumaHasta(wsX As Worksheet, ByVal fila As Long, ByVal mes As Long) As Double
    SumaHasta = Application.WorksheetFunction.Sum(wsX.Range(wsX.Cells(fila, 2), wsX.Cells(fila, mes + 1)))
End Function

' ---- Prepara el marco de una tarjeta KPI (3 columnas de ancho, filas 5 a 8) ----
Sub PrepararKPI(ws As Worksheet, ByVal col As Long)
    ws.Range(ws.Cells(5, col), ws.Cells(5, col + 2)).Merge
    ws.Range(ws.Cells(6, col), ws.Cells(7, col + 2)).Merge
    ws.Range(ws.Cells(8, col), ws.Cells(8, col + 2)).Merge
    With ws.Range(ws.Cells(5, col), ws.Cells(8, col + 2))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(191, 191, 191)
    End With
    ws.Cells(5, col).Font.Size = 10
    ws.Cells(5, col).Font.Color = RGB(89, 89, 89)
    ws.Cells(6, col).Font.Size = 20
    ws.Cells(6, col).Font.Bold = True
    ws.Cells(8, col).Font.Size = 9
End Sub

' ---- Rellena una tarjeta KPI ----
Sub EscribirKPI(ws As Worksheet, ByVal col As Long, titulo As String, ByVal valor As Double, formato As String, subtitulo As String, ByVal desvPct As Double)
    ws.Cells(5, col).Value = titulo
    ws.Cells(6, col).Value = valor
    ws.Cells(6, col).NumberFormat = formato
    ws.Cells(8, col).Value = subtitulo
    ws.Range(ws.Cells(5, col), ws.Cells(8, col + 2)).Interior.Color = ColorSemaforo(desvPct)
End Sub

' ---- Crea un gráfico enlazado a unos rangos, en la zona indicada ----
Sub CrearGrafico(ws As Worksheet, nombre As String, zona As Range, titulo As String, ByVal tipo As Long, nombres As Variant, valores As Variant, ejeX As Range, colores As Variant)
    Dim co As ChartObject
    Dim i As Long

    Set co = ws.ChartObjects.Add(zona.Left, zona.Top, zona.Width, zona.Height)
    co.Name = nombre                                ' Nombre para encontrarlo después

    With co.Chart
        For i = 0 To UBound(nombres)
            With .SeriesCollection.NewSeries
                .Name = nombres(i)
                .Values = valores(i)
                .XValues = ejeX
            End With
        Next i
        .ChartType = tipo
        For i = 0 To UBound(nombres)
            .SeriesCollection(i + 1).Format.Fill.ForeColor.RGB = colores(i)
            .SeriesCollection(i + 1).Format.Line.ForeColor.RGB = colores(i)
        Next i
        .HasTitle = True
        .ChartTitle.Text = titulo
        .HasLegend = (UBound(nombres) > 0)
        If .HasLegend Then .Legend.Position = xlLegendPositionBottom
    End With
End Sub

' ==== MACRO 1: construye el cuadro de mando (se ejecuta una vez) ====
Sub CrearCuadroMando()
    Dim ws As Worksheet, wsD As Worksheet, wsP As Worksheet, wsR As Worksheet
    Dim btn As Shape
    Dim m As Long, n As Long

    If Not HojaExiste("Presupuesto") Or Not HojaExiste("Real") Then
        MsgBox "Ejecuta primero CrearDatosPresupuesto.", vbExclamation
        Exit Sub
    End If
    Set wsP = ThisWorkbook.Sheets("Presupuesto")
    Set wsR = ThisWorkbook.Sheets("Real")

    ' 1. Hoja auxiliar con los datos de los gráficos
    Set wsD = CrearHoja("DatosCuadro")
    wsD.Range("A1:E1").Value = Array("Mes", "Ventas ppto", "Ventas real", "Resultado acum. ppto", "Resultado acum. real")
    wsD.Range("G1:H1").Value = Array("Epígrafe", "Desv. acumulada")
    wsD.Range("J1").Value = "Meses con datos"
    wsD.Range("A1:J1").Font.Bold = True
    wsD.Range("B2:H13").NumberFormat = "#,##0"

    ' Lista de meses que tienen datos reales (para el desplegable)
    For m = 1 To 12
        If Not IsEmpty(wsR.Cells(2, m + 1).Value) Then
            n = n + 1
            wsD.Cells(n + 1, 10).Value = wsP.Cells(1, m + 1).Value
        End If
    Next m

    ' 2. Hoja del cuadro de mando
    Set ws = CrearHoja("CuadroMando")
    ws.Move Before:=ThisWorkbook.Sheets(1)          ' La ponemos la primera
    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B:P").ColumnWidth = 11
    ws.Columns("E").ColumnWidth = 2                 ' Separadores entre tarjetas
    ws.Columns("I").ColumnWidth = 2
    ws.Columns("M").ColumnWidth = 2

    ws.Range("B2").Font.Size = 18
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Color = RGB(31, 78, 121)

    ' Selector de mes con lista desplegable
    ws.Rows(3).RowHeight = 24
    ws.Range("B3").Value = "Mes:"
    ws.Range("B3").Font.Bold = True
    With ws.Range("C3")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=DatosCuadro!$J$2:$J$" & (n + 1)
        .Value = wsD.Cells(n + 1, 10).Value         ' Último mes con datos
        .Interior.Color = RGB(255, 242, 204)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .BorderAround LineStyle:=xlContinuous, Weight:=xlThin
    End With

    ' Botón que ejecuta ActualizarCuadro
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("D3").Left + 5, ws.Range("D3").Top + 2, 90, 20)
    btn.Name = "btnActualizar"
    btn.Fill.ForeColor.RGB = RGB(31, 78, 121)
    btn.Line.Visible = msoFalse
    btn.TextFrame2.TextRange.Text = "Actualizar"
    btn.TextFrame2.TextRange.Font.Bold = msoTrue
    btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    btn.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    btn.OnAction = "ActualizarCuadro"

    ' 3. Marcos de las 4 tarjetas KPI (columnas B, F, J y N)
    ws.Rows("6:7").RowHeight = 20
    PrepararKPI ws, 2
    PrepararKPI ws, 6
    PrepararKPI ws, 10
    PrepararKPI ws, 14

    ' 4. Gráficos enlazados a DatosCuadro
    CrearGrafico ws, "grafVentas", ws.Range("B10:H25"), "Ventas mensuales (€)", xlColumnClustered, Array("Presupuesto", "Real"), Array(wsD.Range("B2:B13"), wsD.Range("C2:C13")), wsD.Range("A2:A13"), Array(RGB(166, 166, 166), RGB(31, 78, 121))
    CrearGrafico ws, "grafResultado", ws.Range("J10:P25"), "Resultado acumulado (€)", xlLineMarkers, Array("Presupuesto", "Real"), Array(wsD.Range("D2:D13"), wsD.Range("E2:E13")), wsD.Range("A2:A13"), Array(RGB(166, 166, 166), RGB(31, 78, 121))
    CrearGrafico ws, "grafDesv", ws.Range("B27:P42"), "Desviación acumulada por epígrafe (€)", xlBarClustered, Array("Desviación"), Array(wsD.Range("H2:H6")), wsD.Range("G2:G6"), Array(RGB(31, 78, 121))
   ws.ChartObjects("grafDesv").Chart.Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
   
    ' 5. Primera actualización
    ActualizarCuadro
End Sub

' ==== MACRO 2: actualiza cifras y gráficos según el mes elegido ====
Sub ActualizarCuadro()
    Dim ws As Worksheet, wsD As Worksheet, wsP As Worksheet, wsR As Worksheet
    Dim posicion As Variant
    Dim mes As Long, m As Long, i As Long, p As Long
    Dim acumP As Double, acumR As Double, resPA As Double
    Dim ventasRA As Double, ventasPA As Double, aprovRA As Double, aprovPA As Double
    Dim resRM As Double, resPM As Double
    Dim margenR As Double, margenP As Double
    Dim desv As Double, peorDesv As Double, peorEpi As String

    If Not HojaExiste("CuadroMando") Or Not HojaExiste("DatosCuadro") Then
        MsgBox "Ejecuta primero CrearCuadroMando.", vbExclamation
        Exit Sub
    End If
    Set ws = ThisWorkbook.Sheets("CuadroMando")
    Set wsD = ThisWorkbook.Sheets("DatosCuadro")
    Set wsP = ThisWorkbook.Sheets("Presupuesto")
    Set wsR = ThisWorkbook.Sheets("Real")

    ' 1. Averiguar el número de mes a partir del texto del desplegable
    posicion = Application.Match(ws.Range("C3").Value, wsP.Range("B1:M1"), 0)
    If IsError(posicion) Then
        MsgBox "Elige un mes de la lista.", vbExclamation
        Exit Sub
    End If
    mes = posicion
    If IsEmpty(wsR.Cells(2, mes + 1).Value) Then
        MsgBox "No hay datos reales de ese mes.", vbExclamation
        Exit Sub
    End If

    ' 2. Series mensuales para los gráficos (los epígrafes están en las filas 2 a 6)
    wsD.Range("C2:C13,E2:E13").ClearContents        ' Vacías de verdad, para que no pinte ceros
    For m = 1 To 12
        wsD.Cells(m + 1, 1).Value = wsP.Cells(1, m + 1).Value
        wsD.Cells(m + 1, 2).Value = wsP.Cells(2, m + 1).Value
        acumP = acumP + Application.WorksheetFunction.Sum(wsP.Range(wsP.Cells(2, m + 1), wsP.Cells(6, m + 1)))
        wsD.Cells(m + 1, 4).Value = acumP
        If m <= mes Then
            wsD.Cells(m + 1, 3).Value = wsR.Cells(2, m + 1).Value
            acumR = acumR + Application.WorksheetFunction.Sum(wsR.Range(wsR.Cells(2, m + 1), wsR.Cells(6, m + 1)))
            wsD.Cells(m + 1, 5).Value = acumR
        End If
    Next m
    resPA = wsD.Cells(mes + 1, 4).Value             ' Resultado presupuestado hasta el mes

    ' 3. Desviación acumulada por epígrafe
    For i = 2 To 6
        wsD.Cells(i, 7).Value = wsP.Cells(i, 1).Value
        desv = SumaHasta(wsR, i, mes) - SumaHasta(wsP, i, mes)
        wsD.Cells(i, 8).Value = desv
        If desv < peorDesv Then
            peorDesv = desv
            peorEpi = wsP.Cells(i, 1).Value
        End If
    Next i

    ' 4. KPIs
    ventasRA = SumaHasta(wsR, 2, mes)
    ventasPA = SumaHasta(wsP, 2, mes)
    aprovRA = SumaHasta(wsR, 3, mes)
    aprovPA = SumaHasta(wsP, 3, mes)
    margenR = (ventasRA + aprovRA) / ventasRA
    margenP = (ventasPA + aprovPA) / ventasPA
    resRM = Application.WorksheetFunction.Sum(wsR.Range(wsR.Cells(2, mes + 1), wsR.Cells(6, mes + 1)))
    resPM = Application.WorksheetFunction.Sum(wsP.Range(wsP.Cells(2, mes + 1), wsP.Cells(6, mes + 1)))

    EscribirKPI ws, 2, "VENTAS ACUMULADAS", ventasRA, "#,##0 ""€""", "vs ppto: " & Format(ventasRA / ventasPA - 1, "+0.0%;-0.0%"), ventasRA / ventasPA - 1
    EscribirKPI ws, 6, "MARGEN BRUTO", margenR, "0.0%", "Ppto: " & Format(margenP, "0.0%") & " (" & Format((margenR - margenP) * 100, "+0.0;-0.0") & " p.p.)", (margenR - margenP) / margenP
    EscribirKPI ws, 10, "RESULTADO ACUMULADO", acumR, "#,##0 ""€""", "vs ppto: " & Format((acumR - resPA) / Abs(resPA), "+0.0%;-0.0%"), (acumR - resPA) / Abs(resPA)
    EscribirKPI ws, 14, "RESULTADO DEL MES", resRM, "#,##0 ""€""", "vs ppto: " & Format((resRM - resPM) / Abs(resPM), "+0.0%;-0.0%"), (resRM - resPM) / Abs(resPM)

    ' 5. Barras verdes o rojas según el signo de la desviación
    With ws.ChartObjects("grafDesv").Chart.SeriesCollection(1)
        For p = 1 To 5
            If wsD.Cells(p + 1, 8).Value < 0 Then
                .Points(p).Format.Fill.ForeColor.RGB = RGB(192, 0, 0)
            Else
                .Points(p).Format.Fill.ForeColor.RGB = RGB(0, 128, 0)
            End If
        Next p
    End With

    ' 6. Título y comentario
    ws.Range("B2").Value = "CUADRO DE MANDO · ACUMULADO ENE - " & UCase(ws.Range("C3").Value)
    If peorEpi <> "" Then
        ws.Range("B44").Value = "Foco: la mayor desviación desfavorable acumulada está en " & peorEpi & " (" & Format(peorDesv, "#,##0") & " €)."
    Else
        ws.Range("B44").Value = "Todas las partidas van mejor que el presupuesto."
    End If
    ws.Range("B44").Font.Italic = True
End Sub
