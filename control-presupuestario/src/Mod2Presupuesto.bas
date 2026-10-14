Attribute VB_Name = "Mod2Presupuesto"
Option Explicit

' Desviación desfavorable que se tolera antes de ponerse en rojo (5 %)
Public Const UMBRAL_VIGILAR As Double = 0.05

' ---- Prepara una hoja con epígrafes en filas y meses en columnas ----
Sub PrepararMatriz(ws As Worksheet, meses As Variant, epigrafes As Variant)
    Dim i As Long
    ws.Range("A1").Value = "Epígrafe"
    ws.Range("B1:M1").Value = meses
    For i = 0 To UBound(epigrafes)
        ws.Cells(i + 2, 1).Value = epigrafes(i)
    Next i
    ws.Range("A1:M1").Font.Bold = True
    ws.Range("B2:M10").NumberFormat = "#,##0;[Red]-#,##0"
    ws.Columns("A").AutoFit
    ws.Columns("B:M").ColumnWidth = 11
End Sub

' ---- Macro 1: crea el presupuesto anual y el real hasta septiembre ----
Sub CrearDatosPresupuesto()
    Dim wsP As Worksheet, wsR As Worksheet
    Dim meses As Variant, epigrafes As Variant
    Dim estacionalidad As Variant, factorReal As Variant
    Dim m As Integer
    Dim ventasP As Double, ventasR As Double

    meses = Array("Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic")
    epigrafes = Array("Importe neto de la cifra de negocios", "Aprovisionamientos", "Gastos de personal", "Otros gastos de explotación", "Amortización del inmovilizado")

    ' Peso de cada mes sobre la venta media (más ventas en verano)
    estacionalidad = Array(0.8, 0.85, 0.95, 1, 1.05, 1.1, 1.2, 1.15, 1, 0.95, 0.95, 1)
    ' Cumplimiento real de ventas frente a presupuesto, de enero a septiembre
    factorReal = Array(0.95, 0.97, 1.02, 1.05, 0.98, 1.03, 1.08, 1.06, 0.92)

    Set wsP = CrearHoja("Presupuesto")
    Set wsR = CrearHoja("Real")
    PrepararMatriz wsP, meses, epigrafes
    PrepararMatriz wsR, meses, epigrafes

    For m = 0 To 11
        ' Presupuesto: 1,4 M de ventas al año, compras al 55 % de las ventas
        ventasP = 1400000 / 12 * estacionalidad(m)
        wsP.Cells(2, m + 2).Value = Round(ventasP, 0)
        wsP.Cells(3, m + 2).Value = Round(-0.55 * ventasP, 0)
        wsP.Cells(4, m + 2).Value = -28000
        wsP.Cells(5, m + 2).Value = -3500
        wsP.Cells(6, m + 2).Value = -2700

        ' Real: solo los meses cerrados. Compras al 57 % y refuerzo de personal en julio y agosto
        If m <= UBound(factorReal) Then
            ventasR = ventasP * factorReal(m)
            wsR.Cells(2, m + 2).Value = Round(ventasR, 0)
            wsR.Cells(3, m + 2).Value = Round(-0.57 * ventasR, 0)
            wsR.Cells(4, m + 2).Value = IIf(m = 6 Or m = 7, -31000, -28000)
            wsR.Cells(5, m + 2).Value = -3850
            wsR.Cells(6, m + 2).Value = -2700
        End If
    Next m
End Sub

' ---- Escribe real, presupuesto y desviaciones; devuelve la desviación en % ----
Function EscribirDesviacion(ws As Worksheet, ByVal fila As Long, ByVal col As Long, ByVal valorReal As Double, ByVal ppto As Double) As Double
    Dim desvPct As Double
    ws.Cells(fila, col).Value = valorReal
    ws.Cells(fila, col + 1).Value = ppto
    ws.Cells(fila, col + 2).Value = valorReal - ppto
    If ppto <> 0 Then
        desvPct = (valorReal - ppto) / Abs(ppto)
        ws.Cells(fila, col + 3).Value = desvPct
    End If
    EscribirDesviacion = desvPct
End Function

' ---- Pinta el semáforo según la desviación ----
Sub PintarSemaforo(celda As Range, ByVal desvPct As Double)
    If desvPct >= 0 Then
        celda.Value = "Favorable"
        celda.Interior.Color = RGB(198, 239, 206)
    ElseIf desvPct >= -UMBRAL_VIGILAR Then
        celda.Value = "Vigilar"
        celda.Interior.Color = RGB(255, 235, 156)
    Else
        celda.Value = "Desfavorable"
        celda.Interior.Color = RGB(255, 199, 206)
    End If
    celda.HorizontalAlignment = xlCenter
End Sub

' ==== MACRO PRINCIPAL ====
Sub AnalisisDesviaciones()
    Dim wsP As Worksheet, wsR As Worksheet, ws As Worksheet
    Dim respuesta As String, nombreMes As String, peorEpi As String
    Dim mes As Integer
    Dim i As Long, fila As Long, ultimaFila As Long
    Dim rm As Double, pm As Double, ra As Double, pa As Double
    Dim rmTot As Double, pmTot As Double, raTot As Double, paTot As Double
    Dim desvPct As Double, peorDesv As Double
    Dim grafico As ChartObject

    If Not HojaExiste("Presupuesto") Or Not HojaExiste("Real") Then
        MsgBox "Ejecuta primero CrearDatosPresupuesto.", vbExclamation
        Exit Sub
    End If
    Set wsP = ThisWorkbook.Sheets("Presupuesto")
    Set wsR = ThisWorkbook.Sheets("Real")

    ' 1. Preguntar el mes y validar la respuesta
    respuesta = InputBox("¿Qué mes quieres analizar? (1 a 12)", "Real vs presupuesto", "9")
    If respuesta = "" Then Exit Sub                     ' Ha pulsado Cancelar
    If Not IsNumeric(respuesta) Then
        MsgBox "Escribe un número del 1 al 12.", vbExclamation
        Exit Sub
    End If
    mes = CInt(respuesta)
    If mes < 1 Or mes > 12 Then
        MsgBox "El mes debe estar entre 1 y 12.", vbExclamation
        Exit Sub
    End If
    If IsEmpty(wsR.Cells(2, mes + 1).Value) Then
        MsgBox "Todavía no hay datos reales del mes " & mes & ".", vbExclamation
        Exit Sub
    End If
    nombreMes = wsP.Cells(1, mes + 1).Value

    ' 2. Preparar la hoja de resultados
    Set ws = CrearHoja("Desviaciones")
    ws.Range("A1").Value = "REAL VS PRESUPUESTO - " & UCase(nombreMes)
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("B3").Value = "MES: " & UCase(nombreMes)
    ws.Range("F3").Value = "ACUMULADO ENE - " & UCase(nombreMes)
    ws.Range("A3:J3").Font.Bold = True
    ws.Range("A4:J4").Value = Array("Epígrafe", "Real", "Ppto", "Desv. €", "Desv. %", "Real", "Ppto", "Desv. €", "Desv. %", "Semáforo")
    With ws.Range("A4:J4")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
    End With

    ' 3. Recorrer los epígrafes
    ultimaFila = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row
    fila = 5
    For i = 2 To ultimaFila
        rm = wsR.Cells(i, mes + 1).Value
        pm = wsP.Cells(i, mes + 1).Value
        ra = Application.WorksheetFunction.Sum(wsR.Range(wsR.Cells(i, 2), wsR.Cells(i, mes + 1)))
        pa = Application.WorksheetFunction.Sum(wsP.Range(wsP.Cells(i, 2), wsP.Cells(i, mes + 1)))

        ws.Cells(fila, 1).Value = wsP.Cells(i, 1).Value
        EscribirDesviacion ws, fila, 2, rm, pm
        desvPct = EscribirDesviacion(ws, fila, 6, ra, pa)
        PintarSemaforo ws.Cells(fila, 10), desvPct

        ' Guardar la peor desviación acumulada en euros
        If ra - pa < peorDesv Then
            peorDesv = ra - pa
            peorEpi = wsP.Cells(i, 1).Value
        End If

        rmTot = rmTot + rm
        pmTot = pmTot + pm
        raTot = raTot + ra
        paTot = paTot + pa
        fila = fila + 1
    Next i

    ' 4. Fila de resultado
    ws.Cells(fila, 1).Value = "RESULTADO"
    EscribirDesviacion ws, fila, 2, rmTot, pmTot
    desvPct = EscribirDesviacion(ws, fila, 6, raTot, paTot)
    PintarSemaforo ws.Cells(fila, 10), desvPct
    ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 10)).Font.Bold = True
    ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 10)).Borders(xlEdgeTop).LineStyle = xlContinuous

    ' 5. Formatos
    ws.Range("B5:D" & fila).NumberFormat = "#,##0;[Red]-#,##0"
    ws.Range("F5:H" & fila).NumberFormat = "#,##0;[Red]-#,##0"
    ws.Range("E5:E" & fila).NumberFormat = "+0.0%;[Red]-0.0%"
    ws.Range("I5:I" & fila).NumberFormat = "+0.0%;[Red]-0.0%"
    ws.Columns("A").AutoFit
    ws.Columns("B:J").ColumnWidth = 12

    ' 6. Comentario automático
    fila = fila + 2
    If peorEpi <> "" Then
        ws.Cells(fila, 1).Value = "Mayor desviación desfavorable acumulada: " & peorEpi & " (" & Format(peorDesv, "#,##0") & " €)."
    Else
        ws.Cells(fila, 1).Value = "Todas las partidas van mejor que el presupuesto."
    End If
    ws.Cells(fila, 1).Font.Italic = True

    ' 7. Gráfico de ventas mensuales
    Set grafico = ws.ChartObjects.Add(Left:=ws.Range("A" & fila + 2).Left, Top:=ws.Range("A" & fila + 2).Top, Width:=600, Height:=260)
    With grafico.Chart
        With .SeriesCollection.NewSeries
            .Name = "Presupuesto"
            .Values = wsP.Range("B2:M2")
            .XValues = wsP.Range("B1:M1")
        End With
        With .SeriesCollection.NewSeries
            .Name = "Real"
            .Values = wsR.Range("B2:M2")
            .XValues = wsP.Range("B1:M1")
        End With
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Ventas mensuales: real vs presupuesto"
    End With

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    MsgBox "Análisis de " & nombreMes & " generado.", vbInformation
End Sub
