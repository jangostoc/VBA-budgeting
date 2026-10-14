Attribute VB_Name = "modConclusiones"
Option Explicit

' =====================================================================
'  modConclusiones
'  Crea la hoja "Conclusiones": informe de gestión del acumulado
'  real vs presupuesto, redactado automáticamente con las cifras.
'    1. Resumen ejecutivo
'    2. Puente del resultado (efecto volumen, efecto precio y resto)
'    3. Análisis por partida, con semáforo
'    4. Previsión de cierre del ejercicio (dos escenarios)
'    5. Puntos de acción
'
'  Usa:  CrearHoja, HojaExiste   (modUtilidades)
'        SumaHasta, ColorSemaforo (modCuadroMando)
'        UMBRAL_VIGILAR          (modPresupuesto, declarada Public)
'
'  Macro que hay que ejecutar:  GenerarConclusionesPpto
' =====================================================================

Private Const FILA_VENTAS As Long = 2       ' Fila de la cifra de negocios en Presupuesto y Real
Private Const FILA_APROV As Long = 3        ' Fila de los aprovisionamientos

' ---------------------------------------------------------------------
'  MACRO PRINCIPAL
' ---------------------------------------------------------------------
Public Sub GenerarConclusionesPpto()
    Dim wsP As Worksheet, wsR As Worksheet, ws As Worksheet
    Dim mes As Long, ultimaFila As Long, i As Long, fila As Long
    Dim mesesDesf As Long
    Dim nombreMes As String, periodo As String, nombreEpi As String
    Dim peorMes As String, peorConcepto As String, texto As String
    Dim vR As Double, vP As Double, aR As Double, aP As Double
    Dim resR As Double, resP As Double, resAnualP As Double
    Dim resMesR As Double, resMesP As Double, vMesR As Double, vMesP As Double
    Dim mbR As Double, mbP As Double
    Dim efVolumen As Double, efPrecio As Double
    Dim desv As Double, desvPct As Double, pptoEpi As Double, suma As Double
    Dim peorImporte As Double, peorDesvMes As Double
    Dim dMes As Double, dAcum As Double
    Dim ventasRestP As Double, previsionA As Double, previsionB As Double
    Dim acciones As Collection
    Dim accion As Variant

    If Not HojaExiste("Presupuesto") Or Not HojaExiste("Real") Then
        MsgBox "Ejecuta primero CrearDatosPresupuesto.", vbExclamation
        Exit Sub
    End If
    Set wsP = ThisWorkbook.Sheets("Presupuesto")
    Set wsR = ThisWorkbook.Sheets("Real")

    mes = MesDelInforme(wsP, wsR)
    If mes = 0 Then
        MsgBox "No hay datos reales en la hoja Real.", vbExclamation
        Exit Sub
    End If
    nombreMes = wsP.Cells(1, mes + 1).Value
    periodo = wsP.Cells(1, 2).Value & " - " & nombreMes
    ultimaFila = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row

    Application.ScreenUpdating = False

    ' -----------------------------------------------------------------
    '  1. CÁLCULOS
    ' -----------------------------------------------------------------
    vR = SumaHasta(wsR, FILA_VENTAS, mes)
    vP = SumaHasta(wsP, FILA_VENTAS, mes)
    aR = SumaHasta(wsR, FILA_APROV, mes)
    aP = SumaHasta(wsP, FILA_APROV, mes)
    vMesR = wsR.Cells(FILA_VENTAS, mes + 1).Value
    vMesP = wsP.Cells(FILA_VENTAS, mes + 1).Value

    For i = 2 To ultimaFila
        resR = resR + SumaHasta(wsR, i, mes)
        resP = resP + SumaHasta(wsP, i, mes)
        resAnualP = resAnualP + SumaHasta(wsP, i, 12)
        resMesR = resMesR + wsR.Cells(i, mes + 1).Value
        resMesP = resMesP + wsP.Cells(i, mes + 1).Value
    Next i

    ' Margen bruto = (ventas + aprovisionamientos) / ventas  (aprov. va en negativo)
    mbR = DivSegura(vR + aR, vR)
    mbP = DivSegura(vP + aP, vP)

    ' Descomposición exacta de la desviación del margen bruto:
    '   volumen = ventas de más (o de menos) x margen bruto presupuestado
    '   precio  = compras reales - compras que "tocaban" con las ventas reales
    efVolumen = (vR - vP) * mbP
    efPrecio = aR - aP * DivSegura(vR, vP)

    ' Previsión de cierre
    ventasRestP = SumaHasta(wsP, FILA_VENTAS, 12) - vP
    previsionA = resR + (resAnualP - resP)
    previsionB = previsionA + (DivSegura(aR, vR) - DivSegura(aP, vP)) * ventasRestP

    ' -----------------------------------------------------------------
    '  2. HOJA
    ' -----------------------------------------------------------------
    Set ws = CrearHoja("Conclusiones")
    ws.Tab.Color = RGB(47, 117, 181)
    ws.Cells.Font.Name = "Calibri"
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 3
    ws.Columns("C").ColumnWidth = 72
    ws.Columns("D:E").ColumnWidth = 17

    fila = 2
    InfTitulo ws, fila, "INFORME DE GESTIÓN · ACUMULADO " & UCase(periodo), 18
    InfParrafo ws, fila, "Real frente a presupuesto. Informe generado automáticamente el " & _
        Format(Now, "dd/mm/yyyy hh:nn") & " con la macro GenerarConclusionesPpto."
    ws.Cells(fila - 1, 3).Font.Italic = True
    ws.Cells(fila - 1, 3).Font.Color = RGB(110, 110, 110)
    fila = fila + 1

    ' -----------------------------------------------------------------
    '  3. RESUMEN EJECUTIVO
    ' -----------------------------------------------------------------
    InfTitulo ws, fila, "1. Resumen ejecutivo", 13

    InfVineta ws, fila, "Las ventas acumuladas suman " & Eur(vR) & ", un " & _
        Pct(DivSegura(vR - vP, vP)) & " frente a los " & Eur(vP) & " presupuestados."

    InfVineta ws, fila, "El resultado acumulado es de " & Eur(resR) & " frente a " & _
        Eur(resP) & " presupuestados: una desviación de " & EurSigno(resR - resP) & _
        " (" & Pct(DivSegura(resR - resP, Abs(resP))) & ")."

    InfVineta ws, fila, "El margen bruto (ventas menos compras) se sitúa en el " & _
        Format(mbR, "0.0%") & ", " & PuntosTexto(mbR - mbP) & " respecto al " & _
        Format(mbP, "0.0%") & " presupuestado."

    dMes = DivSegura(resMesR - resMesP, Abs(resMesP))
    dAcum = DivSegura(resR - resP, Abs(resP))
    If dMes < dAcum Then
        texto = "En " & nombreMes & " la desviación del resultado se amplía hasta el " & _
            Pct(dMes) & " (frente al " & Pct(dAcum) & " acumulado): la tendencia empeora."
    Else
        texto = "En " & nombreMes & " la desviación del resultado es del " & Pct(dMes) & _
            " (frente al " & Pct(dAcum) & " acumulado): la tendencia mejora."
    End If
    InfVineta ws, fila, texto
    fila = fila + 1

    ' -----------------------------------------------------------------
    '  4. PUENTE DEL RESULTADO
    ' -----------------------------------------------------------------
    InfTitulo ws, fila, "2. ¿Qué explica la desviación del resultado?", 13
    InfParrafo ws, fila, "Puente desde el resultado presupuestado hasta el real. Ventas y compras " & _
        "se separan en efecto volumen (vender más o menos de lo previsto) y efecto precio " & _
        "(comprar más caro o más barato en proporción a lo vendido)."

    InfCabecera ws, fila, Array("Concepto", "Importe", "Efecto")
    InfFilaPuente ws, fila, "Resultado presupuestado", resP, True
    InfFilaPuente ws, fila, "Efecto volumen de ventas", efVolumen, False
    InfFilaPuente ws, fila, "Efecto precio de las compras", efPrecio, False
    ComprobarPeor "Efecto volumen de ventas", efVolumen, peorConcepto, peorImporte
    ComprobarPeor "Efecto precio de las compras", efPrecio, peorConcepto, peorImporte
    suma = resP + efVolumen + efPrecio

    ' El resto de partidas (de personal en adelante) van directas
    For i = FILA_APROV + 1 To ultimaFila
        nombreEpi = CStr(wsP.Cells(i, 1).Value)
        desv = SumaHasta(wsR, i, mes) - SumaHasta(wsP, i, mes)
        InfFilaPuente ws, fila, nombreEpi, desv, False
        ComprobarPeor nombreEpi, desv, peorConcepto, peorImporte
        suma = suma + desv
    Next i
    InfFilaPuente ws, fila, "Resultado real", resR, True
    fila = fila + 1

    ' Control: el puente debe llegar exactamente al resultado real
    If Abs(suma - resR) > 1 Then
        InfParrafo ws, fila, "Atención: el puente no cuadra con el resultado real (diferencia de " & _
            EurSigno(suma - resR) & "). Revisa la estructura de las hojas Presupuesto y Real."
        ws.Cells(fila - 1, 3).Font.Color = RGB(192, 0, 0)
    End If

    If resR - resP < 0 And peorImporte < 0 Then
        InfParrafo ws, fila, "El principal factor es «" & peorConcepto & "» (" & EurSigno(peorImporte) & _
            "), que por sí solo equivale al " & Format(peorImporte / (resR - resP), "0%") & _
            " de la desviación total del resultado."
    ElseIf resR - resP >= 0 Then
        InfParrafo ws, fila, "El resultado supera el presupuesto en " & EurSigno(resR - resP) & "."
    End If

    If efVolumen > 0 And efPrecio < 0 Then
        InfParrafo ws, fila, "Vender más de lo previsto ha aportado " & EurSigno(efVolumen) & _
            " de margen, pero el encarecimiento de las compras (" & Format(-aR / vR, "0.0%") & _
            " de las ventas frente al " & Format(-aP / vP, "0.0%") & " presupuestado) lo ha absorbido con creces."
    End If
    fila = fila + 1

    ' -----------------------------------------------------------------
    '  5. ANÁLISIS POR PARTIDA
    ' -----------------------------------------------------------------
    InfTitulo ws, fila, "3. Análisis por partida", 13
    For i = 2 To ultimaFila
        nombreEpi = CStr(wsP.Cells(i, 1).Value)
        pptoEpi = SumaHasta(wsP, i, mes)
        desv = SumaHasta(wsR, i, mes) - pptoEpi
        desvPct = DivSegura(desv, Abs(pptoEpi))
        peorMes = AnalizarMeses(wsP, wsR, i, mes, mesesDesf, peorDesvMes)

        If Abs(desv) < 1 Then
            texto = nombreEpi & ": en línea con el presupuesto."
        Else
            texto = nombreEpi & ": " & EurSigno(desv) & " (" & Pct(desvPct) & "), " & _
                EstadoSemaforo(desvPct) & ". Desfavorable en " & mesesDesf & " de " & mes & " meses"
            If peorMes <> "" Then
                texto = texto & "; el peor mes fue " & peorMes & " (" & EurSigno(peorDesvMes) & ")"
            End If
            texto = texto & "."
        End If
        InfVineta ws, fila, texto, ColorSemaforo(desvPct)
    Next i
    fila = fila + 1

    ' -----------------------------------------------------------------
    '  6. PREVISIÓN DE CIERRE
    ' -----------------------------------------------------------------
    InfTitulo ws, fila, "4. Previsión de cierre del ejercicio", 13
    If mes < 12 Then
        InfCabecera ws, fila, Array("Escenario", "Resultado anual", "vs presupuesto")
        InfFilaPrevision ws, fila, "Presupuesto anual aprobado", resAnualP, 0, True
        InfFilaPrevision ws, fila, "A. Real hasta " & nombreMes & " + presupuesto de " & _
            wsP.Cells(1, mes + 2).Value & " a " & wsP.Cells(1, 13).Value, previsionA, _
            DivSegura(previsionA - resAnualP, Abs(resAnualP)), False
        InfFilaPrevision ws, fila, "B. Escenario A manteniendo el coste actual de las compras", _
            previsionB, DivSegura(previsionB - resAnualP, Abs(resAnualP)), False
        fila = fila + 1

        texto = "Si el resto del año se cumple el presupuesto, el ejercicio cerraría con " & _
            Eur(previsionA) & " (" & Pct(DivSegura(previsionA - resAnualP, Abs(resAnualP))) & _
            " frente al presupuesto)."
        If efPrecio < 0 Then
            texto = texto & " Si las compras siguen en el " & Format(-aR / vR, "0.0%") & _
                " de las ventas, el cierre bajaría a " & Eur(previsionB) & "."
        End If
        InfParrafo ws, fila, texto
    Else
        InfParrafo ws, fila, "El ejercicio está cerrado: el resultado anual real es de " & Eur(resR) & _
            " frente a " & Eur(resAnualP) & " presupuestados."
    End If
    fila = fila + 1

    ' -----------------------------------------------------------------
    '  7. PUNTOS DE ACCIÓN (se generan según lo que se ha detectado)
    ' -----------------------------------------------------------------
    InfTitulo ws, fila, "5. Puntos de acción", 13
    Set acciones = New Collection

    If efPrecio < 0 Then
        texto = "Compras: el coste está en el " & Format(-aR / vR, "0.0%") & " de las ventas frente al " & _
            Format(-aP / vP, "0.0%") & " presupuestado. Revisar precios y condiciones con proveedores"
        If mes < 12 Then texto = texto & "; de mantenerse, restaría " & Eur(previsionA - previsionB) & " más hasta diciembre"
        acciones.Add texto & "."
    End If

    If DivSegura(vMesR - vMesP, Abs(vMesP)) < -UMBRAL_VIGILAR Then
        acciones.Add "Comercial: las ventas de " & nombreMes & " quedan un " & _
            Pct(DivSegura(vMesR - vMesP, Abs(vMesP))) & " por debajo del presupuesto. " & _
            "Confirmar si es un hecho puntual o un cambio de tendencia."
    End If

    For i = FILA_APROV + 1 To ultimaFila
        desv = SumaHasta(wsR, i, mes) - SumaHasta(wsP, i, mes)
        If desv < -1 Then
            nombreEpi = CStr(wsP.Cells(i, 1).Value)
            peorMes = AnalizarMeses(wsP, wsR, i, mes, mesesDesf, peorDesvMes)
            acciones.Add nombreEpi & ": desviación acumulada de " & EurSigno(desv) & " (" & _
                Pct(DivSegura(desv, Abs(SumaHasta(wsP, i, mes)))) & "), desfavorable en " & _
                mesesDesf & " de " & mes & " meses. Analizar la causa y si es recurrente."
        End If
    Next i

    If acciones.Count = 0 Then
        acciones.Add "No hay desviaciones relevantes: mantener el seguimiento mensual."
    End If
    For Each accion In acciones
        InfVineta ws, fila, CStr(accion)
    Next accion
    fila = fila + 1

    ' Nota final
    InfParrafo ws, fila, "Criterio del semáforo: favorable si la desviación es positiva o nula; " & _
        "a vigilar si es desfavorable hasta un " & Format(UMBRAL_VIGILAR, "0%") & _
        "; desfavorable por encima de ese umbral."
    With ws.Cells(fila - 1, 3).Font
        .Size = 9
        .Italic = True
        .Color = RGB(110, 110, 110)
    End With

    ' -----------------------------------------------------------------
    '  8. ACABADO
    ' -----------------------------------------------------------------
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Select
    Application.ScreenUpdating = True

    MsgBox "Informe de conclusiones generado (acumulado " & periodo & ").", vbInformation
End Sub

' ---------------------------------------------------------------------
'  MES DEL INFORME
'  Usa el mes elegido en el CuadroMando (celda C3) si existe y tiene datos.
'  Si no, el último mes con datos reales. Devuelve 0 si no hay ninguno.
'  Es Public porque también la usa la portada.
' ---------------------------------------------------------------------
Public Function MesDelInforme(wsP As Worksheet, wsR As Worksheet) As Long
    Dim posicion As Variant
    Dim m As Long

    If HojaExiste("CuadroMando") Then
        posicion = Application.Match(ThisWorkbook.Sheets("CuadroMando").Range("C3").Value, wsP.Range("B1:M1"), 0)
        If Not IsError(posicion) Then
            If Not IsEmpty(wsR.Cells(FILA_VENTAS, posicion + 1).Value) Then
                MesDelInforme = posicion
                Exit Function
            End If
        End If
    End If

    For m = 12 To 1 Step -1
        If Not IsEmpty(wsR.Cells(FILA_VENTAS, m + 1).Value) Then
            MesDelInforme = m
            Exit Function
        End If
    Next m
End Function

' ---------------------------------------------------------------------
'  ANÁLISIS MES A MES DE UNA PARTIDA
'  Devuelve el nombre del peor mes y, por ByRef, cuántos meses fueron
'  desfavorables y la peor desviación mensual.
' ---------------------------------------------------------------------
Private Function AnalizarMeses(wsP As Worksheet, wsR As Worksheet, ByVal fila As Long, _
                               ByVal mes As Long, ByRef mesesDesf As Long, _
                               ByRef peorDesv As Double) As String
    Dim m As Long
    Dim d As Double

    mesesDesf = 0
    peorDesv = 0
    For m = 1 To mes
        d = wsR.Cells(fila, m + 1).Value - wsP.Cells(fila, m + 1).Value
        If d < 0 Then mesesDesf = mesesDesf + 1
        If d < peorDesv Then
            peorDesv = d
            AnalizarMeses = wsP.Cells(1, m + 1).Value
        End If
    Next m
End Function

' Guarda el concepto con el importe más negativo
Private Sub ComprobarPeor(concepto As String, ByVal importe As Double, _
                          ByRef peorConcepto As String, ByRef peorImporte As Double)
    If importe < peorImporte Then
        peorImporte = importe
        peorConcepto = concepto
    End If
End Sub

' ---------------------------------------------------------------------
'  FORMATO DEL INFORME
' ---------------------------------------------------------------------
Private Sub InfTitulo(ws As Worksheet, ByRef fila As Long, texto As String, ByVal tamano As Integer)
    With ws.Cells(fila, 2)
        .Value = texto
        .Font.Bold = True
        .Font.Size = tamano
        .Font.Color = RGB(31, 56, 100)
    End With
    ws.Rows(fila).RowHeight = tamano * 1.8
    If tamano < 15 Then
        With ws.Range(ws.Cells(fila, 2), ws.Cells(fila, 5)).Borders(xlEdgeBottom)
            .Color = RGB(47, 117, 181)
            .Weight = xlMedium
        End With
    End If
    fila = fila + 1
End Sub

Private Sub InfParrafo(ws As Worksheet, ByRef fila As Long, texto As String)
    With ws.Cells(fila, 3)
        .Value = texto
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    ws.Rows(fila).AutoFit
    fila = fila + 1
End Sub

' Viñeta: punto azul o, si se indica un color, un cuadrado de semáforo
Private Sub InfVineta(ws As Worksheet, ByRef fila As Long, texto As String, _
                      Optional ByVal colorMarca As Long = -1)
    If colorMarca = -1 Then
        With ws.Cells(fila, 2)
            .Value = "•"
            .Font.Bold = True
            .Font.Color = RGB(47, 117, 181)
            .HorizontalAlignment = xlCenter
        End With
    Else
        ws.Cells(fila, 2).Interior.Color = colorMarca
    End If
    ws.Cells(fila, 2).VerticalAlignment = xlTop
    With ws.Cells(fila, 3)
        .Value = texto
        .WrapText = True
        .VerticalAlignment = xlTop
        .IndentLevel = 1
    End With
    ws.Rows(fila).AutoFit
    fila = fila + 1
End Sub

Private Sub InfCabecera(ws As Worksheet, ByRef fila As Long, titulos As Variant)
    With ws.Range(ws.Cells(fila, 3), ws.Cells(fila, 3 + UBound(titulos)))
        .Value = titulos
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 56, 100)
    End With
    ws.Range(ws.Cells(fila, 4), ws.Cells(fila, 3 + UBound(titulos))).HorizontalAlignment = xlRight
    fila = fila + 1
End Sub

Private Sub InfFilaPuente(ws As Worksheet, ByRef fila As Long, concepto As String, _
                          ByVal importe As Double, ByVal esTotal As Boolean)
    ws.Cells(fila, 3).Value = concepto
    ws.Cells(fila, 4).Value = importe
    If esTotal Then
        ws.Cells(fila, 4).NumberFormat = "#,##0 ""€"""
        With ws.Range(ws.Cells(fila, 3), ws.Cells(fila, 5))
            .Font.Bold = True
            .Interior.Color = RGB(242, 246, 252)
        End With
    Else
        ws.Cells(fila, 3).IndentLevel = 2
        ws.Cells(fila, 4).NumberFormat = "+#,##0 ""€"";[Red]-#,##0 ""€"";0 ""€"""
        If Abs(importe) < 0.5 Then
            ws.Cells(fila, 5).Value = "Sin desviación"
        ElseIf importe > 0 Then
            ws.Cells(fila, 5).Value = "Favorable"
            ws.Cells(fila, 5).Font.Color = RGB(0, 128, 0)
        Else
            ws.Cells(fila, 5).Value = "Desfavorable"
            ws.Cells(fila, 5).Font.Color = RGB(192, 0, 0)
        End If
        ws.Cells(fila, 5).HorizontalAlignment = xlRight
    End If
    With ws.Range(ws.Cells(fila, 3), ws.Cells(fila, 5)).Borders(xlEdgeBottom)
        .Color = RGB(217, 225, 242)
        .Weight = xlThin
    End With
    fila = fila + 1
End Sub

Private Sub InfFilaPrevision(ws As Worksheet, ByRef fila As Long, concepto As String, _
                             ByVal importe As Double, ByVal variacion As Double, ByVal esPpto As Boolean)
    ws.Cells(fila, 3).Value = concepto
    ws.Cells(fila, 4).Value = importe
    ws.Cells(fila, 4).NumberFormat = "#,##0 ""€"""
    If esPpto Then
        ws.Range(ws.Cells(fila, 3), ws.Cells(fila, 5)).Interior.Color = RGB(242, 246, 252)
    Else
        ws.Cells(fila, 5).Value = variacion
        ws.Cells(fila, 5).NumberFormat = "+0.0%;[Red]-0.0%;0.0%"
    End If
    With ws.Range(ws.Cells(fila, 3), ws.Cells(fila, 5)).Borders(xlEdgeBottom)
        .Color = RGB(217, 225, 242)
        .Weight = xlThin
    End With
    fila = fila + 1
End Sub

' ---------------------------------------------------------------------
'  UTILIDADES DE TEXTO
' ---------------------------------------------------------------------
Private Function Eur(ByVal x As Double) As String
    Eur = Format(x, "#,##0") & " €"
End Function

Private Function EurSigno(ByVal x As Double) As String
    EurSigno = Format(x, "+#,##0;-#,##0;0") & " €"
End Function

Private Function Pct(ByVal x As Double) As String
    Pct = Format(x, "+0.0%;-0.0%;0.0%")
End Function

Private Function PuntosTexto(ByVal x As Double) As String
    PuntosTexto = Format(x * 100, "+0.0;-0.0;0.0") & " p.p."
End Function

' División que devuelve 0 en lugar de dar error si el divisor es 0
Private Function DivSegura(ByVal num As Double, ByVal den As Double) As Double
    If den <> 0 Then DivSegura = num / den
End Function

' Mismo criterio que ColorSemaforo, pero en palabras
Private Function EstadoSemaforo(ByVal desvPct As Double) As String
    If desvPct >= 0 Then
        EstadoSemaforo = "favorable"
    ElseIf desvPct >= -UMBRAL_VIGILAR Then
        EstadoSemaforo = "a vigilar"
    Else
        EstadoSemaforo = "desfavorable"
    End If
End Function
