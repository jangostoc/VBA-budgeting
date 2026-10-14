Attribute VB_Name = "Mod4Analisis"
Option Explicit

' ---- Datos de ejemplo del ejercicio anterior ----
Sub CrearDatos2024()
    Dim ws As Worksheet
    Dim datos As Variant
    Dim i As Long

    Set ws = CrearHoja("SumasSaldos2024")

    datos = Array( _
        Array(100, "Capital social", -300000), _
        Array(112, "Reserva legal", -40000), _
        Array(170, "Deudas l/p con entidades de crédito", -250000), _
        Array(211, "Construcciones", 400000), _
        Array(217, "Equipos para procesos de información", 40000), _
        Array(281, "Amort. acumulada inmovilizado material", -60000), _
        Array(300, "Mercaderías", 100000), _
        Array(400, "Proveedores", -130000), _
        Array(430, "Clientes", 150000), _
        Array(475, "HP acreedora por conceptos fiscales", -15000), _
        Array(572, "Bancos", 240000), _
        Array(600, "Compras de mercaderías", 600000), _
        Array(621, "Arrendamientos", 38000), _
        Array(640, "Sueldos y salarios", 230000), _
        Array(642, "Seguridad Social a cargo de la empresa", 72000), _
        Array(681, "Amortización inmovilizado material", 25000), _
        Array(700, "Ventas de mercaderías", -1100000))

    ws.Range("A1:C1").Value = Array("Cuenta", "Descripción", "Saldo")
    For i = 0 To UBound(datos)
        ws.Cells(i + 2, 1).Value = datos(i)(0)
        ws.Cells(i + 2, 2).Value = datos(i)(1)
        ws.Cells(i + 2, 3).Value = datos(i)(2)
    Next i

    ws.Range("A1:C1").Font.Bold = True
    ws.Columns("C").NumberFormat = "#,##0;[Red]-#,##0"
    ws.Columns("A:C").AutoFit
End Sub

' ---- Carga la hoja Mapeo en dos diccionarios ----
Sub CargarMapeo(mapEpi As Object, masaDeEpi As Object)
    Dim wsMap As Worksheet
    Dim i As Long, ultimaFila As Long
    Dim epi As String

    Set wsMap = ThisWorkbook.Sheets("Mapeo")
    ultimaFila = wsMap.Cells(wsMap.Rows.Count, "A").End(xlUp).Row

    For i = 2 To ultimaFila
        epi = wsMap.Cells(i, "C").Value
        mapEpi(CStr(wsMap.Cells(i, "A").Value)) = epi
        masaDeEpi(epi) = wsMap.Cells(i, "B").Value
    Next i
    masaDeEpi("Resultado del ejercicio") = "PN"
End Sub

' ---- Suma los saldos de una hoja por epígrafe (Nothing si hay cuentas sin mapear) ----
Function CalcularTotales(nombreHoja As String, mapEpi As Object, masaDeEpi As Object) As Object
    Dim ws As Worksheet
    Dim totales As Object
    Dim k As Variant
    Dim i As Long, ultimaFila As Long
    Dim cuenta As String, sinMapear As String
    Dim resultado As Double

    ' Todos los epígrafes empiezan a cero, en el orden del Mapeo
    Set totales = CreateObject("Scripting.Dictionary")
    For Each k In masaDeEpi.Keys
        totales(k) = 0
    Next k

    Set ws = ThisWorkbook.Sheets(nombreHoja)
    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 2 To ultimaFila
        cuenta = Left(CStr(ws.Cells(i, "A").Value), 3)
        If mapEpi.Exists(cuenta) Then
            totales(mapEpi(cuenta)) = totales(mapEpi(cuenta)) + ws.Cells(i, "C").Value
        Else
            sinMapear = sinMapear & cuenta & " "
        End If
    Next i

    If sinMapear <> "" Then
        MsgBox "En " & nombreHoja & " hay cuentas sin mapear: " & sinMapear, vbCritical
        Set CalcularTotales = Nothing
        Exit Function
    End If

    For Each k In totales.Keys
        If masaDeEpi(k) = "PYG" Then resultado = resultado + totales(k)
    Next k
    totales("Resultado del ejercicio") = resultado

    Set CalcularTotales = totales
End Function

' ---- Total de una masa, con el signo de presentación ----
Function TotalMasa(masa As String, signo As Integer, masaDeEpi As Object, totales As Object) As Double
    Dim k As Variant
    For Each k In totales.Keys
        If masaDeEpi(k) = masa Then TotalMasa = TotalMasa + totales(k) * signo
    Next k
End Function

' ---- Escribe una fila: año actual, año anterior, variaciones y pesos ----
Sub EscribirFila(ws As Worksheet, fila As Long, texto As String, _
                 v1 As Double, v0 As Double, base1 As Double, base0 As Double, negrita As Boolean)
    ws.Cells(fila, 1).Value = texto
    ws.Cells(fila, 2).Value = v1
    ws.Cells(fila, 3).Value = v0
    ws.Cells(fila, 4).Value = v1 - v0                                ' Var. €
    If v0 <> 0 Then ws.Cells(fila, 5).Value = (v1 - v0) / Abs(v0)   ' Var. %
    If base1 <> 0 Then ws.Cells(fila, 6).Value = v1 / base1         ' Vertical año actual
    If base0 <> 0 Then ws.Cells(fila, 7).Value = v0 / base0         ' Vertical año anterior
    ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 7)).Font.Bold = negrita
End Sub

' ---- Escribe una masa y sus epígrafes en formato comparativo ----
Sub EscribirMasaComparada(ws As Worksheet, ByRef fila As Long, titulo As String, _
                          masa As String, signo As Integer, masaDeEpi As Object, _
                          tot1 As Object, tot0 As Object, base1 As Double, base0 As Double)
    Dim k As Variant

    EscribirFila ws, fila, titulo, TotalMasa(masa, signo, masaDeEpi, tot1), _
                 TotalMasa(masa, signo, masaDeEpi, tot0), base1, base0, True
    fila = fila + 1

    For Each k In tot1.Keys
        If masaDeEpi(k) = masa Then
            EscribirFila ws, fila, "     " & k, tot1(k) * signo, tot0(k) * signo, base1, base0, False
            fila = fila + 1
        End If
    Next k
End Sub

' ---- Escribe un ratio de los dos años ----
Sub EscribirRatio(ws As Worksheet, ByRef fila As Long, nombre As String, _
                  r1 As Double, r0 As Double, formato As String)
    ws.Cells(fila, 1).Value = nombre
    ws.Cells(fila, 2).Value = r1
    ws.Cells(fila, 3).Value = r0
    ws.Cells(fila, 4).Value = r1 - r0
    ws.Range(ws.Cells(fila, 2), ws.Cells(fila, 4)).NumberFormat = formato
    fila = fila + 1
End Sub

' ==== MACRO PRINCIPAL ====
Sub AnalisisComparativo()
    Dim mapEpi As Object, masaDeEpi As Object
    Dim tot1 As Object, tot0 As Object          ' 1 = 2025, 0 = 2024
    Dim ws As Worksheet
    Dim fila As Long, filaFinEstados As Long
    Dim ac1 As Double, ac0 As Double, anc1 As Double, anc0 As Double
    Dim pn1 As Double, pn0 As Double, pnc1 As Double, pnc0 As Double
    Dim pc1 As Double, pc0 As Double
    Dim activo1 As Double, activo0 As Double
    Dim ventas1 As Double, ventas0 As Double, res1 As Double, res0 As Double

    If Not HojaExiste("SumasSaldos") Or Not HojaExiste("SumasSaldos2024") _
       Or Not HojaExiste("Mapeo") Then
        MsgBox "Faltan hojas: SumasSaldos, SumasSaldos2024 o Mapeo.", vbExclamation
        Exit Sub
    End If

    ' 1. Cargar mapeo y totales de cada año
    Set mapEpi = CreateObject("Scripting.Dictionary")
    Set masaDeEpi = CreateObject("Scripting.Dictionary")
    CargarMapeo mapEpi, masaDeEpi

    Set tot1 = CalcularTotales("SumasSaldos", mapEpi, masaDeEpi)
    If tot1 Is Nothing Then Exit Sub
    Set tot0 = CalcularTotales("SumasSaldos2024", mapEpi, masaDeEpi)
    If tot0 Is Nothing Then Exit Sub

    ' 2. Magnitudes clave de cada año
    anc1 = TotalMasa("ANC", 1, masaDeEpi, tot1):  anc0 = TotalMasa("ANC", 1, masaDeEpi, tot0)
    ac1 = TotalMasa("AC", 1, masaDeEpi, tot1):    ac0 = TotalMasa("AC", 1, masaDeEpi, tot0)
    pn1 = TotalMasa("PN", -1, masaDeEpi, tot1):   pn0 = TotalMasa("PN", -1, masaDeEpi, tot0)
    pnc1 = TotalMasa("PNC", -1, masaDeEpi, tot1): pnc0 = TotalMasa("PNC", -1, masaDeEpi, tot0)
    pc1 = TotalMasa("PC", -1, masaDeEpi, tot1):   pc0 = TotalMasa("PC", -1, masaDeEpi, tot0)
    activo1 = anc1 + ac1:                          activo0 = anc0 + ac0
    ventas1 = -tot1("Importe neto de la cifra de negocios")
    ventas0 = -tot0("Importe neto de la cifra de negocios")
    res1 = -tot1("Resultado del ejercicio"):      res0 = -tot0("Resultado del ejercicio")

    ' 3. Preparar la hoja
    Set ws = CrearHoja("Analisis")
    ws.Columns("B:D").NumberFormat = "#,##0;[Red]-#,##0"
    ws.Columns("E:G").NumberFormat = "0.0%"

    ws.Range("A1").Value = "ANÁLISIS COMPARATIVO 2025 - 2024"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ws.Range("A3:G3").Value = Array("Epígrafe", "2025", "2024", "Var. €", "Var. %", _
                                    "% s/total 2025", "% s/total 2024")
    ws.Range("B3:C3").NumberFormat = "0"
    With ws.Range("A3:G3")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
    End With

    ' 4. Balance (vertical sobre total activo)
    fila = 4
    EscribirMasaComparada ws, fila, "A) ACTIVO NO CORRIENTE", "ANC", 1, masaDeEpi, tot1, tot0, activo1, activo0
    EscribirMasaComparada ws, fila, "B) ACTIVO CORRIENTE", "AC", 1, masaDeEpi, tot1, tot0, activo1, activo0
    EscribirFila ws, fila, "TOTAL ACTIVO", activo1, activo0, activo1, activo0, True
    fila = fila + 2
    EscribirMasaComparada ws, fila, "A) PATRIMONIO NETO", "PN", -1, masaDeEpi, tot1, tot0, activo1, activo0
    EscribirMasaComparada ws, fila, "B) PASIVO NO CORRIENTE", "PNC", -1, masaDeEpi, tot1, tot0, activo1, activo0
    EscribirMasaComparada ws, fila, "C) PASIVO CORRIENTE", "PC", -1, masaDeEpi, tot1, tot0, activo1, activo0
    EscribirFila ws, fila, "TOTAL PN Y PASIVO", pn1 + pnc1 + pc1, pn0 + pnc0 + pc0, activo1, activo0, True

    ' 5. PyG (vertical sobre ventas)
    fila = fila + 2
    EscribirMasaComparada ws, fila, "RESULTADO DE EXPLOTACIÓN", "PYG", -1, masaDeEpi, tot1, tot0, ventas1, ventas0
    EscribirFila ws, fila, "RESULTADO DEL EJERCICIO", res1, res0, ventas1, ventas0, True
    filaFinEstados = fila

    ' Semáforo en Var. %: rojo si baja, verde si sube
    With ws.Range("E4:E" & filaFinEstados).FormatConditions
        .Delete
        .Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0").Font.Color = RGB(192, 0, 0)
        .Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0").Font.Color = RGB(0, 128, 0)
    End With

    ' 6. Ratios
    fila = fila + 3
    ws.Cells(fila, 1).Value = "RATIOS"
    ws.Cells(fila, 1).Font.Bold = True
    ws.Cells(fila, 1).Font.Size = 14
    fila = fila + 1
    ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 4)).Value = Array("Ratio", "2025", "2024", "Var.")
    ws.Range(ws.Cells(fila, 2), ws.Cells(fila, 3)).NumberFormat = "0"
    With ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 4))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
    End With
    fila = fila + 1

    EscribirRatio ws, fila, "Liquidez (AC / PC)", ac1 / pc1, ac0 / pc0, "0.00"
    EscribirRatio ws, fila, "Prueba ácida ((AC - Existencias) / PC)", _
        (ac1 - tot1("Existencias")) / pc1, (ac0 - tot0("Existencias")) / pc0, "0.00"
    EscribirRatio ws, fila, "Endeudamiento (Pasivo / PN)", (pnc1 + pc1) / pn1, (pnc0 + pc0) / pn0, "0.00"
    EscribirRatio ws, fila, "Solvencia (Activo / Pasivo)", activo1 / (pnc1 + pc1), activo0 / (pnc0 + pc0), "0.00"
    EscribirRatio ws, fila, "ROE (Resultado / PN)", res1 / pn1, res0 / pn0, "0.0%"
    EscribirRatio ws, fila, "ROA (Resultado / Activo)", res1 / activo1, res0 / activo0, "0.0%"
    EscribirRatio ws, fila, "Margen neto (Resultado / Ventas)", res1 / ventas1, res0 / ventas0, "0.0%"

    ' 7. Acabado
    ws.Columns("A").AutoFit
    ws.Columns("B:G").ColumnWidth = 14
    ws.Activate
    ActiveWindow.DisplayGridlines = False

    MsgBox "Análisis generado.", vbInformation
End Sub

' ==================================================================
' CONCLUSIONES: redacta el análisis con las cifras calculadas
' ==================================================================

' ---- Utilidades de formato para la hoja de conclusiones ----
Sub EscribirTitulo(ws As Worksheet, ByRef fila As Long, texto As String, tamano As Integer)
    ws.Cells(fila, 2).Value = texto
    ws.Cells(fila, 2).Font.Bold = True
    ws.Cells(fila, 2).Font.Size = tamano
    ws.Cells(fila, 2).Font.Color = RGB(31, 78, 121)
    fila = fila + 1
End Sub

Sub EscribirParrafo(ws As Worksheet, ByRef fila As Long, texto As String)
    With ws.Cells(fila, 2)
        .Value = texto
        .WrapText = True                ' Ajustar texto
        .VerticalAlignment = xlTop
    End With
    ws.Rows(fila).AutoFit
    fila = fila + 1
End Sub

Sub EscribirCabecera(ws As Worksheet, ByRef fila As Long, titulos As Variant)
    With ws.Range(ws.Cells(fila, 2), ws.Cells(fila, 2 + UBound(titulos)))
        .Value = titulos
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
    End With
    fila = fila + 1
End Sub

' ---- Macro principal ----
Sub GenerarConclusiones()
    Const INCN As String = "Importe neto de la cifra de negocios"
    Const CAJA As String = "Efectivo y otros activos líquidos equivalentes"
    Const AMORT As String = "Amortización del inmovilizado"

    Dim mapEpi As Object, masaDeEpi As Object, tot1 As Object, tot0 As Object
    Dim ws As Worksheet
    Dim fila As Long, filaIni As Long, i As Long
    Dim k As Variant, conceptos As Variant, importes As Variant
    Dim ventas1 As Double, ventas0 As Double, res1 As Double, res0 As Double
    Dim anc1 As Double, anc0 As Double, ac1 As Double, ac0 As Double
    Dim pn1 As Double, pn0 As Double, pnc1 As Double, pnc0 As Double
    Dim pc1 As Double, pc0 As Double, caja1 As Double, caja0 As Double
    Dim amort1 As Double, dividendos As Double, capex As Double, varCirculante As Double
    Dim efecto As Double, peorEfecto As Double, mejorEfecto As Double
    Dim peor As String, mejor As String
    Dim sumaPuente As Double, mayorSalida As Double, conceptoMayorSalida As String

    If Not HojaExiste("SumasSaldos") Or Not HojaExiste("SumasSaldos2024") _
       Or Not HojaExiste("Mapeo") Then
        MsgBox "Faltan hojas: SumasSaldos, SumasSaldos2024 o Mapeo.", vbExclamation
        Exit Sub
    End If

    ' 1. Cargar datos (reutilizamos las funciones de la lección 5)
    Set mapEpi = CreateObject("Scripting.Dictionary")
    Set masaDeEpi = CreateObject("Scripting.Dictionary")
    CargarMapeo mapEpi, masaDeEpi
    Set tot1 = CalcularTotales("SumasSaldos", mapEpi, masaDeEpi)
    If tot1 Is Nothing Then Exit Sub
    Set tot0 = CalcularTotales("SumasSaldos2024", mapEpi, masaDeEpi)
    If tot0 Is Nothing Then Exit Sub

    ' 2. Magnitudes
    ventas1 = -tot1(INCN):                         ventas0 = -tot0(INCN)
    res1 = -tot1("Resultado del ejercicio"):       res0 = -tot0("Resultado del ejercicio")
    anc1 = TotalMasa("ANC", 1, masaDeEpi, tot1):   anc0 = TotalMasa("ANC", 1, masaDeEpi, tot0)
    ac1 = TotalMasa("AC", 1, masaDeEpi, tot1):     ac0 = TotalMasa("AC", 1, masaDeEpi, tot0)
    pn1 = TotalMasa("PN", -1, masaDeEpi, tot1):    pn0 = TotalMasa("PN", -1, masaDeEpi, tot0)
    pnc1 = TotalMasa("PNC", -1, masaDeEpi, tot1):  pnc0 = TotalMasa("PNC", -1, masaDeEpi, tot0)
    pc1 = TotalMasa("PC", -1, masaDeEpi, tot1):    pc0 = TotalMasa("PC", -1, masaDeEpi, tot0)
    caja1 = tot1(CAJA):                            caja0 = tot0(CAJA)
    amort1 = tot1(AMORT)

    ' Componentes del flujo de caja
    dividendos = pn0 - (pn1 - res1)                 ' PN que no explica el resultado
    capex = (anc1 - anc0) + amort1                  ' Inversión bruta en inmovilizado
    varCirculante = -((ac1 - caja1) - (ac0 - caja0)) + (pc1 - pc0)

    ' 3. Preparar la hoja
    Set ws = CrearHoja("Conclusiones")
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 80
    ws.Columns("C:E").ColumnWidth = 16
    ws.Activate
    ActiveWindow.DisplayGridlines = False

    fila = 2
    EscribirTitulo ws, fila, "CONCLUSIONES DEL ANÁLISIS 2025 - 2024", 16
    fila = fila + 1

    ' ---- Apartado 1: margen ----
    EscribirTitulo ws, fila, "1. Evolución del margen neto", 12
    EscribirParrafo ws, fila, "Las ventas crecen un " & Format(ventas1 / ventas0 - 1, "0.0%") & _
        ", pero el margen neto pasa del " & Format(res0 / ventas0, "0.0%") & _
        " al " & Format(res1 / ventas1, "0.0%") & ". Peso de cada gasto sobre las ventas:"

    EscribirCabecera ws, fila, Array("Gasto", "% s/ventas 2024", "% s/ventas 2025", "Efecto en margen")
    filaIni = fila
    For Each k In tot1.Keys
        If masaDeEpi(k) = "PYG" And k <> INCN Then
            efecto = -(tot1(k) / ventas1 - tot0(k) / ventas0)
            ws.Cells(fila, 2).Value = k
            ws.Cells(fila, 3).Value = tot0(k) / ventas0
            ws.Cells(fila, 4).Value = tot1(k) / ventas1
            ws.Cells(fila, 5).Value = efecto

            ' Detectar el gasto que más resta y el que más suma
            If peor = "" Or efecto < peorEfecto Then
                peor = k
                peorEfecto = efecto
            End If
            If mejor = "" Or efecto > mejorEfecto Then
                mejor = k
                mejorEfecto = efecto
            End If
            fila = fila + 1
        End If
    Next k
    ws.Cells(fila, 2).Value = "Margen neto"
    ws.Cells(fila, 3).Value = res0 / ventas0
    ws.Cells(fila, 4).Value = res1 / ventas1
    ws.Cells(fila, 5).Value = res1 / ventas1 - res0 / ventas0
    ws.Range(ws.Cells(fila, 2), ws.Cells(fila, 5)).Font.Bold = True
    ws.Range(ws.Cells(filaIni, 3), ws.Cells(fila, 4)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(filaIni, 5), ws.Cells(fila, 5)).NumberFormat = "+0.0%;[Red]-0.0%"
    fila = fila + 2

    EscribirParrafo ws, fila, "El factor que más resta margen es «" & peor & "» (" & _
        Format(peorEfecto, "+0.0%;-0.0%") & "), porque su peso sobre las ventas aumenta."
    If mejorEfecto > 0 Then
        EscribirParrafo ws, fila, "Lo compensa parcialmente «" & mejor & "» (" & _
            Format(mejorEfecto, "+0.0%;-0.0%") & "), que crece menos que la facturación: " & _
            "se aprovechan economías de escala."
    End If
    If tot1("Existencias") > tot0("Existencias") Then
        EscribirParrafo ws, fila, "A revisar: las existencias aumentan " & _
            Format(tot1("Existencias") - tot0("Existencias"), "#,##0") & _
            " €, así que no todo lo comprado se ha consumido. Con datos reales habría que " & _
            "comprobar la variación de existencias (cuenta 610), que reduciría el coste."
    End If
    fila = fila + 1

' ---- Apartado 2: liquidez ----
    EscribirTitulo ws, fila, "2. Liquidez", 12
    EscribirParrafo ws, fila, "La liquidez (AC / PC) pasa de " & Format(ac0 / pc0, "0.00") & _
        " a " & Format(ac1 / pc1, "0.00") & ". El activo corriente varía en " & _
        Format(ac1 - ac0, "#,##0") & " € y el pasivo corriente en " & _
        Format(pc1 - pc0, "#,##0") & " €. Dentro del activo corriente, la caja pasa de " & _
        Format(caja0, "#,##0") & " € a " & Format(caja1, "#,##0") & " €."
    fila = fila + 1

    ' ---- Apartado 3: conciliación de la caja ----
    EscribirTitulo ws, fila, "3. ¿Por qué cambia la caja?", 12
    conceptos = Array("Resultado + amortización", _
                      "Variación del circulante (clientes, existencias, proveedores, HP)", _
                      "Inversión en inmovilizado", _
                      "Variación de la deuda a largo plazo", _
                      "Dividendos (estimados)")
    importes = Array(res1 + amort1, varCirculante, -capex, pnc1 - pnc0, -dividendos)

    EscribirCabecera ws, fila, Array("Concepto", "Importe")
    filaIni = fila
    For i = 0 To UBound(conceptos)
        ws.Cells(fila, 2).Value = conceptos(i)
        ws.Cells(fila, 3).Value = importes(i)
        sumaPuente = sumaPuente + importes(i)
        If importes(i) < mayorSalida Then
            mayorSalida = importes(i)
            conceptoMayorSalida = conceptos(i)
        End If
        fila = fila + 1
    Next i
    ws.Cells(fila, 2).Value = "Variación de caja"
    ws.Cells(fila, 3).Value = sumaPuente
    ws.Range(ws.Cells(fila, 2), ws.Cells(fila, 3)).Font.Bold = True
    ws.Range(ws.Cells(filaIni, 3), ws.Cells(fila, 3)).NumberFormat = "#,##0;[Red]-#,##0"
    fila = fila + 2

    ' Control: la conciliación debe coincidir con la caja real
    If Abs(sumaPuente - (caja1 - caja0)) < 0.01 Then
        EscribirParrafo ws, fila, "La conciliación cuadra con la variación real de la caja. " & _
            "La mayor salida es «" & conceptoMayorSalida & "» (" & _
            Format(-mayorSalida, "#,##0") & " €)."
    Else
        EscribirParrafo ws, fila, "Atención: la conciliación no cuadra con la caja real " & _
            "(diferencia de " & Format(sumaPuente - (caja1 - caja0), "#,##0") & _
            " €). Puede haber ampliaciones de capital u otros movimientos no contemplados."
        ws.Cells(fila - 1, 2).Font.Color = RGB(192, 0, 0)
    End If
    EscribirParrafo ws, fila, "Nota: los dividendos se estiman como la variación del patrimonio " & _
        "neto que no explica el resultado del ejercicio (se supone que no hay ampliaciones de capital)."

    MsgBox "Conclusiones generadas.", vbInformation
End Sub
