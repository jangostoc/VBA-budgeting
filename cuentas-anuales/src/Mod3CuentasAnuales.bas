Attribute VB_Name = "Mod3CuentasAnuales"
Option Explicit

' ---- Macro 1: crea la tabla de mapeo cuenta -> masa -> epígrafe ----
Sub CrearMapeo()
    Dim ws As Worksheet
    Dim datos As Variant
    Dim i As Long

    Set ws = CrearHoja("Mapeo")

    ' Masas: ANC, AC (activo) / PN, PNC, PC (pasivo) / PYG (resultados)
    datos = Array( _
        Array(211, "ANC", "Inmovilizado material"), _
        Array(217, "ANC", "Inmovilizado material"), _
        Array(281, "ANC", "Inmovilizado material"), _
        Array(300, "AC", "Existencias"), _
        Array(430, "AC", "Deudores comerciales y otras cuentas a cobrar"), _
        Array(572, "AC", "Efectivo y otros activos líquidos equivalentes"), _
        Array(100, "PN", "Capital"), _
        Array(112, "PN", "Reservas"), _
        Array(170, "PNC", "Deudas a largo plazo"), _
        Array(400, "PC", "Acreedores comerciales y otras cuentas a pagar"), _
        Array(475, "PC", "Acreedores comerciales y otras cuentas a pagar"), _
        Array(700, "PYG", "Importe neto de la cifra de negocios"), _
        Array(600, "PYG", "Aprovisionamientos"), _
        Array(640, "PYG", "Gastos de personal"), _
        Array(642, "PYG", "Gastos de personal"), _
        Array(621, "PYG", "Otros gastos de explotación"), _
        Array(681, "PYG", "Amortización del inmovilizado"))

    ws.Range("A1:C1").Value = Array("Cuenta", "Masa", "Epígrafe")
    For i = 0 To UBound(datos)
        ws.Cells(i + 2, 1).Value = datos(i)(0)
        ws.Cells(i + 2, 2).Value = datos(i)(1)
        ws.Cells(i + 2, 3).Value = datos(i)(2)
    Next i

    ws.Range("A1:C1").Font.Bold = True
    ws.Columns("A:C").AutoFit
End Sub

' ---- Escribe una masa (ej. ACTIVO CORRIENTE) con sus epígrafes y devuelve su total ----
Function EscribirMasa(ws As Worksheet, ByRef fila As Long, titulo As String, _
                      masa As String, signo As Integer, _
                      masaDeEpi As Object, totales As Object) As Double
    Dim k As Variant
    Dim total As Double
    Dim filaTitulo As Long

    filaTitulo = fila
    ws.Cells(fila, 1).Value = titulo
    fila = fila + 1

    For Each k In totales.Keys
        If masaDeEpi(k) = masa Then
            ws.Cells(fila, 1).Value = "     " & k
            ws.Cells(fila, 2).Value = totales(k) * signo
            total = total + totales(k) * signo
            fila = fila + 1
        End If
    Next k

    ws.Cells(filaTitulo, 2).Value = total
    ws.Range(ws.Cells(filaTitulo, 1), ws.Cells(filaTitulo, 2)).Font.Bold = True
    EscribirMasa = total
End Function

' ---- Macro 2: genera balance y PyG con formato de cuentas anuales ----
Sub GenerarCuentasAnuales()
    Dim wsSS As Worksheet, wsMap As Worksheet, ws As Worksheet
    Dim mapEpi As Object        ' cuenta    -> epígrafe
    Dim masaDeEpi As Object     ' epígrafe  -> masa
    Dim totales As Object       ' epígrafe  -> importe acumulado
    Dim i As Long, ultimaFila As Long, fila As Long
    Dim cuenta As String, epi As String, sinMapear As String
    Dim k As Variant
    Dim resultado As Double, totalActivo As Double, totalPasivo As Double

    If Not HojaExiste("SumasSaldos") Or Not HojaExiste("Mapeo") Then
        MsgBox "Faltan hojas. Ejecuta antes CrearDatosEjemplo y CrearMapeo.", vbExclamation
        Exit Sub
    End If

    Set mapEpi = CreateObject("Scripting.Dictionary")
    Set masaDeEpi = CreateObject("Scripting.Dictionary")
    Set totales = CreateObject("Scripting.Dictionary")

    ' 1. Cargar el mapeo en memoria
    Set wsMap = ThisWorkbook.Sheets("Mapeo")
    ultimaFila = wsMap.Cells(wsMap.Rows.Count, "A").End(xlUp).Row
    For i = 2 To ultimaFila
        cuenta = CStr(wsMap.Cells(i, "A").Value)
        epi = wsMap.Cells(i, "C").Value
        mapEpi(cuenta) = epi
        masaDeEpi(epi) = wsMap.Cells(i, "B").Value
        totales(epi) = 0
    Next i

    ' 2. Acumular los saldos en su epígrafe
    Set wsSS = ThisWorkbook.Sheets("SumasSaldos")
    ultimaFila = wsSS.Cells(wsSS.Rows.Count, "A").End(xlUp).Row
    For i = 2 To ultimaFila
        cuenta = Left(CStr(wsSS.Cells(i, "A").Value), 3)
        If mapEpi.Exists(cuenta) Then
            epi = mapEpi(cuenta)
            totales(epi) = totales(epi) + wsSS.Cells(i, "C").Value
        Else
            sinMapear = sinMapear & cuenta & " "
        End If
    Next i

    ' Control: ninguna cuenta puede quedarse fuera
    If sinMapear <> "" Then
        MsgBox "Cuentas sin mapear: " & sinMapear & vbCrLf & _
               "Añádelas a la hoja Mapeo y vuelve a ejecutar.", vbCritical
        Exit Sub
    End If

    ' 3. Calcular el resultado y llevarlo al patrimonio neto
    For Each k In totales.Keys
        If masaDeEpi(k) = "PYG" Then resultado = resultado + totales(k)
    Next k
    masaDeEpi("Resultado del ejercicio") = "PN"
    totales("Resultado del ejercicio") = resultado

    ' 4. Escribir los estados
    Set ws = CrearHoja("CuentasAnuales")

    ws.Range("A1").Value = "BALANCE"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    fila = 3
    totalActivo = EscribirMasa(ws, fila, "A) ACTIVO NO CORRIENTE", "ANC", 1, masaDeEpi, totales)
    totalActivo = totalActivo + EscribirMasa(ws, fila, "B) ACTIVO CORRIENTE", "AC", 1, masaDeEpi, totales)
    ws.Cells(fila, 1).Value = "TOTAL ACTIVO (A+B)"
    ws.Cells(fila, 2).Value = totalActivo
    ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 2)).Font.Bold = True

    fila = fila + 2
    totalPasivo = EscribirMasa(ws, fila, "A) PATRIMONIO NETO", "PN", -1, masaDeEpi, totales)
    totalPasivo = totalPasivo + EscribirMasa(ws, fila, "B) PASIVO NO CORRIENTE", "PNC", -1, masaDeEpi, totales)
    totalPasivo = totalPasivo + EscribirMasa(ws, fila, "C) PASIVO CORRIENTE", "PC", -1, masaDeEpi, totales)
    ws.Cells(fila, 1).Value = "TOTAL PATRIMONIO NETO Y PASIVO (A+B+C)"
    ws.Cells(fila, 2).Value = totalPasivo
    ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 2)).Font.Bold = True

    fila = fila + 3
    ws.Cells(fila, 1).Value = "CUENTA DE PÉRDIDAS Y GANANCIAS"
    ws.Cells(fila, 1).Font.Bold = True
    ws.Cells(fila, 1).Font.Size = 14
    fila = fila + 2
    EscribirMasa ws, fila, "A) RESULTADO DE EXPLOTACIÓN", "PYG", -1, masaDeEpi, totales
    ws.Cells(fila, 1).Value = "RESULTADO DEL EJERCICIO"
    ws.Cells(fila, 2).Value = -resultado
    ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 2)).Font.Bold = True

    ' Formato
    ws.Columns("B").NumberFormat = "#,##0;[Red]-#,##0"
    ws.Columns("A").AutoFit
    ws.Columns("B").ColumnWidth = 14

    ' Control de cuadre
    If Abs(totalActivo - totalPasivo) < 0.01 Then
        MsgBox "Cuentas anuales generadas. Activo = PN + Pasivo.", vbInformation
    Else
        MsgBox "¡Descuadre de " & Format(totalActivo - totalPasivo, "#,##0.00") & "!", vbCritical
    End If
End Sub
