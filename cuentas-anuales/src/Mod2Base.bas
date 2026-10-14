Attribute VB_Name = "Mod2Base"
Option Explicit

' ---- Utilidad: borra una hoja si existe y la crea de nuevo ----
Function CrearHoja(nombre As String) As Worksheet
    Application.DisplayAlerts = False       ' Que no pregunte "¿seguro que quiere eliminar?"
    On Error Resume Next                    ' Si la hoja no existe, ignora el error
    ThisWorkbook.Sheets(nombre).Delete
    On Error GoTo 0                         ' Vuelve a avisar de errores
    Application.DisplayAlerts = True

    Set CrearHoja = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    CrearHoja.Name = nombre
End Function

' ---- Utilidad: devuelve True si la hoja existe ----
Function HojaExiste(nombre As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nombre)
    On Error GoTo 0
    HojaExiste = Not ws Is Nothing
    
End Function

' ---- Macro 1: genera un balance de sumas y saldos de ejemplo ----
Sub CrearDatosEjemplo()
    Dim ws As Worksheet
    Dim datos As Variant
    Dim i As Long

    Set ws = CrearHoja("SumasSaldos")

    datos = Array( _
        Array(100, "Capital social", -300000), _
        Array(112, "Reserva legal", -60000), _
        Array(170, "Deudas l/p con entidades de crédito", -200000), _
        Array(211, "Construcciones", 400000), _
        Array(217, "Equipos para procesos de información", 50000), _
        Array(281, "Amort. acumulada inmovilizado material", -90000), _
        Array(300, "Mercaderías", 120000), _
        Array(400, "Proveedores", -150000), _
        Array(430, "Clientes", 180000), _
        Array(475, "HP acreedora por conceptos fiscales", -20000), _
        Array(572, "Bancos", 220000), _
        Array(600, "Compras de mercaderías", 700000), _
        Array(621, "Arrendamientos", 40000), _
        Array(640, "Sueldos y salarios", 250000), _
        Array(642, "Seguridad Social a cargo de la empresa", 80000), _
        Array(681, "Amortización inmovilizado material", 30000), _
        Array(700, "Ventas de mercaderías", -1250000))

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

' ---- Macro 2: agrupa los saldos por grupo del PGC ----
Sub ResumenPorGrupos()
    Dim wsDatos As Worksheet, wsRes As Worksheet
    Dim totales(1 To 7) As Double          ' Una "caja" por cada grupo del 1 al 7
    Dim nombres As Variant
    Dim ultimaFila As Long, i As Long
    Dim grupo As Integer
    Dim sumaTotal As Double

   If Not HojaExiste("SumasSaldos") Then
        MsgBox "No existe la hoja 'SumasSaldos'. Ejecuta primero CrearDatosEjemplo.", vbExclamation
   Exit Sub
   End If

    Set wsDatos = ThisWorkbook.Sheets("SumasSaldos")
    ultimaFila = wsDatos.Cells(wsDatos.Rows.Count, "A").End(xlUp).Row

    ' Recorremos las cuentas y sumamos cada saldo en la caja de su grupo
    For i = 2 To ultimaFila
        grupo = CInt(Left(CStr(wsDatos.Cells(i, "A").Value), 1))   ' Primer dígito
        totales(grupo) = totales(grupo) + wsDatos.Cells(i, "C").Value
    Next i

    ' Escribimos el resumen
    Set wsRes = CrearHoja("Resumen")
    nombres = Array("", "Financiación básica", "Inmovilizado", "Existencias", _
                    "Acreedores y deudores", "Cuentas financieras", _
                    "Compras y gastos", "Ventas e ingresos")

    wsRes.Range("A1:C1").Value = Array("Grupo", "Descripción", "Saldo")
    For grupo = 1 To 7
        wsRes.Cells(grupo + 1, 1).Value = grupo
        wsRes.Cells(grupo + 1, 2).Value = nombres(grupo)
        wsRes.Cells(grupo + 1, 3).Value = totales(grupo)
        sumaTotal = sumaTotal + totales(grupo)
    Next grupo

    ' Controles
    wsRes.Range("B10").Value = "Comprobación (debe ser 0)"
    wsRes.Range("C10").Value = sumaTotal
    wsRes.Range("B11").Value = "Resultado del ejercicio"
    wsRes.Range("C11").Value = -(totales(6) + totales(7))

    wsRes.Range("A1:C1").Font.Bold = True
    wsRes.Range("B10:C11").Font.Bold = True
    wsRes.Columns("C").NumberFormat = "#,##0;[Red]-#,##0"
    wsRes.Columns("A:C").AutoFit

    If Abs(sumaTotal) < 0.01 Then
        MsgBox "El balance cuadra. Resultado: " & Format(-(totales(6) + totales(7)), "#,##0"), vbInformation
    Else
        MsgBox "¡Descuadre de " & Format(sumaTotal, "#,##0.00") & "!", vbCritical
    End If
End Sub
