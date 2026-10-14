Attribute VB_Name = "Mod1Utilidades"
Option Explicit

' ---- Borra una hoja si existe y la crea de nuevo ----
Function CrearHoja(nombre As String) As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(nombre).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set CrearHoja = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    CrearHoja.Name = nombre
End Function

' ---- Devuelve True si la hoja existe ----
Function HojaExiste(nombre As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nombre)
    On Error GoTo 0
    HojaExiste = Not ws Is Nothing
End Function
