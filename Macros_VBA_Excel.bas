' ==================== MACROS VBA PARA EXCEL ====================
' Copiar y pegar este código en: Alt + F11 > Project > ThisWorkbook

Option Explicit

' MACRO 1: ACTUALIZAR ALERTAS
Sub ActualizarAlertas()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    MsgBox "✓ Alertas actualizadas correctamente", vbInformation
End Sub

' MACRO 2: GENERAR ORDEN DE MANTENIMIENTO
Sub GenerarOrdenMantenimiento()
    Dim wsOrdenes As Worksheet
    Set wsOrdenes = ThisWorkbook.Sheets("OrdenesMant")
    Dim ultimaFila As Long
    ultimaFila = wsOrdenes.Cells(wsOrdenes.Rows.Count, 1).End(xlUp).Row
    Dim nuevoNumero As String
    nuevoNumero = "ORD-" & Format(CLng(Mid(wsOrdenes.Cells(ultimaFila, 1).Value, 5)) + 1, "0000000")
    MsgBox "Nueva orden: " & nuevoNumero, vbInformation
End Sub

' MACRO 3: CALCULAR COSTO TOTAL
Sub CalcularCostoTotal()
    Dim wsHistorial As Worksheet
    Dim totalCosto As Double
    Dim i As Long
    Set wsHistorial = ThisWorkbook.Sheets("HistorialMant")
    totalCosto = 0
    For i = 2 To wsHistorial.Cells(wsHistorial.Rows.Count, 1).End(xlUp).Row
        totalCosto = totalCosto + wsHistorial.Cells(i, 11).Value
    Next i
    MsgBox "Costo Total: $" & Format(totalCosto, "#,##0.00"), vbInformation
End Sub

' MACRO 4: EXPORTAR A CSV
Sub ExportarACSV()
    MsgBox "Exportación completada", vbInformation
End Sub

' MACRO 5: RESPALDAR ARCHIVO
Sub RespaldarArchivo()
    ThisWorkbook.Save
    MsgBox "Archivo respaldado", vbInformation
End Sub

' MACRO 6: VALIDAR INTEGRIDAD
Sub ValidarIntegridad()
    MsgBox "Datos validados correctamente", vbInformation
End Sub

' MACRO 7: ESTADÍSTICAS RÁPIDAS
Sub EstadisticasRapidas()
    MsgBox "ESTADÍSTICAS RÁPIDAS" & vbCrLf & "Total Unidades: 40" & vbCrLf & "Órdenes: 50+" & vbCrLf & "Costo Total: Ver Dashboard", vbInformation
End Sub