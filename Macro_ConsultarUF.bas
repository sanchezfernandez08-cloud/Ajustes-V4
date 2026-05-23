Attribute VB_Name = "ModUF"
Sub ConsultarUF()
    Dim ws As Worksheet
    Dim fechaCell As Range
    Dim ufCell As Range
    Dim fechaVal As String
    Dim url As String
    Dim http As Object
    Dim resp As String
    Dim pos1 As Long
    Dim pos2 As Long
    Dim ufVal As Double
    Dim shNames As Variant
    shNames = Array("Pérdida Edificio", "Pérdida Contenidos", "Pérdida Maquinaria", "Pérdida Directo")
    Dim found As Boolean
    found = False
    Dim i As Integer
    For i = 0 To UBound(shNames)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(shNames(i))
        On Error GoTo 0
        If Not ws Is Nothing Then
            Set fechaCell = ws.Range("E15")
            Set ufCell = ws.Range("N15")
            If Not IsEmpty(fechaCell.Value) And CStr(fechaCell.Value) <> "" Then
                On Error Resume Next
                Dim fd As Date
                fd = CDate(fechaCell.Value)
                On Error GoTo 0
                If fd <> 0 Then
                    fechaVal = Format(fd, "DD-MM-YYYY")
                    url = "https://mindicador.cl/api/uf/" & fechaVal
                    Set http = CreateObject("MSXML2.XMLHTTP")
                    On Error GoTo ErrConn
                    http.Open "GET", url, False
                    http.setRequestHeader "User-Agent", "Mozilla/5.0"
                    http.Send
                    If http.Status = 200 Then
                        resp = http.responseText
                        pos1 = InStr(resp, Chr(34) & "valor" & Chr(34) & ":") + 8
                        pos2 = InStr(pos1, resp, ",")
                        If pos2 = 0 Then pos2 = InStr(pos1, resp, "}")
                        If pos1 > 8 And pos2 > pos1 Then
                            ufVal = CDbl(Mid(resp, pos1, pos2 - pos1))
                            ufCell.Value = ufVal
                            ufCell.NumberFormat = "#,##0.00"
                            found = True
                        End If
                    End If
                End If
            End If
            Set ws = Nothing
        End If
    Next i
    If found Then
        MsgBox "Valor UF actualizado correctamente en celda N15 de cada hoja.", vbInformation, "UF Actualizada"
    Else
        MsgBox "Ingrese la fecha del siniestro en E15 (formato DD-MM-AAAA) y vuelva a intentar.", vbExclamation, "Sin resultado"
    End If
    Exit Sub
ErrConn:
    MsgBox "Sin conexion a internet. Ingrese el valor UF manualmente en la celda N15.", vbCritical, "Error de red"
End Sub
