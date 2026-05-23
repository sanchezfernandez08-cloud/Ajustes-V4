Attribute VB_Name = "GenerarCarta"

' ═══════════════════════════════════════════════════════════════════
' A Plus Ajustadores SpA
' Macro: GenerarCartaAjuste
' Lee los datos de la planilla activa y genera una Carta de Ajuste
' en Word (.docx) y PDF en la misma carpeta del archivo Excel.
' ═══════════════════════════════════════════════════════════════════

Sub GenerarCartaAjuste()

    ' ── Verificar que Word esté instalado ──────────────────────────
    Dim oWord As Object
    On Error Resume Next
    Set oWord = CreateObject("Word.Application")
    On Error GoTo 0
    If oWord Is Nothing Then
        MsgBox "Microsoft Word no está instalado o no está disponible." & Chr(10) & _
               "Instale Word para usar esta función.", vbCritical, "Word no encontrado"
        Exit Sub
    End If

    ' ── Determinar hoja activa de trabajo ──────────────────────────
    Dim wsActiva As Worksheet
    Dim sheetName As String
    sheetName = ActiveSheet.Name

    ' Validar que sea una hoja de ajuste válida
    Dim validSheets As Variant
    validSheets = Array("Pérdida Edificio", "Pérdida Contenidos", "Pérdida Maquinaria", "Pérdida Directo")
    Dim esValida As Boolean: esValida = False
    Dim i As Integer
    For i = 0 To UBound(validSheets)
        If sheetName = validSheets(i) Then esValida = True
    Next i

    If Not esValida Then
        ' Usar Resumen para leer datos generales y todas las secciones
        Dim respuesta As Integer
        respuesta = MsgBox("Está en la hoja '" & sheetName & "'." & Chr(10) & Chr(10) & _
                           "La carta se generará con TODAS las secciones disponibles." & Chr(10) & _
                           "¿Desea continuar?", vbYesNo + vbQuestion, "Generar Carta Completa")
        If respuesta = vbNo Then Exit Sub
    End If

    ' ── Leer datos del encabezado del siniestro ────────────────────
    ' Los datos se leen de la hoja Edificio (o la activa si es ajuste)
    Dim wsData As Worksheet
    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets("Pérdida Edificio")
    On Error GoTo 0
    If wsData Is Nothing Then Set wsData = ActiveSheet

    Dim liquidador  As String: liquidador  = Trim(CStr(wsData.Cells(9, 3).Value))
    Dim asegurado   As String: asegurado   = Trim(CStr(wsData.Cells(10, 3).Value))
    Dim direccion   As String: direccion   = Trim(CStr(wsData.Cells(11, 3).Value))
    Dim fechaSin    As String: fechaSin    = Trim(CStr(wsData.Cells(12, 3).Value))
    Dim nroSin      As String: nroSin      = Trim(CStr(wsData.Cells(9, 12).Value))
    Dim nroPoliza   As String: nroPoliza   = Trim(CStr(wsData.Cells(10, 12).Value))
    Dim nroLiquid   As String: nroLiquid   = Trim(CStr(wsData.Cells(11, 12).Value))
    Dim evento      As String: evento      = Trim(CStr(wsData.Cells(12, 12).Value))
    Dim valorUF     As String: valorUF     = Trim(CStr(wsData.Cells(15, 14).Value))

    ' Defaults si están vacíos
    If liquidador = "" Or liquidador = "None" Then liquidador = "[Liquidador]"
    If asegurado  = "" Or asegurado  = "None" Then asegurado  = "[Asegurado]"
    If direccion  = "" Or direccion  = "None" Then direccion  = "[Dirección]"
    If fechaSin   = "" Or fechaSin   = "None" Then fechaSin   = "[Fecha del Siniestro]"
    If nroSin     = "" Or nroSin     = "None" Then nroSin     = "[N° Siniestro]"
    If nroPoliza  = "" Or nroPoliza  = "None" Then nroPoliza  = "[N° Póliza]"
    If nroLiquid  = "" Or nroLiquid  = "None" Then nroLiquid  = "[N° Liquidación]"
    If evento     = "" Or evento     = "None" Then evento     = "[Evento]"
    If valorUF    = "" Or valorUF    = "None" Or _
       InStr(valorUF, "Actualice") > 0 Then valorUF = "[Valor UF]"

    ' ── Recopilar secciones disponibles ───────────────────────────
    Dim secciones(3) As String, titulosSec(3) As String
    Dim numSec As Integer: numSec = 0
    Dim secNames(3) As String
    secNames(0) = "Pérdida Edificio"
    secNames(1) = "Pérdida Contenidos"
    secNames(2) = "Pérdida Maquinaria"
    secNames(3) = "Pérdida Directo"
    Dim secLabels(3) As String
    secLabels(0) = "EDIFICIO"
    secLabels(1) = "CONTENIDOS"
    secLabels(2) = "MAQUINARIA Y EQUIPOS"
    secLabels(3) = "AJUSTE DIRECTO"

    ' ── Leer pérdidas del Resumen ──────────────────────────────────
    Dim wsRes As Worksheet
    On Error Resume Next
    Set wsRes = ThisWorkbook.Sheets("Resumen de Pérdidas")
    On Error GoTo 0

    Dim perdidaTotal As String: perdidaTotal = ""
    If Not wsRes Is Nothing Then
        Dim vTotal As Variant
        vTotal = wsRes.Cells(13, 2).Value
        If Not IsError(vTotal) And vTotal <> "" And vTotal <> 0 Then
            perdidaTotal = Format(vTotal, "0.00") & " UF"
        End If
    End If

    ' ── Rutas de salida ────────────────────────────────────────────
    Dim carpeta As String
    carpeta = ThisWorkbook.Path
    If Right(carpeta, 1) <> "\" And Right(carpeta, 1) <> "/" Then carpeta = carpeta & "\"

    Dim sufijo As String
    sufijo = ""
    If nroSin <> "[N° Siniestro]" And nroSin <> "" Then sufijo = "_Sin" & nroSin

    Dim rutaWord As String: rutaWord = carpeta & "Carta_Ajuste" & sufijo & ".docx"
    Dim rutaPDF  As String: rutaPDF  = carpeta & "Carta_Ajuste" & sufijo & ".pdf"

    ' ── Crear documento Word ───────────────────────────────────────
    Dim oDoc As Object
    oWord.Visible = False
    Set oDoc = oWord.Documents.Add

    ' Configurar página A4
    With oDoc.PageSetup
        .PaperSize = 9          ' wdPaperA4
        .TopMargin    = oWord.InchesToPoints(1.18)
        .BottomMargin = oWord.InchesToPoints(0.98)
        .LeftMargin   = oWord.InchesToPoints(1.18)
        .RightMargin  = oWord.InchesToPoints(0.98)
    End With

    ' ── Estilos base ──────────────────────────────────────────────
    Dim AZUL_OSC As Long: AZUL_OSC = RGB(24, 51, 99)    ' #183363
    Dim AZUL     As Long: AZUL     = RGB(51, 95, 179)   ' #335FB3
    Dim NARANJA  As Long: NARANJA  = RGB(255, 118, 64)  ' #FF7640
    Dim CELESTE  As Long: CELESTE  = RGB(180, 205, 224) ' #B4CDE0
    Dim GRIS     As Long: GRIS     = RGB(175, 173, 168) ' #AFADA8

    ' ── Helper: agregar párrafo ───────────────────────────────────
    Dim oSel As Object
    Set oSel = oWord.Selection

    ' ── ENCABEZADO DEL DOCUMENTO ──────────────────────────────────
    ' Intentar insertar logo si existe
    Dim logoPath As String
    logoPath = carpeta & "logo_aplus.png"
    ' También buscar en rutas comunes
    If Dir(logoPath) = "" Then
        logoPath = ThisWorkbook.Path & "\logo_aplus.png"
    End If

    ' Header con tabla 2 columnas: logo | datos empresa
    Dim oHeader As Object
    Set oHeader = oDoc.Sections(1).Headers(1)  ' wdHeaderFooterPrimary = 1
    oHeader.Range.Delete

    Dim oHdrTable As Object
    Set oHdrTable = oHeader.Range.Tables.Add(oHeader.Range, 1, 2)

    With oHdrTable
        .Columns(1).Width = oWord.InchesToPoints(3)
        .Columns(2).Width = oWord.InchesToPoints(3.76)
        .Borders.Enable = False

        ' Celda izquierda: logo o texto empresa
        With .Cell(1, 1).Range
            .ParagraphFormat.Alignment = 1  ' wdAlignParagraphLeft
            If Dir(logoPath) <> "" Then
                .InlineShapes.AddPicture logoPath, False, True
                .InlineShapes(1).Width  = oWord.InchesToPoints(1.8)
                .InlineShapes(1).Height = oWord.InchesToPoints(1.1)
            Else
                .Text = "A PLUS AJUSTADORES"
                .Font.Bold = True
                .Font.Size = 14
                .Font.Color = AZUL_OSC
            End If
        End With

        ' Celda derecha: datos de contacto
        With .Cell(1, 2).Range
            .ParagraphFormat.Alignment = 2  ' wdAlignParagraphRight
            .Text = "A Plus Ajustadores SpA" & Chr(13) & _
                    "Marchant Pereira N° 367, Of. 204" & Chr(13) & _
                    "Providencia, Santiago" & Chr(13) & _
                    "+56 2 2 484 1515  |  aplusajustadores.cl"
            .Font.Size = 8
            .Font.Color = GRIS
        End With
    End With

    ' Línea separadora naranja bajo el encabezado
    Dim oBorderPara As Object
    Set oBorderPara = oHeader.Range.Paragraphs.Add
    With oBorderPara.Range
        .Text = ""
        With .ParagraphFormat.Borders(3)  ' wdBorderBottom
            .LineStyle = 1        ' wdLineStyleSingle
            .LineWidth = 18       ' wdLineWidth225pt
            .Color = NARANJA
        End With
    End With

    ' ── PIE DE PÁGINA ─────────────────────────────────────────────
    Dim oFtr As Object
    Set oFtr = oDoc.Sections(1).Footers(1)
    With oFtr.Range
        .Text = "Marchant Pereira N° 367, Of. 204, Providencia, Santiago  |  " & _
                "Fono: +56 2 2 484 1515  |  aplusajustadores.cl"
        .Font.Size = 7
        .Font.Color = GRIS
        .ParagraphFormat.Alignment = 3  ' wdAlignParagraphCenter
        With .ParagraphFormat.Borders(3)
            .LineStyle = 1
            .LineWidth = 6
            .Color = CELESTE
        End With
    End With

    ' Número de página
    oFtr.Range.Paragraphs(1).Range.InsertAfter "   Página "
    oFtr.Range.Fields.Add oFtr.Range.Characters.Last, 33  ' wdFieldPage

    ' ── CUERPO DEL DOCUMENTO ──────────────────────────────────────
    Dim oRange As Object
    Set oRange = oDoc.Content
    oRange.Collapse 1  ' wdCollapseStart

    ' ── Fecha y N° carta ──────────────────────────────────────────
    Call AgregarParrafo(oDoc, "Santiago, " & Format(Now, "DD de MMMM de YYYY"), _
                        False, 10, RGB(0,0,0), 1, 0, 60)
    Call AgregarParrafo(oDoc, "N° " & IIf(nroLiquid<>"[N° Liquidación]", nroLiquid, "—"), _
                        False, 10, RGB(0,0,0), 1, 0, 120)

    ' ── Destinatario ──────────────────────────────────────────────
    Call AgregarParrafo(oDoc, asegurado, True, 11, AZUL_OSC, 1, 0, 40)
    Call AgregarParrafo(oDoc, direccion, False, 10, RGB(0,0,0), 1, 0, 40)
    Call AgregarParrafo(oDoc, "Presente", False, 10, RGB(0,0,0), 1, 0, 160)

    ' ── Bloque Ref ────────────────────────────────────────────────
    Call AgregarParrafo(oDoc, "Ref.:", True, 10, AZUL_OSC, 1, 0, 80)

    ' Tabla de referencia
    Dim oRefTable As Object
    Set oRefTable = oDoc.Tables.Add(oDoc.Bookmarks("\EndOfDoc").Range, 6, 2)
    oRefTable.Borders.Enable = False
    oRefTable.PreferredWidthType = 3  ' wdPreferredWidthPercent
    oRefTable.PreferredWidth = 100
    oRefTable.Columns(1).Width = oWord.InchesToPoints(2.2)
    oRefTable.Columns(2).Width = oWord.InchesToPoints(4.56)

    Dim campos(5) As String, valores(5) As String
    campos(0) = "Siniestro N°":   valores(0) = nroSin
    campos(1) = "Póliza N°":      valores(1) = nroPoliza
    campos(2) = "Liquidación N°": valores(2) = nroLiquid
    campos(3) = "Evento":         valores(3) = evento
    campos(4) = "Fecha Siniestro":valores(4) = fechaSin
    campos(5) = "Liquidador":     valores(5) = liquidador

    For i = 0 To 5
        With oRefTable.Cell(i+1, 1).Range
            .Text = campos(i) & ":"
            .Font.Bold = True
            .Font.Size = 10
            .Font.Color = AZUL
        End With
        With oRefTable.Cell(i+1, 2).Range
            .Text = valores(i)
            .Font.Size = 10
        End With
    Next i

    ' Espacio post-ref
    Call AgregarParrafo(oDoc, "", False, 6, RGB(0,0,0), 1, 0, 0)

    ' ── Saludo ────────────────────────────────────────────────────
    Call AgregarParrafo(oDoc, "De nuestra consideración:", False, 10, RGB(0,0,0), 1, 160, 120)

    ' ── Párrafo 1 ─────────────────────────────────────────────────
    Call AgregarParrafo(oDoc, _
        "Luego de haber revisado los antecedentes remitidos a esta oficina, informamos que hemos " & _
        "finalizado el ajuste de pérdidas del siniestro de la referencia.", _
        False, 10, RGB(0,0,0), 3, 0, 120)

    ' ── Párrafo 2 con monto destacado ─────────────────────────────
    Dim oParaMonto As Object
    Set oParaMonto = oDoc.Paragraphs.Add(oDoc.Bookmarks("\EndOfDoc").Range)
    oParaMonto.Range.ParagraphFormat.Alignment = 3  ' justified
    oParaMonto.Range.ParagraphFormat.SpaceAfter = 160
    oParaMonto.Range.Text = "El resultado de nuestro análisis nos ha permitido concluir que la pérdida " & _
        "indemnizable de responsabilidad de la póliza en estudio asciende a la suma única y final de "
    oParaMonto.Range.Font.Size = 10

    Dim oRanMonto As Object
    Set oRanMonto = oParaMonto.Range
    oRanMonto.Collapse 2  ' wdCollapseEnd
    oRanMonto.Text = "UF " & perdidaTotal
    oRanMonto.Font.Bold = True
    oRanMonto.Font.Color = NARANJA
    oRanMonto.Font.Size = 11
    Set oRanMonto = oParaMonto.Range
    oRanMonto.Collapse 2
    oRanMonto.Text = "."
    oRanMonto.Font.Bold = False
    oRanMonto.Font.Color = RGB(0,0,0)
    oRanMonto.Font.Size = 10

    ' ── SECCIÓN: DETERMINACIÓN DE PÉRDIDAS ────────────────────────
    Call AgregarParrafo(oDoc, "", False, 4, RGB(0,0,0), 1, 80, 0)

    ' Banda de título de sección
    Dim oSecBand As Object
    Set oSecBand = oDoc.Paragraphs.Add(oDoc.Bookmarks("\EndOfDoc").Range)
    With oSecBand.Range
        .Text = "DETERMINACIÓN Y AJUSTE DE LA PÉRDIDA"
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(255,255,255)
        .ParagraphFormat.Alignment = 1   ' centered
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 0
        With .ParagraphFormat.Shading
            .Texture = 0
            .ForegroundPatternColor = AZUL_OSC
            .BackgroundPatternColor = AZUL_OSC
        End With
        With .ParagraphFormat.Borders(3)  ' bottom
            .LineStyle = 1
            .LineWidth = 12
            .Color = NARANJA
        End With
    End With

    ' ── Tabla de partidas por sección ─────────────────────────────
    Dim j As Integer
    For j = 0 To 2
        Dim wsSeccion As Worksheet
        On Error Resume Next
        Set wsSeccion = ThisWorkbook.Sheets(secNames(j))
        On Error GoTo 0
        If wsSeccion Is Nothing Then GoTo NextSeccion

        ' Recopilar partidas no vacías
        Dim partidas() As Variant
        Dim nPart As Integer: nPart = 0
        Dim r As Integer
        For r = 19 To 68
            Dim descP As String
            descP = Trim(CStr(wsSeccion.Cells(r, 3).Value))
            If descP <> "" And descP <> "None" And descP <> "0" Then
                nPart = nPart + 1
            End If
        Next r

        If nPart = 0 Then GoTo NextSeccion

        ' Encabezado de sección
        Call AgregarParrafo(oDoc, "", False, 4, RGB(0,0,0), 1, 120, 0)
        Dim oSecTitle As Object
        Set oSecTitle = oDoc.Paragraphs.Add(oDoc.Bookmarks("\EndOfDoc").Range)
        With oSecTitle.Range
            .Text = "Ajuste " & secLabels(j)
            .Font.Bold = True
            .Font.Size = 10
            .Font.Color = AZUL_OSC
            .ParagraphFormat.SpaceBefore = 120
            .ParagraphFormat.SpaceAfter = 40
            With .ParagraphFormat.Borders(3)
                .LineStyle = 1
                .LineWidth = 6
                .Color = CELESTE
            End With
        End With

        ' ── Tabla de partidas ──────────────────────────────────────
        ' Columnas: Ítem | Descripción | Cant | Unidad | V.Unit($) | Total Reclamo | Dep% | Total Ajustado | Obs
        Dim nCols As Integer: nCols = 9
        Dim oPartTable As Object
        Set oPartTable = oDoc.Tables.Add(oDoc.Bookmarks("\EndOfDoc").Range, 1 + nPart, nCols)

        ' Anchos en puntos (total ~6.76 inch = 487 pts)
        Dim colW(8) As Long
        colW(0) = 36:  colW(1) = 158: colW(2) = 32: colW(3) = 46
        colW(4) = 60:  colW(5) = 68:  colW(6) = 32: colW(7) = 68: colW(8) = 24

        Dim c2 As Integer
        For c2 = 1 To nCols
            oPartTable.Columns(c2).Width = colW(c2-1)
        Next c2

        ' Header row
        Dim hdrLabels(8) As String
        hdrLabels(0) = "Ítem":   hdrLabels(1) = "Descripción"
        hdrLabels(2) = "Cant.":  hdrLabels(3) = "Unidad"
        hdrLabels(4) = "V.Unit($)": hdrLabels(5) = "Total Reclamo ($)"
        hdrLabels(6) = "Dep.%":  hdrLabels(7) = "Total Ajustado ($)": hdrLabels(8) = "Obs."

        For c2 = 1 To nCols
            With oPartTable.Cell(1, c2).Range
                .Text = hdrLabels(c2-1)
                .Font.Bold = True
                .Font.Size = 7
                .Font.Color = RGB(255,255,255)
                .ParagraphFormat.Alignment = 1
                With .ParagraphFormat.Shading
                    .Texture = 0
                    .BackgroundPatternColor = AZUL_OSC
                    .ForegroundPatternColor = AZUL_OSC
                End With
            End With
        Next c2

        ' Filas de datos
        Dim rowNum As Integer: rowNum = 2
        For r = 19 To 68
            descP = Trim(CStr(wsSeccion.Cells(r, 3).Value))
            If descP <> "" And descP <> "None" And descP <> "0" Then
                Dim bgRow As Long
                bgRow = IIf(rowNum Mod 2 = 0, RGB(245,246,250), RGB(255,255,255))

                Dim itemV  As String: itemV  = CStr(wsSeccion.Cells(r,1).Value) & IIf(wsSeccion.Cells(r,2).Value<>"","."+CStr(wsSeccion.Cells(r,2).Value),"")
                Dim cantRV As String: cantRV = CStr(wsSeccion.Cells(r,5).Value)
                Dim unitRV As String: unitRV = CStr(wsSeccion.Cells(r,6).Value)
                Dim vuniRV As String
                If IsNumeric(wsSeccion.Cells(r,7).Value) Then
                    vuniRV = "$ " & Format(wsSeccion.Cells(r,7).Value, "#,##0")
                Else
                    vuniRV = ""
                End If
                Dim totRV  As String
                If IsNumeric(wsSeccion.Cells(r,8).Value) Then
                    totRV = "$ " & Format(wsSeccion.Cells(r,8).Value, "#,##0")
                Else
                    totRV = ""
                End If
                Dim depV   As String
                If IsNumeric(wsSeccion.Cells(r,14).Value) And wsSeccion.Cells(r,14).Value <> 0 Then
                    depV = Format(wsSeccion.Cells(r,14).Value * 100, "0") & "%"
                Else
                    depV = "0%"
                End If
                Dim totAjV As String
                If IsNumeric(wsSeccion.Cells(r,15).Value) Then
                    totAjV = "$ " & Format(wsSeccion.Cells(r,15).Value, "#,##0")
                Else
                    totAjV = ""
                End If
                Dim codObsV As String: codObsV = CStr(wsSeccion.Cells(r,17).Value)
                If codObsV = "None" Then codObsV = ""

                Dim rowVals(8) As String
                rowVals(0)=itemV: rowVals(1)=descP: rowVals(2)=cantRV: rowVals(3)=unitRV
                rowVals(4)=vuniRV: rowVals(5)=totRV: rowVals(6)=depV: rowVals(7)=totAjV: rowVals(8)=codObsV

                For c2 = 1 To nCols
                    With oPartTable.Cell(rowNum, c2).Range
                        .Text = rowVals(c2-1)
                        .Font.Size = 7
                        .Font.Color = IIf(c2=7, NARANJA, IIf(c2=8, RGB(30,107,30), RGB(26,26,26)))
                        .Font.Bold = (c2=8)
                        .ParagraphFormat.Alignment = IIf(c2=1 Or c2=3 Or c2=6 Or c2=9, 1, IIf(c2=2, 0, 2))
                        With .ParagraphFormat.Shading
                            .Texture = 0
                            .BackgroundPatternColor = bgRow
                            .ForegroundPatternColor = bgRow
                        End With
                    End With
                Next c2

                rowNum = rowNum + 1
            End If
        Next r

        ' Bordes de la tabla
        With oPartTable.Borders
            .InsideLineStyle = 1
            .InsideLineWidth = 2
            .InsideColor = RGB(220,228,235)
            .OutsideLineStyle = 1
            .OutsideLineWidth = 6
            .OutsideColor = AZUL_OSC
        End With

        ' Resumen de la sección
        Dim totUF_sec As String: totUF_sec = ""
        Dim vPerd As Variant: vPerd = wsSeccion.Cells(77, 15).Value
        If IsNumeric(vPerd) And vPerd <> 0 Then
            totUF_sec = Format(vPerd, "0.00") & " UF"
        End If

        If totUF_sec <> "" Then
            Call AgregarParrafo(oDoc, _
                "Pérdida Determinada " & secLabels(j) & ":  " & totUF_sec, _
                True, 9, NARANJA, 2, 80, 120)
        End If

NextSeccion:
        Set wsSeccion = Nothing
    Next j

    ' ── Observaciones generales ────────────────────────────────────
    Dim hayObs As Boolean: hayObs = False
    For r = 81 To 86
        If Trim(CStr(wsData.Cells(r,1).Value)) <> "" And _
           Trim(CStr(wsData.Cells(r,1).Value)) <> "None" Then
            hayObs = True
        End If
    Next r

    If hayObs Then
        Call AgregarParrafo(oDoc, "", False, 4, RGB(0,0,0), 1, 120, 0)
        Dim oObsTitle As Object
        Set oObsTitle = oDoc.Paragraphs.Add(oDoc.Bookmarks("\EndOfDoc").Range)
        With oObsTitle.Range
            .Text = "Observaciones a la Pérdida Determinada"
            .Font.Bold = True: .Font.Size = 10: .Font.Color = AZUL_OSC
            .ParagraphFormat.SpaceBefore = 120: .ParagraphFormat.SpaceAfter = 40
            With .ParagraphFormat.Borders(3)
                .LineStyle = 1: .LineWidth = 6: .Color = CELESTE
            End With
        End With

        For r = 81 To 86
            Dim obsText As String
            obsText = Trim(CStr(wsData.Cells(r,1).Value))
            If obsText <> "" And obsText <> "None" Then
                Call AgregarParrafo(oDoc, obsText, False, 9, RGB(26,26,26), 3, 60, 60)
            End If
        Next r
    End If

    ' ── Resumen de pérdidas ────────────────────────────────────────
    Call AgregarParrafo(oDoc, "", False, 4, RGB(0,0,0), 1, 120, 0)
    Dim oResTitle As Object
    Set oResTitle = oDoc.Paragraphs.Add(oDoc.Bookmarks("\EndOfDoc").Range)
    With oResTitle.Range
        .Text = "Resumen de Pérdidas"
        .Font.Bold = True: .Font.Size = 10: .Font.Color = AZUL_OSC
        .ParagraphFormat.SpaceBefore = 120: .ParagraphFormat.SpaceAfter = 40
        With .ParagraphFormat.Borders(3)
            .LineStyle = 1: .LineWidth = 6: .Color = CELESTE
        End With
    End With

    ' Tabla resumen
    Dim oResTable As Object
    Set oResTable = oDoc.Tables.Add(oDoc.Bookmarks("\EndOfDoc").Range, 6, 2)
    oResTable.Columns(1).Width = oWord.InchesToPoints(4.5)
    oResTable.Columns(2).Width = oWord.InchesToPoints(2.26)
    oResTable.Borders.Enable = True
    With oResTable.Borders
        .InsideLineStyle = 1: .InsideLineWidth = 2: .InsideColor = RGB(220,228,235)
        .OutsideLineStyle = 1: .OutsideLineWidth = 6: .OutsideColor = AZUL_OSC
    End With

    ' Leer valores del resumen
    Dim resLabels(5) As String, resVals(5) As String, resBold(5) As Boolean, resHigh(5) As Boolean
    Dim deducible As String: deducible = ""
    If Not wsRes Is Nothing Then
        Dim rr As Integer
        For rr = 10 To 12
            Dim v As Variant: v = wsRes.Cells(rr, 2).Value
            If IsNumeric(v) And v <> 0 Then
                resVals(rr-10) = Format(v, "0.00") & " UF"
            Else
                resVals(rr-10) = "0.00 UF"
            End If
            resLabels(rr-10) = "Pérdida Determinada " & wsRes.Cells(rr, 1).Value
        Next rr
        ' Total
        Dim vTot As Variant: vTot = wsRes.Cells(13,2).Value
        resLabels(3) = "Total Pérdida Determinada"
        resVals(3) = IIf(IsNumeric(vTot), Format(vTot,"0.00")&" UF","—")
        resBold(3) = True
    Else
        resLabels(0)="Pérdida Edificio": resLabels(1)="Pérdida Contenidos"
        resLabels(2)="Pérdida Maquinaria": resLabels(3)="Total Pérdida Determinada"
    End If

    ' Deducible desde hoja Edificio
    Dim vDed As Variant: vDed = wsData.Cells(76,15).Value
    If IsNumeric(vDed) And vDed <> 0 Then
        deducible = Format(vDed,"0.00") & " UF"
    Else
        deducible = "0.00 UF"
    End If
    resLabels(4) = "(–) Deducible": resVals(4) = deducible: resBold(4) = False
    resLabels(5) = "PÉRDIDA INDEMNIZABLE": resVals(5) = perdidaTotal: resBold(5) = True: resHigh(5) = True

    For i = 0 To 5
        Dim bgRes As Long
        bgRes = IIf(resHigh(i), RGB(255,240,234), IIf(resBold(i), RGB(238,243,251), IIf(i Mod 2=0, RGB(245,246,250), RGB(255,255,255)))

        With oResTable.Cell(i+1,1).Range
            .Text = resLabels(i)
            .Font.Bold = resBold(i): .Font.Size = 10
            .Font.Color = IIf(resHigh(i), NARANJA, IIf(resBold(i), AZUL_OSC, RGB(26,26,26)))
            With .ParagraphFormat.Shading
                .Texture = 0: .BackgroundPatternColor = bgRes: .ForegroundPatternColor = bgRes
            End With
        End With
        With oResTable.Cell(i+1,2).Range
            .Text = resVals(i)
            .Font.Bold = resBold(i): .Font.Size = 10
            .Font.Color = IIf(resHigh(i), NARANJA, IIf(resBold(i), AZUL_OSC, RGB(26,26,26)))
            .ParagraphFormat.Alignment = 2  ' right
            With .ParagraphFormat.Shading
                .Texture = 0: .BackgroundPatternColor = bgRes: .ForegroundPatternColor = bgRes
            End With
        End With
    Next i

    ' ── Párrafos de cierre ────────────────────────────────────────
    Call AgregarParrafo(oDoc, "", False, 4, RGB(0,0,0), 1, 160, 0)
    Call AgregarParrafo(oDoc, _
        "En consecuencia, en los próximos días procederemos a la emisión de nuestro informe final " & _
        "de liquidación, el que despacharemos a la Compañía de Seguros recomendando la indemnización " & _
        "por la suma antes propuesta.", False, 10, RGB(0,0,0), 3, 0, 120)
    Call AgregarParrafo(oDoc, _
        "Solicitamos nos indiquen a la brevedad cualquier consulta, aclaración u observación al " & _
        "respecto, a fin de aclarar cualquier duda que existiere.", False, 10, RGB(0,0,0), 3, 0, 200)
    Call AgregarParrafo(oDoc, "Sin otro particular, les saluda muy atentamente,", _
                        False, 10, RGB(0,0,0), 1, 0, 600)

    ' ── Firma ─────────────────────────────────────────────────────
    Call AgregarParrafo(oDoc, liquidador, True, 11, AZUL_OSC, 1, 0, 40)
    Call AgregarParrafo(oDoc, "A Plus Ajustadores SpA", False, 10, RGB(0,0,0), 1, 0, 40)
    If liquidador <> "[Liquidador]" Then
        Dim emailLiquidador As String
        emailLiquidador = LCase(Replace(Split(liquidador, " ")(0), "", ""))
        Call AgregarParrafo(oDoc, "Email: " & LCase(Split(liquidador," ")(0)) & "." & _
                            LCase(Split(liquidador," ")(UBound(Split(liquidador," ")))) & _
                            "@aplusajustadores.cl", False, 9, AZUL, 1, 0, 0)
    End If

    ' ── Guardar Word ──────────────────────────────────────────────
    oDoc.SaveAs2 rutaWord, 16   ' wdFormatDocumentDefault = 16

    ' ── Guardar PDF ───────────────────────────────────────────────
    oDoc.ExportAsFixedFormat rutaPDF, 17   ' wdExportFormatPDF = 17

    oDoc.Close False
    oWord.Quit

    MsgBox "Carta de Ajuste generada exitosamente:" & Chr(10) & Chr(10) & _
           "Word: " & rutaWord & Chr(10) & _
           "PDF:  " & rutaPDF & Chr(10) & Chr(10) & _
           "Los archivos se guardaron en la misma carpeta del Excel.", _
           vbInformation, "Carta Generada"

    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not oDoc Is Nothing Then oDoc.Close False
    If Not oWord Is Nothing Then oWord.Quit
    MsgBox "Error al generar la carta: " & Err.Description, vbCritical, "Error"
End Sub

' ── Helper: agregar párrafo formateado ────────────────────────────
Private Sub AgregarParrafo(oDoc As Object, texto As String, bold As Boolean, _
                            sz As Integer, color As Long, align As Integer, _
                            spaceBefore As Integer, spaceAfter As Integer)
    Dim oPar As Object
    Set oPar = oDoc.Paragraphs.Add(oDoc.Bookmarks("\EndOfDoc").Range)
    With oPar.Range
        .Text = texto
        .Font.Bold = bold
        .Font.Size = sz
        .Font.Color = color
        .ParagraphFormat.Alignment = align - 1  ' 0=left,1=center,2=right,3=just
        .ParagraphFormat.SpaceBefore = spaceBefore
        .ParagraphFormat.SpaceAfter = spaceAfter
    End With
End Sub
