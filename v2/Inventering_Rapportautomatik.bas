Attribute VB_Name = "Inventering_Rapportautomatik"
Option Explicit

Private vHeaders As Object, vData As Variant, vSkip As Object, vIDs As Object
Private vCats As Collection, vGroups As Collection, vReuse As Collection, vUnits As Collection
Private vIssues As Collection, vDecisions As Object, vRuleIDs As Object
Private vGroupAll As Object, vGroupArea As Object, vCurrentArea As String
Private vLast As Long, vCalcLast As Long
Private vTokens As Object, vGroupConfigs As Object


' INVENTERING | Vectorworks Excel-export -> rapport
' Importera modulen i Excel VBA (Alt+F11 > Arkiv > Importera fil).
' Spara arbetsboken som .xlsm. Kör InstalleraKnappar en gång.
' Ändringar förs INTE tillbaka till Vectorworks av den här modulen.
' Mängder och sparprocent i befintliga rapportposter följer Rådata automatiskt.
' Kategorier, nya littera och utfällbara objekt uppdateras med UppdateraRapporter.
' Rapportmall innehaller alla grundformat och kolumnbredder for genererade rapporter.

Private Const SHEET_RAW As String = "Rådata"
Private Const SHEET_RULES As String = "Regler"
Private Const SHEET_PREF As String = "Rapportval"
Private Const SHEET_SUM As String = "Sammanställning"
Private Const GENERATED_PREFIX As String = "OMR - "
Private Const LIVE_LAST_ROW As Long = 2000
Private mOutput() As Variant, mFormula() As Boolean, mRowType() As Long
Private mRaw As Variant, mRules As Variant
Private mPhotoJobs As Collection
Private mRulesByRow As Object, mUnits() As String
Private mHeadingOrder As Object
Private mUseConditionRules As Boolean
Private mRawLast As Long
Private mOmitLit As String, mOmitMaterial As String
Private mTemplateStyles(1 To 14, 1 To 12) As String

Public Sub InstalleraKnappar()
    Dim ws As Worksheet, S As Shape, i As Long
    Set ws = ThisWorkbook.Worksheets(SHEET_PREF)
    For i = ws.Shapes.count To 1 Step -1
        If Left$(ws.Shapes(i).name, Len("InventeringBtn_")) = "InventeringBtn_" Then ws.Shapes(i).Delete
    Next i
    Set S = ws.Shapes.AddShape(5, ws.Range("H4").Left, ws.Range("H4").Top, 205, 36)
    With S
        .name = "InventeringBtn_Import"
        .TextFrame.Characters.text = "1   Importera från Vectorworks"
        .OnAction = "ImporteraFranVectorworks"
        .Fill.ForeColor.RGB = RGB(25, 50, 76)
        .Line.Visible = 0
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    End With
    Set S = ws.Shapes.AddShape(5, ws.Range("H7").Left, ws.Range("H7").Top, 205, 36)
    With S
        .name = "InventeringBtn_Rapporter"
        .TextFrame.Characters.text = "2   Uppdatera rapporter"
        .OnAction = "UppdateraRapporter"
        .Fill.ForeColor.RGB = RGB(23, 110, 116)
        .Line.Visible = 0
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    End With
    Set ws = ThisWorkbook.Worksheets(SHEET_RULES)
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len("InventeringRule_")) = "InventeringRule_" Then ws.Shapes(i).Delete
    Next i
    Set s = ws.Shapes.AddShape(5, ws.Range("M1").Left, ws.Range("M1").Top, 150, 28)
    s.Name = "InventeringRule_Add": s.TextFrame.Characters.Text = "Lägg till regel": s.OnAction = "LaggTillRegel"
    Set s = ws.Shapes.AddShape(5, ws.Range("Q1").Left, ws.Range("Q1").Top, 150, 28)
    s.Name = "InventeringRule_Copy": s.TextFrame.Characters.Text = "Duplicera regel": s.OnAction = "DupliceraRegel"
    Set ws = ThisWorkbook.Worksheets(SHEET_PREF)
    ws.Range("H:H").ColumnWidth = 28
    MsgBox "Import- och rapportknappar samt två regelknappar har skapats.", vbInformation
End Sub

Private Function LastRowIn(ByVal ws As Worksheet, ByVal col As Long) As Long
    LastRowIn = ws.Cells(ws.rows.count, col).End(xlUp).Row
End Function

Private Function S(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        S = ""
    Else
        S = Trim$(CStr(v))
    End If
End Function

Private Function N(ByVal v As Variant) As Double
    If IsNumeric(v) Then N = CDbl(v) Else N = 0#
End Function

Private Function IsYes(ByVal value As Variant) As Boolean
    IsYes = (StrComp(S(value), "Ja", vbTextCompare) = 0)
End Function

Private Function Excluded(ByVal ws As Worksheet, ByVal rowNo As Long) As Boolean
    Dim omitLit As String, omitMaterial As String
    omitLit = mOmitLit
    omitMaterial = mOmitMaterial
    Excluded = False
    If vSkip.Exists(CStr(rowNo)) Then Excluded = True: Exit Function
    If Len(omitLit) > 0 Then
        If InStr(1, S(mRaw(rowNo, 1)), omitLit, vbTextCompare) > 0 Then Excluded = True
    End If
    If Len(omitMaterial) > 0 Then
        If StrComp(S(mRaw(rowNo, 2)), omitMaterial, vbTextCompare) = 0 Then Excluded = True
    End If
End Function

Private Function ImagePath(ByVal imageName As String) As String
    Dim base As String
    If Len(imageName) = 0 Then Exit Function
    If InStr(1, imageName, ":") > 0 Or Left$(imageName, 2) = "\\" Then
        ImagePath = imageName
        Exit Function
    End If
    base = S(ThisWorkbook.Worksheets(SHEET_PREF).Range("B11").Value2)
    If Len(base) = 0 Then Exit Function
    If Right$(base, 1) <> "\" And Right$(base, 1) <> "/" Then base = base & "\"
    ImagePath = base & imageName
End Function







Private Function FormulaText(ByVal value As String) As String
    FormulaText = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function



Private Function GetReportOption(ByVal area As String) As Variant
    Dim ws As Worksheet, i As Long, last As Long, target As String, found As Long
    Set ws = ThisWorkbook.Worksheets(SHEET_PREF)
    last = LastRowIn(ws, 1)
    If Len(area) = 0 Then target = "ALLA DELOMRÅDEN" Else target = "STANDARD DELOMRÅDE"
    For i = 21 To last
        If StrComp(S(ws.Cells(i, 1).Value2), target, vbTextCompare) = 0 Then found = i
    Next i
    If Len(area) > 0 Then
        For i = 21 To last
            If StrComp(S(ws.Cells(i, 1).Value2), area, vbTextCompare) = 0 Then found = i
        Next i
    End If
    If found = 0 Then
        GetReportOption = Array(3, False, False, True)
    Else
        GetReportOption = Array(Application.Max(1, Application.Min(3, CLng(N(ws.Cells(found, 2).Value2)))), _
               IsYes(ws.Cells(found, 3).Value2), IsYes(ws.Cells(found, 4).Value2), IsYes(ws.Cells(found, 5).Value2))
    End If
End Function

Private Function SafeSheetName(ByVal area As String, ByVal taken As Object) As String
    Dim invalid As Variant, ch As Variant, root As String, tryName As String, num As Long, suffix As String
    invalid = Array("/", "\", "?", "*", "[", "]", ":")
    root = Trim$(area)
    If Len(root) = 0 Then root = "Utan delområde"
    For Each ch In invalid: root = Replace(root, CStr(ch), "_"): Next ch
    root = GENERATED_PREFIX & root
    root = Left$(root, 31)
    tryName = root
    num = 1
    Do While taken.Exists(UCase$(tryName))
        num = num + 1
        suffix = "_" & CStr(num)
        tryName = Left$(root, 31 - Len(suffix)) & suffix
    Loop
    taken(UCase$(tryName)) = True
    SafeSheetName = tryName
End Function

' Gemensam design styrs helt fran de formaterade exempelraderna i Rapportmall.
' Redigera rad 1/2/4/6/8/10/12/14 och kolumnbredder A:L i mallen.
Private Sub PrepareReportData()
    Dim tmpl As Worksheet, ws As Worksheet, sc As Range, tc As Style
    Dim t As Variant, c As Long, edge As Variant, styleName As String
    Dim i As Long, area As String, rule As Variant
    VLoad
    If N(ThisWorkbook.Worksheets(SHEET_PREF).Range("B26").Value2) <> 1 Then Err.Raise vbObjectError + 340, , "Endast grundrapport nivå 1 är implementerad. Sätt Rapportval!B26 till 1."
    With ThisWorkbook.Worksheets(SHEET_PREF)
        mOmitLit = S(.Range("B6").Value2)
        mOmitMaterial = S(.Range("B7").Value2)
    End With
    Set mRulesByRow = CreateObject("Scripting.Dictionary")
    Set mHeadingOrder = CreateObject("Scripting.Dictionary")
    ReDim mUnits(1 To Application.Max(2, mRawLast))
    For i = 2 To mRawLast
        If Not vSkip.Exists(CStr(i)) Then
            rule = VCategory(i)
            mRulesByRow.Add CStr(i), rule
            RememberHeadingOrder rule
            mUnits(i) = VUnit(i, CStr(rule(3)))
            vDecisions(CStr(i)) = VDecision(i)
            If Excluded(ThisWorkbook.Worksheets(SHEET_RAW), i) Then VIssue "Filtrerat", "Info", i, "Uteslutet enligt Rapportval."
        End If
    Next i
    VBuildCalc
    VControl
    Set tmpl = ThisWorkbook.Worksheets("Rapportmall")
    For Each t In Array(1, 2, 4, 6, 8, 10, 12, 14)
        For c = 1 To 12
            Set sc = tmpl.Cells(CLng(t), c)
            styleName = "Inventering_Auto_" & t & "_" & c
            Set tc = Nothing
            On Error Resume Next
            Set tc = ThisWorkbook.Styles(styleName)
            On Error GoTo 0
            If tc Is Nothing Then Set tc = ThisWorkbook.Styles.Add(styleName)
            mTemplateStyles(CLng(t), c) = styleName
            tc.IncludeAlignment = True
            tc.IncludeBorder = True
            tc.IncludeFont = True
            tc.IncludeNumber = True
            tc.IncludePatterns = True
            tc.IncludeProtection = True
        tc.NumberFormat = sc.NumberFormat
        tc.HorizontalAlignment = sc.HorizontalAlignment
        tc.VerticalAlignment = sc.VerticalAlignment
        tc.WrapText = sc.WrapText
        tc.Orientation = sc.Orientation
        tc.IndentLevel = sc.IndentLevel
        tc.ShrinkToFit = sc.ShrinkToFit
        tc.Locked = sc.Locked
        tc.FormulaHidden = sc.FormulaHidden
        With tc.Font
            .name = sc.Font.name
            .Size = sc.Font.Size
            .Bold = sc.Font.Bold
            .Italic = sc.Font.Italic
            .Underline = sc.Font.Underline
            .Strikethrough = sc.Font.Strikethrough
            .Color = sc.Font.Color
        End With
        With tc.Interior
            .Color = sc.Interior.Color
            .PatternColor = sc.Interior.PatternColor
            .Pattern = sc.Interior.Pattern
        End With
        For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlDiagonalDown, xlDiagonalUp)
            With tc.Borders(edge)
                If sc.Borders(edge).LineStyle = xlLineStyleNone Then
                    .LineStyle = xlLineStyleNone
                Else
                    .Weight = sc.Borders(edge).Weight
                    .Color = sc.Borders(edge).Color
                    .LineStyle = sc.Borders(edge).LineStyle
                End If
            End With
        Next edge
        Next c
    Next t
End Sub

Private Sub RememberHeadingOrder(ByVal rule As Variant)
    Dim level As Long, path As String, order As Double
    order = N(rule(4))
    For level = 0 To 2
        path = path & Chr$(30) & UCase$(CStr(rule(level)))
        If Not mHeadingOrder.Exists(path) Then
            mHeadingOrder.Add path, order
        ElseIf order < CDbl(mHeadingOrder(path)) Then
            mHeadingOrder(path) = order
        End If
    Next level
End Sub

Private Function HeadingSortKey(ByVal rule As Variant) As String
    Dim level As Long, path As String, result As String, title As String
    For level = 0 To 2
        title = UCase$(CStr(rule(level)))
        path = path & Chr$(30) & title
        result = result & Format$(CDbl(mHeadingOrder(path)), "000000") & "|" & title & "|"
    Next level
    HeadingSortKey = result
End Function

Private Sub CopyTemplateRow(ByVal dest As Worksheet, ByVal outRow As Long, ByVal templateRow As Long)
    Dim c As Long
    ' Rapportmall är en formatmall. Exempelvärden och exempelformler får aldrig bli rapportdata.
    mRowType(outRow) = templateRow
    For c = 1 To 12
        mOutput(outRow, c) = Empty
        mFormula(outRow, c) = False
    Next c
End Sub

Private Sub ApplyRowFormats(ByVal ws As Worksheet, ByVal selected As Range, ByVal templateRow As Long)
    Dim c As Long
    For c = 1 To 12
        Intersect(selected, ws.Columns(c)).Style = mTemplateStyles(templateRow, c)
    Next c
    selected.RowHeight = ThisWorkbook.Worksheets("Rapportmall").rows(templateRow).RowHeight
End Sub

Private Sub FlushReport(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal photos As Object)
    Dim block() As Variant, r As Long, c As Long, t As Variant, batch As Range
    Dim count As Long, sc As Range, merged As Range, job As Variant, tick As Double
    Dim tmpl As Worksheet, photoRows As Collection
    Dim firstFormula As Long, k As Long, formulas() As Variant
    tick = Timer
    ReDim block(1 To lastRow, 1 To 12)
    ws.Range("A1").Resize(lastRow, 12).NumberFormat = "General"
    For c = 1 To 12
        Set batch = Nothing
        count = 0
        For r = 1 To lastRow
            If Not mFormula(r, c) Then
                block(r, c) = mOutput(r, c)
                ' Bevara text-ID och littera utan att lägga till apostrofer i innehållet.
                If VarType(block(r, c)) = vbString Then
                    If Len(block(r, c)) > 0 Then
                        If batch Is Nothing Then
                            Set batch = ws.Cells(r, c)
                        Else
                            Set batch = Application.Union(batch, ws.Cells(r, c))
                        End If
                        count = count + 1
                        If count = 50 Then
                            batch.NumberFormat = "@"
                            Set batch = Nothing
                            count = 0
                        End If
                    End If
                End If
            End If
        Next r
        If Not batch Is Nothing Then batch.NumberFormat = "@"
    Next c
    ws.Range("A1").Resize(lastRow, 12).Value2 = block
    ' Skriv endast riktiga formler med Formula, i sammanhängande kolumnblock.
    For c = 1 To 12
        r = 1
        Do While r <= lastRow
            If mFormula(r, c) Then
                firstFormula = r
                Do While r <= lastRow
                    If Not mFormula(r, c) Then Exit Do
                    r = r + 1
                Loop
                ReDim formulas(1 To r - firstFormula, 1 To 1)
                For k = firstFormula To r - 1
                    formulas(k - firstFormula + 1, 1) = mOutput(k, c)
                Next k
                With ws.Cells(firstFormula, c).Resize(r - firstFormula, 1)
                    .NumberFormat = "General"
                    .Formula = formulas
                    If IsNull(.HasFormula) Then Err.Raise vbObjectError + 211, "FlushReport", "Alla rapportformler kunde inte skrivas."
                    If Not .HasFormula Then Err.Raise vbObjectError + 212, "FlushReport", "Rapportformler blev text."
                End With
            Else
                r = r + 1
            End If
        Loop
    Next c
    LoggaTid "  Blockskrivning", tick
    Set tmpl = ThisWorkbook.Worksheets("Rapportmall")
    For Each t In Array(1, 2, 4, 6, 8, 10, 12, 14)
        Set batch = Nothing
        count = 0
        For r = 1 To lastRow
            If mRowType(r) = CLng(t) Then
                If batch Is Nothing Then
                    Set batch = ws.Range("A" & r & ":L" & r)
                Else
                    Set batch = Application.Union(batch, ws.Range("A" & r & ":L" & r))
                End If
                count = count + 1
                If count = 50 Then
                    ApplyRowFormats ws, batch, CLng(t)
                    Set batch = Nothing
                    count = 0
                End If
            End If
        Next r
        If Not batch Is Nothing Then ApplyRowFormats ws, batch, CLng(t)
        For c = 1 To 12
            Set sc = tmpl.Cells(CLng(t), c)
            If sc.MergeCells Then
                Set merged = sc.MergeArea
                If merged.Row = CLng(t) And merged.Column = c And merged.rows.count = 1 _
                   And c + merged.Columns.count - 1 <= 12 Then
                    For r = 1 To lastRow
                        If mRowType(r) = CLng(t) Then ws.Cells(r, c).Resize(1, merged.Columns.count).Merge
                    Next r
                End If
            End If
        Next c
    Next t
    LoggaTid "  Formatering", tick
    For Each job In mPhotoJobs
        Set photoRows = job(1)
        PutPhotoLink ws.Cells(CLng(job(0)), 10), photoRows, photos
    Next job
    LoggaTid "  Bildlänkar", tick
End Sub

Private Sub QueuePhoto(ByVal rowNo As Long, ByVal rows As Collection)
    mPhotoJobs.Add Array(rowNo, rows)
End Sub

Private Sub PaintTitle(ByVal ws As Worksheet, ByVal area As String)
    Dim tmpl As Worksheet, c As Long, headers As Variant
    On Error Resume Next
    Set tmpl = ThisWorkbook.Worksheets("Rapportmall")
    On Error GoTo 0
    If tmpl Is Nothing Then Err.Raise vbObjectError + 210, "PaintTitle", "Fliken Rapportmall saknas. Importera den nya Excelmallen innan du kor VBA-modulen."
    For c = 1 To 12
        ws.Columns(c).ColumnWidth = tmpl.Columns(c).ColumnWidth
    Next c
    CopyTemplateRow ws, 1, 1
    CopyTemplateRow ws, 2, 2
    CopyTemplateRow ws, 4, 4
    If Len(area) = 0 Then
        mFormula(1, 1) = False: mOutput(1, 1) = "INVENTERING  /  ALLA DELOMRÅDEN"
    Else
        mFormula(1, 1) = False: mOutput(1, 1) = "INVENTERING  /  " & UCase$(area)
    End If
    mFormula(2, 1) = False: mOutput(2, 1) = "LEVANDE MÄNGDER  ·  Återbruk i projekt · kör UppdateraRapporter efter ändrat beslut eller regel"
    headers = Array("LITTRERING", "MATERIALTYP", "m²", "lpm", "Antal", _
       "Återbruk i projekt", "Ej aktiverat", "SKICK", "KOMMENTAR", "BILDER", "DELOMRÅDE", "INFO-ID-REFERENS")
    For c = 1 To 12
        mOutput(4, c) = headers(c - 1)
        mFormula(4, c) = False
    Next c
    ws.Outline.summaryRow = xlAbove
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA3
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$1:$4"
        .LeftMargin = Application.CentimetersToPoints(0.7)
        .RightMargin = Application.CentimetersToPoints(0.7)
        .TopMargin = Application.CentimetersToPoints(0.9)
        .BottomMargin = Application.CentimetersToPoints(0.9)
        .CenterFooter = "&P / &N"
    End With
End Sub

Private Sub CategoryHeading(ByVal ws As Worksheet, ByVal rowNo As Long, _
                            ByVal title As String, ByVal level As Long)
    If Len(title) = 0 Then Exit Sub
    CopyTemplateRow ws, rowNo, 4 + 2 * level
    mFormula(rowNo, 1) = False: mOutput(rowNo, 1) = UCase$(title)
End Sub

Private Sub SortKeys(ByRef arr As Variant, ByVal keyStrings As Object, _
                     ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long, pivot As String, temp As Variant
    i = first: j = last
    pivot = CStr(keyStrings(arr((first + last) \ 2)))
    Do While i <= j
        Do While StrComp(CStr(keyStrings(arr(i))), pivot, vbTextCompare) < 0: i = i + 1: Loop
        Do While StrComp(CStr(keyStrings(arr(j))), pivot, vbTextCompare) > 0: j = j - 1: Loop
        If i <= j Then
            temp = arr(i): arr(i) = arr(j): arr(j) = temp
            i = i + 1: j = j - 1
        End If
    Loop
    If first < j Then SortKeys arr, keyStrings, first, j
    If i < last Then SortKeys arr, keyStrings, i, last
End Sub

Private Function RoundedQuantity(ByVal nValue As Double, ByVal unit As String) As Double
    Dim stepValue As Double
    If unit = "lpm" Then
        stepValue = N(ThisWorkbook.Worksheets(SHEET_PREF).Range("B13").Value2)
    ElseIf unit = "m²" Then
        stepValue = N(ThisWorkbook.Worksheets(SHEET_PREF).Range("B14").Value2)
    End If
    If stepValue > 0 Then
        RoundedQuantity = Int(nValue / stepValue + 0.5) * stepValue
    Else
        RoundedQuantity = nValue
    End If
End Function


' LEVANDE RAPPORTFORMEL: summera utan att VBA behöver köras efter mängdändringar.
' Rader med samma littera/material och samma ursprungliga CLASS + DELOMRÅDE ingår
' direkt, även när ytterligare objekt tillkommer. Klassificering och nya grupper
' uppdateras däremot först vid knapptryckning.
Private Function XlText(ByVal value As String) As String
    Dim v As String
    v = Replace(value, "~", "~~")
    v = Replace(v, "*", "~*")
    v = Replace(v, "?", "~?")
    XlText = Chr$(34) & Replace(v, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Function RngRaw(ByVal columnLetter As String) As String
    Dim canonical As Variant, index As Long, col As Long, letter As String
    canonical = Array("LITTRERING", "Materialtyp", "DELOMRÅDE", "SKICK", "KOMMENTAR", "ANDELSPARAS", "FOTOFILNAMN", "INFO-ID-REFERENS", "ANTAL", "AREA", "LÄNGD", "CLASS", "Återbruk i projekt")
    index = Asc(UCase$(columnLetter)) - Asc("A")
    col = CLng(vHeaders(VNorm(canonical(index))))
    letter = Split(ThisWorkbook.Worksheets(SHEET_RAW).Cells(1, col).Address, "$")(1)
    RngRaw = "'Rådata'!$" & letter & "$2:$" & letter & "$" & Application.Max(2, vLast)
End Function

Private Function FindByID(ByVal rawCol As String, ByVal outRow As Long) As String
    Dim expr As String
    expr = "INDEX(" & RngRaw(rawCol) & ",MATCH($L" & outRow & "," & RngRaw("H") & ",0))"
    FindByID = "IFERROR(IF(LEN(" & expr & ")=0,""""," & expr & "),"""")"
End Function



Private Function ExactConditionCriteria(ByVal conditionExpr As String) As String
    ' COUNTIFS tolkar annars *, ? och inledande jämförelsetecken som villkor.
    ExactConditionCriteria = """=""&SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(" & _
        conditionExpr & ",""~"",""~~""),""*"",""~*""),""?"",""~?"")"
End Function



Private Sub LiveReportGroup(ByVal ws As Worksheet, ByVal wsRaw As Worksheet, _
                            ByVal headerRow As Long, ByVal rows As Collection, _
                            ByVal lastDetail As Long, ByVal unit As String, _
                            ByVal splitComment As Boolean, ByVal splitPercent As Boolean)
    Dim qty As String, mainCol As Long, stepCell As String, rowIx As Long, col As Long
    qty = VGroupFormula(rows, "QTY")
    Select Case unit
        Case "lpm": mainCol = 4: stepCell = "'Rapportval'!$B$13"
        Case "m²": mainCol = 3: stepCell = "'Rapportval'!$B$14"
        Case Else: mainCol = 5: stepCell = "0"
    End Select
    mFormula(headerRow, mainCol) = True
    mOutput(headerRow, mainCol) = "=IF(" & stepCell & ">0,ROUND((" & qty & ")/" & stepCell & ",0)*" & stepCell & "," & qty & ")"
    mFormula(headerRow, 6) = False: mOutput(headerRow, 6) = vDecisions(CStr(rows(1)))
    mFormula(headerRow, 7) = False: mOutput(headerRow, 7) = Empty
    If rows.Count > 1 Then
        mOutput(headerRow, 9) = rows.Count & " objekt · öppna (+)"
        For rowIx = headerRow + 1 To lastDetail
            For col = 3 To 5
                mOutput(rowIx, col) = Empty: mFormula(rowIx, col) = False
            Next col
            mFormula(rowIx, mainCol) = True
            Select Case unit
                Case "lpm": mOutput(rowIx, mainCol) = "=" & FindByID("K", rowIx)
                Case "m²": mOutput(rowIx, mainCol) = "=" & FindByID("J", rowIx)
                Case Else: mOutput(rowIx, mainCol) = "=" & FindByID("I", rowIx)
            End Select
            mOutput(rowIx, 6) = vDecisions(CStr(rows(rowIx - headerRow)))
            mOutput(rowIx, 7) = Empty
        Next rowIx
    End If
End Sub

Private Function PhotoNames(ByVal text As String, ByVal re As Object) As Collection
    Dim result As New Collection, seen As Object, matches As Object
    Dim m As Object, part As Variant, filename As String
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Set matches = re.Execute(text)
    If matches.count > 0 Then
        For Each m In matches
            filename = Trim$(CStr(m.SubMatches(0)))
            If Len(filename) > 0 Then
                If Not seen.Exists(filename) Then
                    seen.Add filename, True
                    result.Add filename
                End If
            End If
        Next m
    Else
        For Each part In Split(text, ",")
            filename = Trim$(CStr(part))
            If Len(filename) > 0 Then
                If Not seen.Exists(filename) Then
                    seen.Add filename, True
                    result.Add filename
                End If
            End If
        Next part
    End If
    Set PhotoNames = result
End Function

Private Sub BuildPhotoLookup(ByRef photos As Object)
    Dim raw As Worksheet, re As Object, names As Collection
    Dim data As Variant, i As Long, last As Long
    Set photos = CreateObject("Scripting.Dictionary")
    Set raw = ThisWorkbook.Worksheets(SHEET_RAW)
    last = mRawLast
    If last < 2 Then Exit Sub
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "FOTO\s*\(([^)]*)\)"
    re.Global = True
    re.IgnoreCase = True
    ' Läs en hel kolumn en gång; rad 1 inkluderas så att även ett objekt ger en array.
    data = mRaw
    For i = 2 To last
        Set names = PhotoNames(S(data(i, 7)), re)
        If names.count > 0 Then photos.Add CStr(i), names
    Next i
End Sub

Private Sub PutPhotoLink(ByVal target As Range, ByVal rows As Collection, _
                         ByVal photos As Object)
    Dim v As Variant, filename As Variant, names As Collection
    Dim seen As Object, firstName As String, path As String, label As String
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    For Each v In rows
        If photos.Exists(CStr(v)) Then
            Set names = photos(CStr(v))
            For Each filename In names
                If Not seen.Exists(CStr(filename)) Then
                    seen.Add CStr(filename), True
                    If Len(firstName) = 0 Then firstName = CStr(filename)
                End If
            Next filename
        End If
    Next v
    If seen.count = 0 Then Exit Sub
    label = firstName & " (" & seen.count & " bilder)"
    If seen.count = 1 Then label = firstName & " (1 bild)"
    path = ImagePath(firstName)
    If Len(path) = 0 Then
        target.value = label & " - ange fotomapp i Rapportval!B11"
    Else
        target.Worksheet.Hyperlinks.Add Anchor:=target, Address:=path, TextToDisplay:=label
    End If
End Sub


Private Sub RenderReport(ByVal outputSheet As Worksheet, ByVal areaFilter As String, _
                         ByVal photos As Object)
    Dim raw As Worksheet, optionRow As Variant, rule As Variant
    Dim levels As Long, separateComments As Boolean, separatePct As Boolean, showImages As Boolean
    Dim groups As Object, details As Object, totals As Object, saved As Object, missing As Object
    Dim meta As Object, sortValues As Object, groupKey As String, unit As String, arr As Variant
    Dim i As Long, pos As Long, startDetail As Long, lastDetail As Long
    Dim amount As Double, pct As Double, lit As String, mat As String, area As String
    Dim valComment As String, condition As String, key As Variant, keys As Variant
    Dim rows As Collection, detailsRow As Variant, info As Variant
    Dim l1 As String, l2 As String, l3 As String, prev1 As String, prev2 As String, prev3 As String
    Dim prevGroup As String, nRecords As Long, rawRow As Long, rawAmount As Double, rawPct As Double
    Dim rowNo As Long, startRow As Long, sameCondition As Boolean, firstCondition As String, checkRow As Variant
    Dim singleRow As Collection, summaryRow As Long
    Dim buildTick As Double
    buildTick = Timer
    vCurrentArea = areaFilter
    ReDim mOutput(1 To 5 * Application.Max(2, mRawLast) + 10, 1 To 12)
    ReDim mFormula(1 To UBound(mOutput, 1), 1 To 12)
    ReDim mRowType(1 To UBound(mOutput, 1))
    Set mPhotoJobs = New Collection
    Set raw = ThisWorkbook.Worksheets(SHEET_RAW)
    optionRow = GetReportOption(areaFilter)
    levels = CLng(optionRow(0))
    separateComments = CBool(optionRow(1))
    separatePct = CBool(optionRow(2))
    showImages = CBool(optionRow(3))
    Set groups = CreateObject("Scripting.Dictionary")
    Set totals = CreateObject("Scripting.Dictionary")
    Set saved = CreateObject("Scripting.Dictionary")
    Set missing = CreateObject("Scripting.Dictionary")
    Set meta = CreateObject("Scripting.Dictionary")
    Set sortValues = CreateObject("Scripting.Dictionary")
    PaintTitle outputSheet, areaFilter
    For i = 2 To mRawLast
        area = S(mRaw(i, 3))
        If Len(area) = 0 Then area = "Utan delområde"
        If Len(areaFilter) = 0 Or StrComp(area, areaFilter, vbTextCompare) = 0 Then
            If Not Excluded(raw, i) Then
                lit = S(mRaw(i, 1))
                mat = S(mRaw(i, 2))
                condition = S(mRaw(i, 4))
                valComment = S(mRaw(i, 5))
                rule = mRulesByRow(CStr(i))
                unit = mUnits(i)
                Select Case unit
                    Case "lpm": amount = N(mRaw(i, 11))
                    Case "m²": amount = N(mRaw(i, 10))
                    Case Else: amount = N(mRaw(i, 9))
                End Select
                If Len(areaFilter) = 0 Then groupKey = vGroupAll(CStr(i)) Else groupKey = vGroupArea(CStr(i))
                If Not groups.Exists(groupKey) Then
                    Set rows = New Collection
                    groups.Add groupKey, rows
                    totals.Add groupKey, 0#
                    saved.Add groupKey, 0#
                    missing.Add groupKey, 0&
                    meta.Add groupKey, Array(CStr(rule(0)), CStr(rule(1)), CStr(rule(2)), lit, mat, unit, condition, area)
                    sortValues.Add groupKey, HeadingSortKey(rule) & UCase$(lit) & "|" & _
                       UCase$(mat) & "|" & UCase$(area) & "|" & groupKey
                End If
                Set rows = groups(groupKey)
                rows.Add i
                totals(groupKey) = CDbl(totals(groupKey)) + amount
                If IsNumeric(mRaw(i, 6)) And Len(S(mRaw(i, 6))) > 0 Then
                    pct = CDbl(mRaw(i, 6))
                    saved(groupKey) = CDbl(saved(groupKey)) + amount * pct
                Else
                    missing(groupKey) = CLng(missing(groupKey)) + 1
                End If
            End If
        End If
    Next i
    pos = 5
    If groups.count = 0 Then
        mFormula(pos, 1) = False: mOutput(pos, 1) = "Inga objekt med de valda kriterierna."
        FlushReport outputSheet, pos, photos
        Exit Sub
    End If
    keys = groups.keys
    If groups.count > 1 Then SortKeys keys, sortValues, LBound(keys), UBound(keys)
    prev1 = "": prev2 = "": prev3 = ""
    For Each key In keys
        info = meta(key)
        l1 = CStr(info(0)): l2 = CStr(info(1)): l3 = CStr(info(2))
        If l1 <> prev1 Then
            Debug.Print "  Kategori [" & outputSheet.name & "]: " & l1
            CategoryHeading outputSheet, pos, l1, 1: pos = pos + 1
            prev1 = l1: prev2 = "": prev3 = ""
        End If
        If levels >= 2 Then
            If l2 <> prev2 And Len(l2) > 0 Then
                CategoryHeading outputSheet, pos, l2, 2: pos = pos + 1
                prev2 = l2: prev3 = ""
            End If
        End If
        If levels >= 3 Then
            If l3 <> prev3 And Len(l3) > 0 Then
                CategoryHeading outputSheet, pos, l3, 3: pos = pos + 1
                prev3 = l3
            End If
        End If
        Set rows = groups(key)
        nRecords = rows.count
        summaryRow = pos
        CopyTemplateRow outputSheet, pos, 12
        mFormula(pos, 1) = False: mOutput(pos, 1) = CStr(info(3))
        mFormula(pos, 2) = False: mOutput(pos, 2) = CStr(info(4))
        firstCondition = S(mRaw(CLng(rows(1)), 4))
        sameCondition = True
        For Each checkRow In rows
            If StrComp(firstCondition, S(mRaw(CLng(checkRow), 4)), vbTextCompare) <> 0 Then sameCondition = False
        Next checkRow
        If sameCondition Then
            mFormula(pos, 8) = False: mOutput(pos, 8) = firstCondition
        Else
            mFormula(pos, 8) = False: mOutput(pos, 8) = "Varierande"
        End If
        If nRecords = 1 Then
            rawRow = CLng(rows(1))
            mFormula(pos, 9) = False: mOutput(pos, 9) = S(mRaw(rawRow, 5))
            mFormula(pos, 11) = False: mOutput(pos, 11) = S(mRaw(rawRow, 3))
            mFormula(pos, 12) = False: mOutput(pos, 12) = mRaw(rawRow, 8)
        Else
            mFormula(pos, 9) = False: mOutput(pos, 9) = nRecords & " objekt · se detaljrader (+)"
            If Len(areaFilter) > 0 Then mFormula(pos, 11) = False: mOutput(pos, 11) = areaFilter
        End If
        Select Case CStr(info(5))
            Case "m²": mFormula(pos, 3) = False: mOutput(pos, 3) = RoundedQuantity(CDbl(totals(key)), "m²")
            Case "lpm": mFormula(pos, 4) = False: mOutput(pos, 4) = RoundedQuantity(CDbl(totals(key)), "lpm")
            Case Else: mFormula(pos, 5) = False: mOutput(pos, 5) = CDbl(totals(key))
        End Select
        If CDbl(totals(key)) > 0 And CLng(missing(key)) = 0 Then
            mFormula(pos, 6) = False: mOutput(pos, 6) = CDbl(saved(key)) / CDbl(totals(key))
            mFormula(pos, 7) = False: mOutput(pos, 7) = RoundedQuantity(CDbl(saved(key)), CStr(info(5)))
        End If
        If showImages Then QueuePhoto pos, rows
        If nRecords > 1 Then
            startDetail = pos + 1
            For Each detailsRow In rows
                rawRow = CLng(detailsRow)
                pos = pos + 1
                CopyTemplateRow outputSheet, pos, 14
                rawAmount = 0#
                Select Case CStr(info(5))
                    Case "m²"
                        rawAmount = N(mRaw(rawRow, 10))
                        mFormula(pos, 3) = False: mOutput(pos, 3) = rawAmount
                    Case "lpm"
                        rawAmount = N(mRaw(rawRow, 11))
                        mFormula(pos, 4) = False: mOutput(pos, 4) = rawAmount
                    Case Else
                        rawAmount = N(mRaw(rawRow, 9))
                        mFormula(pos, 5) = False: mOutput(pos, 5) = rawAmount
                End Select
                mFormula(pos, 1) = False: mOutput(pos, 1) = "  >> " & S(mRaw(rawRow, 1))
                mFormula(pos, 2) = False: mOutput(pos, 2) = S(mRaw(rawRow, 2))
                mFormula(pos, 8) = False: mOutput(pos, 8) = S(mRaw(rawRow, 4))
                mFormula(pos, 9) = False: mOutput(pos, 9) = S(mRaw(rawRow, 5))
                mFormula(pos, 11) = False: mOutput(pos, 11) = S(mRaw(rawRow, 3))
                mFormula(pos, 12) = False: mOutput(pos, 12) = mRaw(rawRow, 8)
                If IsNumeric(mRaw(rawRow, 6)) And Len(S(mRaw(rawRow, 6))) > 0 Then
                    rawPct = CDbl(mRaw(rawRow, 6))
                    mFormula(pos, 6) = False: mOutput(pos, 6) = rawPct
                    mFormula(pos, 7) = False: mOutput(pos, 7) = rawAmount * rawPct
                End If
                If showImages Then
                    Set singleRow = New Collection
                    singleRow.Add rawRow
                    QueuePhoto pos, singleRow
                End If
            Next detailsRow
            lastDetail = pos
            outputSheet.rows(startDetail & ":" & lastDetail).Group
            outputSheet.rows(startDetail & ":" & lastDetail).Hidden = True
        End If
        ' Behåll rubriker/gruppstrukturen statisk men beräkna mängder och detaljer direkt från Rådata.
        LiveReportGroup outputSheet, raw, summaryRow, rows, pos, CStr(info(5)), separateComments, separatePct
        pos = pos + 1
    Next key
    LoggaTid "  Gruppering och innehåll", buildTick
    FlushReport outputSheet, pos - 1, photos
    With outputSheet
        .Range("A4:L" & pos).VerticalAlignment = xlCenter
        .Range("I5:I" & pos).WrapText = True
        .Range("F5:F" & pos).NumberFormat = "General"
        .Columns("G").Hidden = True
        .rows("4:4").AutoFilter
        .PageSetup.PrintArea = "$A$1:$K$" & pos - 1
        ' Radhöjder kan visa tidigare dolda rader. Fäll ihop först EFTER all formatering.
        .Outline.ShowLevels RowLevels:=1
    End With
End Sub

Private Sub LoggaTid(ByVal steg As String, ByRef startTid As Double)
    Dim nu As Double, sekunder As Double
    nu = Timer
    sekunder = nu - startTid
    If sekunder < 0 Then sekunder = sekunder + 86400
    Debug.Print steg & ": " & Format$(sekunder, "0.00") & " sek"
    startTid = nu
End Sub

Public Sub UppdateraRapporter()
    Dim oldCalc As XlCalculation, i As Long, areaSheet As Worksheet, ws As Worksheet
    Dim raw As Worksheet, areas As Object, names As Object, photos As Object
    Dim area As Variant, areaValue As String, name As String, generated As Long
    Dim stegTid As Double, totalTid As Double, oldEvents As Boolean
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    stegTid = Timer
    totalTid = stegTid
    Debug.Print "----- Rapportuppdatering " & Now & " -----"
    Application.StatusBar = "Rapporter: förbereder..."
    oldCalc = Application.Calculation
    If Len(ThisWorkbook.Path) = 0 Then Err.Raise vbObjectError + 320, , "Spara testarbetsboken som XLSM innan uppdatering."
    If MsgBox("Version 2 bygger om rapportbladen. Egna anteckningar och format finns kvar i en automatisk säkerhetskopia men överförs inte automatiskt. Fortsätt?", vbOKCancel + vbInformation) <> vbOK Then Exit Sub
    ThisWorkbook.SaveCopyAs ThisWorkbook.Path & Application.PathSeparator & "Inventering_backup_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CLng(Timer * 10) & ".xlsm"
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    PrepareReportData
    Set raw = ThisWorkbook.Worksheets(SHEET_RAW)
    Set areas = CreateObject("Scripting.Dictionary")
    For i = 2 To mRawLast
        If Not Excluded(raw, i) Then
            areaValue = S(raw.Cells(i, 3).Value2)
            If Len(areaValue) = 0 Then areaValue = "Utan delområde"
            If Not areas.Exists(areaValue) Then areas.Add areaValue, True
        End If
    Next i
    LoggaTid "Förberedelser", stegTid
    Application.StatusBar = "Rapporter: rensar gamla rapporter..."
    ' Endast tidigare genererade OMR-flikar tas bort. Använd inte detta prefix för egna flikar.
    For i = ThisWorkbook.Worksheets.count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        If Left$(ws.name, Len(GENERATED_PREFIX)) = GENERATED_PREFIX Then
            If Left$(S(ws.Range("A2").Value2), Len("GENERERAD AUTOMATISKT")) = "GENERERAD AUTOMATISKT" _
               Or Left$(S(ws.Range("A2").Value2), Len("LEVANDE MÄNGDER")) = "LEVANDE MÄNGDER" Then ws.Delete
        End If
    Next i
    Application.DisplayAlerts = True
    Set ws = ThisWorkbook.Worksheets(SHEET_SUM)
    On Error Resume Next
    ws.UsedRange.ClearOutline
    On Error GoTo Failed
    ws.UsedRange.EntireRow.Hidden = False
    ws.UsedRange.UnMerge
    ws.Cells.Clear
    LoggaTid "Rensa gamla rapporter", stegTid
    Application.StatusBar = "Rapporter: läser bildnamn från rådata..."
    BuildPhotoLookup photos
    LoggaTid "Bildreferenser", stegTid
    Application.StatusBar = "Rapporter: bygger sammanställning..."
    RenderReport ws, "", photos
    LoggaTid "Bygga sammanställning", stegTid
    Set names = CreateObject("Scripting.Dictionary")
    For i = 1 To ThisWorkbook.Worksheets.count
        names(UCase$(ThisWorkbook.Worksheets(i).name)) = True
    Next i
    LoggaTid "Förbereda delområden", stegTid
    For Each area In areas.keys
        Application.StatusBar = "Rapporter: bygger " & CStr(area) & " (" & (generated + 1) & "/" & areas.count & ")..."
        name = SafeSheetName(CStr(area), names)
        Set areaSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        areaSheet.name = name
        RenderReport areaSheet, CStr(area), photos
        generated = generated + 1
        LoggaTid "Bygga delområde " & CStr(area), stegTid
    Next area
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = "Rapporter: slutberäknar formler..."
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    LoggaTid "Slutberäkning", stegTid
    LoggaTid "TOTALT", totalTid
    Application.EnableEvents = oldEvents
    Application.StatusBar = False
    MsgBox "Rapporter uppdaterade. Sammanställning + " & generated & " delområdesflikar skapades. Klicka på + för objektspecifika detaljer.", vbInformation
    Exit Sub
Failed:
    Dim failureText As String
    failureText = Err.Description
    On Error Resume Next
    If Not vIssues Is Nothing Then VIssue "Körfel", "Fel", 0, failureText
    VControl
    Application.EnableEvents = oldEvents
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = oldCalc
    MsgBox "Fel vid rapportgenerering: " & failureText, vbExclamation
End Sub



Public Sub ImporteraFranVectorworks()
    Dim choice As Variant, book As Workbook, src As Worksheet, dest As Worksheet, opened As Workbook
    Dim sh As Object, dh As Object, ids As Object, seen As Object, field As Variant, key As Variant
    Dim sourceData As Variant, existing As Variant, values() As Variant, changed() As Boolean
    Dim sourceLast As Long, destLast As Long, sourceCols As Long, destCols As Long, finalCols As Long, finalLast As Long
    Dim r As Long, c As Long, target As Long, id As String, required As Variant
    Dim keep As Boolean, protected As Boolean, oldValue As Variant, newValue As Variant
    Dim oldCalc As XlCalculation, oldEvents As Boolean, oldScreen As Boolean, oldStatus As Variant
    Dim added As Long, updated As Long, conflicts As Long, blankRow As Boolean, totalRow As Boolean
    Dim sourceSums(0 To 3) As Double, destSums(0 To 3) As Double, sumFields As Variant, j As Long
    Dim destTotal As Boolean, sourceTotal As Boolean, backupPath As String, failure As String, wrote As Boolean
    Dim cancelled As Boolean, successful As Boolean, lo As ListObject, dataTable As ListObject, finalRange As Range
    Dim firstRow As Long, lastRun As Long, textRun As Boolean, chunk() As Variant, k As Long
    Dim audit As Worksheet, issue As Variant, auditData() As Variant, auditRow As Long
    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldStatus = Application.StatusBar
    On Error GoTo Failed
    choice = Application.GetOpenFilename("Excel (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls")
    If VarType(choice) = vbBoolean Then cancelled = True: GoTo CleanUp
    If Len(ThisWorkbook.Path) = 0 Then Err.Raise vbObjectError + 332, , "Spara testarbetsboken före import."
    For Each opened In Application.Workbooks
        If StrComp(opened.FullName, CStr(choice), vbTextCompare) = 0 Then Err.Raise vbObjectError + 342, , "Stäng importfilen först. Den får inte vara målfilen eller redan öppen."
    Next opened
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Import: filöppning..."
    Set book = Workbooks.Open(CStr(choice), ReadOnly:=True, UpdateLinks:=False)
    Set src = book.Worksheets(1): Set dest = ThisWorkbook.Worksheets(SHEET_RAW)
    Set sh = VImportHeaders(src): Set dh = VImportHeaders(dest)
    required = Array("LITTRERING", "Materialtyp", "DELOMRÅDE", "SKICK", "KOMMENTAR", "ANDELSPARAS", "FOTOFILNAMN", "INFO-ID-REFERENS", "ANTAL", "AREA", "LÄNGD", "CLASS")
    For Each field In required
        If Not sh.Exists(VNorm(field)) Then Err.Raise vbObjectError + 330, , "Importkolumn saknas: " & field
        If Not dh.Exists(VNorm(field)) Then Err.Raise vbObjectError + 331, , "Rådatakolumn saknas: " & field
    Next field
    sourceLast = VLastRow(src): destLast = VLastRow(dest)
    sourceCols = src.Cells(1, src.Columns.Count).End(xlToLeft).Column
    destCols = dest.Cells(1, dest.Columns.Count).End(xlToLeft).Column
    sourceData = src.Range(src.Cells(1, 1), src.Cells(Application.Max(2, sourceLast), sourceCols)).Value2
    existing = dest.Range(dest.Cells(1, 1), dest.Cells(Application.Max(2, destLast), destCols)).Value2
    Application.StatusBar = "Import: säkerhetskopiering..."
    backupPath = ThisWorkbook.Path & Application.PathSeparator & "Import_backup_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CLng(Timer * 10) & ".xlsm"
    ThisWorkbook.SaveCopyAs backupPath
    Application.StatusBar = "Import: ID-kontroll och verifiering av exporttotaler..."
    Set vIssues = New Collection
    Set ids = CreateObject("Scripting.Dictionary"): Set seen = CreateObject("Scripting.Dictionary")
    sumFields = Array("INFO-ID-REFERENS", "ANTAL", "AREA", "LÄNGD")
    For r = 2 To sourceLast
        For c = 1 To sourceCols
            If IsError(sourceData(r, c)) Then Err.Raise vbObjectError + 343, , "Excel-felvärde i importen på rad " & r & ", kolumn " & c
        Next c
        If r > 2 Then
            For j = 0 To 3: sourceSums(j) = sourceSums(j) + N(sourceData(r, sh(sumFields(j)))): Next j
        End If
    Next r
    For r = 2 To destLast
        For c = 1 To destCols
            If IsError(existing(r, c)) Then Err.Raise vbObjectError + 344, , "Excel-felvärde i Rådata på rad " & r & ", kolumn " & c
        Next c
        If r > 2 Then
            For j = 0 To 3: destSums(j) = destSums(j) + N(existing(r, dh(sumFields(j)))): Next j
        End If
    Next r
    sourceTotal = sourceLast >= 3: destTotal = destLast >= 3
    For Each field In Array("LITTRERING", "MATERIALTYP", "CLASS")
        If Len(S(sourceData(2, sh(field)))) > 0 Then sourceTotal = False
        If Len(S(existing(2, dh(field)))) > 0 Then destTotal = False
    Next field
    For j = 0 To 3
        If Abs(N(sourceData(2, sh(sumFields(j)))) - sourceSums(j)) > 0.00001 Then sourceTotal = False
        If Abs(N(existing(2, dh(sumFields(j)))) - destSums(j)) > 0.00001 Then destTotal = False
    Next j
    For r = 2 To destLast
        If Not (r = 2 And destTotal) Then
            id = VNorm(existing(r, dh("INFO-ID-REFERENS")))
            If Len(id) > 0 Then
                If ids.Exists(id) Then Err.Raise vbObjectError + 333, , "Dubbelt ID i befintlig Rådata: " & id
                ids.Add id, r
            End If
        End If
    Next r
    For r = 2 To sourceLast
        blankRow = True
        For c = 1 To sourceCols
            If Len(S(sourceData(r, c))) > 0 Then blankRow = False
        Next c
        If Not blankRow And Not (r = 2 And sourceTotal) Then
            If Len(S(sourceData(r, sh("LITTRERING")))) = 0 And Len(S(sourceData(r, sh("MATERIALTYP")))) = 0 And Len(S(sourceData(r, sh("CLASS")))) = 0 Then Err.Raise vbObjectError + 341, , "Oidentifierad importpost på rad " & r
            id = VNorm(sourceData(r, sh("INFO-ID-REFERENS")))
            If Len(id) = 0 Then Err.Raise vbObjectError + 334, , "Importobjekt saknar ID på rad " & r
            If seen.Exists(id) Then Err.Raise vbObjectError + 335, , "Dubbelt ID i importfilen: " & id
            seen.Add id, r
        End If
    Next r
    finalCols = destCols: finalLast = destLast
    For Each field In sh.Keys
        If Not dh.Exists(field) Then finalCols = finalCols + 1: dh.Add field, finalCols
    Next field
    ReDim values(1 To Application.Max(2, destLast + seen.Count), 1 To finalCols)
    ReDim changed(1 To UBound(values, 1), 1 To finalCols)
    keep = IsYes(ThisWorkbook.Worksheets(SHEET_PREF).Range("B5").Value2)
    For Each key In seen.Keys
        r = seen(key)
        If ids.Exists(key) Then
            target = ids(key): updated = updated + 1
        Else
            finalLast = finalLast + 1: target = finalLast: added = added + 1: ids.Add key, target
        End If
        For Each field In sh.Keys
            c = dh(field): newValue = sourceData(r, sh(field)): oldValue = Empty
            If target <= destLast And c <= destCols Then oldValue = existing(target, c)
            protected = InStr("|SKICK|KOMMENTAR|ANDELSPARAS|ÅTERBRUK I PROJEKT|", "|" & field & "|") > 0
            If keep And protected And Len(S(oldValue)) > 0 Then
                If Len(S(newValue)) > 0 And VNorm(newValue) <> VNorm(oldValue) Then
                    conflicts = conflicts + 1
                    VIssue "Importbedömning", "Varning", target, "ID " & key & ": " & field & " behölls. Import: " & S(newValue) & "; befintligt: " & S(oldValue)
                End If
            ElseIf VarType(oldValue) <> VarType(newValue) Or CStr(oldValue) <> CStr(newValue) Then
                values(target, c) = newValue: changed(target, c) = True
            End If
        Next field
    Next key
    If added + updated <> seen.Count Then Err.Raise vbObjectError + 345, , "Antalskontrollen för importen stämmer inte."
    ' Håll den verifierade exporttotalen konsekvent så att den inte blir ett falskt objekt nästa gång.
    If destTotal Then
        For j = 0 To 3
            c = dh(sumFields(j)): newValue = 0#
            For r = 3 To finalLast
                If changed(r, c) Then
                    newValue = CDbl(newValue) + N(values(r, c))
                ElseIf r <= destLast Then
                    newValue = CDbl(newValue) + N(existing(r, c))
                End If
            Next r
            If CDbl(newValue) <> N(existing(2, c)) Then values(2, c) = newValue: changed(2, c) = True
        Next j
    End If
    For Each lo In dest.ListObjects
        If lo.Name = "tblRadata" Then Set dataTable = lo
    Next lo
    ' Behåll befintliga tabellrader även om de är tomma. Utöka högst en gång före skrivning.
    Application.StatusBar = "Import: tabelluppdatering..."
    If Not dataTable Is Nothing Then
        Set finalRange = dest.Range(dest.Cells(1, 1), dest.Cells(Application.Max(finalLast, dataTable.Range.Row + dataTable.Range.Rows.Count - 1), Application.Max(finalCols, dataTable.Range.Columns.Count)))
        If finalRange.Address <> dataTable.Range.Address Then wrote = True: dataTable.Resize finalRange
    End If
    For Each field In sh.Keys
        c = dh(field)
        If c > destCols Then wrote = True: dest.Cells(1, c).Value2 = sourceData(1, sh(field))
    Next field
    ' Skriv bara ändrade celler, i sammanhängande kolumnblock. Orörda formler och egna kolumner bevaras.
    For c = 1 To finalCols
        Application.StatusBar = "Import: skrivning, kolumn " & c & " av " & finalCols & "..."
        r = 2
        Do While r <= finalLast
            If changed(r, c) Then
                firstRow = r: textRun = VarType(values(r, c)) = vbString
                Do While r <= finalLast
                    If Not changed(r, c) Then Exit Do
                    If (VarType(values(r, c)) = vbString) <> textRun Then Exit Do
                    r = r + 1
                Loop
                lastRun = r - 1
                ReDim chunk(1 To lastRun - firstRow + 1, 1 To 1)
                For k = firstRow To lastRun: chunk(k - firstRow + 1, 1) = values(k, c): Next k
                wrote = True
                VImportWriteBlock dest.Cells(firstRow, c).Resize(lastRun - firstRow + 1, 1), chunk, textRun
            Else
                r = r + 1
            End If
        Loop
    Next c
    If vIssues.Count > 0 Then
        Set audit = VSheet("IMPORTKONTROLL"): auditRow = VLastRow(audit) + 1
        ReDim auditData(1 To vIssues.Count, 1 To 3): k = 0
        For Each issue In vIssues
            k = k + 1: auditData(k, 1) = Now: auditData(k, 2) = issue(2): auditData(k, 3) = issue(3)
        Next issue
        audit.Cells(auditRow, 1).Resize(k, 3).Value2 = auditData
    End If
    book.Close False: Set book = Nothing
    successful = True
    GoTo CleanUp
Failed:
    failure = Err.Description
CleanUp:
    ' Återställ alltid, även om filöppning, säkerhetskopia eller blockskrivning misslyckas.
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    Application.StatusBar = "Import: slutberäkning och återställning..."
    Err.Clear
    Application.Calculation = oldCalc
    If Err.Number <> 0 Then failure = failure & " Återställning av beräkningsläge: " & Err.Description
    ' Återgång till automatisk beräkning räknar smutsiga formler; inget extra Calculate-anrop.
    ' Om användaren hade manuellt läge lämnas det manuellt utan framtvingad beräkning.
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.StatusBar = oldStatus
    On Error GoTo 0
    If cancelled Then Exit Sub
    If Len(failure) > 0 Then
        If wrote Then failure = failure & vbCrLf & "Importen kan vara delvis skriven. Återställ vid behov från: " & backupPath
        MsgBox "Import avbruten: " & failure, vbExclamation
    ElseIf successful Then
        MsgBox "Import klar: " & updated & " matchade, " & added & " nya. " & conflicts & " bedömningsskillnader finns i IMPORTKONTROLL. Kör UppdateraRapporter vid ändrad struktur.", vbInformation
    End If
End Sub

Private Sub VImportWriteBlock(ByVal target As Range, ByRef values() As Variant, ByVal textValues As Boolean)
    Dim fmt As Variant, formats() As Variant, i As Long, mixed As Boolean
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    If textValues Then
        fmt = target.NumberFormat: mixed = IsNull(fmt)
        If mixed Then
            ReDim formats(1 To target.Rows.Count)
            For i = 1 To target.Rows.Count: formats(i) = target.Cells(i, 1).NumberFormat: Next i
        End If
        target.NumberFormat = "@"
    End If
    target.Value2 = values
RestoreFormat:
    On Error Resume Next
    If textValues Then
        If mixed Then
            For i = 1 To target.Rows.Count: target.Cells(i, 1).NumberFormat = formats(i): Next i
        Else
            target.NumberFormat = fmt
        End If
    End If
    On Error GoTo 0
    If errorNumber <> 0 Then Err.Raise errorNumber, "VImportWriteBlock", errorText
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    Resume RestoreFormat
End Sub

Private Function VImportHeaders(ByVal ws As Worksheet) As Object
    Dim c As Long, key As String
    Set VImportHeaders = CreateObject("Scripting.Dictionary")
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        key = VNorm(ws.Cells(1, c).Value2)
        If Len(key) > 0 Then
            If VImportHeaders.Exists(key) Then Err.Raise vbObjectError + 336, , "Dubbla kolumnrubriker: " & key
            VImportHeaders.Add key, c
        End If
    Next c
End Function



' Regelstyrning version 2. Beslutsfältet får aldrig härledas ur skick eller procent.
Private Function VNorm(ByVal value As Variant) As String
    Dim t As String
    t = Replace(S(value), ChrW(160), " ")
    t = Replace(Replace(t, vbTab, " "), vbCrLf, " ")
    Do While InStr(t, "  ") > 0: t = Replace(t, "  ", " "): Loop
    VNorm = UCase$(Trim$(t))
End Function

Private Function VPack(ByVal value As String) As String
    VPack = Len(value) & ":" & value
End Function

Private Sub VIssue(ByVal kind As String, ByVal severity As String, ByVal r As Long, ByVal explanation As String, Optional ByVal ruleID As String = "")
    vIssues.Add Array(kind, severity, r, explanation, ruleID)
End Sub

Private Function VSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set VSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If VSheet Is Nothing Then
        Set VSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        VSheet.Name = name
    End If
End Function

Private Function VTable(ByVal name As String) As Collection
    Dim lo As ListObject, lr As ListRow, item As Object, c As Long, found As Boolean
    Set VTable = New Collection
    For Each lo In ThisWorkbook.Worksheets(SHEET_RULES).ListObjects
        If lo.Name = name Then
            found = True
            For Each lr In lo.ListRows
                Set item = CreateObject("Scripting.Dictionary")
                For c = 1 To lo.ListColumns.Count
                    item(VNorm(lo.ListColumns(c).Name)) = lr.Range.Cells(1, c).Value2
                Next c
                If Len(VGet(item, "Regel-ID")) > 0 Then VTable.Add item
            Next lr
        End If
    Next lo
    If Not found Then Err.Raise vbObjectError + 300, "VTable", "Tabellen " & name & " saknas. Använd testarbetsboken för version 2."
End Function

Private Function VGet(ByVal item As Object, ByVal field As String) As String
    If item.Exists(VNorm(field)) Then VGet = S(item(VNorm(field)))
End Function

Private Function VRaw(ByVal r As Long, ByVal field As String) As Variant
    If Not vHeaders.Exists(VNorm(field)) Then Err.Raise vbObjectError + 301, "VRaw", "Okänt rådatafält: " & field
    VRaw = vData(r, CLng(vHeaders(VNorm(field))))
End Function

Private Function VMatch(ByVal item As Object, ByVal n As String, ByVal r As Long) As Boolean
    Dim field As String, method As String, needle As String, value As String
    field = VGet(item, "Fält " & n)
    If Len(field) = 0 Then VMatch = True: Exit Function
    method = VNorm(VGet(item, "Metod " & n))
    needle = VNorm(VGet(item, "Värde " & n))
    value = VNorm(VRaw(r, field))
    Select Case method
        Case "INNEHÅLLER": VMatch = InStr(1, value, needle, vbTextCompare) > 0
        Case "ÄR EXAKT": VMatch = value = needle
        Case "BÖRJAR MED": VMatch = Left$(value, Len(needle)) = needle
        Case "INNEHÅLLER INTE": VMatch = InStr(1, value, needle, vbTextCompare) = 0
        Case Else: Err.Raise vbObjectError + 302, "VMatch", "Ogiltig sökmetod i " & VGet(item, "Regel-ID")
    End Select
End Function

Private Function VMatches(ByVal item As Object, ByVal r As Long) As Boolean
    Dim a As Boolean, b As Boolean, c As Boolean, localArea As String
    If Not IsYes(VGet(item, "Aktiv")) Then Exit Function
    localArea = VGet(item, "Lokalt delområde")
    If Len(localArea) > 0 And VNorm(localArea) <> VNorm(mRaw(r, 3)) Then Exit Function
    If Len(VGet(item, "Fält X")) > 0 Then
        If VMatch(item, "X", r) Then Exit Function
    End If
    a = VMatch(item, "1", r): b = VMatch(item, "2", r): c = VMatch(item, "3", r)
    Select Case VNorm(VGet(item, "Logik"))
        Case "", "1": VMatches = a
        Case "1 OCH 2": VMatches = a And b
        Case "1 ELLER 2": VMatches = a Or b
        Case "1 OCH 2 OCH 3": VMatches = a And b And c
        Case "1 ELLER 2 ELLER 3": VMatches = a Or b Or c
        Case "1 OCH (2 ELLER 3)": VMatches = a And (b Or c)
        Case "(1 OCH 2) ELLER 3": VMatches = (a And b) Or c
        Case "1 ELLER (2 OCH 3)": VMatches = a Or (b And c)
        Case "(1 ELLER 2) OCH 3": VMatches = (a Or b) And c
        Case Else: Err.Raise vbObjectError + 303, "VMatches", "Ogiltig logik i " & VGet(item, "Regel-ID")
    End Select
End Function

Private Sub VValidate(ByVal entries As Collection)
    Dim ids As Object, item As Variant, part As Variant, field As String, method As String, logic As String
    Set ids = CreateObject("Scripting.Dictionary")
    For Each item In entries
        field = VNorm(VGet(item, "Regel-ID"))
        If ids.Exists(field) Then Err.Raise vbObjectError + 304, "VValidate", "Dubbelt regel-ID: " & field
        ids.Add field, True
        If IsYes(VGet(item, "Aktiv")) Then
            If Not IsNumeric(VGet(item, "Prioritet")) Then Err.Raise vbObjectError + 305, "VValidate", "Prioritet saknas i " & field
            For Each part In Array("1", "2", "3", "X")
                field = VGet(item, "Fält " & part)
                If Len(field) > 0 Then
                    If Not vHeaders.Exists(VNorm(field)) Then Err.Raise vbObjectError + 306, "VValidate", "Okänt fält " & field
                    method = VNorm(VGet(item, "Metod " & part))
                    If InStr("|INNEHÅLLER|ÄR EXAKT|BÖRJAR MED|INNEHÅLLER INTE|", "|" & method & "|") = 0 Or Len(method) = 0 Then Err.Raise vbObjectError + 307, "VValidate", "Ogiltig metod i " & VGet(item, "Regel-ID")
                End If
            Next part
            logic = VGet(item, "Logik")
            If InStr(logic, "2") > 0 And Len(VGet(item, "Fält 2")) = 0 Then Err.Raise vbObjectError + 308, "VValidate", "Villkor 2 saknas i " & VGet(item, "Regel-ID")
            If InStr(logic, "3") > 0 And Len(VGet(item, "Fält 3")) = 0 Then Err.Raise vbObjectError + 309, "VValidate", "Villkor 3 saknas i " & VGet(item, "Regel-ID")
        End If
    Next item
End Sub

Private Function VPick(ByVal entries As Collection, ByVal r As Long, ByVal outputs As Variant, Optional ByVal action As String = "") As Object
    Dim item As Variant, best As Object, bestPriority As Double, score As Double
    Dim isLocalRule As Boolean, bestLocal As Boolean, conflict As Boolean, field As Variant, equal As Boolean, allowed As Boolean
    bestPriority = -1E+300
    For Each item In entries
        If Len(action) = 0 Or VNorm(VGet(item, "Åtgärd")) = VNorm(action) Then
            allowed = True
            If Len(VGet(item, "Huvudkategori")) > 0 Then
                allowed = False
                If mRulesByRow.Exists(CStr(r)) Then allowed = VNorm(VGet(item, "Huvudkategori")) = VNorm(mRulesByRow(CStr(r))(0))
            End If
            If allowed Then
                If VMatches(item, r) Then
                    score = CDbl(VGet(item, "Prioritet"))
                    isLocalRule = Len(VGet(item, "Lokalt delområde")) > 0
                    If best Is Nothing Then
                        Set best = item: bestPriority = score: bestLocal = isLocalRule: conflict = False
                    ElseIf (isLocalRule And Not bestLocal) Or (isLocalRule = bestLocal And score > bestPriority) Then
                        Set best = item: bestPriority = score: bestLocal = isLocalRule: conflict = False
                    ElseIf isLocalRule = bestLocal And score = bestPriority Then
                        equal = True
                        For Each field In outputs
                            If VNorm(VGet(item, CStr(field))) <> VNorm(VGet(best, CStr(field))) Then equal = False
                        Next field
                        VIssue "Regelkonflikt", IIf(equal, "Varning", "Fel"), r, "Samma prioritet: " & VGet(best, "Regel-ID") & " / " & VGet(item, "Regel-ID"), VGet(item, "Regel-ID")
                        If Not equal Then conflict = True
                    End If
                End If
            End If
        End If
    Next item
    If conflict Then Err.Raise vbObjectError + 310, "VPick", "Motstridiga regler. Se KONTROLL; befintliga rapporter är inte ändrade."
    Set VPick = best
End Function

Private Function VCategory(ByVal r As Long) As Variant
    Dim item As Object
    Set item = VPick(vCats, r, Array("Rubriknivå 1", "Rubriknivå 2", "Rubriknivå 3", "Enhet", "Ordning"))
    If item Is Nothing Then
        VIssue "Ingen regel", "Varning", r, "Objektet redovisas under ÖVRIGT / OKLASSIFICERAT."
        vRuleIDs(CStr(r)) = ""
        VCategory = Array("ÖVRIGT", "OKLASSIFICERAT", "", "Auto", 999, False)
    Else
        vRuleIDs(CStr(r)) = VGet(item, "Regel-ID")
        VCategory = Array(VGet(item, "Rubriknivå 1"), VGet(item, "Rubriknivå 2"), VGet(item, "Rubriknivå 3"), VGet(item, "Enhet"), N(VGet(item, "Ordning")), False)
    End If
End Function

Private Function VDecision(ByVal r As Long) As String
    Dim item As Variant, source As String, result As String, found As Boolean
    source = VNorm(mRaw(r, 13))
    For Each item In vReuse
        If IsYes(VGet(item, "Aktiv")) And VNorm(VGet(item, "Rådatavärde")) = source Then
            If found Then Err.Raise vbObjectError + 311, "VDecision", "Flera återbruksregler för samma värde."
            result = VNorm(VGet(item, "Rapportvärde")): found = True
        End If
    Next item
    If Not found Or (result <> "JA" And result <> "NEJ" And result <> "EJ FASTSTÄLLT") Then
        result = "EJ FASTSTÄLLT"
        VIssue "Återbruksbeslut", "Varning", r, "Okänt värde: " & source
    ElseIf Len(source) = 0 Then
        VIssue "Återbruksbeslut", "Varning", r, "Uppgift saknas; redovisas som EJ FASTSTÄLLT."
    End If
    VDecision = result
End Function

Private Function VGroupKey(ByVal r As Long, ByVal area As String) As String
    Dim base As Object, special As Object, selected As Object, field As Variant
    Dim rule As Variant, fields As Variant, result As String, signature As String, bucket As String
    fields = Array("LITTRERING", "Materialtyp", "SKICK", "KOMMENTAR", "CLASS", "DELOMRÅDE")
    Set base = VPick(vGroups, r, fields, "Grundregel")
    Set special = VPick(vGroups, r, fields, "Komplettera")
    Dim replacement As Object
    Set replacement = VPick(vGroups, r, fields, "Ersätt")
    If Not special Is Nothing And Not replacement Is Nothing Then
        If CDbl(VGet(special, "Prioritet")) = CDbl(VGet(replacement, "Prioritet")) Then Err.Raise vbObjectError + 312, "VGroupKey", "Komplettera och Ersätt har samma prioritet."
        If CDbl(VGet(replacement, "Prioritet")) > CDbl(VGet(special, "Prioritet")) Then Set special = Nothing Else Set replacement = Nothing
    End If
    If Not replacement Is Nothing Then Set base = replacement
    Set selected = CreateObject("Scripting.Dictionary")
    For Each field In fields
        selected(CStr(field)) = False
        If Not base Is Nothing Then selected(CStr(field)) = IsYes(VGet(base, CStr(field)))
        If Not special Is Nothing Then
            If IsYes(VGet(special, CStr(field))) Then selected(CStr(field)) = True
        End If
    Next field
    If base Is Nothing And special Is Nothing Then selected("LITTRERING") = True: selected("Materialtyp") = True
    rule = mRulesByRow(CStr(r))
    result = VPack(VNorm(rule(0))) & VPack(VNorm(rule(1))) & VPack(VNorm(rule(2))) & VPack(mUnits(r)) & VPack(vDecisions(CStr(r))) & VPack(VNorm(area))
    bucket = result & VPack(VNorm(mRaw(r, 1)))
    For Each field In fields
        If selected(CStr(field)) Then signature = signature & VPack(CStr(field))
    Next field
    If vGroupConfigs.Exists(bucket) Then
        If vGroupConfigs(bucket) <> signature Then VIssue "Olika gruppregler", "Varning", r, "Samma littera/kategori/beslut får olika grupperingsfält. Kontrollera grund- och specialregler."
    Else
        vGroupConfigs.Add bucket, signature
    End If
    For Each field In fields
        ' Ta med fältidentiteten för att upptäcka olika grupperingskonfigurationer.
        If selected(CStr(field)) Then result = result & VPack(CStr(field)) & VPack(VNorm(VRaw(r, CStr(field))))
    Next field
    VGroupKey = result
End Function

Private Function VLastRow(ByVal ws As Worksheet) As Long
    Dim cell As Range
    Set cell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If cell Is Nothing Then VLastRow = 1 Else VLastRow = cell.Row
End Function

Private Sub VLoad()
    Dim raw As Worksheet, headers As Variant, c As Long, r As Long, j As Long, id As String, rowBlank As Boolean
    Dim lastCol As Long, sums(1 To 4) As Double, indexes As Variant, total As Boolean, value As Variant
    Set vIssues = New Collection
    Set vSkip = CreateObject("Scripting.Dictionary"): Set vIDs = CreateObject("Scripting.Dictionary")
    Set vRuleIDs = CreateObject("Scripting.Dictionary"): Set vDecisions = CreateObject("Scripting.Dictionary")
    Set vTokens = CreateObject("Scripting.Dictionary")
    Set vGroupConfigs = CreateObject("Scripting.Dictionary")
    Set vGroupAll = CreateObject("Scripting.Dictionary"): Set vGroupArea = CreateObject("Scripting.Dictionary")
    Set vHeaders = CreateObject("Scripting.Dictionary")
    Set raw = ThisWorkbook.Worksheets(SHEET_RAW)
    vLast = VLastRow(raw): lastCol = raw.Cells(1, raw.Columns.Count).End(xlToLeft).Column
    vData = raw.Range(raw.Cells(1, 1), raw.Cells(Application.Max(2, vLast), lastCol)).Value2
    For c = 1 To lastCol
        id = VNorm(vData(1, c))
        If Len(id) > 0 Then
            If vHeaders.Exists(id) Then Err.Raise vbObjectError + 313, "VLoad", "Dubbla kolumnrubriker: " & id
            vHeaders.Add id, c
        End If
    Next c
    headers = Array("LITTRERING", "Materialtyp", "DELOMRÅDE", "SKICK", "KOMMENTAR", "ANDELSPARAS", "FOTOFILNAMN", "INFO-ID-REFERENS", "ANTAL", "AREA", "LÄNGD", "CLASS", "Återbruk i projekt")
    ReDim mRaw(1 To Application.Max(2, vLast), 1 To 13)
    For c = 0 To 12
        If Not vHeaders.Exists(VNorm(headers(c))) Then Err.Raise vbObjectError + 314, "VLoad", "Kolumn saknas: " & headers(c)
        For r = 1 To vLast: mRaw(r, c + 1) = VRaw(r, CStr(headers(c))): Next r
    Next c
    mRawLast = vLast
    ' Verifiera exportens totalrad: tomma identitetsfält och alla fyra totalsummor stämmer.
    indexes = Array(8, 9, 10, 11)
    total = vLast >= 3 And Len(S(mRaw(2, 1))) = 0 And Len(S(mRaw(2, 2))) = 0 And Len(S(mRaw(2, 12))) = 0
    If total Then
        For j = 0 To 3
            For r = 3 To vLast: sums(j + 1) = sums(j + 1) + N(mRaw(r, indexes(j))): Next r
            If Abs(N(mRaw(2, indexes(j))) - sums(j + 1)) > 0.00001 Then total = False
        Next j
        If total Then vSkip("2") = True: VIssue "Exporttotal", "Info", 2, "Verifierad totalsumma från exporten; inte ett inventeringsobjekt."
    End If
    For r = 2 To vLast
        For c = 1 To 13
            If IsError(mRaw(r, c)) Then
                vSkip(CStr(r)) = True
                VIssue "Cellfel", "Fel", r, "Excel-felvärde i rådatafält " & headers(c - 1)
            End If
        Next c
        rowBlank = True
        For c = 1 To 13
            If Len(S(mRaw(r, c))) > 0 Then rowBlank = False
        Next c
        If rowBlank Then vSkip(CStr(r)) = True
        If Not vSkip.Exists(CStr(r)) Then
            id = VNorm(mRaw(r, 8))
            If Len(id) = 0 Then
                vSkip(CStr(r)) = True: VIssue "Saknat ID", "Fel", r, "Objektet finns i KONTROLL men kan inte summeras säkert utan ID."
            ElseIf vIDs.Exists(id) Then
                vSkip(CStr(r)) = True: vSkip(CStr(vIDs(id))) = True
                VIssue "Dubbelt ID", "Fel", CLng(vIDs(id)), "ID förekommer flera gånger; inte med i mängdrapporten."
                VIssue "Dubbelt ID", "Fel", r, "ID förekommer flera gånger; inte med i mängdrapporten."
            Else
                vIDs.Add id, r
            End If
            For c = 9 To 11
                value = mRaw(r, c)
                If Len(S(value)) > 0 Then
                    If Not IsNumeric(value) Then
                        vSkip(CStr(r)) = True: VIssue "Ogiltig mängd", "Fel", r, "Icke-numerisk mängd: " & S(value)
                    ElseIf CDbl(value) < 0 Then
                        vSkip(CStr(r)) = True: VIssue "Ogiltig mängd", "Fel", r, "Negativ mängd."
                    End If
                End If
            Next c
        End If
    Next r
    VDescribeRules
    Set vCats = VTable("tblKategorierV2"): Set vGroups = VTable("tblGrupperV2")
    Set vReuse = VTable("tblAterbrukV2"): Set vUnits = VTable("tblEnheterV2")
    VValidate vCats: VValidate vGroups: VValidate vUnits
End Sub

Private Function VUnit(ByVal r As Long, ByVal pref As String) As String
    Dim item As Object
    Set item = VPick(vUnits, r, Array("Enhet"))
    If Not item Is Nothing Then pref = VGet(item, "Enhet")
    Select Case VNorm(pref)
        Case "LPM": VUnit = "lpm"
        Case "M²": VUnit = "m²"
        Case "ANTAL": VUnit = "antal"
        Case "AUTO"
            If N(mRaw(r, 11)) <> 0 Then
                VUnit = "lpm"
            ElseIf N(mRaw(r, 10)) <> 0 Then
                VUnit = "m²"
            Else
                VUnit = "antal"
            End If
        Case Else: Err.Raise vbObjectError + 315, "VUnit", "Ogiltig enhet på rådatarad " & r & ": " & pref
    End Select
End Function

Private Sub VControl()
    Dim ws As Worksheet, result() As Variant, item As Variant, i As Long, r As Long
    Set ws = VSheet("KONTROLL")
    ws.Cells.ClearContents
    ws.Range("A1:J1").Value = Array("Feltyp", "Allvar", "LITTRERING", "DELOMRÅDE", "Materialtyp", "INFO-ID-REFERENS", "Rådatarad", "Förklaring", "Regel-ID", "Källa")
    If vIssues Is Nothing Then Exit Sub
    If vIssues.Count = 0 Then ws.Cells(2, 1).Value = "Inga avvikelser": Exit Sub
    ReDim result(1 To vIssues.Count, 1 To 10)
    For Each item In vIssues
        i = i + 1: r = CLng(item(2))
        result(i, 1) = item(0): result(i, 2) = item(1): result(i, 7) = r
        result(i, 8) = item(3): result(i, 9) = item(4)
        If r >= 2 And r <= mRawLast Then
            result(i, 3) = mRaw(r, 1): result(i, 4) = mRaw(r, 3)
            result(i, 5) = mRaw(r, 2): result(i, 6) = mRaw(r, 8)
        End If
    Next item
    ws.Range("A2").Resize(vIssues.Count, 10).Value2 = result
    i = 1
    For Each item In vIssues
        i = i + 1: r = CLng(item(2))
        If r > 0 Then ws.Hyperlinks.Add ws.Cells(i, 10), "", "'Rådata'!A" & r, , "Visa rådata"
    Next item
    ws.Rows(1).Font.Bold = True: ws.Columns("A:J").ColumnWidth = 19
    ws.Columns("H").ColumnWidth = 70: ws.Columns("H").WrapText = True
End Sub

Private Sub VBuildCalc()
    Dim ws As Worksheet, data() As Variant, r As Long, i As Long, col As String, key As String
    Set ws = VSheet("_InventeringV2")
    ReDim data(1 To Application.Max(2, mRawLast), 1 To 8)
    data(1, 1) = "ID": data(1, 2) = "Grupp alla": data(1, 3) = "Grupp område"
    data(1, 4) = "Mängd": data(1, 5) = "Återbruk": data(1, 6) = "Regel-ID"
    data(1, 7) = "Källrad": data(1, 8) = "Delområde"
    i = 1
    For r = 2 To mRawLast
        If Not vSkip.Exists(CStr(r)) Then
            If Not Excluded(ThisWorkbook.Worksheets(SHEET_RAW), r) Then
                vGroupAll(CStr(r)) = VGroupKey(r, "")
                vGroupArea(CStr(r)) = VGroupKey(r, S(mRaw(r, 3)))
                i = i + 1
                data(i, 1) = mRaw(r, 8): data(i, 2) = VToken(vGroupAll(CStr(r))): data(i, 3) = VToken(vGroupArea(CStr(r)))
                data(i, 5) = vDecisions(CStr(r)): data(i, 6) = vRuleIDs(CStr(r)): data(i, 7) = r: data(i, 8) = mRaw(r, 3)
            End If
        End If
    Next r
    vCalcLast = i
    ws.Cells.ClearContents
    ws.Range("A1").Resize(UBound(data, 1), 8).NumberFormat = "@"
    ws.Range("A1").Resize(UBound(data, 1), 8).Value2 = data
    Dim quantities() As Variant
    If i > 1 Then
        ReDim quantities(1 To i - 1, 1 To 1)
        For r = 2 To i
            Select Case mUnits(CLng(data(r, 7)))
                Case "lpm": col = "K"
                Case "m²": col = "J"
                Case Else: col = "I"
            End Select
            quantities(r - 1, 1) = "=IFERROR(1*INDEX(" & RngRaw(col) & ",MATCH(A" & r & "," & RngRaw("H") & ",0)),0)"
        Next r
        ws.Range("D2:D" & i).NumberFormat = "General"
        ws.Range("D2:D" & i).Formula = quantities
    End If
    ws.Visible = xlSheetVeryHidden
End Sub

Private Function VGroupFormula(ByVal rows As Collection, ByVal metric As String) As String
    Dim key As String, column As String, args As String
    If Len(vCurrentArea) = 0 Then
        key = vGroupAll(CStr(rows(1))): column = "B"
    Else
        key = vGroupArea(CStr(rows(1))): column = "C"
    End If
    key = VToken(key)
    args = "'_InventeringV2'!$" & column & "$2:$" & column & "$" & vCalcLast & "," & ExactConditionCriteria(FormulaText(key))
    If metric = "COUNT" Then
        VGroupFormula = "COUNTIF(" & args & ")"
    Else
        VGroupFormula = "SUMIF(" & args & ",'_InventeringV2'!$D$2:$D$" & vCalcLast & ")"
    End If
End Function

Public Sub LaggTillRegel()
    VNewRule False
End Sub

Public Sub DupliceraRegel()
    VNewRule True
End Sub

Private Sub VNewRule(ByVal duplicate As Boolean)
    Dim lo As ListObject, row As ListRow, source As Range, idCol As Long, id As String, n As Long
    On Error GoTo Failed
    If ActiveSheet.Name <> SHEET_RULES Then Err.Raise vbObjectError + 316, , "Markera en cell i en regeltabell först."
    Set lo = ActiveCell.ListObject
    If lo Is Nothing Then Err.Raise vbObjectError + 317, , "Markera en cell i en regeltabell först."
    idCol = lo.ListColumns("Regel-ID").Index
    If duplicate Then
        If Intersect(ActiveCell, lo.DataBodyRange) Is Nothing Then Err.Raise vbObjectError + 318, , "Markera en regelrad."
        Set source = lo.ListRows(ActiveCell.Row - lo.HeaderRowRange.Row).Range
    End If
    Set row = lo.ListRows.Add
    Dim vc As Long, vf As String, vt As Long
    If lo.ListRows.Count > 1 Then
        For vc = 1 To lo.ListColumns.Count
            vf = "": vt = 0
            On Error Resume Next
            vt = lo.ListRows(lo.ListRows.Count - 1).Range.Cells(1, vc).Validation.Type
            vf = lo.ListRows(lo.ListRows.Count - 1).Range.Cells(1, vc).Validation.Formula1
            On Error GoTo Failed
            If vt = xlValidateList And Len(vf) > 0 Then
                With row.Range.Cells(1, vc).Validation
                    .Delete
                    .Add Type:=xlValidateList, Formula1:=vf
                    .IgnoreBlank = True: .ShowError = True
                End With
            End If
        Next vc
    End If
    If duplicate Then row.Range.Value2 = source.Value2
    n = 1
    Do
        id = lo.Name & "-" & Format$(Now, "yyyymmddhhnnss") & "-" & CLng(Timer * 100) & "-" & n
        If Application.CountIf(lo.ListColumns(idCol).DataBodyRange, id) = 0 Then Exit Do
        n = n + 1
    Loop
    row.Range.Cells(1, idCol).Value2 = id
    row.Range.Cells(1, lo.ListColumns("Aktiv").Index).Value2 = "Nej"
    If Not duplicate Then
        On Error Resume Next
        row.Range.Cells(1, lo.ListColumns("Prioritet").Index).Value2 = 0
        row.Range.Cells(1, lo.ListColumns("Logik").Index).Value2 = "1"
        On Error GoTo Failed
    End If
    VDescribeRules
    row.Range.Cells(1, 1).Select
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
End Sub

Private Function VToken(ByVal key As String) As String
    If Not vTokens.Exists(key) Then vTokens.Add key, "G" & Format$(vTokens.Count + 1, "000000")
    VToken = vTokens(key)
End Function

Private Sub VDescribeRules()
    Dim lo As ListObject, lr As ListRow, dc As ListColumn, item As Object
    Dim c As Long, n As Variant, message As String, field As String
    For Each lo In ThisWorkbook.Worksheets(SHEET_RULES).ListObjects
        If lo.Name = "tblKategorierV2" Or lo.Name = "tblGrupperV2" Or lo.Name = "tblEnheterV2" Then
            Set dc = Nothing
            On Error Resume Next
            Set dc = lo.ListColumns("Beskrivning")
            On Error GoTo 0
            If dc Is Nothing Then Set dc = lo.ListColumns.Add: dc.Name = "Beskrivning"
            For Each lr In lo.ListRows
                Set item = CreateObject("Scripting.Dictionary")
                For c = 1 To lo.ListColumns.Count: item(VNorm(lo.ListColumns(c).Name)) = lr.Range.Cells(1, c).Value2: Next c
                message = "Logik " & VGet(item, "Logik") & ". "
                For Each n In Array("1", "2", "3")
                    field = VGet(item, "Fält " & n)
                    If Len(field) > 0 Then message = message & n & ": " & field & " " & VGet(item, "Metod " & n) & " " & VGet(item, "Värde " & n) & ". "
                Next n
                If Len(VGet(item, "Fält X")) > 0 Then message = message & "Utom " & VGet(item, "Fält X") & " " & VGet(item, "Metod X") & " " & VGet(item, "Värde X") & ". "
                If lo.Name = "tblKategorierV2" Then message = message & "Placera under " & VGet(item, "Rubriknivå 1") & " / " & VGet(item, "Rubriknivå 2") & " / " & VGet(item, "Rubriknivå 3")
                If lo.Name = "tblGrupperV2" Then message = message & VGet(item, "Åtgärd") & " för " & VGet(item, "Huvudkategori")
                If lo.Name = "tblEnheterV2" Then message = message & "Enhet: " & VGet(item, "Enhet")
                lr.Range.Cells(1, dc.Index).Value2 = message
            Next lr
        End If
    Next lo
End Sub
