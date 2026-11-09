Attribute VB_Name = "Inventering_Rapportautomatik"
Option Explicit

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
Private mRaw As Variant, mRules As Variant, mTemplate As Variant
Private mPhotoJobs As Collection
Private mRulesByRow As Object, mUnits() As String
Private mRawLast As Long
Private mOmitLit As String, mOmitMaterial As String
Private mTemplateStyles(1 To 14, 1 To 12) As String

Public Sub InstalleraKnappar()
    Dim ws As Worksheet, s As Shape, i As Long
    Set ws = ThisWorkbook.Worksheets(SHEET_PREF)
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len("InventeringBtn_")) = "InventeringBtn_" Then ws.Shapes(i).Delete
    Next i
    Set s = ws.Shapes.AddShape(5, ws.Range("H4").Left, ws.Range("H4").Top, 205, 36)
    With s
        .Name = "InventeringBtn_Import"
        .TextFrame.Characters.Text = "1   Importera från Vectorworks"
        .OnAction = "ImporteraFranVectorworks"
        .Fill.ForeColor.RGB = RGB(25, 50, 76)
        .Line.Visible = 0
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    End With
    Set s = ws.Shapes.AddShape(5, ws.Range("H7").Left, ws.Range("H7").Top, 205, 36)
    With s
        .Name = "InventeringBtn_Rapporter"
        .TextFrame.Characters.Text = "2   Uppdatera rapporter"
        .OnAction = "UppdateraRapporter"
        .Fill.ForeColor.RGB = RGB(23, 110, 116)
        .Line.Visible = 0
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    End With
    ws.Range("H:H").ColumnWidth = 28
    MsgBox "Två knappar har skapats i fliken Rapportval.", vbInformation
End Sub

Private Function LastRowIn(ByVal ws As Worksheet, ByVal col As Long) As Long
    LastRowIn = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
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

Private Function GetRule(ByVal material As String, ByVal className As String, _
                         ByVal lit As String, ByVal area As String) As Variant
    Dim ws As Worksheet, r As Long, last As Long, bestRow As Long
    Dim bestScore As Double, score As Double, fieldVal As String
    Dim fieldName As String, needle As String, localArea As String
    Set ws = ThisWorkbook.Worksheets(SHEET_RULES)
    last = UBound(mRules, 1)
    bestScore = -1E+30
    For r = 5 To last
        If IsYes(mRules(r, 2)) Then
            fieldName = S(mRules(r, 3))
            needle = S(mRules(r, 4))
            localArea = S(mRules(r, 5))
            If Len(localArea) = 0 Or StrComp(localArea, area, vbTextCompare) = 0 Then
                Select Case UCase$(fieldName)
                    Case "MATERIALTYP": fieldVal = material
                    Case "CLASS": fieldVal = className
                    Case "LITTRERING": fieldVal = lit
                    Case Else: fieldVal = ""
                End Select
                If Len(needle) = 0 Or InStr(1, fieldVal, needle, vbTextCompare) > 0 Then
                    score = N(mRules(r, 1))
                    If Len(localArea) > 0 Then score = score + 1000000#
                    If score > bestScore Then bestScore = score: bestRow = r
                End If
            End If
        End If
    Next r
    If bestRow = 0 Then
        GetRule = Array("ÖVRIGT", "OKLASSIFICERAT", "", "Auto", 999)
    Else
        GetRule = Array(S(mRules(bestRow, 6)), S(mRules(bestRow, 7)), _
                        S(mRules(bestRow, 8)), S(mRules(bestRow, 9)), _
                        N(mRules(bestRow, 10)))
    End If
End Function

Private Function DetectUnit(ByVal ws As Worksheet, ByVal i As Long, ByVal pref As String) As String
    pref = LCase$(S(pref))
    If pref = "m²" Or pref = "lpm" Or pref = "antal" Then
        DetectUnit = pref: Exit Function
    End If
    If N(mRaw(i, 11)) <> 0 Then
        DetectUnit = "lpm"
    ElseIf N(mRaw(i, 10)) <> 0 Then
        DetectUnit = "m²"
    Else
        DetectUnit = "antal"
    End If
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
    Set ws = ThisWorkbook.Worksheets(SHEET_RAW)
    mRawLast = LastRowIn(ws, 1)
    mRaw = ws.Range("A1:L" & Application.Max(2, mRawLast)).Value2
    Set ws = ThisWorkbook.Worksheets(SHEET_RULES)
    mRules = ws.Range("A1:J" & Application.Max(5, LastRowIn(ws, 1))).Value2
    With ThisWorkbook.Worksheets(SHEET_PREF)
        mOmitLit = S(.Range("B6").Value2)
        mOmitMaterial = S(.Range("B7").Value2)
    End With
    Set mRulesByRow = CreateObject("Scripting.Dictionary")
    ReDim mUnits(1 To Application.Max(2, mRawLast))
    Set ws = ThisWorkbook.Worksheets(SHEET_RAW)
    For i = 2 To mRawLast
        area = S(mRaw(i, 3))
        If Len(area) = 0 Then area = "Utan delområde"
        rule = GetRule(S(mRaw(i, 2)), S(mRaw(i, 12)), S(mRaw(i, 1)), area)
        mRulesByRow.Add CStr(i), rule
        mUnits(i) = DetectUnit(ws, i, CStr(rule(3)))
    Next i
    Set tmpl = ThisWorkbook.Worksheets("Rapportmall")
    mTemplate = tmpl.Range("A1:L14").FormulaR1C1
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
            .Name = sc.Font.Name
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

Private Sub CopyTemplateRow(ByVal dest As Worksheet, ByVal outRow As Long, ByVal templateRow As Long)
    Dim c As Long, value As Variant
    mRowType(outRow) = templateRow
    For c = 1 To 12
        value = mTemplate(templateRow, c)
        If VarType(value) = vbString Then
            If Left$(value, 1) = "=" Then
                value = Application.ConvertFormula(value, xlR1C1, xlA1, xlAbsolute, dest.Cells(outRow, c))
                mFormula(outRow, c) = True
            End If
        End If
        mOutput(outRow, c) = value
    Next c
End Sub

Private Sub ApplyRowFormats(ByVal ws As Worksheet, ByVal selected As Range, ByVal templateRow As Long)
    Dim c As Long
    For c = 1 To 12
        Intersect(selected, ws.Columns(c)).Style = mTemplateStyles(templateRow, c)
    Next c
    selected.RowHeight = ThisWorkbook.Worksheets("Rapportmall").Rows(templateRow).RowHeight
End Sub

Private Sub FlushReport(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal photos As Object)
    Dim block() As Variant, r As Long, c As Long, t As Variant, batch As Range
    Dim count As Long, sc As Range, merged As Range, job As Variant, tick As Double
    Dim tmpl As Worksheet, photoRows As Collection
    tick = Timer
    ReDim block(1 To lastRow, 1 To 12)
    For r = 1 To lastRow
        For c = 1 To 12
            block(r, c) = mOutput(r, c)
            ' Formula-skrivning ska inte tolka littera, ID eller text som formler/tal.
            If Not mFormula(r, c) And VarType(block(r, c)) = vbString Then
                If Len(block(r, c)) > 0 Then block(r, c) = "'" & block(r, c)
            End If
        Next c
    Next r
    ws.Range("A1").Resize(lastRow, 12).Formula = block
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
                If merged.Row = CLng(t) And merged.Column = c And merged.Rows.Count = 1 _
                   And c + merged.Columns.Count - 1 <= 12 Then
                    For r = 1 To lastRow
                        If mRowType(r) = CLng(t) Then ws.Cells(r, c).Resize(1, merged.Columns.Count).Merge
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
    mFormula(2, 1) = False: mOutput(2, 1) = "LEVANDE MÄNGDER  ·  Materialinventering / mängder och återbruk"
    headers = Array("LITTRERING", "MATERIALTYP", "m²", "lpm", "Antal", _
       "Sparas %", "Sparas mängd", "SKICK", "KOMMENTAR", "BILDER", "DELOMRÅDE", "INFO-ID-REFERENS")
    For c = 1 To 12
        mOutput(4, c) = headers(c - 1)
        mFormula(4, c) = False
    Next c
    ws.Outline.SummaryRow = xlAbove
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
    RngRaw = "'Rådata'!$" & columnLetter & "$2:$" & columnLetter & "$" & LIVE_LAST_ROW
End Function

Private Function FindByID(ByVal rawCol As String, ByVal outRow As Long) As String
    Dim expr As String
    expr = "INDEX(" & RngRaw(rawCol) & ",MATCH($L" & outRow & "," & RngRaw("H") & ",0))"
    FindByID = "IFERROR(IF(LEN(" & expr & ")=0,""""," & expr & "),"""")"
End Function

Private Function GroupCondition(ByVal sourceRow As Long) As String
    Dim cellRef As String
    ' Samma källrad som gruppens kommentar/procent; inget ID-uppslag behövs.
    cellRef = "'Rådata'!$D$" & sourceRow
    GroupCondition = "IF(LEN(" & cellRef & ")=0,""""," & cellRef & ")"
End Function

Private Function ExactConditionCriteria(ByVal conditionExpr As String) As String
    ' COUNTIFS tolkar annars *, ? och inledande jämförelsetecken som villkor.
    ExactConditionCriteria = """=""&SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(" & _
        conditionExpr & ",""~"",""~~""),""*"",""~*""),""?"",""~?"")"
End Function

Private Function GroupTerms(ByVal wsRaw As Worksheet, ByVal rows As Collection, _
                            ByVal outRow As Long, ByVal metric As String, _
                            ByVal unit As String, ByVal splitComment As Boolean, _
                            ByVal splitPercent As Boolean) As String
    Dim seen As Object, v As Variant, rr As Long, key As String
    Dim cls As String, area As String, args As String, product As String
    Dim term As String, result As String, amtCol As String, countArgs As String
    Dim sourceFirst As Long
    Set seen = CreateObject("Scripting.Dictionary")
    sourceFirst = CLng(rows(1))
    If unit = "lpm" Then
        amtCol = "K"
    ElseIf unit = "m²" Then
        amtCol = "J"
    Else
        amtCol = "I"
    End If
    For Each v In rows
        rr = CLng(v)
        cls = S(mRaw(rr, 12))
        area = S(mRaw(rr, 3))
        key = UCase$(cls) & Chr$(30) & UCase$(area)
        If Not seen.Exists(key) Then
            seen.Add key, True
            args = RngRaw("A") & ",$A" & outRow & "," & RngRaw("B") & ",$B" & outRow & _
                "," & RngRaw("C") & "," & XlText(area) & "," & RngRaw("L") & "," & XlText(cls)
            product = "--(" & RngRaw("A") & "=$A" & outRow & ")*--(" & RngRaw("B") & "=$B" & outRow & ")" & _
                "*--(" & RngRaw("C") & "=" & XlText(area) & ")*--(" & RngRaw("L") & "=" & XlText(cls) & ")"
            If splitComment Then
                args = args & "," & RngRaw("E") & ",'Rådata'!$E$" & sourceFirst
                product = product & "*--(" & RngRaw("E") & "='Rådata'!$E$" & sourceFirst & ")"
            End If
            If splitPercent Then
                args = args & "," & RngRaw("F") & ",'Rådata'!$F$" & sourceFirst
                product = product & "*--(" & RngRaw("F") & "='Rådata'!$F$" & sourceFirst & ")"
            End If
            Select Case metric
                Case "QTY"
                    term = "SUMIFS(" & RngRaw(amtCol) & "," & args & ")"
                Case "COUNT"
                    term = "COUNTIFS(" & args & ")"
                Case "MISSING"
                    term = "COUNTIFS(" & args & "," & RngRaw("F") & "," & Chr$(34) & Chr$(34) & ")"
                Case "SAVED"
                    term = "SUMPRODUCT(" & product & "*" & RngRaw(amtCol) & "*" & RngRaw("F") & ")"
                Case "CONDITION"
                    term = "COUNTIFS(" & args & "," & RngRaw("D") & "," & ExactConditionCriteria(GroupCondition(sourceFirst)) & ")"
                Case Else
                    term = "0"
            End Select
            If Len(result) > 0 Then result = result & "+"
            result = result & term
        End If
    Next v
    If Len(result) = 0 Then result = "0"
    GroupTerms = "(" & result & ")"
End Function

Private Sub LiveReportGroup(ByVal ws As Worksheet, ByVal wsRaw As Worksheet, _
                            ByVal headerRow As Long, ByVal rows As Collection, _
                            ByVal lastDetail As Long, ByVal unit As String, _
                            ByVal splitComment As Boolean, ByVal splitPercent As Boolean)
    Dim qty As String, saveQty As String, countAll As String, missing As String
    Dim sameCondition As String, stepCell As String, rowIx As Long
    Dim rawFirstID As String, mainCol As Long
    qty = GroupTerms(wsRaw, rows, headerRow, "QTY", unit, splitComment, splitPercent)
    saveQty = GroupTerms(wsRaw, rows, headerRow, "SAVED", unit, splitComment, splitPercent)
    countAll = GroupTerms(wsRaw, rows, headerRow, "COUNT", unit, splitComment, splitPercent)
    missing = GroupTerms(wsRaw, rows, headerRow, "MISSING", unit, splitComment, splitPercent)
    sameCondition = GroupTerms(wsRaw, rows, headerRow, "CONDITION", unit, splitComment, splitPercent)
    rawFirstID = S(mRaw(CLng(rows(1)), 8))
    If unit = "lpm" Then
        mainCol = 4
        stepCell = "'Rapportval'!$B$13"
    ElseIf unit = "m²" Then
        mainCol = 3
        stepCell = "'Rapportval'!$B$14"
    Else
        mainCol = 5
        stepCell = "0"
    End If
    mFormula(headerRow, mainCol) = True: mOutput(headerRow, mainCol) = "=IF(" & countAll & "=0,"""",IF(" & stepCell & _
         ">0,ROUND(" & qty & "/" & stepCell & ",0)*" & stepCell & "," & qty & "))"
    mFormula(headerRow, 6) = True: mOutput(headerRow, 6) = "=IF(OR(" & qty & "=0," & missing & ">0),""""," & saveQty & "/" & qty & ")"
    mFormula(headerRow, 7) = True: mOutput(headerRow, 7) = "=IF(OR(" & qty & "=0," & missing & ">0),"""",IF(" & stepCell & _
         ">0,ROUND(" & saveQty & "/" & stepCell & ",0)*" & stepCell & "," & saveQty & "))"
    ' Villkorskolumnen reagerar också om objektens skick ändras.
    mFormula(headerRow, 8) = True: mOutput(headerRow, 8) = "=IF(" & countAll & "=0,"""",IF(" & sameCondition & "=" & countAll & _
          "," & GroupCondition(CLng(rows(1))) & ",""Varierande""))"
    If rows.Count > 1 Then
        mFormula(headerRow, 9) = True: mOutput(headerRow, 9) = "=" & countAll & "&"" objekt · öppna (+)"""
    Else
        mFormula(headerRow, 9) = True: mOutput(headerRow, 9) = "=IF(" & countAll & ">1," & countAll & _
            "&"" objekt · uppdatera för nya rader""," & FindByID("E", headerRow) & ")"
    End If
    If rows.Count = 1 Then
        mFormula(headerRow, 11) = True: mOutput(headerRow, 11) = "=" & FindByID("C", headerRow)
    End If
    ' Länk till rådata via ID för redan genererade enskilda objektrader.
    If rows.Count = 1 Then
        rawFirstID = S(mRaw(CLng(rows(1)), 8))
        mFormula(headerRow, 12) = False: mOutput(headerRow, 12) = mRaw(CLng(rows(1)), 8)
        mFormula(headerRow, 2) = True: mOutput(headerRow, 2) = "=" & FindByID("B", headerRow)
    Else
        For rowIx = headerRow + 1 To lastDetail
            mFormula(rowIx, 1) = True: mOutput(rowIx, 1) = "=IF(" & FindByID("A", rowIx) & "="""","""",""  » ""&" & FindByID("A", rowIx) & ")"
            mFormula(rowIx, 2) = True: mOutput(rowIx, 2) = "=" & FindByID("B", rowIx)
            mFormula(rowIx, 3) = True: mOutput(rowIx, 3) = "=" & FindByID("J", rowIx)
            mFormula(rowIx, 4) = True: mOutput(rowIx, 4) = "=" & FindByID("K", rowIx)
            mFormula(rowIx, 5) = True: mOutput(rowIx, 5) = "=" & FindByID("I", rowIx)
            mFormula(rowIx, 6) = True: mOutput(rowIx, 6) = "=" & FindByID("F", rowIx)
            mFormula(rowIx, 7) = True: mOutput(rowIx, 7) = "=IF(ISNUMBER(F" & rowIx & "),SUM(C" & rowIx & ":E" & rowIx & ")*F" & rowIx & ","""")"
            mFormula(rowIx, 8) = True: mOutput(rowIx, 8) = "=" & FindByID("D", rowIx)
            mFormula(rowIx, 9) = True: mOutput(rowIx, 9) = "=" & FindByID("E", rowIx)
            mFormula(rowIx, 11) = True: mOutput(rowIx, 11) = "=" & FindByID("C", rowIx)
        Next rowIx
    End If
End Sub

Private Function PhotoNames(ByVal text As String, ByVal re As Object) As Collection
    Dim result As New Collection, seen As Object, matches As Object
    Dim m As Object, part As Variant, filename As String
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Set matches = re.Execute(text)
    If matches.Count > 0 Then
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
    last = LastRowIn(raw, 1)
    If last < 2 Then Exit Sub
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "FOTO\s*\(([^)]*)\)"
    re.Global = True
    re.IgnoreCase = True
    ' Läs en hel kolumn en gång; rad 1 inkluderas så att även ett objekt ger en array.
    data = raw.Range("G1:G" & last).Value2
    For i = 2 To last
        Set names = PhotoNames(S(data(i, 1)), re)
        If names.Count > 0 Then photos.Add CStr(i), names
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
    If seen.Count = 0 Then Exit Sub
    label = firstName & " (" & seen.Count & " bilder)"
    If seen.Count = 1 Then label = firstName & " (1 bild)"
    path = ImagePath(firstName)
    If Len(path) = 0 Then
        target.Value = label & " - ange fotomapp i Rapportval!B11"
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
                groupKey = CStr(rule(0)) & Chr$(30) & CStr(rule(1)) & Chr$(30) & CStr(rule(2)) & Chr$(30) _
                    & lit & Chr$(30) & mat & Chr$(30) & unit
                If separateComments Then groupKey = groupKey & Chr$(30) & valComment
                If separatePct Then groupKey = groupKey & Chr$(30) & S(mRaw(i, 6))
                If Not groups.Exists(groupKey) Then
                    Set rows = New Collection
                    groups.Add groupKey, rows
                    totals.Add groupKey, 0#
                    saved.Add groupKey, 0#
                    missing.Add groupKey, 0&
                    meta.Add groupKey, Array(CStr(rule(0)), CStr(rule(1)), CStr(rule(2)), lit, mat, unit, condition, area)
                    sortValues.Add groupKey, Format$(N(rule(4)), "000000") & "|" & UCase$(CStr(rule(0))) & "|" & _
                       UCase$(CStr(rule(1))) & "|" & UCase$(CStr(rule(2))) & "|" & UCase$(lit) & "|" & _
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
    If groups.Count = 0 Then
        mFormula(pos, 1) = False: mOutput(pos, 1) = "Inga objekt med de valda kriterierna."
        FlushReport outputSheet, pos, photos
        Exit Sub
    End If
    keys = groups.Keys
    If groups.Count > 1 Then SortKeys keys, sortValues, LBound(keys), UBound(keys)
    prev1 = "": prev2 = "": prev3 = ""
    For Each key In keys
        info = meta(key)
        l1 = CStr(info(0)): l2 = CStr(info(1)): l3 = CStr(info(2))
        If l1 <> prev1 Then
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
        nRecords = rows.Count
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
            outputSheet.Rows(startDetail & ":" & lastDetail).Group
            outputSheet.Rows(startDetail & ":" & lastDetail).Hidden = True
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
        .Range("F5:F" & pos).NumberFormat = "0%"
        .Rows("4:4").AutoFilter
        .PageSetup.PrintArea = "$A$1:$K$" & pos - 1
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
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    PrepareReportData
    Set raw = ThisWorkbook.Worksheets(SHEET_RAW)
    Set areas = CreateObject("Scripting.Dictionary")
    For i = 2 To LastRowIn(raw, 1)
        If Not Excluded(raw, i) Then
            areaValue = S(raw.Cells(i, 3).Value2)
            If Len(areaValue) = 0 Then areaValue = "Utan delområde"
            If Not areas.Exists(areaValue) Then areas.Add areaValue, True
        End If
    Next i
    LoggaTid "Förberedelser", stegTid
    Application.StatusBar = "Rapporter: rensar gamla rapporter..."
    ' Endast tidigare genererade OMR-flikar tas bort. Använd inte detta prefix för egna flikar.
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        If Left$(ws.Name, Len(GENERATED_PREFIX)) = GENERATED_PREFIX Then
            If Left$(S(ws.Range("A2").Value2), Len("GENERERAD AUTOMATISKT")) = "GENERERAD AUTOMATISKT" _
               Or Left$(S(ws.Range("A2").Value2), Len("LEVANDE MÄNGDER")) = "LEVANDE MÄNGDER" Then ws.Delete
        End If
    Next i
    Application.DisplayAlerts = True
    Set ws = ThisWorkbook.Worksheets(SHEET_SUM)
    On Error Resume Next
    ws.UsedRange.ClearOutline
    On Error GoTo Failed
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
    For i = 1 To ThisWorkbook.Worksheets.Count
        names(UCase$(ThisWorkbook.Worksheets(i).Name)) = True
    Next i
    LoggaTid "Förbereda delområden", stegTid
    For Each area In areas.Keys
        Application.StatusBar = "Rapporter: bygger " & CStr(area) & " (" & (generated + 1) & "/" & areas.Count & ")..."
        name = SafeSheetName(CStr(area), names)
        Set areaSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        areaSheet.Name = name
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
    Application.EnableEvents = oldEvents
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = oldCalc
    MsgBox "Fel vid rapportgenerering: " & Err.Number & " – " & Err.Description, vbExclamation
End Sub

Private Function ExportHeadersMatch(ByVal ws As Worksheet) As Boolean
    ExportHeadersMatch = (StrComp(S(ws.Cells(1, 1).Value2), "LITTRERING", vbTextCompare) = 0 _
                 And StrComp(S(ws.Cells(1, 8).Value2), "INFO-ID-REFERENS", vbTextCompare) = 0)
End Function

Public Sub ImporteraFranVectorworks()
    Dim dialog As FileDialog, srcWB As Workbook, source As Worksheet, candidate As Worksheet, ws As Worksheet
    Dim ids As Object, visited As Object, rowData As Variant, id As String
    Dim srcRow As Long, targetRow As Long, srcLast As Long, lastTarget As Long
    Dim updated As Long, added As Long, missing As Long, duplicate As Long, col As Long
    Dim keepRatings As Boolean, oldCalc As XlCalculation, lo As ListObject
    Dim name As String, oldVal As String, existing As Boolean
    On Error GoTo Failed
    Set dialog = Application.FileDialog(3)
    With dialog
        .Title = "Välj Excel-export från Vectorworks"
        .Filters.Clear
        .Filters.Add "Excel-filer", "*.xlsx;*.xlsm;*.xls"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        name = .SelectedItems(1)
    End With
    If StrComp(name, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        MsgBox "Välj exportfilen från Vectorworks, inte själva rapportarbetsboken.", vbExclamation
        Exit Sub
    End If
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Set srcWB = Workbooks.Open(Filename:=name, UpdateLinks:=0, ReadOnly:=True)
    For Each candidate In srcWB.Worksheets
        If ExportHeadersMatch(candidate) Then
            Set source = candidate
            Exit For
        End If
    Next candidate
    If source Is Nothing Then
        MsgBox "Hittar inte rådatablad med rubrikerna LITTRERING och INFO-ID-REFERENS i kolumn A respektive H.", vbExclamation
        GoTo SafeEnd
    End If
    Set ws = ThisWorkbook.Worksheets(SHEET_RAW)
    Set ids = CreateObject("Scripting.Dictionary")
    Set visited = CreateObject("Scripting.Dictionary")
    lastTarget = LastRowIn(ws, 1)
    If StrComp(S(ThisWorkbook.Worksheets(SHEET_PREF).Range("B16").Value2), "Radera", vbTextCompare) = 0 Then
        ws.Range("A2:L" & Application.Max(2, lastTarget)).ClearContents
        lastTarget = 1
        ThisWorkbook.Worksheets(SHEET_PREF).Range("B16").Value = "Raderade"
    End If
    For targetRow = 2 To lastTarget
        id = UCase$(S(ws.Cells(targetRow, 8).Value2))
        If Len(id) > 0 Then
            If Not ids.Exists(id) Then ids.Add id, targetRow
        End If
    Next targetRow
    keepRatings = IsYes(ThisWorkbook.Worksheets(SHEET_PREF).Range("B5").Value2)
    srcLast = LastRowIn(source, 8)
    For srcRow = 2 To srcLast
        id = UCase$(S(source.Cells(srcRow, 8).Value2))
        If Len(id) = 0 Then
            missing = missing + 1
        ElseIf visited.Exists(id) Then
            duplicate = duplicate + 1
        Else
            visited.Add id, True
            existing = ids.Exists(id)
            If existing Then
                targetRow = CLng(ids(id))
                updated = updated + 1
            Else
                lastTarget = lastTarget + 1
                targetRow = lastTarget
                ids.Add id, targetRow
                added = added + 1
            End If
            For col = 1 To 12
                If keepRatings And existing And (col = 4 Or col = 5 Or col = 6) And _
                   Len(S(ws.Cells(targetRow, col).Value2)) > 0 Then
                    ' Behåll redan ifyllt SKICK, KOMMENTAR och ANDELSPARAS.
                Else
                    ws.Cells(targetRow, col).Value2 = source.Cells(srcRow, col).Value2
                End If
            Next col
        End If
    Next srcRow
    On Error Resume Next
    Set lo = ws.ListObjects("tblRadata")
    If Not lo Is Nothing Then lo.Resize ws.Range("A1:L" & Application.Max(2, lastTarget))
    On Error GoTo Failed
    ws.Range("F2:F" & Application.Max(2, lastTarget)).NumberFormat = "0%"
SafeEnd:
    If Not srcWB Is Nothing Then srcWB.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.Calculation = oldCalc
    If Not source Is Nothing Then
        UppdateraRapporter
        MsgBox "Import klar. Uppdaterade: " & updated & ". Nya: " & added & _
           ". Saknade ID: " & missing & ". Dubbletter i export: " & duplicate & ".", vbInformation
    End If
    Exit Sub
Failed:
    On Error Resume Next
    If Not srcWB Is Nothing Then srcWB.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.Calculation = oldCalc
    MsgBox "Import misslyckades: " & Err.Description, vbExclamation
End Sub

