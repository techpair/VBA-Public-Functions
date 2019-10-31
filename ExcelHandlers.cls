VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelHandlers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Function AlignLeft(rng As range)

'***********************************************************************************
'Purpose:Aligns a range to left
'Inputs: worksheet and starting row number
'Outputs:row number of an instance of a string
'***********************************************************************************

    With rng
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Function

Function BorderizeAndCleanRng(wsDest As Worksheet, rowToBorderize As Long, lastColToBord As Long)
    
    With wsDest
        
        .Activate
        
        With .range(.Cells(rowToBorderize, 1), .Cells(rowToBorderize, lastColToBord)).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Color = RGB(0, 0, 0)
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        
        With .range(.Cells(rowToBorderize, 1), .Cells(rowToBorderize, lastColToBord)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(0, 0, 0)
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        
    End With
    
End Function


Function PlotFormula(sh As Worksheet, strDestAddress As String, strFormula As String, Optional strAutoFillAddress As String, Optional convertToValues As Boolean)
        
    On Error GoTo errorline
        
    With sh
    
        .range(strDestAddress).Formula = strFormula
        
        If strAutoFillAddress <> vbNullString Then
            .range(strDestAddress).AutoFill .range(strAutoFillAddress), xlFillValues
            .range(strAutoFillAddress).Calculate
            
        End If
        
        If convertToValues = True Then
            .range(strAutoFillAddress).Value = .range(strAutoFillAddress).Value
        End If
        
    End With
    Exit Function
errorline:
    MsgBox "Error while plotting formula in worksheet: " & sh.Name & " to cell range " & strDestAddress & " with the formula " & strFormula
    End
End Function

Function ValidateColumnsInExcelWorksheet(wb As Workbook, ws As Worksheet, strKeywords As Variant, Optional xlLkAt As XlLookAt = xlPart)
    On Error GoTo errorline
            
    If IsEmpty(strKeywords) = False Then
        
        Dim inputCtr As Long
        Dim strWsName As String
        Dim strKeyword As String
        Dim searchRng As Variant
            
            
            strWsName = ws.Name
            
            For inputCtr = LBound(strKeywords, 1) To UBound(strKeywords, 1)
                    
                    strKeyword = CStr(strKeywords(inputCtr))
                    
                    On Error Resume Next
                    Set wsWorking = wb.Sheets(strWsName)
                    
                        If Err.Number <> 0 Then
                        
                            MsgBox "Missing worksheet name: '" & strWsName & "'" & vbNewLine & vbNewLine _
                            & "Workbook name: '" & wb.Name & "'" & vbNewLine & vbNewLine _
                            & "Please upload the correct '" & wb.Name & "' file. This will stop the macro.", vbExclamation, "Opening input file"
                            Err.Clear
                            CloseWorkbooks Array(wb)
                            End
                            
                        End If
                    On Error GoTo 0
        
                    With wsWorking
                    
                        Set searchRng = .range(.Cells(1, 1), .Cells(.UsedRange.Rows.Count, .UsedRange.Columns.Count)).Find(What:=strKeyword, LookIn:=xlValues, LookAt:=xlLkAt)
                        If searchRng Is Nothing Then
                        
                            MsgBox "Missing column header: '" & strKeyword & "'" & vbNewLine & vbNewLine _
                            & "Worksheet name: '" & strWsName & "'" & vbNewLine & vbNewLine _
                            & "Workbook name: '" & wb.Name & "'" & vbNewLine & vbNewLine _
                            & "Please upload the correct '" & wb.Name & "' file. This will stop the macro.", vbExclamation, "Opening input file"
                            CloseWorkbooks Array(wb)
                            Err.Clear
                            End
                        
                        End If
                    
                    End With
            
            Next inputCtr
    End If
    
    Exit Function
errorline:
     ErrorHandler.ErrorHandler "Error while validating columns in excel worksheet " & ws.Name
End Function



Function PlotRankToDiscontinuingData(ws As Worksheet, ctrCol As Long, destCol As Long, rowStart As Long)
    
    Dim runningRow As Long, ctr As Long
    
    'this assumes that the discontinuing data is sorted already
    
    With ws
        Do While runningRow <> .Cells.Rows.Count
            If runningRow = 0 Then runningRow = rowStart
            ctr = ctr + 1
            
            If InStr(1, .Cells(runningRow, ctrCol).Value, "Total", vbTextCompare) = 0 Then
            .Cells(runningRow, destCol).Value = ctr
            End If
            
            If .Cells(runningRow + 1, ctrCol).Value = "" Then
                runningRow = .Cells(runningRow, ctrCol).End(xlDown).row
            Else
                runningRow = runningRow + 1
            End If
        Loop
    End With
    
End Function

Function AdjustColumnSizes(ws As Worksheet, columnSizes As Variant)
    
    Dim columnSize As Variant, longColSize As Double
    Dim colCtr As Long
    
    With ws
        For Each columnSize In columnSizes
            If IsNumeric(columnSize) = True Then
                longColSize = columnSize
                colCtr = colCtr + 1
                ws.Columns(colCtr).ColumnWidth = longColSize
                
            End If
        Next columnSize
    End With
    
End Function



Function FormatRange(ws As Worksheet, rngToFormat As range, colorOfFont As Long, fontSizeOfRange As Long, strFontName As String, _
    Optional border1 As XlBordersIndex, Optional border2 As XlBordersIndex, Optional border3 As XlBordersIndex, Optional border4 As XlBordersIndex, _
    Optional border1style As XlLineStyle, Optional border2style As XlLineStyle, Optional border3style As XlLineStyle, Optional border4style As XlLineStyle, _
    Optional border1weight As XlBorderWeight, Optional border2weight As XlBorderWeight, Optional border3weight As XlBorderWeight, Optional border4weight As XlBorderWeight, _
    Optional blnIsBold As Boolean = False, Optional blnIsWrapText As Boolean = False, Optional blnIsMerged As Boolean = False, _
    Optional border1color As Long, Optional border2color As Long, Optional border3color As Long, Optional border4color As Long, Optional mergeCenter As Boolean = False)    'changing font-colors and font-sizes
    
    With ws
        
        With .range(rngToFormat.Address).SpecialCells(xlCellTypeVisible)
            
            With .Font
                .Color = colorOfFont '-16776961
                '.TintAndShade = 0
            
                .Name = strFontName '"Calibri"
                .Size = fontSizeOfRange '10
                .Bold = blnIsBold
                
                '.Strikethrough = False
    '            .Superscript = False
    '            .Subscript = False
    '            .OutlineFont = False
    '            .Shadow = False
    '            .Underline = xlUnderlineStyleNone
    '            .Color = -16776961
    '            .TintAndShade = 0
    '            .ThemeFont = xlThemeFontMinor
    
            End With
            
            .WrapText = blnIsWrapText
            .MergeCells = blnIsMerged
            If mergeCenter = True Then .HorizontalAlignment = xlCenter
            
            If border1 <> 0 And border1style <> 0 Then
                
                With .Borders(border1)
                    .LineStyle = border1style
                    .Weight = border1weight
                    .Color = border1color
                End With
                
            End If
            
            If border2 <> 0 And border2style <> 0 Then
                With .Borders(border2)
                    .LineStyle = border1style
                    .Weight = border1weight
                    .Color = border2color
                End With
            End If
            
            If border3 <> 0 And border3style <> 0 Then
                With .Borders(border3)
                    .LineStyle = border1style
                    .Weight = border1weight
                    .Color = border3color
                End With
            End If
            
            If border4 <> 0 And border4style <> 0 Then
                With .Borders(border4)
                    .LineStyle = border1style
                    .Weight = border1weight
                    .Color = border4color
                End With
            End If
            
            
            
            
            
        End With
        
        
'        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'        Selection.Borders(xlEdgeLeft).LineStyle = xlNone
'        Selection.Borders(xlEdgeTop).LineStyle = xlNone
'
'        With Selection.Borders(xlEdgeBottom)
'            .LineStyle = xlContinuous
'            .ThemeColor = 2
'            .TintAndShade = 0
'            .Weight = xlThin
'        End With
'
'        Selection.Borders(xlEdgeRight).LineStyle = xlNone
'        Selection.Borders(xlInsideVertical).LineStyle = xlNone
'        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
'
'        Range("C13").Select
        
    End With
    
End Function




Sub ProcessRankLevelData(ws As Worksheet, refCol As Long, startRow As Long, lastRow As Long, columnsToAnalyse As Variant, processesToAnalyse As Variant)
    
    Dim arrayColumns As Variant, verticalCtr As Long, arrayColumn As Variant, columnNum As Long, arrayProcess As Variant, strProcess As Variant, arrCtr As Long
    Dim lastRangeRow As Long
    
    arrayColumns = columnsToAnalyse
    arrayProcess = processesToAnalyse
    
    With ws
        For verticalCtr = startRow To lastRow
            lastRangeRow = 0
            
            For arrCtr = LBound(arrayColumns, 1) To UBound(arrayColumns, 1)
                columnNum = arrayColumns(arrCtr)
                
                strProcess = arrayProcess(arrCtr)
                
                If columnNum <> 0 Then
                    Select Case strProcess
                        Case "clear"
                            .Cells(verticalCtr, columnNum).ClearContents
                            
                        Case "equal below"
                            .Cells(verticalCtr, columnNum).Value = .Cells(verticalCtr + 1, columnNum).Value
                        Case "sum rank group"
                            lastRangeRow = .Cells(verticalCtr, refCol).End(xlDown).row - 1
                            
                            If lastRangeRow = .Cells.Rows.Count - 1 Then
                                lastRangeRow = lastRow '.Cells(verticalCtr, columnNum).End(xlDown).row
                            End If
                            
                            .Cells(verticalCtr, columnNum).Value = Application.WorksheetFunction.Sum( _
                                .range(.Cells(verticalCtr + 1, columnNum), .Cells(lastRangeRow, columnNum)))
                            
                    End Select
                End If
                
            Next arrCtr
            
             If verticalCtr = .Cells.Rows.Count - 1 Then
                verticalCtr = lastRow '.Cells(verticalCtr, columnNum).End(xlDown).row
             End If
            
            
            verticalCtr = .Cells(verticalCtr, refCol).End(xlDown).row - 1  ' minus 1 to handle the "Next" command below or else it will skip the ranks row
'
             If verticalCtr = .Cells.Rows.Count - 1 Then
                verticalCtr = lastRow '.Cells(verticalCtr, columnNum).End(xlDown).row
             End If
'
            If verticalCtr = .Cells.Rows.Count Then
                Exit For
            End If
            
        Next verticalCtr
    End With
End Sub


Function AlignRange(ws As Worksheet, rng As range, vAlign As Variant, hAlign As Variant, _
    Optional blnIsWrapText As Boolean = False, Optional blnIsMerged As Boolean = False)
                
    With ws.range(rng.Address)
            
        .VerticalAlignment = vAlign
        .HorizontalAlignment = hAlign
        .WrapText = blnIsWrapText
        .MergeCells = blnIsMerged
        
    End With
    
End Function

Function NumberFormatRange(ws As Worksheet, lastCol As Long, lastRow As Long, startRow As Long, keywords As Variant, formats As Variant, Optional colCtrr As Long = 1)
    
    Dim keyword As String, formatt As String, keywordCtr As Long
    Dim colCtr As Long
    
    With ws
        For colCtr = colCtrr To lastCol
            
            For keywordCtr = LBound(keywords, 1) To UBound(keywords, 1)
                
                keyword = keywords(keywordCtr)
                formatt = formats(keywordCtr)
                    
                On Error Resume Next
                
                    If .Application.WorksheetFunction.CountIf(.Columns(colCtr), keyword) > 0 Then
                        .range(.Cells(startRow, colCtr), .Cells(lastRow, colCtr)).NumberFormat = formatt
                        '.range(rng.Address).NumberFormat = formatt
                    
                    End If
                    
                    If .Application.WorksheetFunction.CountIf(.Columns(colCtr), "*" & keyword & "*") > 0 Then
                        .range(.Cells(startRow, colCtr), .Cells(lastRow, colCtr)).NumberFormat = formatt
                    End If
                    
                    
                On Error GoTo 0
                Err.Clear
                                
            Next keywordCtr
        Next colCtr
    End With
    
End Function

Function BorderizeRng(ws As Worksheet, rngToFormat As range, _
    border1 As XlBordersIndex, border1style As XlLineStyle, border1weight As XlBorderWeight, border1color As Long)
    
    With ws
        
        With .range(rngToFormat.Address).SpecialCells(xlCellTypeVisible)
            
            
            If border1 <> 0 And border1style <> 0 Then
                
                With .Borders(border1)
                    .LineStyle = border1style
                    .Weight = border1weight
                    .Color = border1color
                End With
                
            End If
            
        End With
        
        
    End With
    
End Function

Function CloseAllWb(workbooknames As Variant)
    
    Dim wb As Workbook, var As Variant
    
    For Each var In workbooknames
        On Error Resume Next
        Set wb = Workbooks(var)
        CloseWorkbooks Array(wb)
    Next var
    
    On Error GoTo 0
    Err.Clear

End Function


Function ConsolidateSameColumns(wb1 As Workbook, arrayWb As Variant, arrayWs As Variant, ByRef arrayColumns As Variant, Optional plotInputTagging As Boolean = False, _
    Optional plotGroupTagging As Boolean = False) As Long
    
    Dim lastRow As Long, startRow As Long, refCol As Long, destStartRow As Long, destLastRow As Long, destRefCol As Long
    Dim refColForLastRow As Long, destRefColForLastRow As Long
    
    Dim wbCtr As Long, keywordCtr As Long, strKeyword As String
    Dim wb As Workbook, ws As Worksheet, destWb As Workbook, destWs As Worksheet
    Dim destLastCol As Long
    Dim inputCount As Long
    
    Dim strWorksheetName As String
    Dim configFilenameRow  As Long
    Dim inputTaggingCol As Long
    Dim groupTaggingCol As Long
    Dim groupRefDtCol As Long
    Dim strGroupName As String
    Dim mktValCol As Long, dblSharePrice As Double
    
On Error GoTo errorline
    
    'This function assumes you already know the worksheet names to combine.
    wb1.Sheets.Add
    Set destWb = wb1
    Set destWs = destWb.ActiveSheet
    'Workbooks.Add
    
    'Set destWb = ActiveWorkbook
    
    'Set destWs = destWb.Sheets(1)
    destWs.Name = "Conso"
        
    With destWs
        .range(.Cells(1, LBound(arrayColumns) + 1), .Cells(1, UBound(arrayColumns) + 1)).Value = arrayColumns   'build the destination column headers
        
        If plotInputTagging = True Then
            destLastCol = .UsedRange.Columns.Count + 1
            .Cells(1, destLastCol).Value = "Input Files"
            inputTaggingCol = destLastCol
                    
        End If
        
        If plotGroupTagging = True Then
            destLastCol = destLastCol + 1
            .Cells(1, destLastCol).Value = "Grouping"
            groupTaggingCol = destLastCol
            
            destLastCol = destLastCol + 1
            .Cells(1, destLastCol).Value = "Group Reference Date"
            groupRefDtCol = destLastCol
            
            
            destLastCol = destLastCol + 1
            .Cells(1, destLastCol).Value = "Market Value"
            mktValCol = destLastCol
            
        End If
        
    End With
    
    For wbCtr = LBound(arrayWb, 1) To UBound(arrayWb, 1)    'looping through the selected filenames
        
        Set wb = Workbooks(CStr(arrayWb(wbCtr)))
            
        ExlHandler.SetSheet wb, CStr(arrayWs(wbCtr)), ws   'setting sheet
        
        'Set ws = wb.Sheets(CStr(arrayWs(wbCtr)))
        
        
            With ws
                startRow = FindRow(wb, ws, .UsedRange, CStr(arrayColumns(LBound(arrayColumns, 1))), xlWhole)
                refCol = FindColumn(wb, ws, .UsedRange, CStr(arrayColumns(LBound(arrayColumns, 1))), xlWhole)
                refColForLastRow = FindColumn_continuous(wb, ws, .UsedRange, "Investment Manager", xlWhole)
                
                If refColForLastRow <> 0 Then
                    lastRow = .Cells(.Cells.Rows.Count, refColForLastRow).End(xlUp).row
                Else
                    lastRow = .Cells(.Cells.Rows.Count, refCol).End(xlUp).row
                End If
                
            End With
            
            With destWs
                destStartRow = FindRow(destWb, destWs, .UsedRange, CStr(arrayColumns(LBound(arrayColumns, 1))), xlWhole)
                
                destRefColForLastRow = FindColumn_continuous(destWb, destWs, .UsedRange, "Investment Manager", xlWhole)  'FindColumn(destWB, destWs, .UsedRange, CStr(arrayColumns(LBound(arrayColumns, 1))), xlWhole)
                
                destRefCol = FindColumn(destWb, destWs, .UsedRange, CStr(arrayColumns(LBound(arrayColumns, 1))), xlWhole)
                
                If destRefColForLastRow <> 0 Then
                    destLastRow = .Cells(.Cells.Rows.Count, destRefColForLastRow).End(xlUp).row
                Else
                    destLastRow = .Cells(.Cells.Rows.Count, destRefCol).End(xlUp).row
                End If
                
                
                If destLastRow = .Cells.Rows.Count Then destLastRow = destStartRow
            End With
            
            'copy-paste data
            
            configFilenameRow = 0
            strGroupName = vbNullString
            
            For keywordCtr = LBound(arrayColumns, 1) To UBound(arrayColumns, 1) 'loop through column arrays for matching
                strKeyword = arrayColumns(keywordCtr)
                
                If strKeyword = "Previous Shares" Then
                    
                    On Error Resume Next
                    configFilenameRow = configSheet.Application.WorksheetFunction.Match(CStr(arrayWs(wbCtr)), configSheet.Columns("A:A"), 0)
                    On Error GoTo 0
                    Err.Clear
                    
                    'this will list all the previous shares of the uploaded input files if they are available in the config tab
                    
                    If configFilenameRow > 0 Then
                        strKeyword = "Shares " & Format(configSheet.Cells(configFilenameRow, 3).Value, "DD/MM/YYYY")
                    Else
                        ErrorHandler.ErrorHandler "Macro cannot find the previous date value of the filename: " & CStr(arrayWs(wbCtr)) & vbNewLine & vbNewLine & _
                        "Please re-configure the input files' setup in the macro file tab 'Input File Config'."
                    End If
                    
                    
                End If
                
                refCol = FindColumn(wb, ws, ws.UsedRange, strKeyword, xlWhole)
                
                With ws
                    .range(.Cells(startRow + 1, refCol), .Cells(lastRow, refCol)).Copy
                End With
                
                destWb.Activate
                
                With destWs
                    .Activate
                    destRefCol = FindColumn_continuous(destWb, destWs, .UsedRange, strKeyword, xlWhole)
                    
                    If destRefCol = 0 Then
                        destLastCol = destLastCol + 1
                        destRefCol = destLastCol
                        .Cells(destStartRow, destRefCol).Value = strKeyword
                    End If
                    
                    .Cells(destLastRow + 1, destRefCol).Select
                    ActiveSheet.Paste
                    
                    
                    If destLastRow = .Cells.Rows.Count Then
                        MsgBox "There are no more rows to copy the data into. Will exit consolidating files with same columns. Please check column " & destRefCol
                        Exit Function
                    End If
                    
                End With
                
            Next keywordCtr
            
            'plot input file tagging
            
            With destWs
                
                If plotInputTagging = True And .Cells(.Cells.Rows.Count, destRefCol).End(xlUp).row <> .Cells.Rows.Count Then
                    .range(.Cells(destLastRow + 1, inputTaggingCol), .Cells(.Cells(.Cells.Rows.Count, destRefCol).End(xlUp).row, inputTaggingCol)).Value = ws.Name
                End If
                                
                
                inputCount = inputCount + 1
                destLastRow = .Cells(.Cells.Rows.Count, destRefCol).End(xlUp).row
            End With
            
        Set wb = Nothing
        Set ws = Nothing
       
    Next wbCtr
        Set destWb = Nothing
        Set destWs = Nothing
        
        ConsolidateSameColumns = inputCount
    Exit Function
errorline:
    ErrorHandler.ErrorHandler "Error while consolidating same columns."
    
End Function

Sub CopyWorksheet(destWb As Workbook, arraySrcFilenames As Variant, arrayWsNames As Variant, strAfterThisSheet As String, Optional blnHide As Boolean = False, _
    Optional ByRef wsToSave As Worksheet, Optional shouldSaveSheet As Boolean)
    
    Dim varSrcFilename As Variant, strSrcFilename As String, strWsName As String
    Dim ctr As Long
    Dim wbSrc As Workbook
    
    On Error Resume Next
    
    For ctr = LBound(arraySrcFilenames) To UBound(arraySrcFilenames)
        varSrcFilename = arraySrcFilenames(ctr)
        strSrcFilename = varSrcFilename
        strSrcFilename = Split(strSrcFilename, ".")(0)
        
        
        strWsName = arrayWsNames(ctr)
        
        Set wbSrc = Workbooks(strSrcFilename)
        
        If strWsName <> vbNullString Then
            
            wbSrc.Sheets(strWsName).Copy After:=destWb.Sheets(strAfterThisSheet)
            
            If shouldSaveSheet = True Then
                Set wsToSave = ActiveSheet 'destWb.Sheets(strWsName)
            End If
            
        Else
            wbSrc.Sheets(strSrcFilename).Copy After:=destWb.Sheets(strAfterThisSheet)
            
            If shouldSaveSheet = True Then
                Set wsToSave = ActiveSheet 'destWb.Sheets(strSrcFilename)
            End If
            
        End If
        
        
        
        If blnHide = True Then ActiveSheet.Visible = xlSheetHidden
        
    Next ctr
    
    Exit Sub
    
errorline:
    ErrorHandler.ErrorHandler "Error while copying worksheet. Please check programming function used: CopyWorksheet"
    
End Sub


Sub ClearFreezePanes()
' Takes off freeze frames
    If ActiveWindow.FreezePanes = True Then
        ActiveWindow.FreezePanes = False
    End If
End Sub
 Sub ClearNames()
' Delete all names in name manager

    Dim n As Name
    On Error Resume Next
    For Each n In ActiveWorkbook.Names
            n.Delete
    Next n
    On Error GoTo 0

End Sub

Sub ClearConditionalFormatting()
'
' Delete all conditional formatting
'

    On Error Resume Next
    Cells.FormatConditions.Delete
    On Error GoTo 0

End Sub
Sub CopyPasteValues(ws As Worksheet, rng As range)
'
' Copy > paste > values - literally
' If rng is given, then...  well, duh.  If not, it'll do the entire sheet.
'
    'PrintToLog ("Calling CopyPasteValues")
    
    On Error GoTo errorline
    
    ws.Activate
    If rng Is Nothing Then
        ws.Cells.Select
    Else
        rng.Select
    End If
    
    With ws.range(rng.Address)
        .Copy
        .PasteSpecial xlPasteValues
    End With
    
    Application.CutCopyMode = False
    
    Exit Sub
errorline:
    ErrorHandler.ErrorHandler "Error while copy-pasting into values the sheet: " & ws.Name & " with the range of cells " & rng.Address
    
End Sub


Function GetLastVisibleColumn(ws As Worksheet, Optional row As Integer = 1) As Long
'
' Finds last visible column
' If row number is given, it'll check against that
' Else, it'll default to row 1
'
    Dim ILastCol As Long
    'PrintToLog ("Calling DeclareLastVisibleColumn")
    ILastCol = ws.Cells(row, ws.Columns.Count).End(xlToLeft).Column
    GetLastVisibleColumn = ILastCol
End Function


Sub StripAccents(Optional rng As range)
'
' Stoles RemoveAccents macro from Kerry-Berry
' Then pimped it to remove other characters too
' And told it to explicilty LookAt:=xlPart
' AND THEN gave it the optional argument of a range - oh yes
'
    If rng Is Nothing Then ' there is definitely a better way of doing that...
        Cells.Select
    Else
        rng.Select
    End If
    With Selection
    .Replace What:="à", replacement:="a", MatchCase:=False, LookAt:=xlPart
    .Replace What:="á", replacement:="a", MatchCase:=False, LookAt:=xlPart
    .Replace What:="â", replacement:="a", MatchCase:=False, LookAt:=xlPart
    .Replace What:="ã", replacement:="a", MatchCase:=False, LookAt:=xlPart
    .Replace What:="ä", replacement:="a", MatchCase:=False, LookAt:=xlPart
    .Replace What:="å", replacement:="a", MatchCase:=False, LookAt:=xlPart
    .Replace What:="æ", replacement:="ae", MatchCase:=False, LookAt:=xlPart
    .Replace What:="ç", replacement:="c", MatchCase:=False, LookAt:=xlPart
    .Replace What:="è", replacement:="e", MatchCase:=False, LookAt:=xlPart
    .Replace What:="é", replacement:="e", MatchCase:=False, LookAt:=xlPart
    .Replace What:="ê", replacement:="e", MatchCase:=False, LookAt:=xlPart
    .Replace What:="ë", replacement:="e", MatchCase:=False, LookAt:=xlPart
    .Replace What:="ì", replacement:="i", MatchCase:=False, LookAt:=xlPart
    .Replace What:="í", replacement:="i", MatchCase:=False, LookAt:=xlPart
    .Replace What:="î", replacement:="i", MatchCase:=False, LookAt:=xlPart
    .Replace What:="ï", replacement:="i", MatchCase:=False, LookAt:=xlPart
    .Replace What:="ñ", replacement:="n", MatchCase:=False, LookAt:=xlPart
    .Replace What:="ò", replacement:="o", MatchCase:=False, LookAt:=xlPart
    .Replace What:="ó", replacement:="o", MatchCase:=False, LookAt:=xlPart
    .Replace What:="ô", replacement:="o", MatchCase:=False, LookAt:=xlPart
    .Replace What:="õ", replacement:="o", MatchCase:=False, LookAt:=xlPart
    .Replace What:="ö", replacement:="o", MatchCase:=False, LookAt:=xlPart
    .Replace What:="ø", replacement:="o", MatchCase:=False, LookAt:=xlPart
    .Replace What:="ù", replacement:="u", MatchCase:=False, LookAt:=xlPart
    .Replace What:="ú", replacement:="u", MatchCase:=False, LookAt:=xlPart
    .Replace What:="û", replacement:="u", MatchCase:=False, LookAt:=xlPart
    .Replace What:="ü", replacement:="u", MatchCase:=False, LookAt:=xlPart
    .Replace What:="ž", replacement:="z", MatchCase:=False, LookAt:=xlPart
    End With
End Sub

Sub RemoveCharacters(ws As Worksheet, Optional rng As range)
' Calls RemoveAccents macro from Kerry-Berry
' Then pimped to remove other characters too
' Told it to explicilty LookAt:=xlPart
' AND gave it the optional argument of a range - oh yes


    With rng
        .Replace What:=" ", replacement:="", MatchCase:=False, LookAt:=xlPart
        .Replace What:=",", replacement:="", MatchCase:=False, LookAt:=xlPart
        .Replace What:=".", replacement:="", MatchCase:=False, LookAt:=xlPart
        .Replace What:="/", replacement:="", MatchCase:=False, LookAt:=xlPart
        .Replace What:="-", replacement:="", MatchCase:=False, LookAt:=xlPart
        .Replace What:=":", replacement:="", MatchCase:=False, LookAt:=xlPart
        .Replace What:="'", replacement:="", MatchCase:=False, LookAt:=xlPart
        .Replace What:=";", replacement:="", MatchCase:=False, LookAt:=xlPart
        .Replace What:="#", replacement:="", MatchCase:=False, LookAt:=xlPart
        .Replace What:="(", replacement:="", MatchCase:=False, LookAt:=xlPart
        .Replace What:=")", replacement:="", MatchCase:=False, LookAt:=xlPart
    End With
End Sub


Sub DeleteSheet(sh As String, wb As Workbook, Optional saveDeletion As Boolean, Optional ByRef ws As Worksheet)
'
' If the sheet exists
' DELETE IT! XD
'
    'PrintToLog ("Calling DeleteSheet")
    If DoesSheetExist(sh, wb) Then
        Application.DisplayAlerts = False
        wb.Sheets(sh).Delete
        Application.DisplayAlerts = True
        If saveDeletion = True Then
            Set ws = Nothing
        End If
    End If
End Sub

Function DoesSheetExist(sh As String, wb As Workbook) As Boolean
' TRUE if sheet exists
' FALSE if sheet does not exist
    'PrintToLog ("Calling DoesSheetExist")
    On Error Resume Next
    DoesSheetExist = (wb.Sheets(sh).Name <> "")
    On Error GoTo 0
End Function


Sub UnhideAllRowsAndColumns(ws As Worksheet)
' Unhides all columns and all rows
    ws.Cells.EntireColumn.Hidden = False
    ws.Cells.EntireRow.Hidden = False
End Sub




Sub UnhideAllSheets(wb As Workbook)

'Unhide all sheets in workbook.

    Dim ws As Worksheet
    
    For Each ws In wb.Worksheets
    
    ws.Visible = xlSheetVisible
    
    Next ws

End Sub


Function CopyPasteFilteredRng(srcWs As Worksheet, destWs As Worksheet, firstFilterRow As Long, tenthFilteredRow As Long, srcCol As Long, destRngStr As String)    'universal sub procedure
    
    '****************************************************************************************************************************
    'Purpose:Copy-pastes filtered range into an unfiltered worksheet range
    'Inputs: source & destination worksheet variables, first & tenth filtered row, source column and destination range address
    'Outputs:Copied range from filtered range source
    '***************************************************************************************************************************
    
    On Error Resume Next
    
    Dim srcRng As range, destRng As range, srcCtr As Integer, destCtr As Integer, srcRngEl As range, destRngEl As range
    
    With srcWs
        .Activate
        .range(.Cells(firstFilterRow, srcCol), .Cells(tenthFilteredRow, srcCol)).Select 'if top 10
             
        Selection.Copy
        
        Set srcRng = Selection.SpecialCells(xlCellTypeVisible)
        
        With destWs
           .range(destRngStr).PasteSpecial xlPasteValues
           Set destRng = .range(destRngStr)
        End With
    End With
    
    Application.CutCopyMode = False
    
    If Err.Number <> 0 Then
        Err.Clear
        
        For Each srcRngEl In srcRng
            srcCtr = srcCtr + 1
            destCtr = 0
            For Each destRngEl In destRng
                destCtr = destCtr + 1
                
                If srcCtr = destCtr Then
                    destWs.range(destRngEl.Address).Value = srcWs.range(srcRngEl.Address).Value
                    Exit For
                End If
                
            Next destRngEl
        Next srcRngEl
        
        On Error GoTo 0
        
    End If
    
End Function

Function SetSheet(wbToCheck As Workbook, strWsNameToCheck As String, ByRef wsVarToSet As Worksheet, Optional varArrayInputFiles As Variant)
    
    '****************************************************************************************************
    'Purpose:Sets a worksheet variable and Validates if the string used as parameter for the sheet exists
    'Inputs: workbook variable to validate, worksheet string name, worksheet variable to set
    'Outputs:worksheet variable or error that the macro did not find the worksheet name
    '****************************************************************************************************
    
    On Error Resume Next
    Set wsVarToSet = wbToCheck.Sheets(strWsNameToCheck)
    
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "The macro did not find this worksheet--> " & strWsNameToCheck & " in this workbook --> " & wbToCheck.Name & ". Please rename your worksheet."
        
        If varArrayInputFiles <> Empty Then
            CloseWorkbooks varArrayInputFiles
        End If
        
        End
    End If
    
End Function

Function FindRow(ByRef wb As Workbook, ByRef ws As Worksheet, rng As range, val As String, Optional xlLkAt As XlLookAt, Optional varArrayInputFiles As Variant, Optional direct As XlSearchDirection) As Long

    '***********************************************************************************
    'Purpose:Finds the row number of the searched string
    'Inputs: ws variable
    'Outputs:row number or error message box
    '***********************************************************************************
    
    On Error Resume Next
    
    Dim c As range, firstAddress As String
    
    If xlLkAt = 0 Then xlLkAt = xlWhole     'if parameter not specified, look for the whole text
    
        
    
    With ws
        With ws.range(rng.Address)
            Set c = .Find(val, LookIn:=xlValues, LookAt:=xlLkAt, MatchCase:=False, SearchDirection:=direct)
            If Not c Is Nothing Then
                FindRow = c.row
            Else
                MsgBox "Cannot find the Column of the value: " & val & " in this worksheet " & ws.Name & " in this workbook " & wb.Name & " Please amend the tab names then try again."
                'Optional varArrayInputFiles As Variant
                If varArrayInputFiles <> Empty Then
                    CloseWorkbooks varArrayInputFiles
                End If
                End
            End If
        End With
    End With
    Exit Function
errorline:
    MsgBox "Error while finding row of the value: " & val & " in this worksheet " & ws.Name & " in this workbook " & wb.Name
End Function

Function DoesRowExist(ByRef wb As Workbook, ByRef ws As Worksheet, rng As range, val As String, Optional xlLkAt As XlLookAt, Optional varArrayInputFiles As Variant, Optional direct As XlSearchDirection) As Boolean

    '***********************************************************************************
    'Purpose:Finds the row number of the searched string
    'Inputs: ws variable
    'Outputs:row number or error message box
    '***********************************************************************************
    
    On Error Resume Next
    
    Dim c As range, firstAddress As String
    
    If xlLkAt = 0 Then xlLkAt = xlWhole     'if parameter not specified, look for the whole text
    
        
    
    With ws
        With ws.range(rng.Address)
            Set c = .Find(val, LookIn:=xlValues, LookAt:=xlLkAt, MatchCase:=False, SearchDirection:=direct)
            If Not c Is Nothing Then
                DoesRowExist = True
            Else
                Err.Clear
                On Error GoTo 0
                DoesRowExist = False
                Exit Function
            End If
        End With
    End With
    Exit Function
errorline:
    MsgBox "Error while finding row of the value: " & val & " in this worksheet " & ws.Name & " in this workbook " & wb.Name
End Function

Function FindColumn(ByRef wb As Workbook, ByRef ws As Worksheet, rng As range, val As String, Optional xlLkAt As XlLookAt, Optional varArrayInputFiles As Variant, Optional direct As XlSearchDirection) As Long

    '***********************************************************************************
    'Purpose:Finds the Column number of the searched string
    'Inputs: ws variable
    'Outputs:Column number or error message box
    '***********************************************************************************
    
    On Error Resume Next
    
    Dim c As range, firstAddress As String
    
    If xlLkAt = 0 Then xlLkAt = xlWhole     'if parameter not specified, look for the whole text
    
'    If direct = 0 Then
'        direct = xlNext
'    Else
'        direct = xlPrevious
'    End If
    
    With ws
        With ws.range(rng.Address)
            
            Set c = .Find(val, LookIn:=xlValues, LookAt:=xlLkAt, MatchCase:=False, SearchDirection:=direct)
            If Not c Is Nothing Then
                FindColumn = c.Column
            Else
                MsgBox "Cannot find the Column of the value: " & val & " in this worksheet " & ws.Name & " in this workbook " & wb.Name & " Please amend the tab names then try again."
                'Optional varArrayInputFiles As Variant
                If varArrayInputFiles <> Empty Then
                    CloseWorkbooks varArrayInputFiles
                End If
                End
            End If
        End With
    End With
    Exit Function
errorline:
    MsgBox "Error while finding Column of the value: " & val & " in this worksheet " & ws.Name & " in this workbook " & wb.Name
End Function

Function FindColumn_continuous(ByRef wb As Workbook, ByRef ws As Worksheet, rng As range, val As String, Optional xlLkAt As XlLookAt, Optional varArrayInputFiles As Variant, Optional direct As XlSearchDirection) As Long

    '***********************************************************************************
    'Purpose:Finds the Column number of the searched string
    'Inputs: ws variable
    'Outputs:Column number or error message box
    '***********************************************************************************
    
    On Error Resume Next
    
    Dim c As range, firstAddress As String
    
    If xlLkAt = 0 Then xlLkAt = xlWhole     'if parameter not specified, look for the whole text
    
'    If direct = 0 Then
'        direct = xlNext
'    Else
'        direct = xlPrevious
'    End If
    
    With ws
        With ws.range(rng.Address)
            
            Set c = .Find(val, LookIn:=xlValues, LookAt:=xlLkAt, MatchCase:=False, SearchDirection:=direct)
            If Not c Is Nothing And val <> vbNullString Then
                FindColumn_continuous = c.Column
            Else
                FindColumn_continuous = 0
            End If
        End With
    End With
    Exit Function
errorline:
    MsgBox "Error while finding Column of the value: " & val & " in this worksheet " & ws.Name & " in this workbook " & wb.Name
End Function


Function GetRowNumberAfterNumericCellValue(ws As Worksheet, Lrow As Long, strDir As String, refCol As Long) As Long
    
    '***********************************************************************************
    'Purpose:Returns the next row of an instance of a string
    'Inputs: worksheet and starting row number
    'Outputs:row number of an instance of a string
    '***********************************************************************************
    
    Dim lastRow As Long, varArray As Variant, ctr As Long
    
    varArray = ws.UsedRange.Value
    
    If strDir = "Upwards" Then
        lastRow = 1
        
        For ctr = Lrow - 1 To lastRow Step -1
            If IsNumeric(varArray(ctr, refCol)) Then
                NextRowOfNumericString = ctr
                Exit For
            End If
        Next ctr
        
    Else
        lastRow = ws.UsedRange.Rows.Count
        
        For ctr = Lrow + 1 To lastRow
            ctr = ctr + 1
            If IsNumeric(varArray(ctr, refCol)) Then
                NextRowOfNumericString = ctr + 1
                Exit For
            End If
        Next ctr
        
    End If
    
    
    
End Function


Function CloseWorkbooks(varArray As Variant)
    
    On Error Resume Next
    
    Dim varElement As Variant
    
    For Each varElement In varArray
          With varElement
               
               Dim str As String
               str = varElement.Name
               
               Workbooks(str).Close False
               Set Worksheets(str) = Nothing
             
            End With
    Next varElement
    
    
    Err.Clear
    
    Application.Calculation = xlCalculationAutomatic
    
End Function


Function ReturnNameFromRightViaCodeFromLeft(wb As Workbook, ws As Worksheet, strCodes As Variant) As Variant
    Dim rowNum As Long, colNum As Long, wsArray As Variant
    Dim arrayRowCtr As Long, arrayColCtr As Long
    Dim arrayContainer As Variant
    Dim strCode As String
    Dim wsNm As String
    
    arrayContainer = strCodes
    
    'get the code from left most then return the value on its next right
    
    With ws
        For arrayRowCtr = LBound(strCodes, 1) To UBound(strCodes, 1)
            strCode = strCodes(arrayRowCtr)
            
            rowNum = ExlHandler.FindRow(wb, ws, .UsedRange, strCode)
            colNum = ExlHandler.FindColumn(wb, ws, .UsedRange, strCode)
            wsNm = .Cells(rowNum, .Cells(rowNum, colNum).End(xlToRight).Column).Value
            arrayContainer(arrayRowCtr) = wsNm
            
        Next arrayRowCtr
        
        
        
    End With
    GetNameViaCode = arrayContainer
End Function



Sub FilterandInsertWeightingsSymbol2(wb As Workbook, sh As Worksheet, startRow As Long, lastRow As Long, plotCol As Long)
    
    Dim ctr As Long, strNumfin As String, strNumHIS As String, strPlotStr
    
    wb.Activate
    sh.Activate
    
    Cells.Replace "#N/A", "-"
    
    For ctr = startRow To lastRow
        
        strPlotStr = Cells(ctr, plotCol).Value
        
        
        If strPlotStr <> vbNullString Then
            If strPlotStr = "1" Or strPlotStr = "-1" Or strPlotStr = "0" Then Cells(ctr, plotCol).NumberFormat = " [GREEN]" & ChrW(&H25B2) & ";[RED]" & ChrW(&H25BC) & ";[BLUE]" & ChrW(&H25CF)
            If strPlotStr = "2" Or strPlotStr = "-2" Then Cells(ctr, plotCol).NumberFormat = " [GREEN]" & ChrW(&H25B2) & ChrW(&H25B2) & ";[RED]" & ChrW(&H25BC) & ChrW(&H25BC) & ";[BLUE]" & ChrW(&H25CF)
            If strPlotStr = "3" Or strPlotStr = "-3" Then Cells(ctr, plotCol).NumberFormat = " [GREEN]" & ChrW(&H25B2) & ChrW(&H25B2) & ChrW(&H25B2) & ";[RED]" & ChrW(&H25BC) & ChrW(&H25BC) & ChrW(&H25BC) & ";[BLUE]" & ChrW(&H25CF)
            If strPlotStr = "-" Or strPlotStr = "#N/A" Then Cells(ctr, plotCol).Value = "n/a": Cells(ctr, plotCol).Font.Color = RGB(0, 0, 0)
            
            If strPlotStr = "n/a" Then
                Cells(ctr, plotCol).Font.Color = RGB(0, 0, 0)
                Cells(ctr, plotCol).NumberFormat = "General"
            End If
            
        End If
    Next ctr
    
        
End Sub

Public Function GetLastRow(ws As Worksheet, Optional longRow As Long = 1, Optional longCol As Long = 1) As Long

    GetLastRow = ws.Cells(Rows.Count, longCol).End(xlUp).row

End Function


Public Function GetLastCol(ws As Worksheet, Optional longRow As Long = 1) As Long

    GetLastCol = ws.Cells(longRow, Columns.Count).End(xlToLeft).Column

End Function

Public Function GetFirstFilteredRow(ws As Worksheet) As Long
    On Error Resume Next
    GetFirstFilteredRow = ws.AutoFilter.range.Offset(1, 0).SpecialCells(xlCellTypeVisible).row
    On Error GoTo 0
    Err.Clear
End Function


Public Function GetLastFilteredRow(ws As Worksheet, longRow As Long, Optional longCol As Long = 2) As Long
    
    On Error Resume Next
    GetLastFilteredRow = ws.Cells(longRow, longCol).End(xlDown).row
    On Error GoTo 0
    Err.Clear

End Function


Function SetConditionalFormatting(ws As Worksheet, rg As range)

    Dim cond1 As FormatCondition, cond2 As FormatCondition, cond3 As FormatCondition, cond4 As FormatCondition

    'clear any existing conditional formatting
    ws.range(ws.range(rg.Address).Address).FormatConditions.Delete
     
    'define the rule for each conditional format
    Set cond1 = ws.range(rg.Address).FormatConditions.Add(xlCellValue, xlGreater, 0)
    Set cond2 = ws.range(rg.Address).FormatConditions.Add(xlCellValue, xlLess, 0)
    Set cond3 = ws.range(rg.Address).FormatConditions.Add(xlCellValue, xlEqual, 0)
    Set cond4 = ws.range(rg.Address).FormatConditions.Add(xlCellValue, xlEqual, "#N/A")
     
    'define the format applied for each conditional format
    With cond1
    .Font.Color = RGB(0, 175, 80)
    End With
     
    With cond2
    .Font.Color = RGB(255, 0, 0)
    End With
     
    With cond3
    .Font.Color = RGB(0, 112, 192)
    End With
    
    With cond4
    .Font.Color = RGB(0, 0, 0)
    End With
 
End Function


Public Function CopyFromExcelToPowerPointTable(ws As Worksheet, sourceRange As range, slideNum As Integer, strTableName As String, longFRow As Long, strPPTtempfilename As String, Optional longFCol As Long = 1)
        
    'Set pptApp = CreateObject("PowerPoint.Application")
        
    On Error Resume Next
    pptApp.Presentations(strPPTtempfilename).Windows(1).Activate
    If Err.Number <> 0 Then
            ErrorHandler.ErrorHandler "Error while activating the powerpoint filename: " & strPPTtempfilename & ". Please check the powerpoint template if it's opened or not corrupted." & vbNewLine & _
           "Error in function: " & "ProcessSlide3"
    End If
    On Error GoTo 0
    
    On Error Resume Next
    pptApp.ActiveWindow.View.gotoslide (slideNum)
    If Err.Number <> 0 Then
            ErrorHandler.ErrorHandler "Error while activating the powerpoint slide number: " & slideNum & ". Please check the powerpoint template if it's opened or not corrupted." & vbNewLine & _
           "Error in function: " & "ProcessSlide3"
    End If
    On Error GoTo 0
    
    Set pptSlide = pptApp.ActiveWindow.View.slide
            
    'Copy from source and paste to powerpoint table
    ws.range(sourceRange.Address).Copy
    
    On Error Resume Next
    pptSlide.Shapes(strTableName).Table.cell(longFRow, longFCol).Select
    If Err.Number <> 0 Then
         ErrorHandler.ErrorHandler "Error while selecting the table from the powerpoint template. Table Name: " & strTableName & ". Please check the powerpoint template if the table name is present in slide " & slideNum & _
         " of the powerpoint template." & vbNewLine & _
        "Error in function: " & "ProcessSlide3"
    End If
    On Error GoTo 0
    
    On Error Resume Next
    pptApp.ActiveWindow.View.Paste
    If Err.Number <> 0 Then
            ErrorHandler.ErrorHandler "Error while pasting values on table: " & strTableName & ". Please check the powerpoint template if it's opened or not corrupted." & vbNewLine & _
           "Error in function: " & "ProcessSlide3"
    End If
    On Error GoTo 0
    
    Application.CutCopyMode = False

End Function





Function FindFirstDate(wb As Workbook, ws As Worksheet, keyword As String, Optional strFormat As String = "MMM YYYY", Optional chooseCol1 As Long = 1) As String
    
    Dim startRow As Long, startCol As Long
    Dim searchRng As range, searchElement As range
    
    With ws
        wb.Activate
        .Activate
        .Cells(1, 1).Select
        startRow = ExlHandler.FindRow(wb, ws, .UsedRange, keyword, xlPart, direct:=xlNext)
        startCol = ExlHandler.FindColumn(wb, ws, .UsedRange, keyword, xlPart, direct:=xlNext)
        
        If startRow - 1 = 0 Then
            startRow = 1
        Else
            startRow = startRow - 1
        End If
        
        If chooseCol1 = 1 Then startCol = 1
            
        
        Set searchRng = .range(.Cells(startRow, startCol), .Cells(.UsedRange.Rows.Count, .UsedRange.Columns.Count))
        
        For Each searchElement In searchRng
            If IsDate(searchElement) = True Then
                
                FindFirstDate = Format(searchElement.Value, strFormat)
                Exit Function
                
            End If
        Next searchElement
        
        
    End With
    
    
    ErrorHandler.ErrorHandler "Error while finding date in this workbook " & wb.Name & " in this worksheet " & ws.Name & ". Program will abend." & vbNewLine & _
        "Error in function: " & "FindFirstDate"
    
    
End Function



Sub DeleteChartsExcept(ws As Worksheet, strChartNm As String, Optional strChrtNm2 As String, Optional strChrtNm3 As String, Optional strChrtNm4 As String)
    
    Dim ASheet As Worksheet
    Dim AChart As ChartObject
    Dim concStr As String
    
      concStr = "-" & strChartNm & "-" & strChrtNm2 & "-" & strChrtNm3 & "-" & strChrtNm4

      Application.DisplayAlerts = False
      Application.ScreenUpdating = False
              
      '** Now scan and delete any ChartSheets ****
      For Each AChart In ws.ChartObjects
        
        If InStr(1, concStr, "-" & AChart.Name & "-", vbTextCompare) = 0 Then
            AChart.Delete
        End If
      Next
    
      Application.DisplayAlerts = True
      Application.ScreenUpdating = True

End Sub


Sub TransferRangeData(srcWs As Worksheet, destWs As Worksheet, srcKeywordStart As String, srcKeywordLast As String, destKeywordStart As String, destKeywordLast As String, destWb As Workbook, srcWb As Workbook, _
    destKeywordsArray As Variant, srcKeywordsArray As Variant, Optional formulaString As String, Optional formulaColStr1 As String, Optional formulaColStr2 As String)
    
    Dim srcRngRowCount As Long, destRngRowCount As Long, rowDifference As Long
    Dim srcStartRow As Long, srcLastRow As Long, destStartRow As Long, destLastRow As Long
    
    Dim destKeywordLong1 As Long, srcKeywordLong1 As Long, destKeywordLong2 As Long, srcKeywordLong2 As Long, srcKeywordLong3 As Long, destKeywordLong3 As Long, srcKeywordLong4 As Long, _
      destKeywordLong4 As Long, srcKeywordLong5 As Long, destKeywordLong5 As Long, srcKeywordLong6 As Long, destKeywordLong6 As Long, srcKeywordLong7 As Long, destKeywordLong7 As Long
    
    Dim destKeywordLong As Long, srcKeywordLong As Long
    
    With srcWs
        
        srcStartRow = ExlHandler.FindRow(srcWb, srcWs, .UsedRange, srcKeywordStart, xlWhole) + 1
        srcLastRow = ExlHandler.FindRow(srcWb, srcWs, .UsedRange, srcKeywordLast, xlWhole)
        srcRngRowCount = srcLastRow - srcStartRow + 1
               
    End With
    
    With destWs
        destStartRow = ExlHandler.FindRow(destWb, destWs, .UsedRange, destKeywordStart, xlWhole) + 1
        destLastRow = ExlHandler.FindRow(destWb, destWs, .UsedRange, destKeywordLast, xlWhole) - 1
        destRngRowCount = destLastRow - destStartRow + 1
        
    End With
    
    rowDifference = srcRngRowCount - destRngRowCount
    
    If rowDifference < 0 Then
        'rowDifference = Abs(rowDifference)
        
        'destination has more rows than the source range; delete rows in destination
        
        With destWs
            Application.DisplayAlerts = False
            .Rows(destLastRow & ":" & destLastRow + rowDifference + 1).Delete
            destLastRow = ExlHandler.FindRow(destWb, destWs, .UsedRange, destKeywordLast, xlWhole) - 1
            Application.DisplayAlerts = True
        End With
    ElseIf rowDifference > 0 Then
        'delete unnecessary rows
        
        'source has more rows than the destination; insert rows in destination
        With destWs
            .Rows(destLastRow + 1 & ":" & destLastRow + rowDifference).Insert (xlUp)
            destLastRow = ExlHandler.FindRow(destWb, destWs, .UsedRange, destKeywordLast, xlWhole) - 1
        End With
        
    End If
        
    Dim srcElementStr As String, destElementStr As String, srcArrayCtr As Long
    Dim formulaCol1 As Long, formulaCol2 As Long
    
    'plot columns
    For srcArrayCtr = LBound(srcKeywordsArray, 1) To UBound(srcKeywordsArray, 1)
        srcElementStr = srcKeywordsArray(srcArrayCtr)
        destElementStr = destKeywordsArray(srcArrayCtr)
        
        srcKeywordLong = ExlHandler.FindColumn_continuous(srcWb, srcWs, srcWs.UsedRange, srcElementStr, xlWhole, direct:=xlPrevious)
        destKeywordLong = ExlHandler.FindColumn_continuous(destWb, destWs, destWs.UsedRange, destElementStr, xlWhole, direct:=xlPrevious)
        
        If srcKeywordLong <> 0 And destKeywordLong <> 0 Then
            With srcWs
                .range(.Cells(srcStartRow, srcKeywordLong), .Cells(srcLastRow, srcKeywordLong)).Copy
            End With
            
            With destWs
                .range(.Cells(destStartRow, destKeywordLong), .Cells(destLastRow, destKeywordLong)).PasteSpecial xlPasteValues
            End With
            
            Application.CutCopyMode = False
                        
        End If
        
    Next srcArrayCtr
    
    formulaCol1 = ExlHandler.FindColumn_continuous(destWb, destWs, destWs.UsedRange, formulaColStr1, xlWhole)
    formulaCol2 = ExlHandler.FindColumn_continuous(destWb, destWs, destWs.UsedRange, formulaColStr2, xlWhole)

    'col1 are the values to process
    'col2 the destination column

    With destWs
        Select Case formulaString
            Case "AVERAGE NON ZERO"
                Dim dblAvg As Double
                
                dblAvg = Application.WorksheetFunction.SumIf(.range(.Cells(destStartRow, formulaCol1), .Cells(destLastRow, formulaCol1)), "<>0") / _
                    Application.WorksheetFunction.CountIf(.range(.Cells(destStartRow, formulaCol1), .Cells(destLastRow, formulaCol1)), "<>0")
                
                .range(.Cells(destStartRow, formulaCol2), .Cells(destLastRow, formulaCol2)).Value = dblAvg
                
                
            Case "SUM"
                Dim dblSum As Double
                            
                dblSum = Application.WorksheetFunction.Sum(.range(.Cells(destStartRow, formulaCol1), .Cells(destLastRow, formulaCol1)))
                .range(.Cells(destStartRow, formulaCol2), .Cells(destLastRow, formulaCol2)).Value = dblSum
                            
            Case "AVERAGE"
                Dim dblAv As Double
                            
                dblAv = Application.WorksheetFunction.Average(.range(.Cells(destStartRow, formulaCol1), .Cells(destLastRow, formulaCol1)))
                .range(.Cells(destStartRow, formulaCol2), .Cells(destLastRow, formulaCol2)).Value = dblAv
                            
        End Select
    End With
            
    
   
    
End Sub


Sub RepositionTabs(tabsArray As Variant, destWb As Workbook)
    
    Dim arrayCtr As Long, tabNm As String, shCtr As Long, lastSheetNum As Long
    
    destWb.Activate
    
    For arrayCtr = LBound(tabsArray) To UBound(tabsArray)
        tabNm = tabsArray(arrayCtr)
        
        With destWb
            .Sheets(tabNm).Move before:=.Sheets(arrayCtr + 1)
        End With
        
    Next arrayCtr
    
    lastSheetNum = destWb.Sheets(arrayCtr + 1).Index
    
    If lastSheetNum < destWb.Sheets.Count Then
        
        For shCtr = UBound(tabsArray) + 1 To destWb.Sheets.Count
                
                If InStr(1, destWb.Sheets(shCtr).Name, "Data", vbTextCompare) > 0 Or _
                    InStr(1, destWb.Sheets(shCtr).Name, " by ", vbTextCompare) > 0 Then
                    
                    destWb.Sheets(shCtr).Visible = xlSheetHidden
                    
                End If
        Next shCtr
        
    End If
    
    destWb.Sheets(1).Activate
    
End Sub



Function PlotErrorDataInSheet(srcWb As Workbook, destWb As Workbook, ws As Worksheet, colWithError As Long, strError As String, arrayToGet As Variant, _
    shDest As Worksheet, arrayColsToPasteInto As Variant) 'parameters are ref sheet, error string, column array to get, dest sheet, column array to paste into
        
    Dim firstFilteredRow As Long, lastFilteredRow As Long, longCount As Long, srcCol As Long, destCol As Long, srcColStr As String, destColStr As String
    Dim arrayCtr As Long, destPlotRow As Long
    
    On Error GoTo errorline
    
    With ws
        .UsedRange.AutoFilter colWithError, strError, xlFilterValues
        
             firstFilteredRow = ExlHandler.GetFirstFilteredRow(ws)
            
            If firstFilteredRow > 0 Then
                .Activate
                .Cells(firstFilteredRow, colWithError).Select  'investment manager column
                lastFilteredRow = Selection.End(xlDown).row
    
                If lastFilteredRow > .Cells.Rows.Count - 5 Then lastFilteredRow = firstFilteredRow
    
    
                longCount = WorksheetFunction.Subtotal(3, .range(.Cells(firstFilteredRow, colWithError), .Cells(lastFilteredRow, colWithError)))
                
                If longCount >= 1 Then
                    
                    
                    shDest.Activate
                    
                    For arrayCtr = LBound(arrayToGet) To UBound(arrayToGet)
                        
                        
                        srcColStr = arrayToGet(arrayCtr)
                        destColStr = arrayColsToPasteInto(arrayCtr)
                        
                        srcCol = ExlHandler.FindColumn(srcWb, ws, .UsedRange, srcColStr)
                        destCol = ExlHandler.FindColumn(destWb, shDest, shDest.UsedRange, destColStr)
                            
                        With shDest
                            
                            If destPlotRow = 0 Then
                                destPlotRow = .Cells(.Cells.Rows.Count, destCol).End(xlUp).row + 1
                            End If
                            
                            With ws
                                .range(.Cells(firstFilteredRow, srcCol), .Cells(lastFilteredRow, srcCol)).Copy
                            End With
                            
                            .Cells(destPlotRow, destCol).Select
                            ActiveSheet.Paste
                              
                        End With
                        
                        
                    Next arrayCtr
                    destPlotRow = 0
                End If
            End If
        
    End With
    
    Exit Function
errorline:
    ErrorHandler.ErrorHandler "Error while plotting new positions in the DB."
End Function



Function isThereAnInstanceOfKeyword(ws As Worksheet, rng As range, str As String) As Boolean
    
    Dim countt As Long
    
    With ws
        On Error Resume Next
            countt = .Application.WorksheetFunction.CountIf(rng, str)
        On Error GoTo 0
        Err.Clear
        
        isThereAnInstanceOfKeyword = (countt > 0)
        
    End With
    
End Function

Function SortData(wb As Workbook, ws As Worksheet, rngToSort As range, rngKey As range, sortorder As Long, Optional rngKey2 As range, Optional sortorder2 As Long)
    
    On Error GoTo errorline
    
    'sort data only
    
    wb.Activate
    
    With ws
        .Activate
        .Sort.SortFields.Clear
        
        .Sort.SortFields.Add _
            Key:=rngKey, SortOn:=xlSortOnValues, order:=sortorder, _
            DataOption:=xlSortNormal
        
        If Not rngKey2 Is Nothing Then
             .Sort.SortFields.Add _
            Key:=rngKey2, SortOn:=xlSortOnValues, order:=sortorder2, _
            DataOption:=xlSortNormal
        End If
        
        
        With .Sort
            .SetRange rngToSort
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
    End With
    
    Exit Function
errorline:
    ErrorHandler.ErrorHandler "Error while sorting data of the worksheet " & ws.Name & " of its workbook " & wb.Name
    
End Function


Function RenameInWorksheet(ws As Worksheet, srcArray As Variant, destArray As Variant, rng As range)
    
    Dim ctr As Long, strToReplace As String, strReplacementString As String
    
    With ws
        For ctr = LBound(srcArray, 1) To UBound(srcArray, 1)
            strToReplace = destArray(ctr)
            strReplacementString = srcArray(ctr)
            .range(rng.Address).Replace strToReplace, strReplacementString
        Next ctr
    End With
    
End Function


Function FilterThenCopySheet(srcWb As Workbook, srcSh As Worksheet, varCriteria As Variant, colCriteria As Long, _
    Optional strDestWsName As String, Optional varCriteria2 As Variant, Optional colCriteria2 As Long, Optional lastRowRefCol As Long = 1) As Worksheet
    
    On Error GoTo errorline
    
    'parameters src sheet, columns in array, dest worksheet name
    
    Dim strColName As String, colArrayCtr As Long, srcColNum As Long
    
    Dim longDestColCtr As Long
    
    Dim firstFilteredRow As Long, lastFilteredRow As Long, longCount As Long
    
    Dim draftSh As Worksheet, strCriteria As String, strCriteria2 As String
    
    With srcSh
        
         .AutoFilterMode = False
            
            Select Case True
                
                Case IsMissing(varCriteria) = False And IsMissing(varCriteria2) = False
                    If IsArray(varCriteria) Then
                        
                        If UBound(varCriteria, 1) > 0 Then
                            .UsedRange.AutoFilter colCriteria, Criteria1:=varCriteria, Operator:=xlFilterValues
                        Else
                            strCriteria = CStr(varCriteria(0))
                            
                            If strCriteria = "Just Copy" Then GoTo CopyLine
                            
                            .UsedRange.AutoFilter colCriteria, Criteria1:=strCriteria, Operator:=xlFilterValues
                        End If
                        
                    Else
                        .UsedRange.AutoFilter colCriteria, Criteria1:=CStr(varCriteria), Operator:=xlFilterValues
                    End If
                    
                    If IsArray(varCriteria2) Then
                        
                        If UBound(varCriteria2, 1) > 0 Then
                            .UsedRange.AutoFilter colCriteria2, Criteria1:=varCriteria2, Operator:=xlFilterValues
                        Else
                            strCriteria2 = CStr(varCriteria2(0))
                            .UsedRange.AutoFilter colCriteria2, Criteria1:=strCriteria2, Operator:=xlFilterValues
                        End If
                    Else
                        .UsedRange.AutoFilter colCriteria2, Criteria1:=CStr(varCriteria2), Operator:=xlFilterValues
                    End If
                                    
                Case IsMissing(varCriteria) = False And IsMissing(varCriteria2)
                    
                    If IsArray(varCriteria) Then
                        
                        If UBound(varCriteria, 1) > 0 Then
                            .UsedRange.AutoFilter colCriteria, Criteria1:=varCriteria, Operator:=xlFilterValues
                        Else
                            strCriteria = CStr(varCriteria(0))
                            .UsedRange.AutoFilter colCriteria, Criteria1:=strCriteria, Operator:=xlFilterValues
                        End If
                        
                    Else
                        
                        If varCriteria = "Just Copy" Then strCriteria = varCriteria: GoTo CopyLine
                    
                        .UsedRange.AutoFilter colCriteria, Criteria1:=CStr(varCriteria), Operator:=xlFilterValues
                    End If
                    
            End Select
                               
            firstFilteredRow = ExlHandler.GetFirstFilteredRow(srcSh)
            
            If firstFilteredRow > 0 Then
                .Activate
                .Cells(firstFilteredRow, lastRowRefCol).Select  'investment manager column
                lastFilteredRow = Selection.End(xlDown).row
    
                If lastFilteredRow > .Cells.Rows.Count - 5 Then lastFilteredRow = firstFilteredRow
    
    
                longCount = WorksheetFunction.Subtotal(3, .range(.Cells(firstFilteredRow, lastRowRefCol), .Cells(lastFilteredRow, lastRowRefCol)))
                
                If longCount > 1 Then

CopyLine:
                    
                    If strCriteria = "Just Copy" Then
                        srcSh.AutoFilterMode = False
                        firstFilteredRow = 1
                        lastFilteredRow = .UsedRange.Rows.Count
                    End If
                  
                    Application.DisplayAlerts = False
                    
                    srcWb.Sheets.Add
                    
                  Set draftSh = ActiveSheet
                    .range(.Cells(1, 1), .Cells(1, .UsedRange.Columns.Count)).Copy
                    draftSh.Cells(1, 1).Select
                    ActiveSheet.Paste
                                        
                    .range(.Cells(firstFilteredRow, 1), .Cells(lastFilteredRow, .UsedRange.Columns.Count)).Copy
                    draftSh.Cells(2, 1).Select
                    ActiveSheet.Paste
                                                            
                    Application.CutCopyMode = False
                    Application.DisplayAlerts = True
                    
                    If strDestWsName <> vbNullString Then draftSh.Name = strDestWsName
                    
                    Set FilterThenCopySheet = draftSh
                    
                End If
                                            
        End If  'end of condition checking if there are filtered cells
        .AutoFilterMode = False
    End With
    
    Exit Function
errorline:
    ErrorHandler.ErrorHandler "Error while filtering then copying sheet. " & vbNewLine & vbNewLine & _
        "Source worksheet: " & srcSh.Name & vbNewLine & _
        "Destination sheet: " & draftSh.Name
        
    
End Function

Function InsertSpecialTotalRow(destWb As Workbook, destWs As Worksheet, srcWb As Workbook, srcWs As Worksheet, _
    Optional blnIsPriv As Boolean, Optional blnIsUnallo As Boolean)
    Dim destLastRow As Long, srcLastRow As Long, destPlotRow As Long, destLastRankRow As Long
    
    On Error GoTo errorline
    
    srcLastRow = ExlHandler.FindRow(srcWb, srcWs, srcWs.UsedRange, "TOTAL", xlWhole)
    
    With destWs
        destLastRow = ExlHandler.FindRow(destWb, destWs, .UsedRange, "TOTAL", xlWhole)
        destPlotRow = destLastRow
                
        .Rows(destLastRow).Insert (xlDown)
        destLastRow = destLastRow + 1
        destLastRankRow = .Cells(destPlotRow, 1).End(xlUp).row
        
        .range(.Cells(destPlotRow, 1), .Cells(destPlotRow, .UsedRange.Columns.Count)).Value = _
            srcWs.range(srcWs.Cells(srcLastRow, 1), srcWs.Cells(srcLastRow, .UsedRange.Columns.Count)).Value
        
        .range(.Cells(destLastRankRow, 1), .Cells(destLastRankRow, .UsedRange.Columns.Count)).Copy
        .range(.Cells(destPlotRow, 1), .Cells(destPlotRow, .UsedRange.Columns.Count)).PasteSpecial xlPasteFormats
        
        If blnIsPriv = True Then
            .Cells(destPlotRow, 2).Value = "Private Investors"
            '.range(.Cells(destPlotRow, 3), .Cells(destPlotRow, 5)).ClearContents
        End If
        
        If blnIsUnallo = True Then
            .Cells(destPlotRow, 2).Value = "Unallocated"
            '.range(.Cells(destPlotRow, 3), .Cells(destPlotRow, 5)).ClearContents
        End If
        
        'Update Total Row
        Dim totalRowColCtr As Long
        'destLastRow = destLastRow + 1
        For totalRowColCtr = 1 To .UsedRange.Columns.Count
            If .Cells(destLastRow, totalRowColCtr).Value <> vbNullString And IsNumeric(.Cells(destLastRow, totalRowColCtr).Value) Then
                .Cells(destLastRow, totalRowColCtr).Value = .Cells(destLastRow, totalRowColCtr).Value + .Cells(destPlotRow, totalRowColCtr).Value
            End If
        Next totalRowColCtr
        
    End With
    
    Exit Function
errorline:
    ErrorHandler.ErrorHandler "Error while inserting special total rows"
    
End Function


Function FilterRangeThenDelete(ws As Worksheet, colToFilter As Long, lastRowRefCol As Long, criteriaStr As String, Optional colToFilter2 As Long, Optional criteriaStr2 As String, Optional excludeArray As Variant)

     Dim firstFilteredRow As Long, lastFilteredRow As Long


       With ws
            .AutoFilterMode = False
            
            If criteriaStr <> vbNullString Then
            
            'If IsEmpty(excludeArray) Or UBound(excludeArray) < 1 Then
                .UsedRange.AutoFilter colToFilter, Criteria1:=criteriaStr, Operator:=xlFilterValues
            Else
                
                If VarType(excludeArray) = 8204 Then 'check if string
               
                    .UsedRange.AutoFilter colToFilter, Criteria1:=Application.WorksheetFunction.Transpose(excludeArray), Operator:=xlFilterValues
                Else
                    criteriaStr = excludeArray
                    .UsedRange.AutoFilter colToFilter, Criteria1:=criteriaStr, Operator:=xlFilterValues
                End If
            End If
            
            
            firstFilteredRow = .AutoFilter.range.Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).row
            .Activate
            .Cells(firstFilteredRow, lastRowRefCol).Select  'investment manager column
            lastFilteredRow = Selection.End(xlDown).row

            If lastFilteredRow > .Cells.Rows.Count - 5 Then lastFilteredRow = firstFilteredRow

             Dim longCount As Long

            longCount = WorksheetFunction.Subtotal(3, .range(.Cells(firstFilteredRow, lastRowRefCol), .Cells(lastFilteredRow, lastRowRefCol)))

            If longCount > 1 Then
                Application.DisplayAlerts = False
                .range(.Cells(firstFilteredRow, 1), .Cells(lastFilteredRow, .UsedRange.Columns.Count)).Delete
                Application.DisplayAlerts = True
            End If
            
            
            
             If colToFilter2 <> 0 Then
                FilterRangeThenDelete ws, colToFilter2, lastRowRefCol, criteriaStr2
            End If
            
 End With
            ws.AutoFilterMode = False
            
           
            
End Function

Function FormatTotalRow(ws As Worksheet, lastRow As Long, lastCol As Long, startRow As Long, Optional startCol As Long = 5)
    
    Dim ctr As Long
    
    
    
    With ws
        .Cells(lastRow, 2).Value = "TOTAL"
        ExlHandler.FormatRange ws, .range(.Cells(lastRow, 1), .Cells(lastRow, lastCol)), RGB(0, 0, 0), 7, "Arial", _
        border1:=xlEdgeTop, border1style:=xlContinuous, border1weight:=xlThin, border1color:=RGB(0, 0, 0), _
        border2:=xlEdgeBottom, border2style:=xlContinuous, border2weight:=xlThin, border2color:=RGB(0, 0, 0), blnIsBold:=True
        
        .range(.Cells(startRow, 1), .Cells(lastRow - 1, lastCol)).AutoFilter 1, "<>", xlFilterValues
        
        Dim firstFilteredRow As Long, lastFilteredRow As Long
        
    
            'firstFilteredRow = .AutoFilter.range.Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).row
            firstFilteredRow = startRow
            .Activate
            .Cells(firstFilteredRow, 5).Select  'investment manager column
            lastFilteredRow = Selection.End(xlDown).row
    
            If lastFilteredRow > .Cells.Rows.Count - 5 Then lastFilteredRow = firstFilteredRow
    
             Dim longCount As Long
    
            longCount = WorksheetFunction.Subtotal(3, .range(.Cells(firstFilteredRow, 2), .Cells(lastFilteredRow, 2)))
    
            If longCount >= 1 Then
                For ctr = startCol To lastCol - 1
                    
                    On Error Resume Next
                    
                    If IsNumeric(.Cells(firstFilteredRow, ctr).Value) And .Cells(firstFilteredRow, ctr).Value <> vbNullString Then
                        .Cells(lastRow, ctr).Value = .Application.WorksheetFunction.Subtotal(9, _
                        .range(.Cells(firstFilteredRow, ctr), .Cells(lastFilteredRow, ctr)))
                    End If
                    
                     If Err.Number <> 0 Then
                        ErrorHandler.ErrorHandler "Error while formatting and computing the totals in the consolidated input files. Please check the current input files if there are errors or #N/As in cells."
                    End If
                    
                    On Error GoTo 0
                    Err.Clear
                    
                Next ctr
            End If
        .AutoFilterMode = False
    End With
End Function




