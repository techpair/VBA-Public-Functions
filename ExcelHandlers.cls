VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelHandlers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Function ChangeChartColor(wb As Workbook, ws As Worksheet, chartName As String, Optional paramColors As Variant = Empty)
'
' ChangeChartColor Macro
'

'
    On Error Resume Next
    
    wb.Activate
    ws.Activate
    
    Dim defaultArrayColors As Variant, finalColors As Variant
    
    defaultArrayColors = Array(RGB(0, 84, 159), RGB(0, 152, 195), RGB(183, 2, 110), RGB(247, 142, 85), RGB(2, 195, 98), _
        RGB(102, 102, 102), RGB(147, 169, 207), RGB(255, 79, 79))
    
    If IsEmpty(paramColors) Then
        finalColors = defaultArrayColors
    Else
        finalColors = paramColors
    End If
    
    ws.ChartObjects(chartName).Activate
    
    Dim seriesCtr As Long
    
    For seriesCtr = 1 To ActiveChart.SeriesCollection(1).Points.Count
         
         ActiveChart.SeriesCollection(1).Points(seriesCtr).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = finalColors(seriesCtr - 1)
            .Transparency = 0
            .Solid
        End With
        
    Next seriesCtr
    
End Function


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
                
            
                .Name = strFontName '"Calibri"
                .Size = fontSizeOfRange '10
                .Bold = blnIsBold
                
           
    
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
                                lastRangeRow = .Cells(verticalCtr, columnNum).End(xlDown).row
                            End If
                            
                            .Cells(verticalCtr, columnNum).Value = Application.WorksheetFunction.Sum( _
                                .range(.Cells(verticalCtr + 1, columnNum), .Cells(lastRangeRow, columnNum)))
                            
                    End Select
                End If
                
            Next arrCtr
            
            verticalCtr = .Cells(verticalCtr, refCol).End(xlDown).row - 1  ' minus 1 to handle the "Next" command below or else it will skip the ranks row
            
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


Function ConsolidateSameColumns(arrayWb As Variant, arrayWs As Variant, ByRef arrayColumns As Variant, configSheet As Worksheet, Optional plotInputTagging As Boolean = False, _
    Optional plotGroupTagging As Boolean = False) As Long
    
    Dim lastRow As Long, startRow As Long, refCol As Long, destStartRow As Long, destLastRow As Long, destRefCol As Long
    Dim wbCtr As Long, keywordCtr As Long, strKeyword As String
    Dim wb As Workbook, ws As Worksheet, destWB As Workbook, destWs As Worksheet
    Dim destLastCol As Long
    Dim inputCount As Long
    
    Dim strWorksheetName As String
    Dim configFilenameRow  As Long
    Dim inputTaggingCol As Long
    Dim groupTaggingCol As Long
    Dim groupRefDtCol As Long
    Dim strGroupName As String
    Dim mktValCol As Long, dblSharePrice As Double
    
    'This function assumes you already know the worksheet names to combine.
    
    Workbooks.Add
    Set destWB = ActiveWorkbook
    Set destWs = destWB.Sheets(1)
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
                lastRow = .Cells(.Cells.Rows.Count, refCol).End(xlUp).row
            End With
            
            With destWs
                destStartRow = FindRow(destWB, destWs, .UsedRange, CStr(arrayColumns(LBound(arrayColumns, 1))), xlWhole)
                destRefCol = FindColumn(destWB, destWs, .UsedRange, CStr(arrayColumns(LBound(arrayColumns, 1))), xlWhole)
                destLastRow = .Cells(.Cells.Rows.Count, destRefCol).End(xlUp).row
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
                
                destWB.Activate
                
                With destWs
                    .Activate
                    destRefCol = FindColumn_continuous(destWB, destWs, .UsedRange, strKeyword, xlWhole)
                    
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
                                
                                
                                
                
                If plotGroupTagging = True And .Cells(.Cells.Rows.Count, destRefCol).End(xlUp).row <> .Cells.Rows.Count Then
                    
                    
                    On Error Resume Next
                    configFilenameRow = configSheet.Application.WorksheetFunction.Match(CStr(arrayWs(wbCtr)), configSheet.Columns("A:A"), 0)
                    On Error GoTo 0
                    Err.Clear
                    
                    'this will list all the previous shares of the uploaded input files if they are available in the config tab
                    
                    If configFilenameRow > 0 Then
                        strGroupName = configSheet.Cells(configFilenameRow, 3).Value
                    Else
                        ErrorHandler.ErrorHandler "Macro cannot find the grouping of the filename: " & CStr(arrayWs(wbCtr)) & vbNewLine & vbNewLine & _
                        "Please re-configure the input files' setup in the macro file tab 'Inputs Config with groupings'."
                    End If
                    
                    'plot group names
                    .range(.Cells(destLastRow + 1, groupTaggingCol), .Cells(.Cells(.Cells.Rows.Count, destRefCol).End(xlUp).row, groupTaggingCol)).Value = strGroupName
                    .range(.Cells(destLastRow + 1, groupTaggingCol + 1), .Cells(.Cells(.Cells.Rows.Count, destRefCol).End(xlUp).row, groupTaggingCol + 1)).Value = configSheet.Cells(configFilenameRow, 4).Value
                    
                    'compute market value
                    dblSharePrice = configSheet.Cells(configFilenameRow, 5).Value
                    
                    .Cells(destLastRow + 1, mktValCol).Formula = "=" & .Cells(destLastRow + 1, 4).Address(False, False) & "*" & dblSharePrice
                    
                    .Cells(destLastRow + 1, mktValCol).AutoFill .range(.Cells(destLastRow + 1, mktValCol), .Cells(.Cells(.Cells.Rows.Count, destRefCol).End(xlUp).row, mktValCol))
                    
                    .range(.Cells(destLastRow + 1, mktValCol), .Cells(.Cells(.Cells.Rows.Count, destRefCol).End(xlUp).row, mktValCol)).Value = _
                        .range(.Cells(destLastRow + 1, mktValCol), .Cells(.Cells(.Cells.Rows.Count, destRefCol).End(xlUp).row, mktValCol)).Value
                End If
                
                inputCount = inputCount + 1
                destLastRow = .Cells(.Cells.Rows.Count, destRefCol).End(xlUp).row
            End With
            
        Set wb = Nothing
        Set ws = Nothing
       
    Next wbCtr
        Set destWB = Nothing
        Set destWs = Nothing
        
        ConsolidateSameColumns = inputCount
        
End Function

Sub CopyWorksheet(destWB As Workbook, arraySrcFilenames As Variant, arrayWsNames As Variant, strAfterThisSheet As String)
    
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
        
            wbSrc.Sheets(strWsName).Copy After:=destWB.Sheets(strAfterThisSheet)
        Else
            wbSrc.Sheets(strSrcFilename).Copy After:=destWB.Sheets(strAfterThisSheet)
        End If
        
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
    .Replace What:="", replacement:="z", MatchCase:=False, LookAt:=xlPart
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


Sub DeleteSheet(sh As String, wb As Workbook)
'
' If the sheet exists
' DELETE IT! XD
'
    'PrintToLog ("Calling DeleteSheet")
    If DoesSheetExist(sh, wb) Then
        Application.DisplayAlerts = False
        wb.Sheets(sh).Delete
        Application.DisplayAlerts = True
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

    GetFirstFilteredRow = ws.AutoFilter.range.Offset(1, 0).SpecialCells(xlCellTypeVisible).row

End Function


Public Function GetLastFilteredRow(ws As Worksheet, longRow As Long, Optional longCol As Long = 2) As Long
    
    GetLastFilteredRow = ws.Cells(longRow, longCol).End(xlDown).row
    

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

Function deleteBlankColumnsInRange(wb As Workbook, ws As Worksheet, startRow As Long, lastRow As Long)
    Dim xEndCol As Long
    Dim I As Long
    Dim xDel As Boolean
    
    On Error Resume Next
    
    wb.Activate
    ws.Activate
    
 With ws
    xEndCol = .Range(.Cells(startRow, 1), .Cells(lastRow, .UsedRange.Columns.Count)).Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    If xEndCol = 0 Then
        'MsgBox "There is no data on """ & ws.Name & """ .", vbExclamation
        'Exit Function
    End If
    
    For I = xEndCol To 1 Step -1
        If Application.WorksheetFunction.CountA(.Range(.Cells(startRow, I), .Cells(lastRow, I))) <= 1 Then
            Columns(I).Delete
            xDel = True
        End If
    Next
    
'    If xDel Then
'        MsgBox "All blank and column(s) with only a header row have now been deleted.", vbInformation, "Kutools for Excel"
'    Else
'        MsgBox "There are no Columns to delete as each one has more data (rows) than just a header.", vbExclamation, "Kutools for Excel"
'    End If
    
End With
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


Sub TransferRangeData(srcWs As Worksheet, destWs As Worksheet, srcKeywordStart As String, srcKeywordLast As String, destKeywordStart As String, destKeywordLast As String, destWB As Workbook, srcWb As Workbook, _
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
        destStartRow = ExlHandler.FindRow(destWB, destWs, .UsedRange, destKeywordStart, xlWhole) + 1
        destLastRow = ExlHandler.FindRow(destWB, destWs, .UsedRange, destKeywordLast, xlWhole) - 1
        destRngRowCount = destLastRow - destStartRow + 1
        
    End With
    
    rowDifference = srcRngRowCount - destRngRowCount
    
    If rowDifference < 0 Then
        'rowDifference = Abs(rowDifference)
        
        'destination has more rows than the source range; delete rows in destination
        
        With destWs
            Application.DisplayAlerts = False
            .Rows(destLastRow & ":" & destLastRow + rowDifference + 1).Delete
            destLastRow = ExlHandler.FindRow(destWB, destWs, .UsedRange, destKeywordLast, xlWhole) - 1
            Application.DisplayAlerts = True
        End With
    ElseIf rowDifference > 0 Then
        'delete unnecessary rows
        
        'source has more rows than the destination; insert rows in destination
        With destWs
            .Rows(destLastRow + 1 & ":" & destLastRow + rowDifference).Insert (xlUp)
            destLastRow = ExlHandler.FindRow(destWB, destWs, .UsedRange, destKeywordLast, xlWhole) - 1
        End With
        
    End If
        
    Dim srcElementStr As String, destElementStr As String, srcArrayCtr As Long
    Dim formulaCol1 As Long, formulaCol2 As Long
    
    'plot columns
    For srcArrayCtr = LBound(srcKeywordsArray, 1) To UBound(srcKeywordsArray, 1)
        srcElementStr = srcKeywordsArray(srcArrayCtr)
        destElementStr = destKeywordsArray(srcArrayCtr)
        
        srcKeywordLong = ExlHandler.FindColumn_continuous(srcWb, srcWs, srcWs.UsedRange, srcElementStr, xlWhole, direct:=xlPrevious)
        destKeywordLong = ExlHandler.FindColumn_continuous(destWB, destWs, destWs.UsedRange, destElementStr, xlWhole, direct:=xlPrevious)
        
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
    
    formulaCol1 = ExlHandler.FindColumn_continuous(destWB, destWs, destWs.UsedRange, formulaColStr1, xlWhole)
    formulaCol2 = ExlHandler.FindColumn_continuous(destWB, destWs, destWs.UsedRange, formulaColStr2, xlWhole)

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


Sub RepositionTabs(tabsArray As Variant, destWB As Workbook)
    
    Dim arrayCtr As Long, tabNm As String, shCtr As Long, lastSheetNum As Long
    
    destWB.Activate
    
    For arrayCtr = LBound(tabsArray) To UBound(tabsArray)
        tabNm = tabsArray(arrayCtr)
        
        With destWB
            .Sheets(tabNm).Move After:=.Sheets(arrayCtr + 1)
        End With
        
    Next arrayCtr
    
    lastSheetNum = destWB.Sheets(arrayCtr + 1).Index
    
    If lastSheetNum < destWB.Sheets.Count Then
        
        For shCtr = UBound(tabsArray) + 1 To destWB.Sheets.Count
                
                If InStr(1, destWB.Sheets(shCtr).Name, "Data", vbTextCompare) > 0 Or _
                    InStr(1, destWB.Sheets(shCtr).Name, " by ", vbTextCompare) > 0 Then
                    
                    destWB.Sheets(shCtr).Visible = xlSheetHidden
                    
                End If
        Next shCtr
        
    End If
    
    destWB.Sheets(1).Activate
    
End Sub

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

Function ReplaceKeywordsInSheet(ws As Worksheet, prevKeyword As Variant, newKeyword As Variant)
    
    Dim varCtr As Long, strPrevKeyword As String, strNewKeyword As String
    
    'This function replaces keywords using arrays. It is a requirement that the arrays have the same sizes.
    
    With ws
        
        
        For varCtr = LBound(newKeyword) To UBound(newKeyword)
            strPrevKeyword = prevKeyword(varCtr)
            strNewKeyword = newKeyword(varCtr)
            
            .Cells.Replace strPrevKeyword, strNewKeyword
            
        Next varCtr
        
    End With
    
End Function

Function FindDateAndFormat(ws As Worksheet, formatMode As String, Optional transferWhere As Long = 0) As Range
    
    Dim rngFind As Range, rngWithVals As Range, rngElement As Range, rngFindCol As Long
    
    'this function finds a date range and returns a date range, it's optional to format the found range
    
    With ws
        .Activate
        Set rngWithVals = .UsedRange
        
        For Each rngElement In rngWithVals
            If IsDate(rngElement.Value) Then
                'do function
                                
                Set rngFind = rngElement
                rngFindCol = rngFind.Column
                
                Select Case formatMode
                    Case "unmerge"  'this will unmerge all the rows starting from the start of the found range up to the end
                        .Range(.Cells(rngFind.row, rngFindCol), .Cells(rngFind.row, .UsedRange.Columns.Count)).UnMerge
                    Case "unmerge&transfer"
                        .Range(.Cells(rngFind.row, rngFindCol), .Cells(rngFind.row, .UsedRange.Columns.Count)).UnMerge
                        .Cells(rngFind.row, rngFindCol).Cut
                        .Cells(rngFind.row, rngFindCol + transferWhere).Select
                        ActiveSheet.Paste
                            'transferWhere = 0
                            .Range(.Cells(rngFind.row, rngFindCol + transferWhere), _
                                .Cells(rngFind.row, .UsedRange.Columns.Count)).Merge
                    
                        
                End Select
                
                
                Exit Function
            End If
        Next rngElement
        
    End With
    
End Function


Function CheckDelimiterFormat(wbToCheck As Workbook, wsToCheck As Worksheet, strHeaderKeyword As String, strDelimiter As String, _
    Optional checkLimitRowCount As Long, Optional IncludeHdrRow As Boolean = False, Optional IncludeEndRow As Boolean = False)
    
    Dim startRow As Long, endRow As Long, hdrRow As Long, hdrCol As Long, ctr As Long, strCheck As String
    
    With wsToCheck
        hdrRow = FindRow(wbToCheck, wsToCheck, .UsedRange, strHeaderKeyword)
        hdrCol = FindColumn(strHeaderKeyword, hdrRow)
        
        
        If IncludeHdrRow = True Then
            startRow = hdrRow
        Else
            startRow = hdrRow + 1
        End If
        
        endRow = .Cells(startRow, hdrCol).End(xlDown).row
        
        If IncludeEndRow = False Then endRow = endRow - 1
        
        If checkLimitRowCount = 0 Then checkLimitRowCount = endRow
        
        For ctr = startRow To checkLimitRowCount
            
            strCheck = .Cells(ctr, hdrCol).Value
            
            If InStr(1, strCheck, strDelimiter, vbTextCompare) = 0 Then
                
                MsgBox "The file " & wbToCheck.Name & " with the worksheet " & wsToCheck.Name & " doesn't have the delimiters format of '" & _
                    strDelimiter & "' In column " & GetColumnLetter(CInt(hdrCol)) & ". Please choose the correct input file and try again."
                
                wbToCheck.Activate
                wsToCheck.Activate
                
                End
            End If
            
        Next ctr
        
    End With
    
    'This function checks if a specific column have the required delimiters in which the programmer can set up to which row.
    
End Function

Function EditTextbox(ws As Worksheet, strTxtBxNm As String, strText As String, Optional longTop As Long, _
    Optional longLeft As Long, Optional boldUpToWhere As Integer, Optional fontSize1 As Integer, _
    Optional fontSize1UptoWhere As Integer, Optional fontSize2 As Integer, Optional fontSize2StartWhere As Integer, _
    Optional fontSize2UptoWhere As Integer)

    With ws
        With .Shapes(strTxtBxNm)
            .TextFrame.Characters.Text = strText
        
            If longTop <> 0 Then .Top = longTop
            If longLeft <> 0 Then .Left = longLeft
            
            .TextFrame.Characters.Font.Bold = False
            
            If boldUpToWhere <> 0 Then .TextFrame.Characters(1, boldUpToWhere).Font.Bold = True
            
            If fontSize1 <> 0 And fontSize1UptoWhere <> 0 Then
                .TextFrame.Characters(1, fontSize1UptoWhere).Font.Size = fontSize1
            End If
            
            If fontSize2 <> 0 And fontSize2UptoWhere <> 0 And fontSize2StartWhere <> 0 Then
                .TextFrame.Characters(fontSize2StartWhere, fontSize2UptoWhere).Font.Size = fontSize2
            End If
            
        End With
        
    End With

End Function

Sub FindColumnsThenRemove(wsSource, strFind As String)

Dim rangeFindFF As Range, columnCtr As Long

                
                With wsSource
                  
                    .Activate
                    .AutoFilterMode = False
                    Set rangeFindFF = .UsedRange.Find(strFind, , , xlWhole)
                    
                    If Not rangeFindFF Is Nothing Then
                    
                        columnCtr = rangeFindFF.Column
                        
                        Do
                            
                            'If .Cells(rangeFindFF.row, rangeFindFF.Column).MergeCells = True Then
                                '.Columns(rangeFindFF.Column).UnMerge
                                '.Range(.Cells(rangeFindFF.row, rangeFindFF.Column), .Cells(rangeFindFF.row, rangeFindFF.Column + 1)).EntireColumn.Delete
                            'Else
                                '.Cells(rangeFindFF.row, rangeFindFF.Column).EntireColumn.Delete
                            'End If
                            .Columns(rangeFindFF.Column).EntireColumn.Delete
                            
                            Set rangeFindFF = .UsedRange.Find(strFind, , , xlWhole)
                            
                            If rangeFindFF Is Nothing Then
                            
                                GoTo DoneFinding
                                
                            
                            End If
                    
                        Loop While rangeFindFF.Column <> columnCtr
                    
  
                    End If
                    
                End With
                
DoneFinding:
            

End Sub

Function IsWorkbookProtected(wb As Workbook) As Boolean

With wb
    
   IsWorkbookProtected = (.ProtectWindows Or .ProtectStructure)
    
End With

End Function

Function showNonZeroOnly(ws As Worksheet, rngToPlot As Range)
    
    Dim rng As Range, sEntry As String, dEntryNumber As Double, arr As Variant, longFirstInst As Long
    Dim strRemovedZero As String, strKeywordForPos As String, longKwordPos As Long, strNumFmt As String
    Dim decCtr As Long, decSum As Double
    
    With ws
        For Each rng In rngToPlot
                If IsNumeric(rng) Then
                    sEntry = rng.Value
                    dEntryNumber = CDbl(sEntry)

                    arr = Split(sEntry, ".")
                    If UBound(arr) = 1 Then
                        ' Change NumberFormat in accordance with
                        ' the number of digits after the decimal point
                        'rng.NumberFormat = "0." & String(Len(arr(1)), "0")
                       
                       'FIND(MID(SUBSTITUTE(I171,0,""),2,1),I171)-1
                       
                       'substitute removes the zeros to cut down the leading zeros of the string of the keyword
                       'mid gets the keyword to find to get the starting position
                       'find gets the starting position
                       
                       strNumFmt = "10"
                            
                       strRemovedZero = Replace(Replace(rng.Value, "0", ""), ".", "") 'Application.WorksheetFunction.Substitute(rng.Value, "0", "")
                       strKeywordForPos = Mid(strRemovedZero, 1, 1)
                       
                       Dim wordLength As String
                       
                       wordLength = Len(rng.Value)
                            
                       
                            
                            rng.NumberFormat = "@"
                            rng.FormulaR1C1 = rng.FormulaR1C1
                       
                            
                            rng.NumberFormat = "0." & CStr(CLng(strNumFmt) ^ (wordLength - 2))    '0.0034
                            
                            longKwordPos = Application.WorksheetFunction.Find(strKeywordForPos, rng.Text) - 4
                            
                            If longKwordPos < 2 Then longKwordPos = 2
                            
                            
                            rng.NumberFormat = "0." & Replace(CStr(CLng(strNumFmt) ^ (longKwordPos)), "1", "") & "%"   '0.0034
                            rng.FormulaR1C1 = rng.FormulaR1C1
                            
                            On Error Resume Next
                            'for unchanged exponents
                            If Application.WorksheetFunction.Find("E", rng.Text) > 0 Then
                                                               
                                If Err.Number = 0 Then
                                    rng.Style = "Percent"
                                        'rng.NumberFormat = "0." & CStr(CLng(strNumFmt) ^ (CLng(Split(rng.Text, "E-")(1)) - 2)) & "&"
                                        decSum = 0
        
                                        For decCtr = 1 To wordLength
        
                                            If Mid(rng.Text, decCtr, 1) <> "." And IsNumeric(Mid(rng.Text, decCtr, 1)) Then
                                                decSum = decSum + CDbl(Replace(rng.Text, "%", ""))
        
                                                If decSum > 0 Then
                                                    Exit For
                                                Else
                                                    rng.NumberFormat = "0." & Replace(CStr(CLng(strNumFmt) ^ (decCtr)), "1", "") & "%"
                                                    'rng.NumberFormat = "0." & Replace(CStr(CLng(strNumFmt) ^ (decCtr)), "1", "") & "%"
                                                End If
                                            Else
                                                rng.NumberFormat = "0." & Replace(CStr(CLng(strNumFmt) ^ (decCtr)), "1", "") & "%"
                                            End If
        
                                        Next decCtr
        
                                    On Error GoTo 0
                                    Err.Clear
                                End If
                            End If
                            
                            'for 0.10

                            If Mid(rng.Text, Len(rng.Text) - 1, 1) = 0 And rng.NumberFormat <> "0.00%" And _
                                Mid(rng.Text, Len(rng.Text) - 2, 1) > 0 Then
                                
                                rng.NumberFormat = Mid(CStr(rng.NumberFormat), 1, Len(CStr(rng.NumberFormat)) - 2) & "%"
                                
                            End If
                                
                            
                    End If
                End If
        Next rng
    End With

End Function

Function computeSumOfGroup_double(wb As Workbook, ws As Worksheet, strStartRow As String, strStartCol As String, _
    strSumCol As String, Optional rng As Range, Optional preLongSumCol As Long = 0, Optional preLongSumColAdder As Long)
    
    Dim longStartRow As Long, longStartCol As Long, longSumCol As Long
    Dim ctr As Long, grpStartRow As Long, grpEndRow As Long
    
    With ws
        
        If rng Is Nothing Then Set rng = .UsedRange
        
        longStartRow = SearchKeywordRow(wb, ws, strStartRow, .UsedRange, xlWhole) + 1
        longStartCol = SearchKeywordCol(wb, ws, strStartCol, .UsedRange, xlWhole)
        
        If preLongSumCol = 0 Then
        longSumCol = SearchKeywordCol(wb, ws, strSumCol, .Range(rng.Address), xlWhole)
        Else
            longSumCol = preLongSumCol
        End If
        
        ctr = longStartRow
        
        For ctr = longStartRow To .UsedRange.Rows.Count
            
            grpStartRow = ctr
            
            If .Cells(ctr + 1, longStartCol).Value <> vbNullString Then
                grpEndRow = ctr
                grpStartRow = ctr
            Else    'if next row is blank
                grpEndRow = .Cells(grpStartRow, longStartCol).End(xlDown).row - 1
            End If
            
            If grpEndRow > grpStartRow And .Cells(grpStartRow, longStartCol) <> "TOTAL" Then
                'sum-up group
                
                On Error Resume Next
                                
                    .Cells(grpStartRow, longSumCol).Value = _
                    Application.WorksheetFunction.Sum(.Range(.Cells(grpStartRow + 1, longSumCol), .Cells(grpEndRow, longSumCol)))
                    
                    If preLongSumColAdder = 0 Then preLongSumColAdder = 1
                    
                    .Cells(grpStartRow, longSumCol + preLongSumColAdder).Value = _
                    Application.WorksheetFunction.Sum(.Range(.Cells(grpStartRow + 1, longSumCol + preLongSumColAdder), .Cells(grpEndRow, longSumCol + preLongSumColAdder)))
                    
                    If Err.Number > 0 Then
                        MsgBox "Cannot sum-up this column number: " & strSumCol
                    End If
                    
                On Error GoTo 0
                Err.Clear
                
            End If
            
            ctr = grpEndRow
            
nextCtrLine:
            
        Next
        
    End With
    
End Function


Function ChangeWidthOrAutofit(wb As Workbook, ws As Worksheet, widthDbl As Double, refCol As Long, _
    startRow As Long, lastrow As Long)
    
    Dim hasRow As Long
    
    hashRow = FindRow(wb, ws, .Range(.Cells(startRow, refCol), .Cells(lastrow, refCol)), "*#", xlWhole)
    
    With ws
        .Columns(refCol).ColumnWidth = widthDbl
        If hasRow > 0 Then
            .Columns(refCol).AutoFit
        End If
    
    End With
End Function

Function wsFindDiffMatrix(srcWb As Workbook, srcWs As Worksheet, destWb As Workbook, destWs As Worksheet, srcHdrRow As Long, _
    destHdrRow As Long, destLrow As Long, destHdrColStart As Long, destHdrColEnd As Long, destRefCol, _
    destRowStart As Long, whatColor As Double)

    Dim destHdrColCtr As Long, srcRefCol As Long, srcRefRow As Long, destRowCtr As Long, strDest _
        As String, strSrcNm As String, strDestNm As String
    
    With destWs
        
            For destRowCtr = destRowStart To destLrow
                
                For destHdrColCtr = destHdrColStart To destHdrColEnd
                    strDest = .Cells(destRowCtr, destRefCol).Value
                    strDestNm = .Cells(destHdrRow, destHdrColCtr).Value
                    
                    If Trim(strDest) <> vbNullString And Trim(strDestNm) <> vbNullString Then
                        'get coordinates
                        
                        With srcWs
                            srcRefRow = FindRow(srcWb, srcWs, .UsedRange, strDest, xlPart)
                            
                            If srcRefRow <> 0 Then
                                
                                srcRefCol = FindColumn(srcWb, srcWs, .UsedRange, strDestNm, xlWhole)
                                
                                If srcRefCol <> 0 Then
                                    
                                    If IsNumeric(.Cells(srcRefRow, srcRefCol).Value) And IsNumeric(destWs.Cells(destRowCtr, destHdrColCtr).Value) Then
                                    
                                    'compare numbers
                                    If .Cells(srcRefRow, srcRefCol).Value > _
                                        destWs.Cells(destRowCtr, destHdrColCtr).Value Then
                                        'color yellow
                                        
                                        If destWs.Cells(destRowCtr, destHdrColCtr).Value <> 0 Then
                                        
                                        With destWs.Cells(destRowCtr, destHdrColCtr).Interior
                                            .Pattern = xlSolid
                                            .PatternColorIndex = xlAutomatic
                                            .Color = whatColor
                                            .TintAndShade = 0
                                            .PatternTintAndShade = 0
                                        End With
                                        
                                        Else
                                            destWs.Cells(destRowCtr, destHdrColCtr).NumberFormat = "0.00"
                                            
                                        End If
                                        
                                    Else
                                        Exit For
                                    End If
                                    
                                    Else
                                        Exit For
                                    End If
                                    
                                End If
                                
                            Else
                                Exit For
                            End If
                            
                        End With
                    Else
                        Exit For
                    End If
                    
                Next destHdrColCtr
                
            Next destRowCtr
        
        
        
        
    End With
    
End Function

Function CreateConcatAndLookupValNonContigOutput(wb As Workbook, ws As Worksheet, refOutColBasis As Long, refLookupColBasis As Long, _
    ByRef ConcatCol As Long, ByRef outCol As Long, ByRef sRow As Long, ByRef lrow As Long, arrayStrConcat As Variant, _
    strStartRowHdr As String, Optional strRefLastRowGetter As String, Optional outColPos As Long, Optional concatStyle As String)
    
    Dim arrayStrConcElem As Variant, strFormula As String, strFormula2 As String
    
    With ws
    'initiate variables
        sRow = FindRow(wb, ws, .UsedRange, strStartRowHdr, xlWhole)
        
        'get last row if it's not inputted in the parameter
        If lrow = 0 Then
            'check if a keyword is inputted as reference to get the last row-
            'using a keyword representing the column with non-contiguous cells
            If strRefLastRowGetter <> vbNullString Then
                'if keyword was found, use it to get the reference column then to finally get the last row
                lrow = .Cells(sRow, FindColumn(wb, ws, .UsedRange, strRefLastRowGetter, xlWhole)).End(xlDown).Row
                
                'if last row is too big just get the usedrange last row
                If lrow > .Cells.Rows.Count - 5 Then lrow = .UsedRange.Rows.Count
            Else
                'if no inputted last row and no keyword inputted, just use the usedrange last row
                lrow = .UsedRange.Rows.Count
            End If
        End If
        
        'if the column where there is non-contiguous column is 0 then the macro will just use the first column
        If refOutColBasis = 0 Then refOutColBasis = 1
        
        'if plot column is not defined, plot to the last column
        If outColPos = 0 Then outColPos = .Cells(sRow, .Cells.Columns.Count).End(xlToLeft).Column + 1
        
        If ConcatCol = 0 Then
            'set the plot columns
            ConcatCol = outColPos
        End If
        
        If outCol = 0 Then
            outCol = outColPos + 1
        End If
        
        'plot formulas
        
        'plot concat formula based on array of column names
        
        For Each arrayStrConcElem In arrayStrConcat
            If strFormula = vbNullString Then
                'if formula builder is not initialized
                If CStr(arrayStrConcElem) <> vbNullString Then
                    strFormula = _
                       .Cells(sRow + 1, FindColumn(wb, ws, .UsedRange, CStr(arrayStrConcElem), xlWhole)).Address(False, False)
                    
                Else
                    strFormula = _
                       .Cells(sRow + 1, refLookupColBasis).Address(False, False)
                       
                End If
                
                'if formula builder is not initialized
                If CStr(arrayStrConcElem) <> vbNullString Then
                    strFormula2 = _
                       .Cells(sRow + 1, FindColumn(wb, ws, .UsedRange, CStr(arrayStrConcElem), xlWhole)).Address(False, False)
                    
                Else
                    strFormula2 = _
                       .Cells(sRow + 1, refOutColBasis).Address(False, False)
                       
                End If
                
            Else
                'if formula builder is initialized
                If CStr(arrayStrConcElem) <> vbNullString Then
                    'if formula in array is not blank
                    
                    If concatStyle = vbNullString Then  'if there is no concat style
                    strFormula = strFormula & " & " & _
                    .Cells(sRow + 1, FindColumn(wb, ws, .UsedRange, CStr(arrayStrConcElem), xlWhole)).Address(False, False)
                    ElseIf concatStyle = "()" Then 'if it is parenthesis style
                        strFormula = strFormula & " & " & _
                    """ (""" & " & " & .Cells(sRow + 1, FindColumn(wb, ws, .UsedRange, CStr(arrayStrConcElem), xlWhole)).Address(False, False) & " & " & """)"""
                    
                    End If
                Else
                    'if formula in array is blank
                    
                    If concatStyle = vbNullString Then  'if there is no concat style
                    strFormula = strFormula & " & " & _
                    .Cells(sRow + 1, refLookupColBasis).Address(False, False)
                    ElseIf concatStyle = "()" Then 'if it is parenthesis style
                    strFormula = strFormula & " & " & _
                    """ (""" & " & " & .Cells(sRow + 1, refLookupColBasis).Address(False, False) & " & " & """)"""
                    End If
                    
                End If
                
                
                'if formula builder is initialized
                If CStr(arrayStrConcElem) <> vbNullString Then
                    'if formula in array is not blank
                    
                    If concatStyle = vbNullString Then  'if there is no concat style
                    strFormula2 = strFormula2 & " & " & _
                    .Cells(sRow + 1, FindColumn(wb, ws, .UsedRange, CStr(arrayStrConcElem), xlWhole)).Address(False, False)
                    ElseIf concatStyle = "()" Then 'if it is parenthesis style
                        strFormula2 = strFormula2 & " & " & _
                    """ (""" & " & " & .Cells(sRow + 1, FindColumn(wb, ws, .UsedRange, CStr(arrayStrConcElem), xlWhole)).Address(False, False) & " & " & """)"""
                    
                    End If
                Else
                    'if formula in array is blank
                    
                    If concatStyle = vbNullString Then  'if there is no concat style
                    strFormula2 = strFormula2 & " & " & _
                    .Cells(sRow + 1, refOutColBasis).Address(False, False)
                    ElseIf concatStyle = "()" Then 'if it is parenthesis style
                    strFormula2 = strFormula2 & " & " & _
                    """ (""" & " & " & .Cells(sRow + 1, refOutColBasis).Address(False, False) & " & " & """)"""
                    End If
                    
                End If
                
            End If
        Next arrayStrConcElem
        
        .Range(.Cells(sRow + 1, ConcatCol), .Cells(lrow, ConcatCol)) = "=IF(" & .Cells(sRow + 1, refOutColBasis).Address(False, False) & "=" & """""" & "," & strFormula & "," & strFormula2 & ")"
        
        'plot output formula to solve non-contiguous output column
        '=IF(B5<>"",B5,P4)
        .Range(.Cells(sRow + 1, outCol), .Cells(lrow, outCol)) = "=IF(" & .Cells(sRow + 1, refOutColBasis).Address(False, False) & _
            "<>" & """""" & "," & .Cells(sRow + 1, refOutColBasis).Address(False, False) & "," & _
             .Cells(sRow, outCol).Address(False, False) & ")"
        
        .Range(.Cells(sRow + 1, ConcatCol), .Cells(lrow, outCol)).Value = .Range(.Cells(sRow + 1, ConcatCol), .Cells(lrow, outCol)).Value
        
    End With
    
End Function


Function sumUpRangesRowNumeric(wb As Workbook, ws As Worksheet, rngToSumUp As Range, colToCheck As Long, _
    rowToPlot As Long, strCriteria As String)
    
    Dim startRow As Long, lastRow As Long, startCol As Long, lastCol As Long, coltr As Long
    
    startRow = rngToSumUp.Row
    lastRow = rngToSumUp.Row + rngToSumUp.Rows.Count - 1
    startCol = rngToSumUp.Column
    lastCol = rngToSumUp.Column + rngToSumUp.Columns.Count - 1
    
    With ws
        
        .Range(.Cells(rowToPlot, startCol), .Cells(rowToPlot, lastCol)) = _
            "=IF(ISNUMBER(" & .Cells(startRow, startCol).Address(False, False) & ")," & _
            "SUMIF(" & .Range(.Cells(startRow, colToCheck), .Cells(lastRow, colToCheck)).Address(True, True) & _
            "," & strCriteria & "," & _
            .Range(.Cells(startRow, startCol), .Cells(lastRow, startCol)).Address(True, False) & _
            ")," & """""" & ")"
        
'        For coltr = startCol To lastCol
'            If isnumberic(.Cells(startRow, coltr).Value) And _
'                Trim(.Cells(startRow, coltr).Value) <> vbNullString Then
'
'                '.Cells(rowToPlot,coltr).value =
'
'            End If
'        Next
    End With
    
End Function
