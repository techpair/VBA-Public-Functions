VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileHandlers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function OpenFiles(ByRef wbVariableToOpen As Workbook, strCompleteFilename As Variant, strFilename As String, Optional varWsName As Variant = Empty, Optional varKeyword As Variant, Optional xlLkAt As XlLookAt = xlPart, Optional strInputname As String)

    With Application
        '.DisplayAlerts = False
        .EnableEvents = False
        
        On Error Resume Next
        Workbooks.Open strCompleteFilename, ReadOnly:=True, UpdateLinks:=False
        If Err.Number <> 0 Then
            ErrorHandler.ErrorHandler "Error while opening " & strCompleteFilename & " Program will abend." & vbNewLine & _
            "Error in function: " & "OpenFiles" & vbNewLine
        End If
        
        On Error GoTo 0
        
        
        .DisplayAlerts = True
        .EnableEvents = True
    End With
                
    'Set variables of the investor download
    Set wbVariableToOpen = Workbooks(strFilename)
    
    
    If IsEmpty(varWsName) = False Then  'only useful when you already know the worksheet names of the input files browsed
        'Added by ralph
        Dim inputCtr As Long
        Dim strWsName As String
        Dim strKeyword As String
        Dim searchRng As Variant
                
            For inputCtr = LBound(varWsName, 1) To UBound(varWsName, 1)
                    
                    strWsName = CStr(varWsName(inputCtr))
                    strKeyword = CStr(varKeyword(inputCtr))
                    
                    On Error Resume Next
                    Set wsWorking = wbVariableToOpen.Sheets(strWsName)
                    
                        If Err.Number <> 0 Then
                        
                            MsgBox "Missing worksheet name: '" & strWsName & "'" & vbNewLine _
                            & "Workbook name: '" & strFilename & "'" & vbNewLine _
                            & "Input: '" & strInputname & "'" & vbNewLine & vbNewLine _
                            & "Please upload the correct '" & strInputname & "' file. This will stop the macro.", vbExclamation, "Opening input file"
                            Err.Clear
                            CloseWorkbooks Array(wbMirlin, wbSecIdx, wbPrcVolFile)
                            End
                            
                        End If
                    On Error GoTo 0
        
                    With wsWorking
                    
                        Set searchRng = .range(.Cells(1, 1), .Cells(.UsedRange.Rows.Count, .UsedRange.Columns.Count)).Find(What:=strKeyword, LookIn:=xlValues, LookAt:=xlLkAt)
                        If searchRng Is Nothing Then
                        
                            MsgBox "Missing keyword: '" & strKeyword & "'" & vbNewLine _
                            & "Worksheet name: '" & strWsName & "'" & vbNewLine _
                            & "Workbook name: '" & strFilename & "'" & vbNewLine _
                            & "Input file: '" & strInputname & "'" & vbNewLine & vbNewLine _
                            & "Please upload the correct '" & strInputname & "' file. This will stop the macro.", vbExclamation, "Opening input file"
                            CloseWorkbooks Array(wbMirlin, wbSecIdx, wbPrcVolFile)
                            Err.Clear
                            End
                        
                        End If
                    
                    End With
            
            Next inputCtr
    End If
    
    Workbooks(strFilename).Windows(1).WindowState = xlMinimized
    
End Function


Function SaveReport(strTimeFormat As String, ByRef wbToSave As Workbook, strFilename As String, blnToClose As Boolean, Optional strOutputDir As String, Optional blnMakeVal As Boolean) As Boolean
     
    '****************************************************************************************************
    'Purpose:Saves the report automatically to a directory
    'Inputs: N/A
    'Outputs:Enable excel
    '****************************************************************************************************
     
    Dim strSysTym As String, strOutDir As String
    
    On Error GoTo errorline
    
    If strOutputDir = vbNullString Then strOutputDir = "OUTPUT"
    
    strOutDir = ThisWorkbook.Path & "\" & strOutputDir & "\"
    strSysTym = Format(Now(), strTimeFormat) ')"YYYYMMDD hhmmss")

    If Dir(strOutDir, vbDirectory) = vbNullString Then MkDir strOutDir

    With wbToSave
        Application.DisplayAlerts = False
        
        If blnMakeVal = True Then
            Dim ws As Worksheet
            
            For Each ws In wbToSave.Sheets
                ws.UsedRange.Value = ws.UsedRange.Value
            Next ws
        End If
        
        If strTimeFormat = vbNullString Then strSysTym = vbNullString
        
        On Error Resume Next
        
        Workbooks(Trim(Replace(strFilename, ".", " ") & " " & strSysTym & ".xlsx")).Close False
        On Error GoTo errorline
        Err.Clear
        
        .SaveAs strOutDir & Replace(strFilename, ".", " ") & " " & strSysTym, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        
        Application.DisplayAlerts = True
        
        If blnToClose = True Then
            .Close True
            Set wbToSave = Nothing
        End If
        
        SaveReport = True
        '.Close False
    End With
    
    
    
    Exit Function
errorline:
    MsgBox "Error while saving report", vbExclamation
    SaveReport = False
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
