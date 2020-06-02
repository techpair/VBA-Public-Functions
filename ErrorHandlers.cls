VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorHandlers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False



Function ErrorHandler(strMsg As String)
    MsgBox strMsg & vbNewLine & vbNewLine & _
          "Error Number: " & Err.Number & vbNewLine & vbNewLine & _
          "Error Desription: " & vbNewLine & vbNewLine & _
         Err.Description, vbExclamation
    End
End Function

Sub delimitRow(refCols As String, delimiter As String, srcWs As Worksheet)
    Dim outWs As Worksheet, outWb As Workbook, srcColElement As Variant
    Dim srcRowCtr As Long, delArray As Variant, delColsArray As Variant, delArraySrcElement As Variant
    Dim plotCtr As Long, srcLastCol As Long, secondaryArray As Variant, secondaryArrayCtr As Long
    Dim srcCol As Long, secondaryCol As Long
    
    'this sub procedure creates a file that contains rows based from a delimited column and 2nd level delimited column
    
    Sheet1.Copy
    Set outWb = ActiveWorkbook
    Set outWs = outWb.ActiveSheet
    
    With srcWs
        
        srcLastCol = .UsedRange.Columns.Count
        plotCtr = 2
        
         srcCol = Application.WorksheetFunction.Match(CStr(Split(refCols, " | ")(0)), _
                    .Range(.Cells(1, 1), .Cells(1, srcLastCol)), 0)
                    
        secondaryCol = Application.WorksheetFunction.Match(CStr(Split(refCols, " | ")(1)), _
            .Range(.Cells(1, 1), .Cells(1, srcLastCol)), 0)
               
        
        For srcRowCtr = 2 To .UsedRange.Rows.Count
            'delColsArray = Split(refCols, delimiter)
                
                delArray = Split(CStr(.Cells(srcRowCtr, srcCol).Value), "|")
                
                For Each delArraySrcElement In delArray
                    
                    secondaryArray = Split(CStr(.Cells(srcRowCtr, secondaryCol).Value), "|")
                    For secondaryArrayCtr = LBound(secondaryArray) To UBound(secondaryArray)
                        
                        'copy range first
                        outWs.Range(outWs.Cells(plotCtr, 1), outWs.Cells(plotCtr, srcLastCol)).Value = _
                        .Range(.Cells(srcRowCtr, 1), .Cells(srcRowCtr, srcLastCol)).Value
                        
                        'plot corresponding column value
                        outWs.Cells(plotCtr, srcCol).Value = delArraySrcElement
                        outWs.Cells(plotCtr, secondaryCol).Value = secondaryArray(secondaryArrayCtr)
                        plotCtr = plotCtr + 1
                    
                    Next secondaryArrayCtr
                    
                    
                Next delArraySrcElement
                
                
        Next srcRowCtr
    End With
    
    With outWs
        
    End With
    
End Sub
