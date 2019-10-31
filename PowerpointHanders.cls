VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PowerpointHanders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Function SetColorWeightingsinPowerpoint(iRow As Long, Lrow As Long, iCol1 As Long, iCol2 As Long, strTable As String)
    'set color for two tables
    Dim r As Long

    For r = iRow To Lrow
    
        With pptSlide.Shapes(strTable).Table.cell(r, iCol1).Shape.TextFrame
        
            If .TextRange = 1 Then
                .TextRange = ChrW(&H25B2)
                .TextRange.Font.Color = RGB(0, 175, 80)
            ElseIf .TextRange = 2 Then
                .TextRange = ChrW(&H25B2) & ChrW(&H25B2)
                .TextRange.Font.Color = RGB(0, 175, 80)
            ElseIf .TextRange = 3 Then
                .TextRange = ChrW(&H25B2) & ChrW(&H25B2) & ChrW(&H25B2)
                .TextRange.Font.Color = RGB(0, 175, 80)
            ElseIf .TextRange = -1 Then
                .TextRange = ChrW(&H25BC)
                .TextRange.Font.Color = RGB(255, 0, 0)
            ElseIf .TextRange = -2 Then
                .TextRange = ChrW(&H25BC) & ChrW(&H25BC)
                .TextRange.Font.Color = RGB(255, 0, 0)
            ElseIf .TextRange = -3 Then
                .TextRange = ChrW(&H25BC) & ChrW(&H25BC) & ChrW(&H25BC)
                .TextRange.Font.Color = RGB(255, 0, 0)
            ElseIf .TextRange = 0 Then
                .TextRange = ChrW(&H25CF)
                .TextRange.Font.Color = RGB(0, 112, 192)
            ElseIf .TextRange = "-" Then
                .TextRange = "n/a"
                .TextRange.Font.Color = RGB(0, 0, 0)
            ElseIf .TextRange = "n/a" Then
                .TextRange.Font.Color = RGB(0, 0, 0)
            End If
            
        End With
        
        With pptSlide.Shapes(strTable).Table.cell(r, iCol2).Shape.TextFrame
        
            If .TextRange = 1 Then
                .TextRange = ChrW(&H25B2)
                .TextRange.Font.Color = RGB(0, 175, 80)
            ElseIf .TextRange = 2 Then
                .TextRange = ChrW(&H25B2) & ChrW(&H25B2)
                .TextRange.Font.Color = RGB(0, 175, 80)
            ElseIf .TextRange = 3 Then
                .TextRange = ChrW(&H25B2) & ChrW(&H25B2) & ChrW(&H25B2)
                .TextRange.Font.Color = RGB(0, 175, 80)
            ElseIf .TextRange = -1 Then
                .TextRange = ChrW(&H25BC)
                .TextRange.Font.Color = RGB(255, 0, 0)
            ElseIf .TextRange = -2 Then
                .TextRange = ChrW(&H25BC) & ChrW(&H25BC)
                .TextRange.Font.Color = RGB(255, 0, 0)
            ElseIf .TextRange = -3 Then
                .TextRange = ChrW(&H25BC) & ChrW(&H25BC) & ChrW(&H25BC)
                .TextRange.Font.Color = RGB(255, 0, 0)
            ElseIf .TextRange = 0 Then
                .TextRange = ChrW(&H25CF)
                .TextRange.Font.Color = RGB(0, 112, 192)
            ElseIf .TextRange = "-" Then
                .TextRange = "n/a"
                .TextRange.Font.Color = RGB(0, 0, 0)
            ElseIf .TextRange = "n/a" Then
                .TextRange.Font.Color = RGB(0, 0, 0)
            End If
            
        End With
        


    Next r


End Function

Function SetRowColorinPowerpoint(iRow As Long, Lrow As Long, iCol1 As Long, iCol2 As Long, strTable As String, colStart1 As Long, colEnd1 As Long, _
    colStart2 As Long, colEnd2 As Long, rowColor As Long)

Dim r As Long, colColorCtr As Long

    For r = iRow To Lrow
    
        With pptSlide.Shapes(strTable).Table.cell(r, iCol1).Shape.TextFrame
        
            If .TextRange < 5 Then
                '.TextRange = ChrW(&H25B2)
                '.TextRange.Font.Color = RGB(0, 175, 80)
                 
                For colColorCtr = colStart1 To colEnd1
                    
                    pptSlide.Shapes(strTable).Table.cell(r, colColorCtr).Shape.Fill.ForeColor.RGB = rowColor 'RGB(217, 217, 217)
                    
                Next colColorCtr
                 
                
            End If
            
        End With
        
        With pptSlide.Shapes(strTable).Table.cell(r, iCol2).Shape.TextFrame
        
            If .TextRange < 5 Then
                
                For colColorCtr = colStart2 To colEnd2
                    
                    pptSlide.Shapes(strTable).Table.cell(r, colColorCtr).Shape.Fill.ForeColor.RGB = rowColor 'RGB(217, 217, 217)
                    
                Next colColorCtr
                 
            End If
            
        End With
        

    Next r


End Function



Public Function ChangeDatesinPowerpointTable(strTableName As String, longCurrRow As Long, longCurrCol As Long, longPrevRow As Long, longPrevCol As Long, strDtFormat As String)

    pptSlide.Shapes(strTableName).Table.cell(longCurrRow, longCurrCol).Shape.TextFrame.TextRange = Format(dateCurrent, strDtFormat) ' "dd-MMM-yy")
    pptSlide.Shapes(strTableName).Table.cell(longPrevRow, longPrevCol).Shape.TextFrame.TextRange = Format(datePrevious, strDtFormat) ' "dd-MMM-yy")

End Function


Function HighlightBlanksinPowerpointTable(iRow As Long, Lrow As Long, iCol1 As Long, iCol2 As Long, strTable As String, rowColor As Long)

Dim r As Long, colColorCtr As Long

    For r = iRow To Lrow
    
        With pptSlide.Shapes(strTable).Table.cell(r, iCol1).Shape.TextFrame
            If .TextRange = "" Or .TextRange = 0 Then

                    pptSlide.Shapes(strTable).Table.cell(r, iCol1).Shape.Fill.ForeColor.RGB = rowColor 'RGB(255, 255, 0)
                    .TextRange = ""

            End If
        End With
        
        With pptSlide.Shapes(strTable).Table.cell(r, iCol2).Shape.TextFrame
            If .TextRange = "" Or .TextRange = 0 Then
                    pptSlide.Shapes(strTable).Table.cell(r, iCol2).Shape.Fill.ForeColor.RGB = rowColor ' RGB(255, 255, 0)
                    .TextRange = ""
            End If
        End With
        

    Next r


End Function
