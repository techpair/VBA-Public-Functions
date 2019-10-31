VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PivotTableHandlers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public pvt As PivotTable

Property Get ThePivotTable() As PivotTable
    Set ThePivotTable = pvt
End Property

Sub CreatePivotTable(srcWb As Workbook, srcWs As Worksheet, srcRange As range, strPivotTableName As String, strDestAddress As String)
'PURPOSE: Creates a brand new Pivot table on a new worksheet from data in the ActiveSheet


Dim sht As Worksheet
Dim pvtCache As PivotCache
Dim StartPvt As String
Dim SrcData As String

'Determine the data range you want to pivot
    
  srcWb.Activate
  
  SrcData = srcWs.Name & "!" & range(srcRange.Address).Address(ReferenceStyle:=xlR1C1)

'Create a new worksheet
  Set sht = Sheets.Add

'Where do you want Pivot Table to start?
  StartPvt = sht.Name & "!" & sht.range(strDestAddress).Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    sourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable( _
    tableDestination:=StartPvt, _
    TableName:=strDestAddress)

End Sub

Function CreatePivot(wb As Workbook, destWs As Worksheet, youPivotTblName As String, srcDataws As Worksheet, srcStartRow As Long, _
    srcEndRow As Long, srcStartColumn As Long, srcEndColumn As Long, tableDestWsName As String, _
    tableDestStartRow As Long, taleDestStartColumn As Long, pivotField1 As String, pivotFieldOrientation As XlPivotFieldOrientation, _
    Optional pivotField2 As String, Optional pivotField2Orientation As XlPivotFieldOrientation, _
    Optional pivotField3 As String, Optional pivotField3Orientation As XlPivotFieldOrientation, _
    Optional pivotField4 As String, Optional pivotField4Orientation As XlPivotFieldOrientation, _
    Optional pivotField5 As String, Optional pivotField5Orientation As XlPivotFieldOrientation, _
    Optional pivotField6 As String, Optional pivotField6Orientation As XlPivotFieldOrientation, _
    Optional pivotField7 As String, Optional pivotField7Orientation As XlPivotFieldOrientation, _
    Optional pivotField8 As String, Optional pivotField8Orientation As XlPivotFieldOrientation, _
    Optional pivotField9 As String, Optional pivotField9Orientation As XlPivotFieldOrientation, _
    Optional pivotField10 As String, Optional pivotField10Orientation As XlPivotFieldOrientation, _
    Optional pivotField18 As String, Optional pivotField18Orientation As XlPivotFieldOrientation, _
    Optional pivotFieldTotal As String, Optional removeSubtotals As Boolean, Optional repLayout As XlLayoutRowType)
    
    Dim sourceData As String, tableDestination As String
    Dim PCache As PivotCache
    
    On Error Resume Next
  
    
    sourceData = "'" & srcDataws.Name & "'" & "!" & "R" & srcStartRow & "C" & srcStartColumn & ":R" & srcEndRow & "C" & _
        srcEndColumn
        
    tableDestination = "'" & tableDestWsName & "'" & "!" & "R" & tableDestStartRow & "C" & taleDestStartColumn

   Set PCache = wb.PivotCaches.Create(SourceType:=xlDatabase, sourceData:=sourceData, Version:=xlPivotTableVersion15)
   
   Set pvt = PCache.CreatePivotTable(tableDestination:=tableDestination, TableName:=youPivotTblName, DefaultVersion:=xlPivotTableVersion15)
                                                                        
    With destWs
        .Activate
        .Cells(1, 1).Select
    End With
    
    wb.ShowPivotTableFieldList = True
    
    With destWs.PivotTables(youPivotTblName).PivotFields(pivotField1)
        .Orientation = pivotFieldOrientation
        .Position = 1
    End With
       
       
       
    If pivotField18 <> vbNullString Then
        With destWs.PivotTables(youPivotTblName).PivotFields(pivotField18)
            .Orientation = pivotField18Orientation
            .Position = 1
            .Function = xlSum
        End With
    End If
    
    If pivotField16 <> vbNullString Then
        With destWs.PivotTables(youPivotTblName).PivotFields(pivotField16)
            .Orientation = pivotField16Orientation
            .Position = 1
            .Function = xlSum
        End With
    End If
           
    
    
    If pivotField14 <> vbNullString Then
        With destWs.PivotTables(youPivotTblName).PivotFields(pivotField14)
            .Orientation = pivotField14Orientation
            .Position = 1
            .Function = xlSum
        End With
    End If
    
    If pivotField13 <> vbNullString Then
        With destWs.PivotTables(youPivotTblName).PivotFields(pivotField13)
            .Orientation = pivotField13Orientation
            .Position = 1
            .Function = xlSum
        End With
    End If
    
    If pivotField12 <> vbNullString Then
        With destWs.PivotTables(youPivotTblName).PivotFields(pivotField12)
            .Orientation = pivotField12Orientation
            .Position = 1
            .Function = xlSum
        End With
    End If
    
    If pivotField11 <> vbNullString Then
        With destWs.PivotTables(youPivotTblName).PivotFields(pivotField11)
            .Orientation = pivotField11Orientation
            .Position = 1
            .Function = xlSum
        End With
    End If
    
    If pivotField10 <> vbNullString Then
        With destWs.PivotTables(youPivotTblName).PivotFields(pivotField10)
            .Orientation = pivotField10Orientation
            .Position = 1
            .Function = xlSum
        End With
    End If
    
    If pivotField9 <> vbNullString Then
        With destWs.PivotTables(youPivotTblName).PivotFields(pivotField9)
            .Orientation = pivotField9Orientation
            .Position = 1
            .Function = xlSum
        End With
    End If
    
    If pivotField8 <> vbNullString Then
        With destWs.PivotTables(youPivotTblName).PivotFields(pivotField8)
            .Orientation = pivotField8Orientation
            .Position = 1
            .Function = xlSum
        End With
    End If
    
    If pivotField7 <> vbNullString Then
        With destWs.PivotTables(youPivotTblName).PivotFields(pivotField7)
            .Orientation = pivotField7Orientation
            .Position = 1
            .Function = xlSum
        End With
    End If
    
    If pivotField6 <> vbNullString Then
        With destWs.PivotTables(youPivotTblName).PivotFields(pivotField6)
            .Orientation = pivotField6Orientation
            .Position = 1
            .Function = xlSum
        End With
    End If
    
    If pivotField5 <> vbNullString Then
        With destWs.PivotTables(youPivotTblName).PivotFields(pivotField5)
            .Orientation = pivotField5Orientation
            .Position = 1
            .Function = xlSum
        End With
    End If
    
    If pivotField4 <> vbNullString Then
        With destWs.PivotTables(youPivotTblName).PivotFields(pivotField4)
            .Orientation = pivotField4Orientation
            .Position = 1
            .Function = xlSum
        End With
    End If
    
    If pivotField3 <> vbNullString Then
        With destWs.PivotTables(youPivotTblName).PivotFields(pivotField3)
            .Orientation = pivotField3Orientation
            .Position = 1
            .Function = xlSum
        End With
    End If
    
    If pivotField2 <> vbNullString Then
        With destWs.PivotTables(youPivotTblName).PivotFields(pivotField2)
            .Orientation = pivotField2Orientation
            .Position = 1
            .Function = xlSum
        End With
    End If
    
    
    destWs.range("A3").Select
    
    With destWs.PivotTables(youPivotTblName)
        .InGridDropZones = True
        .RowAxisLayout repLayout
    End With
    
    destWs.range("A7").Select
    
    If removeSutotals = True Then
        destWs.PivotTables(youPivotTblName).PivotFields(pivot3).Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End If
    
    If pivotFieldTotal <> vbNullString Then
        destWs.PivotTables(youPivotTblName).AddDataField destWs.PivotTables( _
            youPivotTblName).PivotFields(pivotFieldTotal), "Sum of " & pivotFieldTotal, xlSum
            
    End If
    
    wb.ShowPivotTableFieldList = False
    
End Function
