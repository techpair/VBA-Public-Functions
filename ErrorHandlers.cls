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

