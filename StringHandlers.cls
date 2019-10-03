VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringHandlers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function GetColumnLetter(iCol As Integer) As String
'
' Feed in a column number and spew out the column letter
'
    Dim sCol As String
    'PrintToLog ("Calling GetColumnLetter")
    If iCol > 26 Then ' spew out two letters i.e. AA onwards
        sCol = Left(Cells(1, iCol).Address(False, False), 2)
    Else ' spew out one letter i.e. A to Z
        sCol = Left(Cells(1, iCol).Address(False, False), 1)
    End If
    GetColumnLetter = sCol
End Function

Function RemoveSpecial(str As String) As String

Dim xChars As String
Dim i As Long
xChars = "#$%()^*&|'/:."

For i = 1 To Len(xChars)
        
        If i <> 11 Then
        str = Replace$(str, Mid$(xChars, i, 1), " ")
        Else
        str = Replace$(str, Mid$(xChars, i, 1), "")
        End If
    
Next

RemoveSpecial = str

End Function
Function IsAlphaNumeric(pValue) As Boolean

   Dim LPos As Integer
   Dim LChar As String
   Dim LValid_Values As String

   'Start at first character in pValue
   LPos = 1

   'Set up values that are considered to be alphanumeric
   LValid_Values = " abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"

   'Test each character in pValue
   While LPos <= Len(pValue)

      'Single character in pValue
      LChar = Mid(pValue, LPos, 1)

      'If character is not alphanumeric, return FALSE
      If InStr(LValid_Values, LChar) = 0 Then
         IsAlphaNumeric = False
         Exit Function
      End If

      'Increment counter
      LPos = LPos + 1

   Wend

   'Value is alphanumeric, return TRUE
   IsAlphaNumeric = True

End Function

Function HasAlphaNumeric(pValue) As Boolean

   Dim LPos As Integer
   Dim LChar As String
   Dim LValid_Values As String
   Dim LValid_Alphabet As String
   Dim LValid_Numbers As String
   Dim v_Alpha As Boolean
   Dim v_Numbers As Boolean
   Dim InValid_Value As String

   'Start at first character in pValue
   LPos = 1

   'Set up values that are considered to be alphanumeric
   LValid_Alphabet = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
   LValid_Numbers = "0123456789"
   InValid_Value = "#$%()^*&|'/:@!"
   v_Alpha = False
   v_Numbers = False
   'Test each character in pValue
   While LPos <= Len(pValue)

      'Single character in pValue
      LChar = Mid(pValue, LPos, 1)

      'If character is not alphanumeric, return FALSE
      If InStr(LValid_Alphabet, LChar) = 0 And InStr(InValid_Value, pValue) > 0 Then
         v_Alpha = False
         'Exit Function
      ElseIf InStr(LValid_Alphabet, LChar) > 0 And InStr(InValid_Value, pValue) = 0 Then
        v_Alpha = True
        GoTo HERE
      End If
      
      'Increment counter
      LPos = LPos + 1

   Wend
   
HERE:

    LPos = 1

   While LPos <= Len(pValue)

      'Single character in pValue
      LChar = Mid(pValue, LPos, 1)

      'If character is not alphanumeric, return FALSE
        If InStr(LValid_Numbers, LChar) = 0 And InStr(InValid_Value, pValue) > 0 Then
         v_Numbers = False
         'Exit Function
      ElseIf InStr(LValid_Numbers, LChar) > 0 And InStr(InValid_Value, pValue) = 0 Then
        v_Numbers = True
        GoTo HERE1
      End If
      

      'Increment counter
      LPos = LPos + 1

   Wend
   
HERE1:
    
    If v_Numbers = True And v_Alpha = True Then
            'Value is alphanumeric, return TRUE
            HasAlphaNumeric = True
    Else
            HasAlphaNumeric = False
    End If

End Function
