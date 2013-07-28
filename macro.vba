
' -------------------------------------------------------------------
' Range関数を行単位で選択しやすくした関数
' ex) RowSelector("A:B", 10) ⇒ Range("A10:B10") と同義
' ex) RowSelector("", 9)     ⇒ Range("9:9")     と同義
' -------------------------------------------------------------------
Private Function RowSelector(str As String, row_ As Long, Optional Sh As Worksheet = Nothing) As Range
    If Sh Is Nothing Then Set Sh = ActiveSheet
    
    If str = Empty Then
        Set RowSelector = Sh.Range("" & row_ & ":" & row_)
        Exit Function
    End If
    
    Dim v As Variant: v = Split(str, ":")
    If UBound(v) + 1 = 2 Then
        Set RowSelector = Sh.Range(v(0) & row_ & ":" & v(1) & row_)
    End If
End Function

