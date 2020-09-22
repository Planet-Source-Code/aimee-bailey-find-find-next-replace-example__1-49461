Attribute VB_Name = "FindMOD"
Public Function FindAndHighlight(txt1 As TextBox, SearchString As String, CaseSensitive As Boolean, Optional StartIndex As Integer)
Dim x As Integer
On Error GoTo err
Dim xSelStart As Integer
Dim xSelLength As Integer

If StartIndex <= 0 Then x = 1 Else x = StartIndex

    If CaseSensitive = True Then
        xSelStart = InStr(x, txt1.Text, SearchString) - 1
    Else
        xSelStart = InStr(x, LCase(txt1.Text), LCase(SearchString)) - 1
    End If

xSelLength = Len(SearchString)

txt1.SelStart = xSelStart
txt1.SelLength = xSelLength
err:
End Function

Public Function FindNextAndHighlight(txt1 As TextBox, SearchString As String, CaseSensitive As Boolean)
Dim x As Integer
On Error GoTo err
Dim xSelStart As Integer
Dim xSelLength As Integer

If txt1.SelStart <= 0 Then
    x = 1 + txt1.SelLength
Else
    x = txt1.SelStart + txt1.SelLength
End If

    If CaseSensitive = True Then
        xSelStart = InStr(x, txt1.Text, SearchString)
    Else
        xSelStart = InStr(x, LCase(txt1.Text), LCase(SearchString))
    End If

xSelLength = Len(SearchString)

Text1.SelStart = xSelStart
Text1.SelLength = xSelLength
err:
End Function

Public Function ReplaceAndHighLight(txt1 As TextBox, ReplaceWith As String)
Dim xSelStart As Integer
Dim xSelLength As Integer
On Error GoTo err

xSelStart = txt1.SelStart
xSelLength = Len(ReplaceWith)

txt1.SelText = ReplaceWith
txt1.SelStart = xSelStart
txt1.SelLength = xSelLength
err:
End Function
