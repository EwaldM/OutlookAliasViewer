Sub ShowAlias()
    Dim Item As Object
    Dim strHeaders As String
    Dim strTo As String
    Dim strEmail As String
    Dim i As Long, iEnd As Long

    If TypeName(ActiveWindow) = "Inspector" Then
        Set Item = ActiveWindow.CurrentItem
    Else
        Set Item = ActiveExplorer.Selection.Item(1)
    End If

    If TypeName(Item) = "MailItem" Then
        On Error Resume Next
        strHeaders = Item.PropertyAccessor.GetProperty( _
            "http://schemas.microsoft.com/mapi/proptag/0x007D001E")

        i = InStr(1, strHeaders, vbCrLf & "To:", vbTextCompare)
        If i > 0 Then
            i = i + 4
            iEnd = InStr(i, strHeaders, vbCrLf)
            If iEnd > 0 Then
                strTo = Trim$(Mid$(strHeaders, i, iEnd - i))
            Else
                strTo = Trim$(Mid$(strHeaders, i))
            End If
        End If
        On Error GoTo 0

        strEmail = ExtractEmail(strTo)
        If Len(strEmail) > 0 Then
            MsgBox strEmail
        Else
            MsgBox "None"
        End If
    End If
End Sub

Private Function ExtractEmail(ByVal s As String) As String
    Dim re As Object
    Dim matches As Object
    Dim m As Object
    Dim result As String

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"
    re.Global = True	' find all addresses

    If re.Test(s) Then
        Set matches = re.Execute(s)
        For Each m In matches
            If Len(result) > 0 Then result = result & ", "
            result = result & m.Value
        Next m
    End If

    ExtractEmail = result
End Function