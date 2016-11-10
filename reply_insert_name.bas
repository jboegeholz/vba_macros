Sub InsertNameInReply()

    Dim Msg As Outlook.MailItem
    Dim MsgReply As Outlook.MailItem
    Dim strGreetName As String
    Dim lGreetType As Long

     ' set reference to open/selected mail item
    On Error Resume Next
    Select Case TypeName(Application.ActiveWindow)
    Case "Explorer"
        Set Msg = ActiveExplorer.Selection.Item(1)
    Case "Inspector"
        Set Msg = ActiveInspector.CurrentItem
    Case Else
    End Select
    On Error GoTo 0

    If Msg Is Nothing Then GoTo ExitProc

    lGreetType = 1

    iPos = Len(Msg.SenderName)
    iPos = iPos - InStr(1, Msg.SenderName, " ")
    strGreetName = Right$(Msg.SenderName, iPos)
    strGreetName = Replace(strGreetName, "(EXT)", "")
    strGreetName = Trim(strGreetName)

    Set MsgReply = Msg.Reply

    With MsgReply
        .HTMLBody = "Hallo " & strGreetName & "," & .HTMLBody
        .Display
    End With

ExitProc:
    Set Msg = Nothing
    Set MsgReply = Nothing
End Sub
