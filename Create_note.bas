Option Explicit

Private Const TRILIUM_API_URL As String = "URL_Trillium/etapi"
Private Const API_TOKEN As String = "Your_token"
Private WithEvents Items As Outlook.Items

Private Sub Application_Startup()
    Dim Ns As Outlook.NameSpace
    Set Ns = Application.GetNamespace("MAPI")
    Set Items = Ns.GetDefaultFolder(olFolderInbox).Items
End Sub

Private Sub Items_ItemChange(ByVal Item As Object)
    If TypeOf Item Is Outlook.MailItem Then
        Dim Mail As Outlook.MailItem
        Set Mail = Item

        If Mail.FlagStatus = olFlagMarked Then
            CreateTriliumNote Mail.Subject, Mail.Body
        End If
    End If
End Sub

Private Sub CreateTriliumNote(ByVal Title As String, ByVal Content As String)
    Dim Http As Object
    Set Http = CreateObject("MSXML2.XMLHTTP")

    Dim Url As String
    Url = TRILIUM_API_URL & "/create-note"

    Dim JsonBody As String
    JsonBody = "{""title"":""" & EscapeJson(Title) & """,""type"":""text"",""content"":""" & EscapeJson(Content) & """,""parentNoteId"":""root""}"

    Http.Open "POST", Url, False
    Http.setRequestHeader "Content-Type", "application/json"
    Http.setRequestHeader "Authorization", API_TOKEN
    Http.Send JsonBody

    If Http.Status <> 201 Then
        MsgBox "Error creating note: " & Http.Status & " - " & Http.responseText, vbCritical
    End If
End Sub


Private Function EscapeJson(ByVal Txt As String) As String
    Txt = Replace(Txt, "\", "\\")
    Txt = Replace(Txt, """", "\""")
    Txt = Replace(Txt, vbCrLf, "\n")
    Txt = Replace(Txt, vbLf, "\n")
    EscapeJson = Txt
End Function
