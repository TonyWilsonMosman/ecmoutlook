Attribute VB_Name = "Module1"
Public UserVar(5) As String
Sub InitGlobals()

'Need help? T.Wilson@mosman.nsw.gov.au -  - Tony Wilson - Mosman Council 2018
'for UserVar(1) enter your email
'for UserVar(2) enter the name of the temporary outlook ECM working folder in your email account tree (Inbox/zCi)
'for UserVar(3) enter the name of your outlook team subfolder (shared/personal)
'for UserVar(4) enter the email address for ECM (ECMConnect@domain.com)
'for UserVar(5) enter the email address for IT help desk, for errors (IThelpDesk@domain.com)

    UserVar(1) = "My Outlook Data File(1)ù"
    UserVar(2) = "folder1"
    UserVar(3) = "folder2"
    UserVar(4) = "ECMConnect@domain.com"
    UserVar(5) = "IThelpDesk@domain.com"

End Sub
Sub Default1()

    FirstMove
    Dim QAP As String
    QAP = " "
    ForwardDefault QAP
    
End Sub
Sub Address()

    FirstMove
    Dim QAP As String
    QAP = "#QAP 69"
    ForwardDefault QAP
    
End Sub
Sub Fire()

    FirstMove
    Dim QAP As String
    QAP = "#QAP 71"
    ForwardDefault QAP
    
End Sub
Sub ResidentSub()

    FirstMove
    Dim QAP As String
    QAP = "#QAP 210"
    ForwardDefault QAP
    
End Sub
Sub Job()

    FirstMove
    Dim QAP As String
    QAP = "#QAP 224"
    ForwardDefault QAP
    
End Sub
Sub ParkingIllegal()

    FirstMove
    Dim QAP As String
    QAP = "#QAP 302"
    ForwardDefault QAP
    
End Sub
Sub UserInput()
On Error GoTo Oups
    
    FirstMove
    Dim userValue As String
    Dim QAP As String
    userValue = InputBox("Enter QAP number")
    QAP = "#QAP " + userValue
    ForwardDefault QAP
    Exit Sub
   
Oups:
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With OutMail
        .To = UserVar(5)
        .Subject = UserVar(1) + " UserInput " + "   - Error Occured - Error Number " & Err.Number
        .Body = Err.Description

        MsgBox "Something went wrong. What did you do???!!!!"
        .Send
    End With

    Set OutApp = Nothing: Set OutMail = Nothing
    End
End Sub
Sub FirstMove()
 On Error GoTo Oups

    Call InitGlobals
    
    Dim objOutlook As Outlook.Application
    Dim objNamespace As Outlook.NameSpace
    Dim objSourceFolder As Outlook.MAPIFolder
    Dim objDestFolder As Outlook.MAPIFolder
    Dim objItem As mailItem
    Dim CopiedItem As mailItem
    Dim oAccount As Outlook.Account
    Dim objCouncilFolder As Outlook.MAPIFolder
    
    Set objOutlook = Application
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objSourceFolder = objNamespace.GetDefaultFolder(olFolderDrafts)
    Set objItem = objOutlook.ActiveExplorer.Selection.Item(1)
    Set oAccount = Application.Session.Accounts(UserVar(1))
    Set objDestFolder = objNamespace.Folders(UserVar(1)).Folders(UserVar(2))
    Set objCouncilFolder = objNamespace.Folders(UserVar(1)).Folders(UserVar(3))
    Set CopiedItem = objItem.Copy
    
    CopiedItem.UnRead = False
    CopiedItem.move objCouncilFolder
    objItem.UnRead = False
    objItem.move objDestFolder
    
    Exit Sub
    
Oups:
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With OutMail
        .To = UserVar(5)
        .Subject = UserVar(1) + " First Move " + "   - Error Occured - Error Number " & Err.Number
        .Body = Err.Description

        MsgBox "Something went wrong. What did you do???!!!!"
        .Send
    End With

    Set OutApp = Nothing: Set OutMail = Nothing
    End
End Sub

Sub ForwardDefault(QAP As String)
 On Error GoTo Oups

    Dim objPropertyAccessor As Outlook.PropertyAccessor
    Dim strHeader As String
    Dim objFileSystem As Object
    Dim strTextFile As String
    Dim objTextFile As Object
    Dim objOutlook As Outlook.Application
    Dim objNamespace As Outlook.NameSpace
    Dim objSourceFolder As Outlook.MAPIFolder
    Dim objDestFolder As Outlook.MAPIFolder
    Dim objItem As mailItem
    Dim ECMaddress As String
    Dim objMail As Outlook.mailItem
    Dim strbody As String
    Dim senderaddress As String
    Dim addresstype As Integer
    Dim oAccount As Outlook.Account
    Dim objCouncilFolder As Outlook.MAPIFolder
    Dim objItems As Outlook.Items
    
    Set objOutlook = Application
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objSourceFolder = objNamespace.GetDefaultFolder(olFolderDrafts)
    Set objDestFolder = objNamespace.Folders(UserVar(1)).Folders(UserVar(2))
    Set objCouncilFolder = objNamespace.Folders(UserVar(1)).Folders(UserVar(3))
    
    ECMaddress = UserVar(4)
  
    Do Until objDestFolder.Items.Count = 0
        Set Obj = objDestFolder.Items(1)
        With Obj
            Set objMail = Obj.Forward
            Set objPropertyAccessor = Obj.PropertyAccessor
            strHeader = objPropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
            objMail.To = ECMaddress
            objMail.Subject = Obj.Subject
            strbody = "<font size=""1"" color=""white"">" + "<br>" + "#ECMBODY" + "<br>" + "#SILENT" + "<br>" + QAP + "<br>" + "#NOREG" + "<br>" + strHeader + "</font>"
            objMail.HTMLBody = objMail.HTMLBody + strbody
            objMail.Send
            Obj.move objCouncilFolder
        End With
    Loop
    
    Set objItem = Nothing
    Set objMail = Nothing
    
    Exit Sub
    
Oups:
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With OutMail
        .To = UserVar(5)
        .Subject = UserVar(1) + " ForwardDefault " + "   - Error Occured - Error Number " & Err.Number
        .Body = Err.Description

        MsgBox "Something went wrong. What did you do???!!!!"
        .Send
    End With

    Set OutApp = Nothing: Set OutMail = Nothing
    End
End Sub
