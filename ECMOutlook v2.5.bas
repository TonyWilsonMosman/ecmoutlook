Attribute VB_Name = "Module1"
Public UserVar(5) As String
Sub InitGlobals()
'=============================
'Under UserVar(1) enter your email account (full email address)
UserVar(1) = "My Outlook Data File(1)ù"
'===============================


'Need help? T.Wilson@mosman.nsw.gov.au -  - Tony Wilson - Mosman Council 2018
'UserVar(2) create a temporary/working folder under your personal email account
'UserVar(3) create a folder under your personal email account where all the emails sent for regostration will be saved for QC (in duplicate)
'for UserVar(4) enter the ECM email address (ECMConnect@domain.com)
'for UserVar(5) enter the email address for IT help desk, for errors (IThelpDesk@domain.com)
    UserVar(2) = "zTemp"
    UserVar(3) = "zCi"
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
    Dim objzTempFolder As Outlook.MAPIFolder
    Dim objzCiFolder As Outlook.MAPIFolder
    Dim oAccount As Outlook.Account
    
    Dim objItem As mailItem
    Dim copiedItem As mailItem
    
    Set objOutlook = Application
    Set oAccount = Application.Session.Accounts(UserVar(1))
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objItem = objOutlook.ActiveExplorer.Selection.Item(1)
    Set copiedItem = objItem.Copy
        
    Set objzTempFolder = objNamespace.Folders(UserVar(1)).Folders(UserVar(2))
    Set objzCiFolder = objNamespace.Folders(UserVar(1)).Folders(UserVar(3))
    
    copiedItem.UnRead = False
    copiedItem.move objzCiFolder
    objItem.UnRead = False
    objItem.move objzTempFolder
    
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
    Dim objOutlook As Outlook.Application
    Dim objItems As Outlook.Items
    Dim objItem As mailItem
    Dim objMail As Outlook.mailItem
    Dim objNamespace As Outlook.NameSpace
    Dim oAccount As Outlook.Account

    Dim strHeader As String
    Dim strbody As String
    
    Dim objzTempFolder As Outlook.MAPIFolder
    Dim objzCiFolder As Outlook.MAPIFolder
    
    Set objOutlook = Application
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objzTempFolder = objNamespace.Folders(UserVar(1)).Folders(UserVar(2))
    Set objzCiFolder = objNamespace.Folders(UserVar(1)).Folders(UserVar(3))

  
    Do Until objzTempFolder.Items.Count = 0
        Set Obj = objzTempFolder.Items(1)
        With Obj
            Set objMail = Obj.Forward
            Set objPropertyAccessor = Obj.PropertyAccessor
            strHeader = objPropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
            objMail.To = UserVar(4)
            objMail.Subject = Obj.Subject
            strbody = "<font size=""1"" color=""white"">" + "<br>" + "#ECMBODY" + "<br>" + "#SILENT" + "<br>" + QAP + "<br>" + "#NOREG" + "<br>" + strHeader + "</font>"
            objMail.HTMLBody = objMail.HTMLBody + strbody
            objMail.Send
            Obj.move objzCiFolder
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
