Public UserVar(2) As String

Sub InitGlobals()
'for UserVar(1) enter the ECM email address (ECMConnect@domain.com)
'for UserVar(2) enter the email address for IT help desk, for errors (IThelpDesk@domain.com)
' also change the name of the folder to your shared account folder at "ozCiFolder", 50 lines bellow

    UserVar(1) = "ECMConnect@domain.com"
    UserVar(2) = "IThelpDesk@domain.com"
    
'Need help? T.Wilson@mosman.nsw.gov.au -  - Tony Wilson - Mosman Council 2018
End Sub

Sub FirstMove()
 On Error GoTo Oups

    Call InitGlobals
    
    Dim oOutlook As Outlook.Application
    Dim oNameSpace As Outlook.NameSpace
    Dim oAccount As Outlook.account
    
    Dim myFolder As Outlook.MAPIFolder
    Dim ozTempFolder As Outlook.MAPIFolder
    Dim ozCiFolder As Outlook.MAPIFolder
    Dim CurrentFolder As Outlook.MAPIFolder
    
    Dim oItem As mailItem
    Dim copiedItem As mailItem
    
    Set oOutlook = Application
    Set oNameSpace = oOutlook.GetNamespace("MAPI")
    Set oAccount = Application.Session.accounts(1)

    Set myFolder = oNameSpace.GetDefaultFolder(olFolderInbox)
    Set CurrentFolder = Application.ActiveExplorer.CurrentFolder
    
    On Error Resume Next
    Set ozTempFolder = myFolder.Folders.Add("zEcmTemp")
    Set ozTempFolder = myFolder.Folders("zEcmTemp")
    
    '================'
                '================'
                        '================'
                                '================'
    'End Folder, change to the name of the shared folder'
    'The example here is //Share account name/Inbox/staff sub folder'
    Set ozCiFolder = oNameSpace.Folders("Share account name").Folders("Inbox").Folders("staff sub folder")
                                '================'
                        '================'
            '================'
    '================'
        
        
    Set oItem = oOutlook.ActiveExplorer.Selection.Item(1)
    Set copiedItem = oItem.Copy
    
    copiedItem.UnRead = False
    If CurrentFolder <> ozCiFolder Then
    copiedItem.move ozCiFolder
    End If
    
    oItem.UnRead = False
    oItem.move ozTempFolder
    
    Exit Sub
    
Oups:
    Set outApp = CreateObject("Outlook.Application")
    Set OutMail = outApp.CreateItem(0)

    With OutMail
        .To = UserVar(2)
        .Subject = (CStr(oAccount)) + " First Move " + "   - Error Occured - Error Number " & Err.Number
        .Body = Err.Description

        MsgBox "Something went wrong. What did you do???!!!!"
        .Send
    End With

    Set outApp = Nothing: Set OutMail = Nothing
    End
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

Sub CourtAttend()

    FirstMove
    Dim QAP As String
    QAP = "#QAP 200"
    ForwardDefault QAP
    
End Sub
Sub InfringMoreInfo()

    FirstMove
    Dim QAP As String
    QAP = "#QAP 329"
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
    Set outApp = CreateObject("Outlook.Application")
    Set OutMail = outApp.CreateItem(0)

    With OutMail
        .To = UserVar(2)
        .Subject = (CStr(oAccount)) + " UserInput " + "   - Error Occured - Error Number " & Err.Number
        .Body = Err.Description

        MsgBox "Something went wrong. What did you do???!!!!"
        .Send
    End With

    Set outApp = Nothing: Set OutMail = Nothing
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
    Dim oAccount As Outlook.account
    Dim myFolder As Outlook.MAPIFolder

    Dim strHeader As String
    Dim strbody As String
    
    Dim objzTempFolder As Outlook.MAPIFolder
    Dim objDELFolder As Outlook.MAPIFolder
        
    Set objOutlook = Application
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set oAccount = Application.Session.accounts(1)
    Set myFolder = objNamespace.GetDefaultFolder(olFolderInbox)
    
    Set objzTempFolder = myFolder.Folders("zEcmTemp")
    Set objDELFolder = objNamespace.Folders(CStr(oAccount)).Folders("Deleted Items")
  

    Do Until objzTempFolder.Items.Count = 0
        Set Obj = objzTempFolder.Items(1)
        With Obj
            Set objMail = Obj.Forward
            Set objPropertyAccessor = Obj.PropertyAccessor
            strHeader = objPropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
            objMail.To = UserVar(1)
            objMail.Subject = Obj.Subject
            strbody = "<font size=""1"" color=""white"">" + "<br>" + "#ECMBODY" + "<br>" + "#SILENT" + "<br>" + QAP + "<br>" + "#NOREG" + "<br>" + strHeader + "</font>"
            objMail.HTMLBody = objMail.HTMLBody + strbody
            
            objMail.SendUsingAccount = objNamespace.accounts.Item(1)
            objMail.Send
            Obj.move objDELFolder
        End With
    Loop
    
    Set objItem = Nothing
    Set objMail = Nothing
    
    Exit Sub
    
Oups:
    Set outApp = CreateObject("Outlook.Application")
    Set OutMail = outApp.CreateItem(0)

    With OutMail
        .To = UserVar(2)
        .Subject = (CStr(oAccount)) + " ForwardDefault " + "   - Error Occured - Error Number " & Err.Number
        .Body = Err.Description

        MsgBox "Something went wrong. What did you do???!!!!"
        .Send
    End With

    Set outApp = Nothing: Set OutMail = Nothing
    End
End Sub