Attribute VB_Name = "Module1"
Option Explicit



Public Sub SetDomain()

    ' From http://slipstick.me/1
    Dim currentExplorer As Explorer
    Dim Selection As Selection
    Dim obj As Object
    Dim objProp As Outlook.UserProperty
    Dim objMail As Object
    Dim strDomain
    Dim myOlApp As New Outlook.Application
    Dim myFolder As Outlook.Folder
    Dim myNameSpace As Outlook.NameSpace
    Dim myDomainFolder As Outlook.MAPIFolder
    Dim myOlSel As Outlook.Selection
    Dim olItems As Outlook.Items
    Dim olItem As Outlook.MailItem
    Dim i As Long, j As Long
    Dim strFolder As String
    Dim vAddr As Variant

    Dim contactFolder As Outlook.MAPIFolder
    Set currentExplorer = Application.ActiveExplorer
    Set Selection = currentExplorer.Selection
   
    On Error Resume Next
    
    

  
    
    Set contactFolder = Session.GetDefaultFolder(olFolderContacts)

    For Each obj In Selection
         Set objMail = obj
         
           strDomain = Right(objMail.SenderEmailAddress, Len(objMail.SenderEmailAddress) - InStr(objMail.SenderEmailAddress, "@"))
         Set objProp = objMail.UserProperties.Add("Domain", olText, True)
         objProp.Value = strDomain
        objMail.Save
        

'List the e-mail addresses to move each separated by "|"

    Debug.Print objProp
  If objProp = "gmail.com" Or objProp = "hotmail.com" Or objProp = "yahoo.com" Or objProp = "shaw.ca" Or objProp = "aol.com" Or objProp = "telus.com" Then
    Else
  Set myNameSpace = myOlApp.GetNamespace("MAPI")
  Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
  Set myDomainFolder = myFolder.Folders.Add(objProp)
    
    vAddr = Split(objMail.SenderEmailAddress, "|")
    strFolder = objProp
    Set myOlSel = currentExplorer.Selection
    Set olItems = Session.GetDefaultFolder(olFolderInbox).Items
    For i = olItems.Count To 1 Step -1        'Check each message in reverse order
        Set olItem = olItems(i)
        For j = 0 To UBound(vAddr)        'Compare the sender e-mail address with the items in the list
            If LCase(olItem.SenderEmailAddress) = LCase(vAddr(j)) Then
                'If a match then move the 'Info' subfolder of Inbox
                olItem.Move Session.GetDefaultFolder(olFolderInbox).Folders(strFolder)
                     
            End If
        Next j
    Next i
 
CleanUp:
    Set olItems = Nothing
    Set olItem = Nothing

        
End If
  Next


  

End Sub










