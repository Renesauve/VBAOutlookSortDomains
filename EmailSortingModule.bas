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
    Dim colRules As Outlook.Rules
    Dim oRule As Outlook.Rule
    Dim colRuleActions As Outlook.RuleActions
    Dim oMoveRuleAction As Outlook.MoveOrCopyRuleAction
    Dim oFromCondition As Outlook.ToOrFromRuleCondition
    Dim oExceptSubject As Outlook.TextRuleCondition
    Dim oMoveTarget As Outlook.Folder
    Dim myRecipient As Outlook.Recipient
    Dim xlApp As Excel.Application
    Dim xlSht As Excel.Worksheet
   
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

    
    
  If objProp = "gmail.com" Or objProp = "hotmail.com" Or objProp = "yahoo.com" Or objProp = "shaw.ca" Or objProp = "aol.com" Or objProp = "telus.com" Then
    Else
  Set myNameSpace = myOlApp.GetNamespace("MAPI")
  Set myRecipient = myNameSpace.CreateRecipient("<TARGETED EMAIL HERE>")
  Set myFolder = myNameSpace.GetSharedDefaultFolder(myRecipient, olFolderInbox)
  Set myDomainFolder = myFolder.Folders.Add(objProp)
 
  
    vAddr = Split(objMail.SenderEmailAddress, "|")
    strFolder = objProp
    Set myOlSel = currentExplorer.Selection
    Set olItems = Session.GetSharedDefaultFolder(myRecipient, olFolderInbox).Items
    

    For i = olItems.Count To 1 Step -1
        Set olItem = olItems(i)
        For j = 0 To UBound(vAddr)
            If LCase(olItem.SenderEmailAddress) = LCase(vAddr(j)) Then
                olItem.Move Session.GetSharedDefaultFolder(myRecipient, olFolderInbox).Folders(strFolder)
                
                

        
        Dim RulesNames As Object
        Set RulesNames = colRules.Item((strDomain + " " + "rule"))
        
        
       For Each RulesNames In colRules
        
        If (RulesNames <> (strDomain + " " + "rule")) Then
        Set oMoveTarget = myFolder.Folders(strDomain)
        Set colRules = Application.Session.DefaultStore.GetRules()
        Set oRule = colRules.Create(objProp + " " + "rule", olRuleReceive)

         Set oFromCondition = oRule.Conditions.From
         With oFromCondition
         .Enabled = True
         
         .Recipients.Add (olItem.SenderEmailAddress)
         
         .Recipients.ResolveAll
         
         End With
         
         Set oMoveRuleAction = oRule.Actions.MoveToFolder
         
         With oMoveRuleAction
         
         .Enabled = True
         
         .Folder = oMoveTarget
         
         End With
         
            Debug.Print "Yay"
        colRules.Save
         Else
            Debug.Print "Failed"

                 End If
             Next
                 
           
            End If
        Next j
    Next i
 
CleanUp:
    Set olItems = Nothing
    Set olItem = Nothing