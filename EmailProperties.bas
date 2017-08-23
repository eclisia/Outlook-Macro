Attribute VB_Name = "EmailProperties"
'***Explication'
'With VBA you can access to the e-mail properties by different way.
'We have choose one method (see the code)
'But Outlook doesn't manage the clipboard function of Windows.
'So you have to add a textbox to copy/paste your data

Public Sub getReceivedEmailProperties()
'Definition of the object
    Dim oMail As MailItem
    Dim myFolder As MAPIFolder
    Dim Expl As Outlook.Explorer
    Dim myOlApp As Outlook.Application
    Dim myNameSpace As NameSpace
    Dim sel As Selection
    
    

'Initialisation des object
    Set myOlApp = Outlook.Application
    Set myNameSpace = myOlApp.GetNamespace("MAPI")
    Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
    Set Expl = Outlook.ActiveExplorer
    Set sel = Expl.Selection


'Loop for each item (mail) of the selection
    For Each oMail In sel
        Dim message As Variant
'Formating of the string
        message = "Cf. e-mail : Subject: " & oMail.Subject & " || From: " & oMail.SenderName & "(" & oMail.SenderEmailAddress & ") || Received : " & oMail.ReceivedTime
        Debug.Print message
' Outlook + VBA doesn't manage the clipboard functions of from Windows.
' Tips and Tricks : create a Form with a textbox.
' Fill the textbox with the data you want to copy/past.
        UserForm1.TextBox1.Text = message
        UserForm1.TextBox1.SetFocus
        UserForm1.TextBox1.SelStart = 0
        UserForm1.TextBox1.SelLength = Len(UserForm1.TextBox1.Text)
        UserForm1.Show
    Next oMail

End Sub

Public Sub getOpenedAndActive_ReceivedEmailProperties()
On Error GoTo GestionErreur
    'Variable Declaration
    Dim myInspector As Outlook.Inspector
    Dim myItem As Outlook.MailItem
    Dim message As String
    
    'Object instanciation
    Set myInspector = Application.ActiveInspector
    
    If Not TypeName(myInspector) = "Nothing" Then
        If TypeName(myInspector.CurrentItem) = "MailItem" Then
            Set myItem = myInspector.CurrentItem    'Get the actual e-mail

            
            message = "Cf. e-mail : Subject: " & myItem.Subject & " || From: " & myItem.SenderEmailAddress & " || Received : " & myItem.ReceivedTime
            Debug.Print message
            
            'Call a display information
            UserForm1.TextBox1.Text = message
            UserForm1.BackColor = RGB(200, 200, 2)
            UserForm1.TextBox1.SetFocus
            UserForm1.TextBox1.SelStart = 0
            UserForm1.TextBox1.SelLength = Len(UserForm1.TextBox1.Text) 'Select all string
            UserForm1.Show
            
        End If
    End If
    
GestionErreur:
    If Err.Number <> 0 Then
        Debug.Print Err.Number & " _ " & Err.Description
        MsgBox "/!\ THIS sub works only on a opened e-mail !", vbCritical, " Erreur"
        Resume
    End If
    Debug.Print Err.Number & " _ " & Err.Description


End Sub

Public Sub getOpenedAndActive_SendEmailProperties()
On Error GoTo GestionErreur
    'Variable Declaration
    Dim myInspector As Outlook.Inspector
    Dim myItem As Outlook.MailItem
    Dim message As String
    
    'Object instanciation
    Set myInspector = Application.ActiveInspector
    
    If Not TypeName(myInspector) = "Nothing" Then
        If TypeName(myInspector.CurrentItem) = "MailItem" Then
            Set myItem = myInspector.CurrentItem    'Get the actual e-mail

'message = "Cf. e-mail : Subject: " & oMail.Subject & " || To: " & oMail.To & " || Send : " & oMail.SentOn
            
            message = "Cf. e-mail : Subject: " & myItem.Subject & " || To: " & myItem.To & " || Send : " & myItem.SentOn
            Debug.Print message
            
            'Call a display information
            UserForm1.TextBox1.Text = message
            UserForm1.BackColor = RGB(200, 2, 2)
            UserForm1.TextBox1.SetFocus
            UserForm1.TextBox1.SelStart = 0
            UserForm1.TextBox1.SelLength = Len(UserForm1.TextBox1.Text) 'Select all string
            UserForm1.Show
            
        End If
    End If
    
GestionErreur:
    If Err.Number <> 0 Then
        Debug.Print Err.Number & " _ " & Err.Description
        MsgBox "/!\ THIS sub works only on a opened e-mail !", vbCritical, " Erreur"
        Resume
    End If
    Debug.Print Err.Number & " _ " & Err.Description


End Sub

Public Sub getSendEmailProperties()
'Definition of the object
    Dim oMail As MailItem
    Dim myFolder As MAPIFolder
    Dim Expl As Outlook.Explorer
    Dim myOlApp As Outlook.Application
    Dim myNameSpace As NameSpace
    Dim sel As Selection
    
    

'Initialisation des object
    Set myOlApp = Outlook.Application
    Set myNameSpace = myOlApp.GetNamespace("MAPI")
    Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
    Set Expl = Outlook.ActiveExplorer
    Set sel = Expl.Selection


'Loop for each item (mail) of the selection
    For Each oMail In sel
        Dim message As Variant
'Formating of the string
        message = "Cf. e-mail : Subject: " & oMail.Subject & " || To: " & oMail.To & " || Send : " & oMail.SentOn
        Debug.Print message
' Outlook + VBA doesn't manage the clipboard functions of from Windows.
' Tips and Tricks : create a Form with a textbox.
' Fill the textbox with the data you want to copy/past.
        UserForm1.TextBox1.Text = message
        UserForm1.TextBox1.SetFocus
        UserForm1.TextBox1.SelStart = 0
        UserForm1.TextBox1.SelLength = Len(UserForm1.TextBox1.Text)
        UserForm1.Show
    Next oMail

End Sub


Public Sub getMailFolder()
' Version 02 :
' Modification date : 23/08/2017
' adding of code to check if more than one mail is selected.
' If yes, so the sub is aborted. Else, the sub run.


'Definition of the object
    Dim oMail As MailItem
    Dim myFolder As MAPIFolder
    Dim Expl As Outlook.Explorer
    Dim myOlApp As Outlook.Application
    Dim myNameSpace As NameSpace
    Dim sel As Selection
    Dim myObject As Object
    
    

'Initialisation des object
    Set myOlApp = Outlook.Application
    Set myNameSpace = myOlApp.GetNamespace("MAPI")
    Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
    Set Expl = Outlook.ActiveExplorer
    Set sel = Expl.Selection
    Set myObject = Application.ActiveWindow
    
    'Check if more than one item is selected.
    If sel.Count > 1 Then
        Debug.Print "More than one e-mail selected, so the sub is aborted"
        Exit Sub
    End If


'Loop for each item (mail) of the selection
    For Each oMail In sel
        Dim message As Variant
'Formating of the string
        Set myObject = myObject.Selection(1)
        Set myFolder = myObject.Parent
        message = "Répertoire du message sélectionné :  " & myFolder.FolderPath
        Debug.Print message
' Outlook + VBA doesn't manage the clipboard functions of from Windows.
' Tips and Tricks : create a Form with a textbox.
' Fill the textbox with the data you want to copy/past.
        UserForm2.caption = "REPERTOIRE DU MESSAGE SELECTIONNE"
        UserForm2.TextBox1.Text = message
        UserForm2.TextBox1.SetFocus
        UserForm2.TextBox1.SelStart = 0
        UserForm2.TextBox1.SelLength = Len(UserForm2.TextBox1.Text)
        UserForm2.Show
        
        UserForm2.CommandButton1.SetFocus
        
    Next oMail

End Sub



Public Sub getOpenedAndActiveEMAILFolder()
'Sub which permits to get the Path of the opened item (i.e. opened e-mail)
'/!\ THIS sub works only on a opened e-mail !

On Error GoTo GestionErreur
    'Variable Declaration
    Dim myInspector As Outlook.Inspector
    Dim myItem As Outlook.MailItem
    Dim strUserFormTitle As String
    Dim message As String
    
    'Object instanciation
    Set myInspector = Application.ActiveInspector
    
    If Not TypeName(myInspector) = "Nothing" Then
        If TypeName(myInspector.CurrentItem) = "MailItem" Then
            Set myItem = myInspector.CurrentItem    'Get the actual e-mail

            
            message = myItem.Parent.FolderPath     'Get the FolderPath of the e-mail
            
            strUserFormTitle = "Répertoire de l'e-mail courant"
            Debug.Print "Répertoire du message sélectionné :  " & message
            DisplayInformationUserForm2 strUserFormTitle, message, RGB(200, 200, 255)   'Call Display function
            
        
        End If
    
    End If
    
GestionErreur:
    If Err.Number <> 0 Then
        Debug.Print Err.Number & " _ " & Err.Description
        MsgBox "/!\ THIS sub works only on a opened e-mail !", vbCritical, " Erreur"
        Resume
    End If
    Debug.Print Err.Number & " _ " & Err.Description


 



End Sub

Public Sub SaveAttachment()
'Sub which permits to get the filename for the attachment of the opened item (i.e. opened e-mail)
'/!\ THIS sub works only on a opened e-mail !

On Error GoTo GestionErreur
    'Variable Declaration
    Dim myInspector As Outlook.Inspector
    Dim myItem As Outlook.MailItem
    Dim myAttachments As Outlook.Attachments
    Dim myAtt As Outlook.Attachment
    Dim strUserFormTitle As String
    
    'Object instanciation
    Set myInspector = Application.ActiveInspector
    
    If Not TypeName(myInspector) = "Nothing" Then
        If TypeName(myInspector.CurrentItem) = "MailItem" Then
            Set myItem = myInspector.CurrentItem
            Set myAttachments = myItem.Attachments
            
            strUserFormTitle = "Nom de la pièce-jointe de l'e-mail ouvert"
            
            'Call the display function through a for-each loop structure
            'to manage all the attachements
            For Each myAtt In myItem.Attachments
                Debug.Print myAtt.DisplayName
                DisplayInformationUserForm2 strUserFormTitle, myAtt.DisplayName, RGB(0, 255, 0)
            Next myAtt
        
        End If
    
    End If
    
GestionErreur:
    If Err.Number <> 0 Then
        Debug.Print Err.Number & " _ " & Err.Description
        MsgBox "/!\ THIS sub works only on a opened e-mail !", vbCritical, " Erreur"
        Resume
    End If
    Debug.Print Err.Number & " _ " & Err.Description


 
End Sub

Private Sub DisplayInformationUserForm2(caption As String, message As String, color As Long)
'The Sub permits to call the UserForm2, modify its look then display information inside.
    
    UserForm2.BackColor = color
    UserForm2.caption = caption
    UserForm2.TextBox1.Text = message
    UserForm2.TextBox1.SetFocus
    UserForm2.TextBox1.SelStart = 0
    UserForm2.TextBox1.SelLength = Len(UserForm2.TextBox1.Text)
    UserForm2.Show
    
    UserForm2.CommandButton1.SetFocus
End Sub


