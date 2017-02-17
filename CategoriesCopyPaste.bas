Attribute VB_Name = "CategoriesCopyPaste"
'Ensemble de Macro qui permet de dupliquer les cat�gories d'un e-mail pour les appliquer � un autre.
'Code librement et largement inspir� de :
' http://www.developpez.net/forums/d739494/logiciels/microsoft-office/outlook/vba-outlook/affecter-categorie-masse-aux-messages-dossier-outlook-2007-a/

'D�claration d'une variable static
    Public categorieTempo As Variant



'Fonction de copie
Sub CopierCat�gorie()
 
'D�claration des Objets et variables
Dim MonApply As Outlook.Application
Dim Expl As Explorer
Dim myNameSpace As NameSpace
Dim myFolder As MAPIFolder
Dim sel As Selection
Dim myItems As Items
Dim xi As Integer
 
'Instance des Objets
        Set MonApply = Outlook.Application 'Application Outlook
 
        'Expl nous donne le dossier courant
        Set Expl = ActiveExplorer
        'Permet d'acc�der � toutes les donn�es Outlook qui y sont stock�es
        Set myNameSpace = MonApply.GetNamespace("MAPI")
        Set sel = Expl.Selection
        
        
        'Loop for each item (mail) of the selection and copy the categories data into a public variant
    For Each oMail In sel
        categorieTempo = oMail.Categories
    Next oMail

 
End Sub


'Fonction qui permet de coller la valeure
Sub CollerCat�gorie()
 
'D�claration des Objets et variables
Dim MonApply As Outlook.Application
Dim Expl As Explorer
Dim myNameSpace As NameSpace
Dim myFolder As MAPIFolder
Dim sel As Selection
Dim myItems As Items
Dim xi As Integer
 
'Instance des Objets
        Set MonApply = Outlook.Application 'Application Outlook
 
        'Expl nous donne le dossier courant
        Set Expl = ActiveExplorer
        'Permet d'acc�der � toutes les donn�es Outlook qui y sont stock�es
        Set myNameSpace = MonApply.GetNamespace("MAPI")
        Set sel = Expl.Selection
        
        
        'Loop for each item (mail) of the selection get the public variant and paste it into the item categories
    For Each oMail In sel
        oMail.Categories = categorieTempo
        oMail.Save
    Next oMail
 

 
End Sub
