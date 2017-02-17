Attribute VB_Name = "CategoriesCopyPaste"
'Ensemble de Macro qui permet de dupliquer les catégories d'un e-mail pour les appliquer à un autre.
'Code librement et largement inspiré de :
' http://www.developpez.net/forums/d739494/logiciels/microsoft-office/outlook/vba-outlook/affecter-categorie-masse-aux-messages-dossier-outlook-2007-a/

'Déclaration d'une variable static
    Public categorieTempo As Variant



'Fonction de copie
Sub CopierCatégorie()
 
'Déclaration des Objets et variables
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
        'Permet d'accéder à toutes les données Outlook qui y sont stockées
        Set myNameSpace = MonApply.GetNamespace("MAPI")
        Set sel = Expl.Selection
        
        
        'Loop for each item (mail) of the selection and copy the categories data into a public variant
    For Each oMail In sel
        categorieTempo = oMail.Categories
    Next oMail

 
End Sub


'Fonction qui permet de coller la valeure
Sub CollerCatégorie()
 
'Déclaration des Objets et variables
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
        'Permet d'accéder à toutes les données Outlook qui y sont stockées
        Set myNameSpace = MonApply.GetNamespace("MAPI")
        Set sel = Expl.Selection
        
        
        'Loop for each item (mail) of the selection get the public variant and paste it into the item categories
    For Each oMail In sel
        oMail.Categories = categorieTempo
        oMail.Save
    Next oMail
 

 
End Sub
