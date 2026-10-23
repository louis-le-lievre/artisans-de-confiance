VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dispoarti 
   Caption         =   "Disponibilité Artisan"
   ClientHeight    =   3948
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7920
   OleObjectBlob   =   "Dispoarti.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Dispoarti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ' Chargement initial : Remplir TOUT au démarrage du formulaire
    ' Pas besoin de sélection manuelle, tout se fait automatiquement
   
    ' Remplir CB_etatarti avec des exemples (adaptez les options)
    CB_etatarti.Clear
    CB_etatarti.AddItem "Actif"
    CB_etatarti.AddItem "Inactif"
    CB_etatarti.AddItem "Tous"
   
    ' Sélectionner "Tous" par défaut dans CB_etatarti
    CB_etatarti.ListIndex = 2  ' Index 2 = "Tous"
   
    ' Remplir directement CB_etat avec VRAI, FAUX, TOUT
    CB_etat.Clear
    CB_etat.AddItem "VRAI"
    CB_etat.AddItem "FAUX"
    CB_etat.AddItem "TOUT"
   
    ' Sélectionner "TOUT" par défaut dans CB_etat
    CB_etat.ListIndex = 2  ' Index 2 = "TOUT"
   
    ' Optionnel : Message de confirmation (à supprimer si inutile)
    ' MsgBox "Formulaire chargé avec options par défaut."
End Sub

Private Sub CB_etatarti_Change()
    ' Déclencheur optionnel : Si changement dans CB_etatarti, rafraîchir CB_etat
    ' (Mais comme tout est déjà rempli au chargement, c'est redondant sauf si besoin)
    If CB_etatarti.Value <> "" Then
        CB_etat.Clear
        CB_etat.AddItem "VRAI"
        CB_etat.AddItem "FAUX"
        CB_etat.AddItem "TOUT"
        CB_etat.ListIndex = 2  ' "TOUT" par défaut
    End If
End Sub

Private Sub CB_etat_Change()
    ' Gestion du changement dans CB_etat (exemple : affichage)
    If CB_etat.Value <> "" Then
        MsgBox "État sélectionné : " & CB_etat.Value
    End If
End Sub
