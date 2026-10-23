VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_GestionClient 
   Caption         =   "UserForm1"
   ClientHeight    =   8904.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11556
   OleObjectBlob   =   "UserForm_GestionClient.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_GestionClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Call ChargerClients("")
    
    ' Configuration de la ListBox des commandes avec 5 colonnes
    With Me.lstCommandes
        .Clear
        .ColumnCount = 5
        .ColumnWidths = "80;90;80;100;120"  ' Ajustez les largeurs si besoin
    End With
    
    ' Chargement des régions dans la ComboBox
    With Me.cboRegion
        .Clear
        .AddItem "Auvergne-Rhône-Alpes"
        .AddItem "Bourgogne-Franche-Comté"
        .AddItem "Bretagne"
        .AddItem "Centre-Val de Loire"
        .AddItem "Corse"
        .AddItem "Grand Est"
        .AddItem "Guadeloupe"
        .AddItem "Guyane"
        .AddItem "Hauts-de-France"
        .AddItem "Île-de-France"
        .AddItem "Martinique"
        .AddItem "Mayotte"
        .AddItem "Normandie"
        .AddItem "Nouvelle-Aquitaine"
        .AddItem "Occitanie"
        .AddItem "Pays de la Loire"
        .AddItem "Provence-Alpes-Côte d'Azur"
        .AddItem "La Réunion"
        .Style = fmStyleDropDownList
        .ListWidth = 220
    End With
End Sub

Private Sub lstClients_Click()
    Dim code As String
    Dim ws As Worksheet
    
    If Me.lstClients.ListIndex = -1 Then Exit Sub
    
    code = Me.lstClients.List(Me.lstClients.ListIndex, 1)  ' id_client
    
    ' Chargement des infos client
    Set ws = Worksheets("CLIENT")
    On Error Resume Next
    Me.txtNom.Text = Application.WorksheetFunction.VLookup(code, ws.Range("A2:I1000"), 2, False)
    Me.txtPrenom.Text = Application.WorksheetFunction.VLookup(code, ws.Range("A2:I1000"), 3, False)
    Me.txtAdresse.Text = Application.WorksheetFunction.VLookup(code, ws.Range("A2:I1000"), 4, False)
    Me.txtVille.Text = Application.WorksheetFunction.VLookup(code, ws.Range("A2:I1000"), 5, False)
    Me.txtTel.Text = Application.WorksheetFunction.VLookup(code, ws.Range("A2:I1000"), 6, False)
    Me.txtEmail.Text = Application.WorksheetFunction.VLookup(code, ws.Range("A2:I1000"), 7, False)
    Me.cboRegion.Value = Application.WorksheetFunction.VLookup(code, ws.Range("A2:I1000"), 9, False)
    On Error GoTo 0
    
    ' Mise à jour de l'historique des commandes avec le nom de l'artisan
    Call ChargerCommandes(code)
End Sub

Sub ChargerCommandes(codeClient As String)
    Dim wsC As Worksheet, wsS As Worksheet, wsA As Worksheet
    Dim lastRow As Long, i As Long
    Dim statutID As String, statutLib As String
    Dim artisanID As String, artisanNom As String
    Dim dateCreation As String
    
    Set wsC = Worksheets("COMMANDE")          ' Feuille commandes
    Set wsS = Worksheets("STATUT_COMMANDE")   ' Feuille statuts
    Set wsA = Worksheets("ARTISAN")           ' Feuille artisans
    
    Me.lstCommandes.Clear
    
    ' En-têtes de la ListBox
    Me.lstCommandes.AddItem
    Me.lstCommandes.List(0, 0) = "Commande"
    Me.lstCommandes.List(0, 1) = "Date création"
    Me.lstCommandes.List(0, 2) = "Montant (€)"
    Me.lstCommandes.List(0, 3) = "Statut"
    Me.lstCommandes.List(0, 4) = "Artisan"
    
    lastRow = wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        If Trim(wsC.Cells(i, 8).Value) = codeClient Then  ' Colonne H = id_client
            
            ' --- ID et nom de l'artisan ---
            artisanID = Trim(wsC.Cells(i, 9).Value)  ' Colonne I = id_artisan
            On Error Resume Next
            artisanNom = Application.WorksheetFunction.VLookup(artisanID, wsA.Range("A:B"), 2, False)
            If Err.Number <> 0 Then artisanNom = "Non assigné"
            On Error GoTo 0
            
            ' --- Statut ---
            statutID = Trim(wsC.Cells(i, 13).Value)  ' Colonne M = id_statut_commande
            On Error Resume Next
            statutLib = Application.WorksheetFunction.VLookup(statutID, wsS.Range("A:B"), 2, False)
            If Err.Number <> 0 Then statutLib = "Inconnu"
            On Error GoTo 0
            
            ' --- Date création ---
            If IsNumeric(wsC.Cells(i, 2).Value) And wsC.Cells(i, 2).Value > 0 Then
                dateCreation = Format(DateSerial(1900, 1, 1) + wsC.Cells(i, 2).Value - 2, "dd/mm/yyyy")
            Else
                dateCreation = "Non définie"
            End If
            
            ' --- Ajout de la ligne dans la ListBox ---
            Me.lstCommandes.AddItem
            Me.lstCommandes.List(Me.lstCommandes.ListCount - 1, 0) = wsC.Cells(i, 1).Value                    ' ID Commande
            Me.lstCommandes.List(Me.lstCommandes.ListCount - 1, 1) = dateCreation                             ' Date
            Me.lstCommandes.List(Me.lstCommandes.ListCount - 1, 2) = Format(wsC.Cells(i, 7).Value, "#,##0 €") ' Montant
            Me.lstCommandes.List(Me.lstCommandes.ListCount - 1, 3) = statutLib                                 ' Statut
            Me.lstCommandes.List(Me.lstCommandes.ListCount - 1, 4) = artisanNom                                ' Nom Artisan
        End If
    Next i
    
    ' Message si aucune commande
    If Me.lstCommandes.ListCount = 1 Then
        Me.lstCommandes.Clear
        Me.lstCommandes.AddItem "Aucune commande trouvée pour ce client"
    End If
End Sub

' === Les autres procédures restent inchangées ===

Private Sub btnModifier_Click()
    Dim code As String
    Dim ws As Worksheet
    Dim f As Range
    Dim rep As Integer
    
    If Me.lstClients.ListIndex = -1 Then
        MsgBox "Veuillez sélectionner un client dans la liste.", vbExclamation
        Exit Sub
    End If
    
    code = Me.lstClients.List(Me.lstClients.ListIndex, 1)
    Set ws = Worksheets("CLIENT")
    Set f = ws.Columns(1).Find(code, LookAt:=xlWhole)
   
    If f Is Nothing Then
        MsgBox "Client introuvable.", vbCritical
        Exit Sub
    End If
    
    rep = MsgBox("Confirmer la modification des informations du client ?", vbYesNo + vbQuestion)
    If rep = vbNo Then Exit Sub
    
    f.Offset(0, 1).Value = Me.txtNom.Text
    f.Offset(0, 2).Value = Me.txtPrenom.Text
    f.Offset(0, 3).Value = Me.txtAdresse.Text
    f.Offset(0, 4).Value = Me.txtVille.Text
    f.Offset(0, 5).Value = Me.txtTel.Text
    f.Offset(0, 6).Value = Me.txtEmail.Text
    f.Offset(0, 8).Value = Me.cboRegion.Value
    
    Call ChargerClients("")
    MsgBox "Client modifié avec succès !", vbInformation
End Sub

Private Sub btnRecherche_Click()
    Call ChargerClients(Me.txtRecherche.Text)
End Sub

Sub ChargerClients(critere As String)
    Dim ws As Worksheet, lastRow As Long, i As Long
    Set ws = Worksheets("CLIENT")
    Me.lstClients.Clear
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        If critere = "" Or InStr(1, ws.Cells(i, 2).Value & " " & ws.Cells(i, 3).Value, critere, vbTextCompare) > 0 Then
            Me.lstClients.AddItem ws.Cells(i, 2).Value & " " & ws.Cells(i, 3).Value
            Me.lstClients.List(Me.lstClients.ListCount - 1, 1) = ws.Cells(i, 1).Value
        End If
    Next i
End Sub

Private Sub btnNouveau_Click()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim code As String
    Dim rep As Integer
    
    Set ws = Worksheets("CLIENT")
    
    If Me.txtNom.Text = "" Or Me.txtPrenom.Text = "" Or Me.txtAdresse.Text = "" Or _
       Me.txtVille.Text = "" Or Me.txtTel.Text = "" Or Me.cboRegion.Value = "" Then
        MsgBox "Veuillez remplir tous les champs obligatoires (y compris la région).", vbExclamation
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    code = "CLI" & Format(lastRow, "000")
    
    rep = MsgBox("Ajouter ce nouveau client ?", vbYesNo + vbQuestion)
    If rep = vbNo Then Exit Sub
    
    lastRow = lastRow + 1
    
    ws.Cells(lastRow, 1).Value = code
    ws.Cells(lastRow, 2).Value = Me.txtNom.Text
    ws.Cells(lastRow, 3).Value = Me.txtPrenom.Text
    ws.Cells(lastRow, 4).Value = Me.txtAdresse.Text
    ws.Cells(lastRow, 5).Value = Me.txtVille.Text
    ws.Cells(lastRow, 6).Value = Me.txtTel.Text
    ws.Cells(lastRow, 7).Value = Me.txtEmail.Text
    ws.Cells(lastRow, 9).Value = Me.cboRegion.Value
    
    Call ChargerClients("")
    
    Me.txtNom.Text = ""
    Me.txtPrenom.Text = ""
    Me.txtAdresse.Text = ""
    Me.txtVille.Text = ""
    Me.txtTel.Text = ""
    Me.txtEmail.Text = ""
    Me.cboRegion.Value = ""
    
    MsgBox "Nouveau client ajouté avec succès !", vbInformation
End Sub
Private Sub btnSupprimer_Click()
    Dim code As String
    Dim wsClient As Worksheet
    Dim wsCommande As Worksheet
    Dim f As Range
    Dim lastRowClient As Long
    Dim i As Long
    Dim aDesCommandes As Boolean
    Dim rep1 As Integer, rep2 As Integer
    
    ' Vérifier qu'un client est sélectionné
    If Me.lstClients.ListIndex = -1 Then
        MsgBox "Veuillez sélectionner un client dans la liste à supprimer.", vbExclamation, "Aucun client sélectionné"
        Exit Sub
    End If
    
    code = Trim(Me.lstClients.List(Me.lstClients.ListIndex, 1))  ' id_client
    
    Set wsClient = Worksheets("CLIENT")
    Set wsCommande = Worksheets("COMMANDE")
    
    ' Rechercher le client dans la feuille CLIENT
    Set f = wsClient.Columns(1).Find(What:=code, LookAt:=xlWhole)
    If f Is Nothing Then
        MsgBox "Client introuvable dans la base.", vbCritical
        Exit Sub
    End If
    
    ' === Vérification : le client a-t-il des commandes ? ===
    aDesCommandes = False
    lastRowClient = wsCommande.Cells(wsCommande.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRowClient
        If Trim(CStr(wsCommande.Cells(i, 8).Value)) = code Then
            aDesCommandes = True
            Exit For
        End If
    Next i
    
    If aDesCommandes Then
        MsgBox "Impossible de supprimer ce client." & vbCrLf & _
               "Il possède une ou plusieurs commandes dans l'historique." & vbCrLf & vbCrLf & _
               "Supprimez ou réaffectez d'abord les commandes associées.", vbCritical, "Suppression interdite"
        Exit Sub
    End If
    
    ' === Double confirmation ===
    rep1 = MsgBox("Êtes-vous sûr de vouloir supprimer le client suivant ?" & vbCrLf & vbCrLf & _
                  "Nom : " & Me.txtNom.Text & " " & Me.txtPrenom.Text & vbCrLf & _
                  "ID : " & code & vbCrLf & vbCrLf & _
                  "Cette action est irréversible !", vbYesNo + vbQuestion + vbDefaultButton2, "Confirmer la suppression")
    If rep1 = vbNo Then Exit Sub
    
    rep2 = MsgBox("Confirmation finale : supprimer définitivement ce client ?", vbYesNo + vbCritical + vbDefaultButton2, "Dernière confirmation")
    If rep2 = vbNo Then Exit Sub
    
    ' === Suppression de la ligne ===
    Application.EnableEvents = False  ' Pour éviter des boucles éventuelles
    f.EntireRow.Delete
    Application.EnableEvents = True
    
    ' Réinitialiser les champs du formulaire
    Me.txtNom.Text = ""
    Me.txtPrenom.Text = ""
    Me.txtAdresse.Text = ""
    Me.txtVille.Text = ""
    Me.txtTel.Text = ""
    Me.txtEmail.Text = ""
    Me.cboRegion.Value = ""
    Me.lstCommandes.Clear
    Me.lstCommandes.AddItem "Aucune commande trouvée pour ce client"
    
    ' Recharger la liste des clients
    Call ChargerClients("")
    
    MsgBox "Le client a été supprimé avec succès.", vbInformation, "Suppression effectuée"
End Sub

