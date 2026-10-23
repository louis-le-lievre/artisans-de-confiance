VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fiche_Four 
   Caption         =   "Fiche Fourniseur"
   ClientHeight    =   7248
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7980
   OleObjectBlob   =   "Fiche_Four.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Fiche_Four"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variables module (ajoutez en haut du module, avant les Subs pour le nouveau UserForm Fournisseur)
Private originalDataFour As Variant
Private fournisseurTable As ListObject ' Pour stocker la table plg_fournisseur
Private produitTable As ListObject ' NOUVEAU : Pour stocker la table plg_produit (pour lookup noms)
Private artisanTable As ListObject ' NOUVEAU : Pour stocker la table plg_artisant (pour lookup noms)
Private acheterTable As ListObject ' NOUVEAU : Pour stocker la table plg_acheter

' Procédure publique pour charger un fournisseur par ID (appelée depuis double-clic externe)
Public Sub LoadFournisseur(fourID As Variant)
    On Error GoTo ErrHandler
  
    If IsEmpty(fourID) Or fourID = "" Then
        MsgBox "ID fournisseur invalide pour chargement.", vbExclamation
        Exit Sub
    End If
  
    ' Chargement direct sans simuler la ComboBox (pour éviter problèmes de sélection)
    Dim selectedID As Variant
    selectedID = CStr(fourID) ' Forcer en String pour cohérence avec CStr dans l'autre form
  
    ' Lookup direct de la ligne dans la table et remplissage des TextBoxes
    Dim rowIndex As Variant
    rowIndex = Application.match(selectedID, fournisseurTable.ListColumns(1).DataBodyRange, 0)
    If IsError(rowIndex) Then
        MsgBox "Fournisseur non trouvé dans la table (ID: " & selectedID & ").", vbExclamation
        Exit Sub
    End If
  
    ' Remplir les TextBoxes (colonnes : 1=ID, 2=Nom, 3=Coor, 4=Tel, 5=Email)
    TB_idfour.Value = selectedID ' Colonne 1 : ID
    TB_nomfour.Value = fournisseurTable.ListColumns(2).DataBodyRange(rowIndex, 1).Value ' Colonne 2 : Nom
    TB_coorfour.Value = fournisseurTable.ListColumns(3).DataBodyRange(rowIndex, 1).Value ' Colonne 3 : Coordonnées
    TB_Telfour.Value = fournisseurTable.ListColumns(4).DataBodyRange(rowIndex, 1).Value ' Colonne 4 : Téléphone
    TB_emailfour.Value = fournisseurTable.ListColumns(5).DataBodyRange(rowIndex, 1).Value ' Colonne 5 : Email
  
    ' Debug : Affiche les valeurs remplies
    Debug.Print "Fournisseur chargé directement - ID: " & selectedID & ", Nom: " & TB_nomfour.Value & ", Tel: " & TB_Telfour.Value
  
    ' Charger la ListBox avec les achats liés à ce fournisseur
    Call FillLBProduitsByFournisseur(selectedID)
  
    ' Sélectionner dans la ComboBox pour cohérence UI (recharge la liste complète)
    Call LoadFullListFour
    CB_recherfour.Value = selectedID
  
    Debug.Print "Fournisseur chargé via LoadFournisseur : ID=" & selectedID
    Exit Sub
  
ErrHandler:
    MsgBox "Erreur dans LoadFournisseur : " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub LB_listeproduit_Click()
 
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler ' Gestion d'erreur globale
    ' Charger la table FOURNISSEUR (existant)
    Dim ws_fourn As Worksheet
    Set ws_fourn = ThisWorkbook.Worksheets("FOURNISSEUR")
    Set fournisseurTable = ws_fourn.ListObjects("plg_fournisseur")
    ' Vérifier si la table existe
    If fournisseurTable Is Nothing Then
        MsgBox "Erreur : La table 'plg_fournisseur' n'existe pas dans la feuille 'FOURNISSEUR'." & vbCrLf & _
               "Vérifiez le nom de la table.", vbCritical
        Exit Sub
    End If
    ' Vérifier s'il y a des données
    If fournisseurTable.ListRows.count = 0 Then
        MsgBox "Aucune donnée dans la table 'plg_fournisseur'.", vbExclamation
        Exit Sub
    End If
    ' Charger la table PRODUIT pour lookup noms
    Dim ws_prod As Worksheet
    Set ws_prod = ThisWorkbook.Worksheets("PRODUIT")
    Set produitTable = ws_prod.ListObjects("plg_produit")
    If produitTable Is Nothing Then
        MsgBox "Erreur : La table 'plg_produit' n'existe pas dans la feuille 'PRODUIT'." & vbCrLf & _
               "Vérifiez le nom de la table.", vbCritical
        Exit Sub
    End If
    If produitTable.ListRows.count = 0 Then
        Debug.Print "Aucune donnée dans la table 'plg_produit' (mais on continue)."
    End If
    '  Charger la table ARTISAN pour lookup noms
    Dim ws_art As Worksheet
    Set ws_art = ThisWorkbook.Worksheets("ARTISAN")
    Set artisanTable = ws_art.ListObjects("plg_artisant")
    If artisanTable Is Nothing Then
        MsgBox "Erreur : La table 'plg_artisant' n'existe pas dans la feuille 'ARTISAN'." & vbCrLf & _
               "Vérifiez le nom de la table.", vbCritical
        Exit Sub
    End If
    If artisanTable.ListRows.count = 0 Then
        Debug.Print "Aucune donnée dans la table 'plg_artisant' (mais on continue)."
    End If
    ' Charger la table ACHETER
    Dim ws_ach As Worksheet
    Set ws_ach = ThisWorkbook.Worksheets("ACHETER")
    Set acheterTable = ws_ach.ListObjects("plg_acheter")
    If acheterTable Is Nothing Then
        MsgBox "Erreur : La table 'plg_acheter' n'existe pas dans la feuille 'ACHETER'." & vbCrLf & _
               "Vérifiez le nom de la table.", vbCritical
        Exit Sub
    End If
    If acheterTable.ListRows.count = 0 Then
        Debug.Print "Aucune donnée dans la table 'plg_acheter' (mais on continue)."
    End If
    ' Préparer un array pour les 2 premières colonnes seulement (ID et Nom) - FOURNISSEUR
    Dim dataArray As Variant
    dataArray = fournisseurTable.ListColumns(1).DataBodyRange.Resize(, 2).Value ' Colonnes 1 et 2 (ID, Nom)
    ' Assurer que c'est un array 2D (fix pour Type Mismatch potentiel)
    If Not IsArray(dataArray) Then
        ReDim dataArray(1 To 1, 1 To 2)
        dataArray(1, 1) = fournisseurTable.ListColumns(1).DataBodyRange.Cells(1, 1).Value
        dataArray(1, 2) = fournisseurTable.ListColumns(2).DataBodyRange.Cells(1, 1).Value
    End If
    ' Stocker les données originales
    originalDataFour = dataArray
    ' Remplir la ComboBox avec les données complètes
    With CB_recherfour
        .Clear
        .ColumnCount = 2
        .BoundColumn = 1 ' .Value renverra toujours l'ID sélectionné
        .ColumnWidths = "50pt;150pt" ' Affiche ID (50pt) + Nom (150pt) pour distinguer les homonymes
        .MatchEntry = 0 ' fmMatchEntryNone : Permet saisie libre pour le filtrage
        .List = dataArray ' Charge l'array complet
    End With
    ' Initialiser la ListBox LB_listeproduit (7 colonnes : Prod ID, Prod Nom, Art ID, Art Nom, Quantité, Date Achat, Montant)
    With LB_listeproduit
        .Clear
        .ColumnCount = 7 ' Sept colonnes
        .ColumnWidths = "40pt;80pt;40pt;50pt;40pt;50pt;50pt" ' Ajusté pour les nouvelles colonnes (ex. : petits pour IDs/num, larges pour noms/dates)
        .AddItem "Aucun fournisseur sélectionné" ' Placeholder initial (sera remplacé par array avec en-têtes)
    End With
    ' Debug : Affiche le nombre d'éléments chargés et type du premier ID
    Debug.Print "ComboBox fournisseur remplie avec " & CB_recherfour.ListCount & " éléments."
    If CB_recherfour.ListCount > 0 Then
        Debug.Print "Premier élément : ID=" & CB_recherfour.List(0, 0) & ", Nom=" & CB_recherfour.List(0, 1)
        Debug.Print "Type du premier ID : " & TypeName(CB_recherfour.List(0, 0))
    End If
    Exit Sub
ErrHandler:
    MsgBox "Erreur " & Err.Number & " : " & Err.Description & vbCrLf & _
           "Ligne : " & Erl, vbCritical
End Sub

' Événement KeyUp pour filtrage dynamique par nom (similaire à CB_recherarti_KeyUp)
Private Sub CB_recherfour_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrHandler
    ' Ignorer si pas de données originales ou si touche spéciale (ex. flèches, Entrée)
    If Not IsArray(originalDataFour) Or UBound(originalDataFour, 1) = 0 Then Exit Sub
    Dim searchText As String
    searchText = UCase(Trim(CB_recherfour.Text)) ' Texte tapé, majuscules pour insensible à la casse
    ' Si texte vide ou effacement, recharger la liste complète
    If Len(searchText) = 0 Then
        Call LoadFullListFour
        Exit Sub
    End If
    ' Filtrer les éléments où le nom commence par searchText (préfixe)
    Dim filteredArray() As Variant
    Dim matchCount As Long
    matchCount = 0
    Dim i As Long
    For i = 1 To UBound(originalDataFour, 1)
        If Len(originalDataFour(i, 2)) > 0 And UCase(Left(originalDataFour(i, 2), Len(searchText))) = searchText Then
            matchCount = matchCount + 1
        End If
    Next i
    If matchCount = 0 Then
        ' Option : Vider la liste pour signaler "pas de match"
        CB_recherfour.Clear
        Debug.Print "Aucun nom de fournisseur correspondant à '" & CB_recherfour.Text & "'"
        Exit Sub
    End If
    ' Construire l'array filtré
    ReDim filteredArray(1 To matchCount, 1 To 2)
    Dim j As Long
    j = 1
    For i = 1 To UBound(originalDataFour, 1)
        If Len(originalDataFour(i, 2)) > 0 And UCase(Left(originalDataFour(i, 2), Len(searchText))) = searchText Then
            filteredArray(j, 1) = originalDataFour(i, 1) ' ID
            filteredArray(j, 2) = originalDataFour(i, 2) ' Nom
            j = j + 1
        End If
    Next i
    ' Charger la liste filtrée
    With CB_recherfour
        .Clear
        .List = filteredArray
    End With
    Debug.Print matchCount & " nom(s) de fournisseur(s) correspondant à '" & searchText & "'"
    Exit Sub
ErrHandler:
    MsgBox "Erreur dans KeyUp : " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

' Procédure helper pour recharger la liste complète des fournisseurs
Private Sub LoadFullListFour()
    With CB_recherfour
        .Clear
        .List = originalDataFour
    End With
    Debug.Print "Liste complète des fournisseurs rechargée (" & UBound(originalDataFour, 1) & " éléments)."
End Sub

Private Sub CB_recherfour_Change()
    On Error GoTo ErrHandler
    If CB_recherfour.ListIndex = -1 Then
        ' Si pas de sélection, vider la ListBox et afficher message
        With LB_listeproduit
            .Clear
            .AddItem "Aucun fournisseur sélectionné"
        End With
        Exit Sub
    End If
    Dim selectedID As Variant ' Accepte String ou Long
    selectedID = CB_recherfour.Value ' Récupère l'ID sélectionné
    ' Après sélection, recharger la liste complète
    Call LoadFullListFour
    ' Lookup de la ligne dans la table et remplissage des TextBoxes
    Dim rowIndex As Variant
    rowIndex = Application.match(selectedID, fournisseurTable.ListColumns(1).DataBodyRange, 0) ' Trouve la ligne (exact match)
    If IsError(rowIndex) Then
        MsgBox "Fournisseur non trouvé dans la table (ID: " & selectedID & ").", vbExclamation
        Exit Sub
    End If
    ' Remplir les TextBoxes (colonnes : 1=ID, 2=Nom, 3=Coor, 4=Tel, 5=Email)
    TB_idfour.Value = selectedID ' Colonne 1 : ID
    TB_nomfour.Value = fournisseurTable.ListColumns(2).DataBodyRange(rowIndex, 1).Value ' Colonne 2 : Nom
    TB_coorfour.Value = fournisseurTable.ListColumns(3).DataBodyRange(rowIndex, 1).Value ' Colonne 3 : Coordonnées
    TB_Telfour.Value = fournisseurTable.ListColumns(4).DataBodyRange(rowIndex, 1).Value ' Colonne 4 : Téléphone
    TB_emailfour.Value = fournisseurTable.ListColumns(5).DataBodyRange(rowIndex, 1).Value ' Colonne 5 : Email
    ' Debug : Affiche les valeurs remplies
    Debug.Print "Fournisseur chargé - ID: " & selectedID & ", Nom: " & TB_nomfour.Value & ", Tel: " & TB_Telfour.Value
  
    ' Remplir la ListBox avec les achats liés à ce fournisseur (toutes les lignes, avec détails et en-têtes)
    Call FillLBProduitsByFournisseur(selectedID)
    Exit Sub
ErrHandler:
    MsgBox "Erreur dans CB_recherfour_Change : " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub UserForm_Click()
    
End Sub


Private Sub TB_nomfour_Change()
   
End Sub

Private Sub TB_coorfour_Change()
    
End Sub

Private Sub TB_Telfour_Change()
    
End Sub

Private Sub TB_emailfour_Change()
    
End Sub

' Procédure pour remplir LB_listeproduit avec TOUS les achats liés au fournisseur (pas uniques, avec quantité, date, montant ET EN-TÊTES)
Private Sub FillLBProduitsByFournisseur(fournisseurID As Variant)
    On Error GoTo ErrHandler
    If IsEmpty(fournisseurID) Or fournisseurID = "" Then
        With LB_listeproduit
            .Clear
            .AddItem "Aucun fournisseur sélectionné"
        End With
        Exit Sub
    End If
    If acheterTable Is Nothing Or produitTable Is Nothing Or artisanTable Is Nothing Then
        With LB_listeproduit
            .Clear
            .AddItem "Tables non disponibles"
        End With
        Exit Sub
    End If
    If acheterTable.ListRows.count = 0 Then
        With LB_listeproduit
            .Clear
            .AddItem "Aucune donnée dans ACHETER"
        End With
        Exit Sub
    End If
    ' Étape 1 : Collecter TOUTES les lignes depuis plg_acheter où id_fournisseur (col 4) = fournisseurID
    Dim matchingRows As Collection
    Set matchingRows = New Collection ' Pour stocker les indices des lignes matching
    Dim i As Long
    For i = 1 To acheterTable.ListRows.count
        If acheterTable.ListColumns(4).DataBodyRange.Cells(i, 1).Value = fournisseurID Then ' Col 4 : id_fournisseur
            matchingRows.Add i ' Ajouter l'indice de la ligne
        End If
    Next i
    If matchingRows.count = 0 Then
        With LB_listeproduit
            .Clear
            .AddItem "Aucun achat pour ce fournisseur"
        End With
        Debug.Print "Aucune ligne d'achat trouvée pour ID fournisseur : " & fournisseurID
        Exit Sub
    End If
    ' Étape 2 : Pour chaque ligne matching, lookup noms et construire l'array (AVEC EN-TÊTES)
    Dim countRows As Long
    countRows = matchingRows.count
    Dim lbData() As Variant
    ReDim lbData(1 To countRows + 1, 1 To 7) ' +1 pour la ligne d'en-têtes ; 7 cols : Prod ID, Prod Nom, Art ID, Art Nom, Quantité, Date Achat, Montant
    ' Ligne 1 : En-têtes (gras ou distincts visuellement si possible ; ici, texte simple)
    lbData(1, 1) = "ID Produit"
    lbData(1, 2) = "Nom Produit"
    lbData(1, 3) = "ID Artisan"
    lbData(1, 4) = "Nom Artisan"
    lbData(1, 5) = "Quantité"
    lbData(1, 6) = "Date Achat"
    lbData(1, 7) = "Montant"
    Dim j As Long
    For j = 1 To countRows
        Dim rowIndex As Long
        rowIndex = matchingRows(j)
        Dim thisProdID As Variant
        Dim thisArtID As Variant
        Dim quantite As Variant
        Dim dateAchat As Variant
        Dim montant As Variant
        thisProdID = acheterTable.ListColumns(1).DataBodyRange.Cells(rowIndex, 1).Value ' Col 1 : id_produit
        thisArtID = acheterTable.ListColumns(2).DataBodyRange.Cells(rowIndex, 1).Value ' Col 2 : id_artisan
        quantite = acheterTable.ListColumns(3).DataBodyRange.Cells(rowIndex, 1).Value ' Col 3 : quantité
        dateAchat = acheterTable.ListColumns(5).DataBodyRange.Cells(rowIndex, 1).Value ' Col 5 : date_achat
        montant = acheterTable.ListColumns(6).DataBodyRange.Cells(rowIndex, 1).Value ' Col 6 : montant_achat
        ' Lookup nom produit (col 2 de plg_produit)
        Dim prodRowIndex As Variant
        prodRowIndex = Application.match(thisProdID, produitTable.ListColumns(1).DataBodyRange, 0)
        Dim nomProd As String
        If IsError(prodRowIndex) Then
            nomProd = "Produit inconnu (ID: " & thisProdID & ")"
            Debug.Print "Nom produit non trouvé pour ID : " & thisProdID
        Else
            nomProd = produitTable.ListColumns(2).DataBodyRange(prodRowIndex, 1).Value
        End If
        ' Lookup nom artisan (col 2 de plg_artisant)
        Dim artRowIndex As Variant
        artRowIndex = Application.match(thisArtID, artisanTable.ListColumns(1).DataBodyRange, 0)
        Dim nomArt As String
        If IsError(artRowIndex) Then
            nomArt = "Artisan inconnu (ID: " & thisArtID & ")"
            Debug.Print "Nom artisan non trouvé pour ID : " & thisArtID
        Else
            nomArt = artisanTable.ListColumns(2).DataBodyRange(artRowIndex, 1).Value
        End If
        ' Remplir la ligne (décalée à partir de la ligne 2)
        Dim dataRow As Long
        dataRow = j + 1
        lbData(dataRow, 1) = thisProdID ' Prod ID
        lbData(dataRow, 2) = nomProd ' Prod Nom
        lbData(dataRow, 3) = thisArtID ' Art ID
        lbData(dataRow, 4) = nomArt ' Art Nom
        lbData(dataRow, 5) = quantite ' Quantité
        lbData(dataRow, 6) = dateAchat ' Date Achat
        lbData(dataRow, 7) = montant ' Montant
    Next j
    ' Remplir la ListBox
    With LB_listeproduit
        .Clear
        .List = lbData
    End With
    Debug.Print countRows & " achat(s) chargé(s) pour fournisseur : " & fournisseurID & " (avec en-têtes)"
    Exit Sub
ErrHandler:
    MsgBox "Erreur dans FillLBProduitsByFournisseur : " & Err.Number & " - " & Err.Description, vbExclamation
    With LB_listeproduit
        .Clear
        .AddItem "Erreur lors du chargement"
    End With
End Sub

