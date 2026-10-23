VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Gestion_arti 
   Caption         =   "Gestion artisan"
   ClientHeight    =   4752
   ClientLeft      =   -156
   ClientTop       =   -468
   ClientWidth     =   7164
   OleObjectBlob   =   "Gestion_arti.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Gestion_arti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' === FONCTIONS UTILITAIRES (déclarées en premier pour éviter erreurs de compilation - Cohérent avec TP2 : Macros utilitaires) ===
' Fonction utilitaire : Récupère l'ID de catégorie à partir du libellé
Private Function GetCategoryID(libelle As String) As String
    Dim wsCat As Worksheet
    Dim tblCat As ListObject
    Dim rngCat As Range
    Dim k As Long
    On Error GoTo ErreurLookupID
    Set wsCat = ThisWorkbook.Worksheets("CATEGORIE_TRAVAUX")
    Set tblCat = wsCat.ListObjects("plg_categorie_travaux")
    If Not tblCat.DataBodyRange Is Nothing Then
        Set rngCat = tblCat.DataBodyRange
    Else
        Set rngCat = tblCat.Range.Offset(1, 0).Resize(tblCat.Range.Rows.count - 1, tblCat.Range.Columns.count)
    End If
    If rngCat.Rows.count = 0 Then
        GetCategoryID = ""
        Exit Function
    End If
    ' Boucle pour trouver l'ID correspondant au libellé (colonne 2)
    For k = 1 To rngCat.Rows.count
        If UCase(Trim(rngCat.Cells(k, 2).Value)) = UCase(Trim(libelle)) Then
            GetCategoryID = Trim(rngCat.Cells(k, 1).Value) ' Retourne l'ID (colonne 1)
            Exit Function
        End If
    Next k
    GetCategoryID = "" ' Non trouvé
    Exit Function
ErreurLookupID:
    GetCategoryID = ""
End Function
' Fonction utilitaire : Récupère le libellé de catégorie à partir de l'ID
Private Function GetCategoryLabel(id As String) As String
    Dim wsCat As Worksheet
    Dim tblCat As ListObject
    Dim rngCat As Range
    Dim k As Long
    On Error GoTo ErreurLookupLabel
    Set wsCat = ThisWorkbook.Worksheets("CATEGORIE_TRAVAUX")
    Set tblCat = wsCat.ListObjects("plg_categorie_travaux")
    If Not tblCat.DataBodyRange Is Nothing Then
        Set rngCat = tblCat.DataBodyRange
    Else
        Set rngCat = tblCat.Range.Offset(1, 0).Resize(tblCat.Range.Rows.count - 1, tblCat.Range.Columns.count)
    End If
    If rngCat.Rows.count = 0 Then
        GetCategoryLabel = ""
        Exit Function
    End If
    ' Boucle pour trouver le libellé correspondant à l'ID (colonne 1)
    For k = 1 To rngCat.Rows.count
        If UCase(Trim(rngCat.Cells(k, 1).Value)) = UCase(Trim(id)) Then
            GetCategoryLabel = Trim(rngCat.Cells(k, 2).Value) ' Retourne le libellé (colonne 2)
            Exit Function
        End If
    Next k
    GetCategoryLabel = "" ' Non trouvé
    Exit Function
ErreurLookupLabel:
    GetCategoryLabel = ""
End Function
' Procédure utilitaire pour actualiser la ListBox
Private Sub ActualiserListBox(Optional searchFilter As String = "")
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim i As Long, j As Long
    Dim feuilleActive As Worksheet
    Dim searchText As String
    Dim dataRow As Long ' Compteur pour les lignes de données (après en-tête)
    Set feuilleActive = ActiveSheet
    On Error GoTo ErreurActualiser
    If searchFilter <> "" Then
        searchText = UCase(Trim(searchFilter))
    Else
        searchText = ""
    End If
    Set ws = ThisWorkbook.Worksheets("ARTISAN")
    ws.Activate
    Set tbl = ws.ListObjects("plg_artisant")
    If Not tbl.DataBodyRange Is Nothing Then
        Set rng = tbl.DataBodyRange
    Else
        Set rng = tbl.Range
        If rng.Rows.count <= 1 Then GoTo RestaurerActualiser
        Set rng = rng.Offset(1, 0).Resize(rng.Rows.count - 1, rng.Columns.count)
    End If
    If rng.Rows.count = 0 Then GoTo RestaurerActualiser
    LB_listearti.Clear
    LB_listearti.ColumnCount = rng.Columns.count
    LB_listearti.ColumnWidths = "50;100;80;80;120;60;100"
    ' Ajoute la ligne d'en-tête manuellement (comme première ligne de la ListBox, si HeaderRowRange existe)
    If Not tbl.HeaderRowRange Is Nothing Then
        LB_listearti.AddItem tbl.HeaderRowRange.Cells(1, 1).Value ' En-tête colonne 1 : ID
        For j = 2 To rng.Columns.count
            LB_listearti.List(0, j - 1) = tbl.HeaderRowRange.Cells(1, j).Value ' En-têtes des autres colonnes
        Next j
        dataRow = 1 ' Index de la ListBox pour les données (après en-tête à l'index 0)
    Else
        dataRow = 0 ' Pas d'en-tête, commence directement à 0
    End If
    ' Peuple les données filtrées
    For i = 1 To rng.Rows.count
        If searchText = "" Or UCase(Left(Trim(rng.Cells(i, 2).Value), Len(searchText))) = searchText Then
            LB_listearti.AddItem rng.Cells(i, 1).Value ' Colonne 1 : ID (nouvelle ligne)
            For j = 2 To rng.Columns.count
                If j = 6 Then ' Colonne 6 : Dispo
                    LB_listearti.List(dataRow, j - 1) = CStr(rng.Cells(i, j).Value)
                ElseIf j = 7 Then ' Colonne 7 : Catégorie (affiche libellé au lieu d'ID)
                    LB_listearti.List(dataRow, j - 1) = GetCategoryLabel(CStr(rng.Cells(i, j).Value))
                Else
                    LB_listearti.List(dataRow, j - 1) = rng.Cells(i, j).Value
                End If
            Next j
            dataRow = dataRow + 1 ' Passe à la ligne suivante pour les données
        End If
    Next i
    LB_listearti.MultiSelect = fmMultiSelectSingle
RestaurerActualiser:
    feuilleActive.Activate
    Exit Sub
ErreurActualiser:
    MsgBox "Erreur lors de l'actualisation : " & Err.Description, vbError
    feuilleActive.Activate
End Sub
' Procédure utilitaire pour vider les TextBoxes et les ComboBox (TP3 : Événement pour reset)
Private Sub ViderTextBoxes()
    TB_idarti.Value = ""
    TB_nomarti.Value = ""
    CB_catarti.Value = "" ' Vide la sélection du ComboBox Cat
    TB_coorarti.Value = ""
    TB_Telarti.Value = ""
    TB_emailarti.Value = ""
    CB_dispoarti.Value = "" ' Vide la sélection du ComboBox Dispo
    TB_recherarti.Value = "" ' Vide la recherche
End Sub
' Bouton Ajouter : Ajoute une nouvelle ligne au tableau
Private Sub CB_ajoutarti_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim nouvelleLigne As ListRow
    Dim feuilleActive As Worksheet
    Dim i As Long
    Dim idExistant As Boolean
    Dim duplicateExistant As Boolean
    Dim rngDonnees As Range
    Dim idCategorie As String ' ID de la catégorie sélectionnée
    Dim dataArray As Variant
    Dim match As Boolean
    ' Validation : Champs obligatoires (TP3 : Saisie contrôlée dans UserForm)
    If Trim(TB_nomarti.Value) = "" Then
        MsgBox "Le nom est obligatoire pour ajouter un artisan.", vbExclamation
        Exit Sub
    End If
    ' Validation ID : Obligatoire et unique
    If Trim(TB_idarti.Value) = "" Then
        MsgBox "L'ID est obligatoire pour ajouter un artisan.", vbExclamation
        Exit Sub
    End If
    ' Validation Catégorie : Obligatoire (via ComboBox - TP3 : Liste déroulante)
    If Trim(CB_catarti.Value) = "" Then
        MsgBox "Sélectionnez une catégorie pour ajouter un artisan.", vbExclamation
        Exit Sub
    End If
    ' Validation Disponibilité : Obligatoire (via ComboBox)
    If Trim(CB_dispoarti.Value) = "" Then
        MsgBox "Sélectionnez la disponibilité (VRAI/FAUX) pour ajouter un artisan.", vbExclamation
        Exit Sub
    End If
    ' Récupère l'ID de la catégorie sélectionnée
    idCategorie = GetCategoryID(CB_catarti.Value)
    If idCategorie = "" Then
        MsgBox "Catégorie sélectionnée non valide. Vérifiez les données.", vbExclamation
        Exit Sub
    End If
    Set feuilleActive = ActiveSheet
    On Error GoTo ErreurAjout
    Set ws = ThisWorkbook.Worksheets("ARTISAN")
    ws.Activate
    Set tbl = ws.ListObjects("plg_artisant")
    ' Récupère les données pour vérifier l'unicité de l'ID et les doublons (TP1 : Plages nommées)
    If Not tbl.DataBodyRange Is Nothing Then
        Set rngDonnees = tbl.DataBodyRange
    Else
        Set rngDonnees = tbl.Range.Offset(1, 0).Resize(tbl.Range.Rows.count - 1, tbl.Range.Columns.count)
    End If
    ' Récupère les données dans un tableau pour optimisation (évite accès multiples à la feuille)
    dataArray = rngDonnees.Value ' Charge tout le range en tableau 2D (1-based)
    ' Vérifie si l'ID existe déjà (boucle sur colonne 1)
    idExistant = False
    For i = 1 To UBound(dataArray, 1)
        If UCase(Trim(dataArray(i, 1))) = UCase(Trim(TB_idarti.Value)) Then
            idExistant = True
            Exit For
        End If
    Next i
    If idExistant Then
        MsgBox "Un artisan avec cet ID existe déjà. Veuillez choisir un ID unique.", vbExclamation
        feuilleActive.Activate
        Exit Sub
    End If
    ' Vérifie si un doublon existe sur les autres caractéristiques (colonnes 2-6 + ID catégorie en 7)
    duplicateExistant = False
    For i = 1 To UBound(dataArray, 1)
        match = True
     
        If UCase(Trim(dataArray(i, 2))) <> UCase(Trim(TB_nomarti.Value)) Then match = False
        If UCase(Trim(dataArray(i, 3))) <> UCase(Trim(TB_coorarti.Value)) Then match = False
        If UCase(Trim(dataArray(i, 4))) <> UCase(Trim(TB_Telarti.Value)) Then match = False
        If UCase(Trim(dataArray(i, 5))) <> UCase(Trim(TB_emailarti.Value)) Then match = False
        If UCase(Trim(dataArray(i, 6))) <> UCase(Trim(CB_dispoarti.Value)) Then match = False
        If UCase(Trim(dataArray(i, 7))) <> UCase(idCategorie) Then match = False
     
        If match Then
            duplicateExistant = True
            Exit For
        End If
    Next i
    If duplicateExistant Then
        MsgBox "Un artisan avec ces mêmes caractéristiques (nom, coordonnées, téléphone, email, disponibilité, catégorie) existe déjà (ID différent). Vérifiez les données.", vbExclamation
        feuilleActive.Activate
        Exit Sub
    End If
    ' Ajoute une nouvelle ligne si ID unique et pas de doublon (TP3 : Interaction bouton pour ajout)
    Set nouvelleLigne = tbl.ListRows.Add
    ' Remplit les colonnes (ordre : 1=ID, 2=Nom, 3=Coor, 4=Tel, 5=Email, 6=Dispo, 7=ID Cat - TP1 : Structure colonnes)
    nouvelleLigne.Range(1, 1).Value = TB_idarti.Value ' Colonne 1 : ID
    nouvelleLigne.Range(1, 2).Value = TB_nomarti.Value ' Colonne 2 : Nom
    nouvelleLigne.Range(1, 3).Value = TB_coorarti.Value ' Colonne 3 : Coor
    nouvelleLigne.Range(1, 4).Value = TB_Telarti.Value ' Colonne 4 : Tel
    nouvelleLigne.Range(1, 5).Value = TB_emailarti.Value ' Colonne 5 : Email
    nouvelleLigne.Range(1, 6).Value = CB_dispoarti.Value ' Colonne 6 : Dispo (VRAI/FAUX via ComboBox)
    nouvelleLigne.Range(1, 7).Value = idCategorie ' Colonne 7 : ID de catégorie
    ' Actualise la ListBox et vide les champs (TP2 : Macro pour mise à jour)
    Call ActualiserListBox
    Call ViderTextBoxes
    MsgBox "Artisan ajouté avec succès.", vbInformation
    feuilleActive.Activate
    Exit Sub
ErreurAjout:
    MsgBox "Erreur lors de l'ajout : " & Err.Description, vbError
    If Not nouvelleLigne Is Nothing Then nouvelleLigne.Delete ' Annule si erreur
    feuilleActive.Activate
End Sub
' Bouton Modifier : Met à jour la ligne sélectionnée
Private Sub CB_modifarti_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim ligneModif As ListRow
    Dim feuilleActive As Worksheet
    Dim idCategorie As String ' ID de la catégorie sélectionnée
    Dim targetID As String
    Dim r As Long
    ' Vérifie sélection (TP3 : Interaction ListBox)
    If LB_listearti.ListIndex = -1 Then
        MsgBox "Sélectionnez un artisan à modifier.", vbExclamation
        Exit Sub
    End If
    ' Validation : Champs obligatoires (TP3 : Saisie contrôlée)
    If Trim(TB_nomarti.Value) = "" Then
        MsgBox "Le nom est obligatoire pour modifier.", vbExclamation
        Exit Sub
    End If
    ' Validation Catégorie : Obligatoire (via ComboBox)
    If Trim(CB_catarti.Value) = "" Then
        MsgBox "Sélectionnez une catégorie pour modifier.", vbExclamation
        Exit Sub
    End If
    ' Validation Disponibilité : Obligatoire (via ComboBox)
    If Trim(CB_dispoarti.Value) = "" Then
        MsgBox "Sélectionnez la disponibilité (VRAI/FAUX) pour modifier.", vbExclamation
        Exit Sub
    End If
    ' Récupère l'ID de la catégorie sélectionnée
    idCategorie = GetCategoryID(CB_catarti.Value)
    If idCategorie = "" Then
        MsgBox "Catégorie sélectionnée non valide. Vérifiez les données.", vbExclamation
        Exit Sub
    End If
    Set feuilleActive = ActiveSheet
    On Error GoTo ErreurModif
    Set ws = ThisWorkbook.Worksheets("ARTISAN")
    ws.Activate
    Set tbl = ws.ListObjects("plg_artisant")
    ' Recherche la ligne par ID (robuste avec filtre)
    targetID = Trim(TB_idarti.Value)
    Set ligneModif = Nothing
    For r = 1 To tbl.ListRows.count
        If UCase(Trim(tbl.ListRows(r).Range(1, 1).Value)) = UCase(targetID) Then
            Set ligneModif = tbl.ListRows(r)
            Exit For
        End If
    Next r
    If ligneModif Is Nothing Then
        MsgBox "Ligne non trouvée pour modification.", vbExclamation
        feuilleActive.Activate
        Exit Sub
    End If
    ' Met à jour les colonnes (TP3 : Mise à jour via UserForm)
    ligneModif.Range(1, 1).Value = TB_idarti.Value
    ligneModif.Range(1, 2).Value = TB_nomarti.Value
    ligneModif.Range(1, 3).Value = TB_coorarti.Value
    ligneModif.Range(1, 4).Value = TB_Telarti.Value
    ligneModif.Range(1, 5).Value = TB_emailarti.Value
    ligneModif.Range(1, 6).Value = CB_dispoarti.Value ' Colonne 6 : Dispo (VRAI/FAUX via ComboBox)
    ligneModif.Range(1, 7).Value = idCategorie ' Colonne 7 : ID de catégorie
    ' Actualise la ListBox (TP2 : Mise à jour macro)
    Call ActualiserListBox
    MsgBox "Artisan modifié avec succès.", vbInformation
    feuilleActive.Activate
    Exit Sub
ErreurModif:
    MsgBox "Erreur lors de la modification : " & Err.Description, vbError
    feuilleActive.Activate
End Sub
' Bouton Supprimer : Supprime la ligne sélectionnée (TP2 : Macro pour suppression)
Private Sub CB_supparti_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim ligneSupp As ListRow
    Dim reponse As VbMsgBoxResult
    Dim feuilleActive As Worksheet
    Dim targetID As String
    Dim r As Long
    ' Vérifie sélection (TP3 : Interaction ListBox)
    If LB_listearti.ListIndex = -1 Then
        MsgBox "Sélectionnez un artisan à supprimer.", vbExclamation
        Exit Sub
    End If
    ' Confirmation (TP3 : Boîte de dialogue)
    reponse = MsgBox("Confirmez la suppression de l'artisan : " & TB_nomarti.Value & " ?", vbYesNo + vbQuestion)
    If reponse = vbNo Then Exit Sub
    Set feuilleActive = ActiveSheet
    On Error GoTo ErreurSupp
    Set ws = ThisWorkbook.Worksheets("ARTISAN")
    ws.Activate
    Set tbl = ws.ListObjects("plg_artisant")
    ' Recherche la ligne par ID (robuste avec filtre)
    targetID = Trim(TB_idarti.Value)
    Set ligneSupp = Nothing
    For r = 1 To tbl.ListRows.count
        If UCase(Trim(tbl.ListRows(r).Range(1, 1).Value)) = UCase(targetID) Then
            Set ligneSupp = tbl.ListRows(r)
            Exit For
        End If
    Next r
    If ligneSupp Is Nothing Then
        MsgBox "Ligne non trouvée pour suppression.", vbExclamation
        feuilleActive.Activate
        Exit Sub
    End If
    ' Supprime la ligne (TP1 : Tables dynamiques)
    ligneSupp.Delete
    ' Actualise et vide (TP2 : Macro pour mise à jour)
    Call ActualiserListBox
    Call ViderTextBoxes
    MsgBox "Artisan supprimé avec succès.", vbInformation
    feuilleActive.Activate
    Exit Sub
ErreurSupp:
    MsgBox "Erreur lors de la suppression : " & Err.Description, vbError
    feuilleActive.Activate
End Sub
' Bouton Réinitialiser : Vide les TextBoxes et désélectionne la ListBox (TP3 : Reset UserForm)
Private Sub CB_teini_Click()
    Call ViderTextBoxes ' Vide tous les TextBoxes
   
    LB_listearti.ListIndex = -1
    
    MsgBox "Formulaire réinitialisé.", vbInformation
End Sub
' Événement Change pour le ComboBox Catégorie : Optionnel (TP3 : Événement Change)
Private Sub CB_catarti_Change()
    If CB_catarti.ListIndex = -1 Then
        
    End If
End Sub
' Événement Change pour le ComboBox Disponibilité : Optionnel (TP3 : Événement Change)
Private Sub CB_dispoarti_Change()
    If CB_dispoarti.ListIndex = -1 Then
        
    End If
End Sub
Private Sub L_titregest_Click()
End Sub
' Mise à jour de LB_listearti_Click : Remplit les ComboBoxes et TextBoxes avec les caractéristiques de l'artisan sélectionné
Private Sub LB_listearti_Click()
    Dim selectedIndex As Integer
    Dim libelleCat As String
    Dim dispoValue As String ' Variable pour la vérification Dispo
    ' Vérifie si un élément est sélectionné (TP3 : Interaction ListBox)
    If LB_listearti.ListIndex = -1 Then
        Exit Sub ' Rien sélectionné, ne fait rien
    End If
    selectedIndex = LB_listearti.ListIndex
    ' Ignore si l'en-tête est sélectionné (index 0)
    If selectedIndex = 0 Then
        Exit Sub
    End If
    ' PAS D'AJUSTEMENT : Les indices des données commencent directement à 1 (en-tête à 0), donc ListIndex correspond déjà à l'index de la ligne de données
    ' Remplit les TextBoxes de base (colonnes 1-5, index 0-4 - TP3 : Saisie depuis ListBox)
    TB_idarti.Value = LB_listearti.List(selectedIndex, 0) ' Colonne 1 : ID
    TB_nomarti.Value = LB_listearti.List(selectedIndex, 1) ' Colonne 2 : Nom
    TB_coorarti.Value = LB_listearti.List(selectedIndex, 2) ' Colonne 3 : Coor
    TB_Telarti.Value = LB_listearti.List(selectedIndex, 3) ' Colonne 4 : Tel
    TB_emailarti.Value = LB_listearti.List(selectedIndex, 4) ' Colonne 5 : Email
    ' Remplit le ComboBox Dispo (colonne 6, index 5) avec vérification
    dispoValue = CStr(LB_listearti.List(selectedIndex, 5))
    If dispoValue = "VRAI" Or dispoValue = "FAUX" Then
        CB_dispoarti.Value = dispoValue
    Else
        CB_dispoarti.Value = "" ' Si invalide, vide
    End If
    ' Remplit le ComboBox Catégorie (colonne 7, index 6) directement avec le libellé affiché dans la ListBox
    libelleCat = CStr(LB_listearti.List(selectedIndex, 6))
    If libelleCat <> "" Then
        CB_catarti.Value = libelleCat ' Sélectionne le libellé dans ComboBox Cat
    Else
        CB_catarti.Value = "" ' Si invalide, vide
    End If
End Sub
Private Sub LB_listearti_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim selectedIndex As Integer
    Dim selectedID As String
    Dim reponse As VbMsgBoxResult
    Dim ficheArti As Object
   
    If LB_listearti.ListIndex = -1 Then
        MsgBox "Sélectionnez un artisan en cliquant dessus avant de double-cliquer.", vbExclamation
        Exit Sub
    End If
   
    selectedIndex = LB_listearti.ListIndex
   
    If selectedIndex = 0 Then Exit Sub ' Ignore en-tête
   
    selectedID = CStr(LB_listearti.List(selectedIndex, 0))
   
    ' NOUVEAU : Pop-up de confirmation avant ouverture de la fiche artisan
    reponse = MsgBox("Voulez-vous ouvrir la fiche de l'artisan ID : " & selectedID & " ?", vbYesNo + vbQuestion, "Confirmation ouverture fiche")
    If reponse = vbNo Then
        Exit Sub
    End If
   
    On Error GoTo ErreurDoubleClick
   
    Application.ScreenUpdating = False
   
    Set ficheArti = VBA.UserForms.Add("Fiche_arti")
    ficheArti.LoadArtisan selectedID
    ficheArti.Show vbModal
   
    Call ActualiserListBox
    Unload ficheArti
    Set ficheArti = Nothing
   
    Application.ScreenUpdating = True
    Exit Sub
   
ErreurDoubleClick:
    MsgBox "Erreur lors de l'ouverture de Fiche_arti : " & Err.Description, vbError
    If Not ficheArti Is Nothing Then Unload ficheArti
    Application.ScreenUpdating = True
End Sub
' Événements vides
Private Sub TB_coorarti_Change()
End Sub
Private Sub TB_emailarti_Change()
End Sub
Private Sub TB_idarti_Change()
End Sub
Private Sub TB_nomarti_Change()
End Sub
Private Sub TB_recherarti_Change()
    Call ActualiserListBox(Trim(TB_recherarti.Value))
End Sub
Private Sub TB_Telarti_Change()
End Sub
' Événement au chargement/activation du UserForm : Remplit la ListBox ET les ComboBox (TP3 : Événement Initialize)
Private Sub UserForm_Initialize()
    Dim feuilleActive As Worksheet ' Pour sauvegarder et restaurer
    Dim wsCat As Worksheet
    Dim tblCat As ListObject
    Dim rngCat As Range
    Dim k As Long
    Dim feuilleActiveCat As Worksheet
    ' Sauvegarde la feuille active actuelle
    Set feuilleActive = ActiveSheet
    ' === PARTIE 2 : Chargement du ComboBox pour les catégories (TP3 : Liste déroulante) ===
    Set feuilleActiveCat = ActiveSheet ' Sauvegarde pour la section catégories
    On Error GoTo ErreurCategorie
    Set wsCat = ThisWorkbook.Worksheets("CATEGORIE_TRAVAUX")
    wsCat.Activate ' Active la feuille pour accès sûr (TP2 : Contexte feuille)
    Set tblCat = wsCat.ListObjects("plg_categorie_travaux")
    ' Récupère les données (colonne 2 pour "libelle_categorie_travaux" - TP1 : Import/Conversion)
    If Not tblCat.DataBodyRange Is Nothing Then
        Set rngCat = tblCat.DataBodyRange
    Else
        Set rngCat = tblCat.Range
        If rngCat.Rows.count <= 1 Then
            MsgBox "Le tableau 'plg_categorie_travaux' est vide sur CATEGORIE_TRAVAUX.", vbExclamation
            GoTo RestaurerCategorie
        End If
        Set rngCat = rngCat.Offset(1, 0).Resize(rngCat.Rows.count - 1, rngCat.Columns.count)
    End If
    If rngCat.Rows.count = 0 Then
        MsgBox "Aucune catégorie disponible sur la feuille CATEGORIE_TRAVAUX.", vbExclamation
        GoTo RestaurerCategorie
    End If
    ' Vide et peuple le ComboBox avec les noms de catégories (colonne 2 : libelle_categorie_travaux)
    CB_catarti.Clear
    For k = 1 To rngCat.Rows.count
        CB_catarti.AddItem rngCat.Cells(k, 2).Value ' Ajoute chaque libellé de catégorie (colonne 2)
    Next k
    ' Configuration : Liste déroulante simple (limite le choix - TP3 : Propriétés ComboBox)
    CB_catarti.Style = fmStyleDropDownList
    CB_catarti.ListIndex = -1 ' Aucune sélection par défaut
RestaurerCategorie:
    feuilleActiveCat.Activate
    On Error GoTo 0 ' Restaure la gestion d'erreurs normale
    ' === PARTIE 3 : Chargement du ComboBox pour la disponibilité (TP3 : Liste déroulante simple) ===
    CB_dispoarti.Clear
    CB_dispoarti.AddItem "VRAI"
    CB_dispoarti.AddItem "FAUX"
    CB_dispoarti.Style = fmStyleDropDownList ' Liste déroulante simple (limite le choix)
    CB_dispoarti.ListIndex = -1 ' Aucune sélection par défaut
    ' Optionnel : Vider TextBoxes au démarrage (TP3 : Initialize pour reset)
    Call ViderTextBoxes
    ' Chargement de la ListBox (tous les artisans)
    Call ActualiserListBox
RestaurerFeuille:
    feuilleActive.Activate
    Exit Sub
ErreurCategorie:
    MsgBox "Erreur lors du chargement des catégories depuis CATEGORIE_TRAVAUX : " & Err.Description, vbError
    GoTo RestaurerCategorie
End Sub
