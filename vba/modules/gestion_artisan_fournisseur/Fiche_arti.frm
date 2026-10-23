VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fiche_arti 
   Caption         =   "Fiche artisan"
   ClientHeight    =   7608
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7968
   OleObjectBlob   =   "Fiche_arti.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Fiche_arti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variables module (ajoutez en haut du module, avant les Subs)
Private originalData As Variant
Private artisanTable As ListObject ' Pour stocker la table et éviter de la recharger
Private possedeTable As ListObject ' Pour stocker la table plg_possede
Private produitTable As ListObject ' Pour stocker la table plg_produit
Private approvisionneTable As ListObject ' Pour stocker la table plg_approvisionne
Private fournisseurTable As ListObject ' NOUVEAU : Pour stocker la table plg_fournisseur
Private currentArtisanID As Variant ' Pour stocker l'ID artisan courant (pour filtrage produit)

'Procédure publique pour charger un artisan par ID (appelée depuis double-clic externe)
Public Sub LoadArtisan(artiID As Variant)
    On Error GoTo ErrHandler
  
    If IsEmpty(artiID) Or artiID = "" Then
        MsgBox "ID artisan invalide pour chargement.", vbExclamation
        Exit Sub
    End If
  
    ' Simuler la sélection dans CB_recherarti (met à jour Value et déclenche Change)
    CB_recherarti.Value = artiID
    Call CB_recherarti_Change ' Déclenche le remplissage des TextBoxes et produits
  
    Debug.Print "Artisan chargé via LoadArtisan : ID=" & artiID
    Exit Sub
  
ErrHandler:
    MsgBox "Erreur dans LoadArtisan : " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub CB_recherproduit_Change()
    On Error GoTo ErrHandler
    If CB_recherproduit.ListIndex = -1 Then
        ' Si pas de sélection, afficher tous les fournisseurs dans LB et vider le stock
        TB_stockprod.Value = ""
        If Not IsEmpty(currentArtisanID) Then Call FillLBFULL(currentArtisanID)
        Exit Sub
    End If
    Dim selectedProdID As Variant
    selectedProdID = CB_recherproduit.Value ' Récupère l'id_produit sélectionné
    ' Si placeholder ou pas d'ID valide, afficher tous et vider stock
    If IsEmpty(selectedProdID) Or selectedProdID = "" Then
        TB_stockprod.Value = ""
        If Not IsEmpty(currentArtisanID) Then Call FillLBFULL(currentArtisanID)
        Exit Sub
    End If
    '* AMÉLIORÉ : Vérif si currentArtisanID est valide avant lookup stock
    If IsEmpty(currentArtisanID) Then
        TB_stockprod.Value = ""
        Exit Sub
    End If
  
    ' Récupérer la quantité du produit pour l'artisan sélectionné depuis plg_possede (colonne C)
    Dim stockQuantity As Variant
    stockQuantity = ""
    Dim i As Long
    For i = 1 To possedeTable.ListRows.count
        If possedeTable.ListColumns(1).DataBodyRange.Cells(i, 1).Value = selectedProdID And _
           possedeTable.ListColumns(2).DataBodyRange.Cells(i, 1).Value = currentArtisanID Then
            stockQuantity = possedeTable.ListColumns(3).DataBodyRange.Cells(i, 1).Value ' Colonne C : quantité
            Exit For
        End If
    Next i
    TB_stockprod.Value = stockQuantity ' Afficher la quantité (ou vide si non trouvée)
    ' Filtrer la ListBox pour ne montrer que les fournisseurs de ce produit
    Call FilterLBByProduct(currentArtisanID, selectedProdID)
    Debug.Print "Fournisseurs filtrés pour produit ID=" & selectedProdID & " et artisan " & currentArtisanID & ", Stock: " & stockQuantity
    Exit Sub
ErrHandler:
    MsgBox "Erreur dans CB_recherproduit_Change : " & Err.Number & " - " & Err.Description, vbExclamation
    TB_stockprod.Value = "" ' Vider en cas d'erreur
End Sub

Private Sub LB_listefourni_Click()

End Sub

Private Sub LB_listefourni_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim selectedIndex As Integer
    Dim selectedFourID As String
    Dim reponse As VbMsgBoxResult
    Dim ficheFour As Object ' Référence à l'UserForm Fiche_Four
   
    ' Vérifie si un élément est sélectionné
    If LB_listefourni.ListIndex = -1 Then
        MsgBox "Sélectionnez un fournisseur en cliquant dessus avant de double-cliquer.", vbExclamation
        Exit Sub
    End If
   
    selectedIndex = LB_listefourni.ListIndex
   
    ' Ignore si l'en-tête est sélectionné (index 0)
    If selectedIndex = 0 Then
        Exit Sub
    End If
   
    ' Récupère l'ID du fournisseur sélectionné (colonne 1, index 0)
    selectedFourID = CStr(LB_listefourni.List(selectedIndex, 0))
   
    ' NOUVEAU : Pop-up de confirmation avant ouverture de la fiche fournisseur
    reponse = MsgBox("Voulez-vous ouvrir la fiche du fournisseur ID : " & selectedFourID & " ?", vbYesNo + vbQuestion, "Confirmation ouverture fiche")
    If reponse = vbNo Then
        Exit Sub
    End If
   
    On Error GoTo ErreurDoubleClick
   
    Application.ScreenUpdating = False ' Fluidité UI lors de l'ouverture
   
    ' Crée et affiche l'UserForm Fiche_Four
    Set ficheFour = VBA.UserForms.Add("Fiche_Four") ' Charge le form par nom (assure-toi qu'il existe dans VBAProject)
   
    ' Charge le fournisseur directement via la méthode publique (adapte si ta procédure d'init est différente)
    ficheFour.LoadFournisseur selectedFourID ' Appel à LoadFournisseur (ajoute cette méthode publique dans Fiche_Four si absente)
   
    ' Affiche le form en mode modal (bloque l'UI jusqu'à fermeture)
    ficheFour.Show vbModal
   
    ' Optionnel : Actualise la ListBox après fermeture (si Fiche_Four permet des edits)
    If Not IsEmpty(currentArtisanID) Then Call FillLBFULL(currentArtisanID)
   
    ' Nettoie la référence
    Unload ficheFour
    Set ficheFour = Nothing
   
    Application.ScreenUpdating = True ' Restaure l'UI
   
    Exit Sub
   
ErreurDoubleClick:
    MsgBox "Erreur lors de l'ouverture de Fiche_Four : " & Err.Description & vbCrLf & _
           "(Vérifiez que l'UserForm 'Fiche_Four' existe et a une méthode LoadFournisseur.)", vbError
    If Not ficheFour Is Nothing Then
        Unload ficheFour
        Set ficheFour = Nothing
    End If
    Application.ScreenUpdating = True ' Restaure même en erreur
End Sub

Private Sub TB_catearti_Change()
    
End Sub

Private Sub TB_coorarti_Change()
    
End Sub

Private Sub TB_dispoarti_Change()
    
End Sub

Private Sub TB_emailarti_Change()
    
End Sub

Private Sub TB_idarti_Change()
    
End Sub

Private Sub TB_nomarti_Change()
    
End Sub

Private Sub TB_Stockprod_Change()
End Sub

Private Sub TB_Telarti_Change()
   
End Sub

Private Sub TextBox4_Change()
    
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler ' Gestion d'erreur globale
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("ARTISAN")
    Set artisanTable = ws.ListObjects("plg_artisant") ' Stocker la table pour réutilisation
    ' Vérifier si la table existe
    If artisanTable Is Nothing Then
        MsgBox "Erreur : La table 'plg_artisant' n'existe pas dans la feuille 'ARTISAN'." & vbCrLf & _
               "Vérifiez le nom de la table.", vbCritical
        Exit Sub
    End If
    ' Vérifier s'il y a des données
    If artisanTable.ListRows.count = 0 Then
        MsgBox "Aucune donnée dans la table 'plg_artisant'.", vbExclamation
        Exit Sub
    End If
    ' Charger la table POSSEDE
    Dim ws_pos As Worksheet
    Set ws_pos = ThisWorkbook.Worksheets("POSSEDE")
    Set possedeTable = ws_pos.ListObjects("plg_possede")
    If possedeTable Is Nothing Then
        MsgBox "Erreur : La table 'plg_possede' n'existe pas dans la feuille 'POSSEDE'." & vbCrLf & _
               "Vérifiez le nom de la table.", vbCritical
        Exit Sub
    End If
    If possedeTable.ListRows.count = 0 Then
        Debug.Print "Aucune donnée dans la table 'plg_possede' (mais on continue)."
    End If
    ' Charger la table PRODUIT
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
    ' Charger la table APPROVISIONNE
    Dim ws_app As Worksheet
    Set ws_app = ThisWorkbook.Worksheets("APPROVISIONNE")
    Set approvisionneTable = ws_app.ListObjects("plg_approvisionne")
    If approvisionneTable Is Nothing Then
        MsgBox "Erreur : La table 'plg_approvisionne' n'existe pas dans la feuille 'APPROVISIONNE'." & vbCrLf & _
               "Vérifiez le nom de la table.", vbCritical
        Exit Sub
    End If
    If approvisionneTable.ListRows.count = 0 Then
        Debug.Print "Aucune donnée dans la table 'plg_approvisionne' (mais on continue)."
    End If
    ' NOUVEAU : Charger la table FOURNISSEUR
    Dim ws_fourn As Worksheet
    Set ws_fourn = ThisWorkbook.Worksheets("FOURNISSEUR")
    Set fournisseurTable = ws_fourn.ListObjects("plg_fournisseur")
    If fournisseurTable Is Nothing Then
        MsgBox "Erreur : La table 'plg_fournisseur' n'existe pas dans la feuille 'FOURNISSEUR'." & vbCrLf & _
               "Vérifiez le nom de la table.", vbCritical
        Exit Sub
    End If
    If fournisseurTable.ListRows.count = 0 Then
        Debug.Print "Aucune donnée dans la table 'plg_fournisseur' (mais on continue)."
    End If
    ' Préparer un array pour les 2 premières colonnes seulement (ID et Nom)
    Dim dataArray As Variant
    dataArray = artisanTable.ListColumns(1).DataBodyRange.Resize(, 2).Value ' Colonnes 1 et 2
    ' Assurer que c'est un array 2D (fix pour Type Mismatch potentiel)
    If Not IsArray(dataArray) Then
        ReDim dataArray(1 To 1, 1 To 2)
        dataArray(1, 1) = artisanTable.ListColumns(1).DataBodyRange.Cells(1, 1).Value
        dataArray(1, 2) = artisanTable.ListColumns(2).DataBodyRange.Cells(1, 1).Value
    End If
    ' Stocker les données originales
    originalData = dataArray
    ' Remplir la ComboBox avec les données complètes
    With CB_recherarti
        .Clear
        .ColumnCount = 2
        .BoundColumn = 1 ' .Value renverra toujours l'ID sélectionné
        .ColumnWidths = "50pt;150pt" ' Affiche ID (50pt) + Nom (150pt) pour distinguer les homonymes
        .MatchEntry = 0 ' fmMatchEntryNone : Permet saisie libre pour le filtrage
        .List = dataArray ' Charge l'array complet
    End With
    ' Initialiser la ComboBox CB_recherproduit (2 colonnes : ID masqué + Nom visible)
    With CB_recherproduit
        .Clear
        .ColumnCount = 2 ' Deux colonnes : id_produit (masqué) + nom_produit (visible)
        .BoundColumn = 1 ' .Value renverra l'id_produit
        .ColumnWidths = "0pt;150pt" ' Masque col1 (ID), affiche col2 (Nom)
        .MatchEntry = 0 ' Permet saisie libre si besoin
        .AddItem "Sélectionnez un artisan d'abord..." ' Placeholder (sera remplacé)
    End With
    ' MODIFIÉ : Initialiser la ListBox LB_listefourni (2 colonnes : ID + Nom fournisseur)
    With LB_listefourni
        .Clear
        .ColumnCount = 2 ' Deux colonnes : id_fournisseur + nom_fournisseur
        .ColumnWidths = "50pt;150pt" ' ID petit, Nom large
        .AddItem "Aucun fournisseur sélectionné" ' Placeholder initial
    End With
    ' Initialiser TB_stockprod vide
    TB_stockprod.Value = ""
    ' Debug : Affiche le nombre d'éléments chargés et type du premier ID
    Debug.Print "ComboBox remplie avec " & CB_recherarti.ListCount & " éléments."
    If CB_recherarti.ListCount > 0 Then
        Debug.Print "Premier élément : ID=" & CB_recherarti.List(0, 0) & ", Nom=" & CB_recherarti.List(0, 1)
        Debug.Print "Type du premier ID : " & TypeName(CB_recherarti.List(0, 0))
    End If
    Exit Sub
ErrHandler:
    MsgBox "Erreur " & Err.Number & " : " & Err.Description & vbCrLf & _
           "Ligne : " & Erl, vbCritical
End Sub

' Événement KeyUp pour filtrage
Private Sub CB_recherarti_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrHandler
    ' Ignorer si pas de données originales ou si touche spéciale (ex. flèches, Entrée)
    If Not IsArray(originalData) Or UBound(originalData, 1) = 0 Then Exit Sub
    Dim searchText As String
    searchText = UCase(Trim(CB_recherarti.Text)) ' Texte tapé, majuscules pour insensible à la casse
    ' Si texte vide ou effacement, recharger la liste complète
    If Len(searchText) = 0 Then
        Call LoadFullList
        Exit Sub
    End If
    ' Filtrer les éléments où le nom commence par searchText (préfixe)
    Dim filteredArray() As Variant
    Dim matchCount As Long
    matchCount = 0
    Dim i As Long
    For i = 1 To UBound(originalData, 1)
        If Len(originalData(i, 2)) > 0 And UCase(Left(originalData(i, 2), Len(searchText))) = searchText Then
            matchCount = matchCount + 1
        End If
    Next i
    If matchCount = 0 Then
        ' Option : Vider la liste ou recharger complète ? Ici, vider pour signaler "pas de match"
        CB_recherarti.Clear
        Debug.Print "Aucun nom correspondant à '" & CB_recherarti.Text & "'"
        Exit Sub
    End If
    ' Construire l'array filtré
    ReDim filteredArray(1 To matchCount, 1 To 2)
    Dim j As Long
    j = 1
    For i = 1 To UBound(originalData, 1)
        If Len(originalData(i, 2)) > 0 And UCase(Left(originalData(i, 2), Len(searchText))) = searchText Then
            filteredArray(j, 1) = originalData(i, 1) ' ID
            filteredArray(j, 2) = originalData(i, 2) ' Nom
            j = j + 1
        End If
    Next i
    ' Charger la liste filtrée
    With CB_recherarti
        .Clear
        .List = filteredArray
    End With
    Debug.Print matchCount & " nom(s) correspondant à '" & searchText & "'"
    Exit Sub
ErrHandler:
    MsgBox "Erreur dans KeyUp : " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

' Procédure helper pour recharger la liste complète
Private Sub LoadFullList()
    With CB_recherarti
        .Clear
        .List = originalData
    End With
    Debug.Print "Liste complète rechargée (" & UBound(originalData, 1) & " éléments)."
End Sub

' Procédure pour remplir LB avec TOUS les fournisseurs des produits de l'artisan
Private Sub FillLBFULL(artisanID As Variant)
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False '* AMÉLIORÉ : Fluidité UI
  
    If IsEmpty(artisanID) Or artisanID = "" Then
        With LB_listefourni
            .Clear
            .AddItem "Aucun artisan sélectionné"
        End With
        Application.ScreenUpdating = True
        Exit Sub
    End If
    If possedeTable Is Nothing Or produitTable Is Nothing Or approvisionneTable Is Nothing Or fournisseurTable Is Nothing Then
        With LB_listefourni
            .Clear
            .AddItem "Tables non disponibles"
        End With
        Application.ScreenUpdating = True
        Exit Sub
    End If
    If possedeTable.ListRows.count = 0 Or approvisionneTable.ListRows.count = 0 Then
        With LB_listefourni
            .Clear
            .AddItem "Aucune donnée"
        End With
        Application.ScreenUpdating = True
        Exit Sub
    End If
    ' Étape 1 : Collecter les id_produit de l'artisan depuis plg_possede (uniques)
    Dim uniqueProds As Object
    Set uniqueProds = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To possedeTable.ListRows.count
        If possedeTable.ListColumns(2).DataBodyRange.Cells(i, 1).Value = artisanID Then
            Dim prodID As Variant
            prodID = possedeTable.ListColumns(1).DataBodyRange.Cells(i, 1).Value
            If Not uniqueProds.Exists(prodID) Then
                uniqueProds.Add prodID, True
            End If
        End If
    Next i
    If uniqueProds.count = 0 Then
        With LB_listefourni
            .Clear
            .AddItem "Aucun produit pour cet artisan"
        End With
        Debug.Print "Aucun produit trouvé pour ID artisan : " & artisanID
        Application.ScreenUpdating = True
        Exit Sub
    End If
    ' Étape 2 : Collecter les id_fournisseur uniques depuis plg_approvisionne pour ces produits
    Dim uniqueFournisseurs As Object
    Set uniqueFournisseurs = CreateObject("Scripting.Dictionary")
    Dim keys As Variant
    keys = uniqueProds.keys
    Dim j As Long
    For j = 0 To UBound(keys)
        Dim thisProdID As Variant
        thisProdID = keys(j)
    
        Dim k As Long
        For k = 1 To approvisionneTable.ListRows.count
            If approvisionneTable.ListColumns(2).DataBodyRange.Cells(k, 1).Value = thisProdID Then
                Dim fourID As Variant
                fourID = approvisionneTable.ListColumns(1).DataBodyRange.Cells(k, 1).Value ' id_fournisseur (col1)
                If Not uniqueFournisseurs.Exists(fourID) Then
                    uniqueFournisseurs.Add fourID, True
                End If
            End If
        Next k
    Next j
    If uniqueFournisseurs.count = 0 Then
        With LB_listefourni
            .Clear
            .AddItem "Aucun fournisseur pour les produits de cet artisan"
        End With
        Debug.Print "Aucun fournisseur trouvé pour les produits de l'artisan : " & artisanID
        Application.ScreenUpdating = True
        Exit Sub
    End If
    ' Étape 3 : Pour chaque id_fournisseur, lookup nom dans plg_fournisseur (col A = id_fournisseur, col B = nom supposé)
    Dim countFourn As Long
    countFourn = uniqueFournisseurs.count
    Dim lbFournisseurs() As Variant
    ReDim lbFournisseurs(1 To countFourn + 1, 1 To 2) ' +1 pour la ligne d'en-têtes ; 2 cols : ID + Nom
    ' Ligne 1 : En-têtes
    lbFournisseurs(1, 1) = "ID Fournisseur"
    lbFournisseurs(1, 2) = "Nom Fournisseur"
    Dim row As Long
    row = 2 ' Commencer les données à la ligne 2
    Dim fourKeys As Variant
    fourKeys = uniqueFournisseurs.keys
    For j = 0 To UBound(fourKeys)
        Dim thisFourID As Variant
        thisFourID = fourKeys(j)
    
        ' Lookup nom_fournisseur
        Dim fourRowIndex As Variant
        fourRowIndex = Application.match(thisFourID, fournisseurTable.ListColumns(1).DataBodyRange, 0)
    
        Dim nomFour As String
        If IsError(fourRowIndex) Then
            nomFour = "Fournisseur inconnu (ID: " & thisFourID & ")"
            Debug.Print "Nom fournisseur non trouvé pour ID : " & thisFourID
        Else
            nomFour = fournisseurTable.ListColumns(2).DataBodyRange(fourRowIndex, 1).Value ' Nom (col B supposé)
        End If
    
        lbFournisseurs(row, 1) = thisFourID ' ID
        lbFournisseurs(row, 2) = nomFour ' Nom
        row = row + 1
    Next j
    ' Remplir la ListBox
    With LB_listefourni
        .Clear
        .List = lbFournisseurs
    End With
    Debug.Print countFourn & " fournisseur(s) chargé(s) (tous) pour artisan : " & artisanID
    Application.ScreenUpdating = True '* AMÉLIORÉ : Restaure UI
    Exit Sub
ErrHandler:
    MsgBox "Erreur dans FillLBFULL : " & Err.Number & " - " & Err.Description, vbExclamation
    With LB_listefourni
        .Clear
        .AddItem "Erreur lors du chargement"
    End With
    Application.ScreenUpdating = True
End Sub

' Procédure pour filtrer LB avec les fournisseurs d'un produit spécifique (avec lookup nom ET EN-TÊTES)
Private Sub FilterLBByProduct(artisanID As Variant, prodID As Variant)
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False '* AMÉLIORÉ : Fluidité UI
  
    If IsEmpty(artisanID) Or IsEmpty(prodID) Then
        With LB_listefourni
            .Clear
            .AddItem "Aucun filtre valide"
        End With
        Application.ScreenUpdating = True
        Exit Sub
    End If
    ' Vérifier si le produit appartient à l'artisan (optionnel, pour cohérence)
    Dim hasProduct As Boolean
    hasProduct = False
    Dim i As Long
    For i = 1 To possedeTable.ListRows.count
        If possedeTable.ListColumns(2).DataBodyRange.Cells(i, 1).Value = artisanID And _
           possedeTable.ListColumns(1).DataBodyRange.Cells(i, 1).Value = prodID Then
            hasProduct = True
            Exit For
        End If
    Next i
    If Not hasProduct Then
        With LB_listefourni
            .Clear
            .AddItem "Ce produit n'appartient pas à cet artisan"
        End With
        Application.ScreenUpdating = True
        Exit Sub
    End If
    If approvisionneTable Is Nothing Or fournisseurTable Is Nothing Then
        With LB_listefourni
            .Clear
            .AddItem "Tables non disponibles"
        End With
        Application.ScreenUpdating = True
        Exit Sub
    End If
    If approvisionneTable.ListRows.count = 0 Then
        With LB_listefourni
            .Clear
            .AddItem "Aucune donnée dans APPROVISIONNE"
        End With
        Application.ScreenUpdating = True
        Exit Sub
    End If
    ' Collecter les id_fournisseur uniques pour ce prodID
    Dim uniqueFournisseurs As Object
    Set uniqueFournisseurs = CreateObject("Scripting.Dictionary")
    Dim k As Long
    For k = 1 To approvisionneTable.ListRows.count
        If approvisionneTable.ListColumns(2).DataBodyRange.Cells(k, 1).Value = prodID Then
            Dim fourID As Variant
            fourID = approvisionneTable.ListColumns(1).DataBodyRange.Cells(k, 1).Value ' id_fournisseur (col1)
            If Not uniqueFournisseurs.Exists(fourID) Then
                uniqueFournisseurs.Add fourID, True
            End If
        End If
    Next k
    If uniqueFournisseurs.count = 0 Then
        With LB_listefourni
            .Clear
            .AddItem "Aucun fournisseur pour ce produit"
        End With
        Debug.Print "Aucun fournisseur trouvé pour produit " & prodID
        Application.ScreenUpdating = True
        Exit Sub
    End If
    ' Étape : Pour chaque id_fournisseur, lookup nom dans plg_fournisseur
    Dim countFourn As Long
    countFourn = uniqueFournisseurs.count
    Dim lbFournisseurs() As Variant
    ReDim lbFournisseurs(1 To countFourn + 1, 1 To 2) ' +1 pour la ligne d'en-têtes ; 2 cols : ID + Nom
    ' Ligne 1 : En-têtes
    lbFournisseurs(1, 1) = "ID Fournisseur"
    lbFournisseurs(1, 2) = "Nom Fournisseur"
    Dim row As Long
    row = 2 ' Commencer les données à la ligne 2
    Dim fourKeys As Variant
    fourKeys = uniqueFournisseurs.keys
    Dim j As Long
    For j = 0 To UBound(fourKeys)
        Dim thisFourID As Variant
        thisFourID = fourKeys(j)
    
        ' Lookup nom_fournisseur
        Dim fourRowIndex As Variant
        fourRowIndex = Application.match(thisFourID, fournisseurTable.ListColumns(1).DataBodyRange, 0)
    
        Dim nomFour As String
        If IsError(fourRowIndex) Then
            nomFour = "Fournisseur inconnu (ID: " & thisFourID & ")"
            Debug.Print "Nom fournisseur non trouvé pour ID : " & thisFourID
        Else
            nomFour = fournisseurTable.ListColumns(2).DataBodyRange(fourRowIndex, 1).Value ' Nom (col B supposé)
        End If
    
        lbFournisseurs(row, 1) = thisFourID ' ID
        lbFournisseurs(row, 2) = nomFour ' Nom
        row = row + 1
    Next j
    ' Remplir la ListBox
    With LB_listefourni
        .Clear
        .List = lbFournisseurs
    End With
    Debug.Print countFourn & " fournisseur(s) chargé(s) (filtré) pour produit " & prodID & " et artisan " & artisanID
    Application.ScreenUpdating = True '* AMÉLIORÉ : Restaure UI
    Exit Sub
ErrHandler:
    MsgBox "Erreur dans FilterLBByProduct : " & Err.Number & " - " & Err.Description, vbExclamation
    With LB_listefourni
        .Clear
        .AddItem "Erreur lors du filtrage"
    End With
    Application.ScreenUpdating = True
End Sub

' Procédure pour charger les noms de produits uniques de l'artisan (remplit seulement CB_recherproduit)
Private Sub LoadArtisanProducts(selectedID As Variant)
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False '* AMÉLIORÉ : Fluidité UI
  
    currentArtisanID = selectedID ' Stocker pour filtrage futur
    If possedeTable Is Nothing Or produitTable Is Nothing Then
        CB_recherproduit.Clear
        CB_recherproduit.AddItem "Tables POSSEDE/PRODUIT non disponibles"
        TB_stockprod.Value = "" ' Vider le stock
        Application.ScreenUpdating = True
        Exit Sub
    End If
    If possedeTable.ListRows.count = 0 Then
        CB_recherproduit.Clear
        CB_recherproduit.AddItem "Aucune donnée dans POSSEDE"
        TB_stockprod.Value = "" ' Vider le stock
        Application.ScreenUpdating = True
        Exit Sub
    End If
    ' Vider le stock lors du changement d'artisan
    TB_stockprod.Value = ""
    ' Vérifier si l'ID est valide (String ou Long)
    Dim idToMatch As Variant
    idToMatch = selectedID
    ' Compter les produits uniques
    Dim productCount As Long
    productCount = 0
    Dim uniqueProds As Object ' Dictionary pour uniques
    Set uniqueProds = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To possedeTable.ListRows.count
        If possedeTable.ListColumns(2).DataBodyRange.Cells(i, 1).Value = idToMatch Then
            Dim prodID As Variant
            prodID = possedeTable.ListColumns(1).DataBodyRange.Cells(i, 1).Value
            If Not uniqueProds.Exists(prodID) Then
                uniqueProds.Add prodID, True
                productCount = productCount + 1
            End If
        End If
    Next i
    If productCount = 0 Then
        CB_recherproduit.Clear
        CB_recherproduit.AddItem "Aucun produit pour cet artisan"
        Debug.Print "Aucun produit trouvé pour ID artisan : " & idToMatch
        Application.ScreenUpdating = True
        Exit Sub
    End If
    ' Construire l'array pour CB_recherproduit (uniques : ID + Nom)
    Dim cbProducts() As Variant
    ReDim cbProducts(1 To productCount, 1 To 2)
    Dim keys As Variant
    keys = uniqueProds.keys
    Dim j As Long
    For j = 0 To UBound(keys)
        Dim thisProdID As Variant
        thisProdID = keys(j)
    
        ' Lookup nom
        Dim prodRowIndex As Variant
        prodRowIndex = Application.match(thisProdID, produitTable.ListColumns(1).DataBodyRange, 0)
    
        Dim nomProd As String
        If IsError(prodRowIndex) Then
            nomProd = "Nom inconnu (ID: " & thisProdID & ")"
            Debug.Print "Nom produit non trouvé pour ID : " & thisProdID
        Else
            nomProd = produitTable.ListColumns(2).DataBodyRange(prodRowIndex, 1).Value
        End If
    
        cbProducts(j + 1, 1) = thisProdID ' ID
        cbProducts(j + 1, 2) = nomProd ' Nom
    Next j
    ' Remplir CB_recherproduit
    With CB_recherproduit
        .Clear
        .List = cbProducts
        .ListIndex = -1 ' Pas de sélection auto
    End With
    ' Par défaut, remplir LB avec tous les fournisseurs
    Call FillLBFULL(idToMatch)
    Debug.Print productCount & " produit(s) unique(s) chargé(s) pour ID artisan : " & idToMatch
    Application.ScreenUpdating = True '* AMÉLIORÉ : Restaure UI
    Exit Sub
ErrHandler:
    MsgBox "Erreur dans LoadArtisanProducts : " & Err.Number & " - " & Err.Description, vbExclamation
    CB_recherproduit.Clear
    CB_recherproduit.AddItem "Erreur lors du chargement des produits"
    TB_stockprod.Value = "" ' Vider en cas d'erreur
    Application.ScreenUpdating = True
End Sub

' Événement Change : AJOUT du remplissage des TextBoxes
Private Sub CB_recherarti_Change()
    On Error GoTo ErrHandler
    If CB_recherarti.ListIndex = -1 Then Exit Sub ' Évite les boucles au chargement
    Dim selectedID As Variant ' Accepte String ou Long
    selectedID = CB_recherarti.Value ' Récupère l'ID sélectionné
    ' Après sélection, recharger la liste complète
    Call LoadFullList
    ' Lookup de la ligne dans la table et remplissage des TextBoxes
    Dim rowIndex As Variant
    rowIndex = Application.match(selectedID, artisanTable.ListColumns(1).DataBodyRange, 0) ' Trouve la ligne (exact match)
    If IsError(rowIndex) Then
        MsgBox "Artisan non trouvé dans la table (ID: " & selectedID & ").", vbExclamation
        Exit Sub
    End If
    ' Remplir les TextBoxes (ajustez les indices de colonnes si l'ordre diffère dans votre table)
    TB_idarti.Value = selectedID ' Colonne 1 : ID
    TB_nomarti.Value = artisanTable.ListColumns(2).DataBodyRange(rowIndex, 1).Value ' Colonne 2 : Nom
    TB_coorarti.Value = artisanTable.ListColumns(3).DataBodyRange(rowIndex, 1).Value ' Colonne 3 : Coordonnées
    TB_Telarti.Value = artisanTable.ListColumns(4).DataBodyRange(rowIndex, 1).Value ' Colonne 4 : Téléphone
    TB_emailarti.Value = artisanTable.ListColumns(5).DataBodyRange(rowIndex, 1).Value ' Colonne 5 : Email
    TB_dispoarti.Value = artisanTable.ListColumns(6).DataBodyRange(rowIndex, 1).Value ' Colonne 6 : Disponibilité
    TB_catearti.Value = artisanTable.ListColumns(7).DataBodyRange(rowIndex, 1).Value ' Colonne 7 : Catégorie
    ' Vider le stock lors du changement d'artisan
    TB_stockprod.Value = ""
    ' Charger les produits correspondants à cet artisan (remplit CB et LB avec tous les fournisseurs par défaut)
    Call LoadArtisanProducts(selectedID)
    ' Debug : Affiche les valeurs remplies
    Debug.Print "Artisan chargé - ID: " & selectedID & ", Nom: " & TB_nomarti.Value & ", Tel: " & TB_Telarti.Value
    Exit Sub
ErrHandler:
    MsgBox "Erreur dans Change : " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

