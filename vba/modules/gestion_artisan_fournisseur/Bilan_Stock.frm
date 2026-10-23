VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Bilan_Stock 
   Caption         =   "Bilan Stock"
   ClientHeight    =   4704
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7920
   OleObjectBlob   =   "Bilan_Stock.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Bilan_Stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variables module (ajoutez en haut du module, avant les Subs pour le nouveau UserForm Stock/Possédé)
Private possedeTable As ListObject
Private produitTable As ListObject
Private artisanTable As ListObject

Private Sub CB_validerseuil_Click()
    On Error GoTo ErrHandler
    
    ' Vérifier si une ligne est sélectionnée (ignorer en-tête à ListIndex=0)
    If LB_listeproduit.ListIndex <= 0 Then
        MsgBox "Sélectionnez une ligne de stock pour modifier le seuil.", vbExclamation
        Exit Sub
    End If
    
    Dim selectedIndex As Long
    selectedIndex = LB_listeproduit.ListIndex ' Direct, car ListIndex=1 pour première ligne de données
    
    Dim prodID As Variant
    prodID = LB_listeproduit.List(selectedIndex, 0) ' Colonne 0 : ID Produit (cachée)
    
    If IsEmpty(prodID) Or prodID = "" Then
        MsgBox "ID produit non valide.", vbExclamation
        Exit Sub
    End If
    
    Dim newSeuil As Variant
    newSeuil = TB_seuilcrit.Value
    If Not IsNumeric(newSeuil) Then
        MsgBox "Le seuil doit être un nombre valide.", vbExclamation
        Exit Sub
    End If
    
    ' Confirmer la mise à jour
    If MsgBox("Mettre à jour le seuil critique du produit '" & TB_nomprod.Value & "' à " & newSeuil & " ?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    ' Trouver la ligne du produit dans plg_produit
    Dim prodRowIndex As Variant
    prodRowIndex = Application.match(prodID, produitTable.ListColumns(1).DataBodyRange, 0)
    
    If IsError(prodRowIndex) Then
        MsgBox "Produit non trouvé dans la table PRODUIT.", vbExclamation
        Exit Sub
    End If
    
    ' Mettre à jour la colonne E (index 5)
    produitTable.ListColumns(5).DataBodyRange(prodRowIndex, 1).Value = newSeuil
    
    ' Confirmer et rafraîchir la liste
    MsgBox "Seuil critique mis à jour avec succès !", vbInformation
    
    ' Rafraîchir la ListBox pour refléter les changements (recharge avec filtre actuel)
    Dim currentFilter As String
    currentFilter = CB_etatstock.Value
    Call FillLBStock(currentFilter)
    
    ' Vider les TextBoxes après mise à jour
    TB_nomprod.Value = ""
    TB_seuilcrit.Value = ""
    
    Debug.Print "Seuil mis à jour pour produit ID " & prodID & " à " & newSeuil
    
    Exit Sub
    
ErrHandler:
    MsgBox "Erreur dans CB_validerseuil_Click : " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub LB_listeproduit_Click()
    On Error GoTo ErrHandler
    
    ' Ignorer si clic sur en-tête (ListIndex=0) ou rien sélectionné (-1)
    If LB_listeproduit.ListIndex <= 0 Then Exit Sub
    
    Dim selectedIndex As Long
    selectedIndex = LB_listeproduit.ListIndex ' Direct, car ListIndex=1 pour première ligne de données (lbData(1))
    
    ' Récupérer le nom du produit (colonne 1)
    TB_nomprod.Value = LB_listeproduit.List(selectedIndex, 1)
    
    ' Récupérer l'ID produit (colonne 0, cachée) pour lookup seuil
    Dim prodID As Variant
    prodID = LB_listeproduit.List(selectedIndex, 0)
    
    If Not IsEmpty(prodID) And prodID <> "" Then
        ' Lookup seuil critique dans plg_produit (colonne 5)
        Dim prodRowIndex As Variant
        prodRowIndex = Application.match(prodID, produitTable.ListColumns(1).DataBodyRange, 0)
        
        If Not IsError(prodRowIndex) Then
            Dim currentSeuil As Variant
            currentSeuil = produitTable.ListColumns(5).DataBodyRange(prodRowIndex, 1).Value
            TB_seuilcrit.Value = currentSeuil
        Else
            TB_seuilcrit.Value = "Non trouvé"
            Debug.Print "Seuil non trouvé pour ID produit : " & prodID
        End If
    Else
        TB_seuilcrit.Value = ""
    End If
    
    Debug.Print "Ligne cliquée - Produit : " & TB_nomprod.Value & ", Seuil : " & TB_seuilcrit.Value
    
    Exit Sub
    
ErrHandler:
    MsgBox "Erreur dans LB_listeproduit_Click : " & Err.Number & " - " & Err.Description, vbExclamation
    TB_nomprod.Value = ""
    TB_seuilcrit.Value = ""
End Sub

Private Sub TB_nomprod_Change()
    
End Sub

Private Sub TB_seuilcrit_Change()

End Sub

Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler
    Dim ws_pos As Worksheet
    Set ws_pos = ThisWorkbook.Worksheets("POSSEDE")
    Set possedeTable = ws_pos.ListObjects("plg_possede")
    If possedeTable Is Nothing Then
        MsgBox "Erreur : La table 'plg_possede' n'existe pas.", vbCritical
        Exit Sub
    End If
    Dim ws_prod As Worksheet
    Set ws_prod = ThisWorkbook.Worksheets("PRODUIT")
    Set produitTable = ws_prod.ListObjects("plg_produit")
    If produitTable Is Nothing Then
        MsgBox "Erreur : La table 'plg_produit' n'existe pas.", vbCritical
        Exit Sub
    End If
    Dim ws_art As Worksheet
    Set ws_art = ThisWorkbook.Worksheets("ARTISAN")
    Set artisanTable = ws_art.ListObjects("plg_artisant")
    If artisanTable Is Nothing Then
        MsgBox "Erreur : La table 'plg_artisant' n'existe pas.", vbCritical
        Exit Sub
    End If
    With LB_listeproduit
        .Clear
        .ColumnCount = 4 ' 4 colonnes : ID Prod (hidden), Nom Prod, Nom Art, Quantité
        .ColumnWidths = "0pt;100pt;100pt;50pt" ' Cacher col 0 (ID)
    End With
    With CB_etatstock
        .Clear
        .AddItem "Tout"
        .AddItem "Critique"
        .AddItem "OK"
        .ListIndex = 0
    End With
    ' Vider initialement les TextBoxes
    TB_nomprod.Value = ""
    TB_seuilcrit.Value = ""
    Call FillLBStock("Tout")
    Exit Sub
ErrHandler:
    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical
End Sub

Private Sub FillLBStock(Optional filterEtat As String = "Tout")
    On Error GoTo LocalErrHandler
    Dim filteredData As Collection
    Set filteredData = New Collection
    ' DÉCLARATIONS UNIQUES (OBLIGATOIRE EN VBA)
    Dim thisProdID As Variant
    Dim thisArtID As Variant
    Dim quantite As Variant
    Dim prodRowIndex As Variant
    Dim artRowIndex As Variant
    Dim seuilCritique As Variant
    Dim i As Long
    For i = 1 To possedeTable.ListRows.count
        thisProdID = possedeTable.ListColumns(1).DataBodyRange.Cells(i, 1).Value
        thisArtID = possedeTable.ListColumns(2).DataBodyRange.Cells(i, 1).Value
        quantite = possedeTable.ListColumns(3).DataBodyRange.Cells(i, 1).Value
        prodRowIndex = Application.match(thisProdID, produitTable.ListColumns(1).DataBodyRange, 0)
        seuilCritique = 0
        If Not IsError(prodRowIndex) Then
            seuilCritique = produitTable.ListColumns(5).DataBodyRange(prodRowIndex, 1).Value
        End If
        Dim includeRow As Boolean
        includeRow = False
        If filterEtat = "Tout" Then
            includeRow = True
        ElseIf filterEtat = "Critique" Then
            If IsNumeric(quantite) And IsNumeric(seuilCritique) Then
                If quantite < seuilCritique Then includeRow = True
            End If
        ElseIf filterEtat = "OK" Then
            If IsNumeric(quantite) And IsNumeric(seuilCritique) Then
                If quantite >= seuilCritique Then includeRow = True
            Else
                includeRow = True
            End If
        End If
        If includeRow Then filteredData.Add i
    Next i
    If filteredData.count = 0 Then
        LB_listeproduit.Clear
        LB_listeproduit.AddItem "Aucun stock " & filterEtat
        Exit Sub
    End If
    Dim lbData() As Variant
    ReDim lbData(0 To filteredData.count, 0 To 3) ' 4 colonnes : ID Prod (0, hidden), Nom Prod (1), Nom Art (2), Quantité (3)
    ' En-têtes (seulement pour colonnes visibles)
    lbData(0, 1) = "Nom Produit"
    lbData(0, 2) = "Nom Artisan"
    lbData(0, 3) = "Quantité Stock"
    Dim j As Long
    For j = 1 To filteredData.count
        Dim rowIndex As Long
        rowIndex = filteredData(j)
        thisProdID = possedeTable.ListColumns(1).DataBodyRange.Cells(rowIndex, 1).Value
        thisArtID = possedeTable.ListColumns(2).DataBodyRange.Cells(rowIndex, 1).Value
        quantite = possedeTable.ListColumns(3).DataBodyRange.Cells(rowIndex, 1).Value
        prodRowIndex = Application.match(thisProdID, produitTable.ListColumns(1).DataBodyRange, 0)
        seuilCritique = 0
        Dim nomProd As String
        If IsError(prodRowIndex) Then
            nomProd = "Produit inconnu (ID: " & thisProdID & ")"
        Else
            nomProd = produitTable.ListColumns(2).DataBodyRange(prodRowIndex, 1).Value
            seuilCritique = produitTable.ListColumns(5).DataBodyRange(prodRowIndex, 1).Value
        End If
        artRowIndex = Application.match(thisArtID, artisanTable.ListColumns(1).DataBodyRange, 0)
        Dim nomArt As String
        If IsError(artRowIndex) Then
            nomArt = "Artisan inconnu (ID: " & thisArtID & ")"
        Else
            nomArt = artisanTable.ListColumns(2).DataBodyRange(artRowIndex, 1).Value
        End If
        lbData(j, 0) = thisProdID ' ID Prod (caché)
        lbData(j, 1) = nomProd ' Nom Prod
        lbData(j, 2) = nomArt ' Nom Art
        If IsNumeric(quantite) And IsNumeric(seuilCritique) Then
            If quantite < seuilCritique Then
                lbData(j, 3) = quantite & " *CRITIQUE*"
            Else
                lbData(j, 3) = quantite
            End If
        Else
            lbData(j, 3) = quantite
        End If
    Next j
    LB_listeproduit.Clear
    LB_listeproduit.List = lbData
    Exit Sub
LocalErrHandler:
    MsgBox "Erreur FillLBStock : " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub CB_etatstock_Change()
    On Error GoTo LocalErrHandler
    Call FillLBStock(CB_etatstock.Value)
    Exit Sub
LocalErrHandler:
    Call FillLBStock("Tout")
End Sub

Private Sub UserForm_Click()
    
End Sub
