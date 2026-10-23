VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Montant_moyen 
   Caption         =   "Montant moyen"
   ClientHeight    =   4992
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7932
   OleObjectBlob   =   "Montant_moyen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Montant_moyen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB_recherregion_Change()
    On Error GoTo ErrHandler
    
    Dim selectedRegion As String
    selectedRegion = CB_recherregion.Value
    
    If selectedRegion = "" Then
        LB_listemoyencate.Clear
        CB_rechercate.Clear
        TB_coutmoyen.Value = ""
        Exit Sub
    End If
    
    ' Effacer la ListBox et la ComboBox avant de les remplir
    LB_listemoyencate.Clear
    CB_rechercate.Clear
    TB_coutmoyen.Value = ""
    
    ' Vérifier l'existence de la feuille COMMANDE
    Dim wsCommande As Worksheet
    On Error Resume Next
    Set wsCommande = ThisWorkbook.Worksheets("COMMANDE")
    On Error GoTo ErrHandler
    If wsCommande Is Nothing Then
        MsgBox "Feuille 'COMMANDE' non trouvée.", vbCritical
        Exit Sub
    End If
    
    ' Vérifier l'existence du tableau plg_commande
    Dim tblCommande As ListObject
    On Error Resume Next
    Set tblCommande = wsCommande.ListObjects("plg_commande")
    On Error GoTo ErrHandler
    If tblCommande Is Nothing Then
        MsgBox "Tableau 'plg_commande' non trouvé dans la feuille 'COMMANDE'. Vérifiez le nom exact.", vbCritical
        Exit Sub
    End If
    
    ' Vérifier si le tableau a des données
    If tblCommande.DataBodyRange Is Nothing Then
        MsgBox "Le tableau 'plg_commande' n'a pas de données.", vbCritical
        Exit Sub
    End If
    
    ' Utiliser un Dictionary pour obtenir les ID_artisant uniques de la colonne J (colonne 10)
    Dim dictArtisans As Object
    Set dictArtisans = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 1 To tblCommande.DataBodyRange.Rows.count
        Dim idArtisant As Variant
        idArtisant = tblCommande.DataBodyRange.Cells(i, 10).Value ' Colonne J = 10
        
        If Not IsEmpty(idArtisant) And Not dictArtisans.Exists(CStr(idArtisant)) Then
            dictArtisans.Add CStr(idArtisant), True
        End If
    Next i
    
    If dictArtisans.count = 0 Then
        MsgBox "Aucun ID artisan trouvé dans la colonne J du tableau 'plg_commande'.", vbExclamation
        Exit Sub
    End If
    
    ' Vérifier l'existence de la feuille ARTISAN
    Dim wsArtisant As Worksheet
    On Error Resume Next
    Set wsArtisant = ThisWorkbook.Worksheets("ARTISAN")
    On Error GoTo ErrHandler
    If wsArtisant Is Nothing Then
        MsgBox "Feuille 'ARTISAN' non trouvée.", vbCritical
        Exit Sub
    End If
    
    ' Parcourir les lignes du tableau plg_artisant via la plage de données
    Dim lastRow As Long
    lastRow = wsArtisant.Cells(wsArtisant.Rows.count, 1).End(xlUp).row
    If lastRow < 2 Then
        MsgBox "Aucune donnée dans la feuille 'ARTISAN' (à partir de la ligne 2).", vbCritical
        Exit Sub
    End If
    
    ' Dictionary pour les IDs d'artisans filtrés par région
    Dim dictFilteredArtisans As Object
    Set dictFilteredArtisans = CreateObject("Scripting.Dictionary")
    
    ' Parcourir à partir de la ligne 2 (en-tête en ligne 1)
    For i = 2 To lastRow
        Dim idArtRow As Variant
        idArtRow = wsArtisant.Cells(i, 1).Value ' Colonne A = ID_artisant
        
        If dictArtisans.Exists(CStr(idArtRow)) Then
            Dim adresse As String
            adresse = CStr(wsArtisant.Cells(i, 3).Value) ' Colonne C = Adresse
            If Not IsEmpty(adresse) And InStr(adresse, ",") > 0 Then
                Dim tabSplit() As String
                tabSplit = Split(adresse, ",")
                Dim valCell As String
                valCell = Trim(tabSplit(1)) ' Partie après la virgule
                If Len(valCell) >= 2 Then
                    Dim dept As String
                    dept = Left(valCell, 2)
                    If IsNumeric(dept) Then
                        Dim extractedRegion As String
                        extractedRegion = TrouverRegion(dept)
                        If StrComp(extractedRegion, selectedRegion, vbTextCompare) = 0 Then
                            LB_listemoyencate.AddItem adresse & " - Dépt: " & dept & " - Région: " & extractedRegion
                            ' Ajouter l'ID artisan filtré
                            dictFilteredArtisans.Add CStr(idArtRow), True
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
    If dictFilteredArtisans.count = 0 Then
        MsgBox "Aucun artisan trouvé pour la région sélectionnée : " & selectedRegion, vbExclamation
        Exit Sub
    End If
    
    ' Maintenant, collecter les id_categorie uniques (colonne L=12) des commandes des artisans filtrés
    Dim dictCategories As Object
    Set dictCategories = CreateObject("Scripting.Dictionary")
    
    For i = 1 To tblCommande.DataBodyRange.Rows.count
        idArtisant = tblCommande.DataBodyRange.Cells(i, 10).Value ' Colonne J
        If dictFilteredArtisans.Exists(CStr(idArtisant)) Then
            Dim idCat As Variant
            idCat = tblCommande.DataBodyRange.Cells(i, 12).Value ' Colonne L = 12
            If Not IsEmpty(idCat) And Not dictCategories.Exists(CStr(idCat)) Then
                dictCategories.Add CStr(idCat), True
            End If
        End If
    Next i
    
    If dictCategories.count = 0 Then
        MsgBox "Aucune catégorie trouvée pour les artisans de la région sélectionnée.", vbExclamation
        Exit Sub
    End If
    
    ' Vérifier l'existence de la feuille CATEGORIE_TRAVAUX
    Dim wsCat As Worksheet
    On Error Resume Next
    Set wsCat = ThisWorkbook.Worksheets("CATEGORIE_TRAVAUX")
    On Error GoTo ErrHandler
    If wsCat Is Nothing Then
        MsgBox "Feuille 'CATEGORIE_TRAVAUX' non trouvée.", vbCritical
        Exit Sub
    End If
    
    ' Parcourir la feuille pour mapper ID (colonne A) vers nom (colonne B)
    Dim lastRowCat As Long
    lastRowCat = wsCat.Cells(wsCat.Rows.count, 1).End(xlUp).row
    If lastRowCat < 2 Then
        MsgBox "Aucune donnée dans la feuille 'CATEGORIE_TRAVAUX' (à partir de la ligne 2).", vbCritical
        Exit Sub
    End If
    
    Dim j As Long
    For Each idCat In dictCategories.keys
        Dim foundName As String
        foundName = "Inconnu" ' Par défaut si non trouvé
        For j = 2 To lastRowCat
            If CStr(wsCat.Cells(j, 1).Value) = CStr(idCat) Then
                foundName = CStr(wsCat.Cells(j, 2).Value) ' Colonne B = Nom travaux
                Exit For ' Arrêter à la première correspondance (assume ID unique)
            End If
        Next j
        ' Ajouter le nom à la ComboBox
        If foundName <> "" Then
            CB_rechercate.AddItem foundName
        End If
    Next idCat
    
    CB_rechercate.ListIndex = -1 ' Pas de sélection par défaut
    
    ' Calculer la moyenne pour la région (toutes catégories si aucune sélectionnée)
    Call CalculerCoutMoyen(dictFilteredArtisans, tblCommande, "")
    
    Debug.Print "ListBox LB_listemoyencate remplie avec " & LB_listemoyencate.ListCount & " adresses pour la région : " & selectedRegion
    Debug.Print "ComboBox CB_rechercate remplie avec " & CB_rechercate.ListCount & " noms de catégories uniques."
    
    Exit Sub
    
ErrHandler:
    MsgBox "Erreur " & Err.Number & " lors du remplissage de la ListBox : " & Err.Description & vbCrLf & "Vérifiez les noms des feuilles et tableaux.", vbCritical
    LB_listemoyencate.Clear
    CB_rechercate.Clear
    TB_coutmoyen.Value = ""
End Sub

Private Sub CB_rechercate_Change()
    On Error GoTo ErrHandler
    
    Dim selectedRegion As String
    selectedRegion = CB_recherregion.Value
    
    Dim selectedCatName As String
    selectedCatName = CB_rechercate.Value
    
    If selectedRegion = "" Or selectedCatName = "" Then Exit Sub
    
    ' Vérifier l'existence de la feuille COMMANDE
    Dim wsCommande As Worksheet
    On Error Resume Next
    Set wsCommande = ThisWorkbook.Worksheets("COMMANDE")
    On Error GoTo ErrHandler
    If wsCommande Is Nothing Then Exit Sub
    
    ' Vérifier l'existence du tableau plg_commande
    Dim tblCommande As ListObject
    On Error Resume Next
    Set tblCommande = wsCommande.ListObjects("plg_commande")
    On Error GoTo ErrHandler
    If tblCommande Is Nothing Then Exit Sub
    
    If tblCommande.DataBodyRange Is Nothing Then Exit Sub
    
    ' Récupérer dictFilteredArtisans (recomputé pour simplicité)
    Dim dictFilteredArtisans As Object
    Set dictFilteredArtisans = CreateObject("Scripting.Dictionary")
    
    ' Réutiliser la logique de filtrage artisans (copié de CB_recherregion_Change pour éviter duplication)
    Dim dictArtisans As Object
    Set dictArtisans = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 1 To tblCommande.DataBodyRange.Rows.count
        Dim idArtisant As Variant
        idArtisant = tblCommande.DataBodyRange.Cells(i, 10).Value
        If Not IsEmpty(idArtisant) And Not dictArtisans.Exists(CStr(idArtisant)) Then
            dictArtisans.Add CStr(idArtisant), True
        End If
    Next i
    
    Dim wsArtisant As Worksheet
    On Error Resume Next
    Set wsArtisant = ThisWorkbook.Worksheets("ARTISAN")
    On Error GoTo ErrHandler
    If wsArtisant Is Nothing Then Exit Sub
    
    Dim lastRow As Long
    lastRow = wsArtisant.Cells(wsArtisant.Rows.count, 1).End(xlUp).row
    If lastRow < 2 Then Exit Sub
    
    For i = 2 To lastRow
        Dim idArtRow As Variant
        idArtRow = wsArtisant.Cells(i, 1).Value
        
        If dictArtisans.Exists(CStr(idArtRow)) Then
            Dim adresse As String
            adresse = CStr(wsArtisant.Cells(i, 3).Value)
            If Not IsEmpty(adresse) And InStr(adresse, ",") > 0 Then
                Dim tabSplit() As String
                tabSplit = Split(adresse, ",")
                Dim valCell As String
                valCell = Trim(tabSplit(1))
                If Len(valCell) >= 2 Then
                    Dim dept As String
                    dept = Left(valCell, 2)
                    If IsNumeric(dept) Then
                        Dim extractedRegion As String
                        extractedRegion = TrouverRegion(dept)
                        If StrComp(extractedRegion, selectedRegion, vbTextCompare) = 0 Then
                            dictFilteredArtisans.Add CStr(idArtRow), True
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
    If dictFilteredArtisans.count = 0 Then Exit Sub
    
    ' Mapper nom catégorie vers ID
    Dim wsCat As Worksheet
    On Error Resume Next
    Set wsCat = ThisWorkbook.Worksheets("CATEGORIE_TRAVAUX")
    On Error GoTo ErrHandler
    If wsCat Is Nothing Then Exit Sub
    
    Dim selectedCatID As String
    selectedCatID = GetIDFromCatName(selectedCatName, wsCat)
    
    If selectedCatID = "" Then Exit Sub
    
    ' Calculer la moyenne pour région + catégorie
    Call CalculerCoutMoyen(dictFilteredArtisans, tblCommande, selectedCatID)
    
    Exit Sub
    
ErrHandler:
    MsgBox "Erreur lors du changement de catégorie : " & Err.Number & " - " & Err.Description, vbCritical
    TB_coutmoyen.Value = ""
End Sub

Private Sub CalculerCoutMoyen(dictFilteredArtisans As Object, tblCommande As ListObject, selectedCatID As String)
    On Error GoTo ErrHandler
    
    Dim totalMontant As Double
    totalMontant = 0
    
    Dim countCommands As Long
    countCommands = 0
    
    Dim i As Long
    For i = 1 To tblCommande.DataBodyRange.Rows.count
        Dim idArt As Variant
        idArt = tblCommande.DataBodyRange.Cells(i, 10).Value ' Col J
        
        If dictFilteredArtisans.Exists(CStr(idArt)) Then
            Dim idCat As Variant
            idCat = tblCommande.DataBodyRange.Cells(i, 12).Value ' Col L
            
            Dim matchCat As Boolean
            matchCat = (selectedCatID = "") Or (CStr(idCat) = selectedCatID)
            
            If matchCat Then
                Dim montant As Variant
                montant = tblCommande.DataBodyRange.Cells(i, 7).Value ' Col G = 7
                If IsNumeric(montant) And montant > 0 Then
                    totalMontant = totalMontant + CDbl(montant)
                    countCommands = countCommands + 1
                End If
            End If
        End If
    Next i
    
    If countCommands > 0 Then
        Dim moyenne As Double
        moyenne = totalMontant / countCommands
        TB_coutmoyen.Value = Format(moyenne, "#,##0.00") & " €"
    Else
        TB_coutmoyen.Value = "Aucune commande"
    End If
    
    Exit Sub
    
ErrHandler:
    TB_coutmoyen.Value = "Erreur calcul"
End Sub

Private Function GetIDFromCatName(catName As String, wsCat As Worksheet) As String
    On Error GoTo ErrHandler
    
    If catName = "" Then
        GetIDFromCatName = ""
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = wsCat.Cells(wsCat.Rows.count, 2).End(xlUp).row ' Basé sur col B
    
    Dim i As Long
    For i = 2 To lastRow
        If CStr(wsCat.Cells(i, 2).Value) = catName Then
            GetIDFromCatName = CStr(wsCat.Cells(i, 1).Value) ' Col A = ID
            Exit Function
        End If
    Next i
    
    GetIDFromCatName = "" ' Non trouvé
    
    Exit Function
    
ErrHandler:
    GetIDFromCatName = ""
End Function

Private Function TrouverRegion(CodeDept As String) As String
    Select Case CodeDept
        Case "75", "77", "78", "91", "92", "93", "94", "95": TrouverRegion = "Île-de-France"
        Case "13", "83", "06", "84", "04", "05": TrouverRegion = "Provence-Alpes-Côte d'Azur"
        Case "69", "01", "38", "73", "74", "42", "26", "07", "15", "03", "63", "43": TrouverRegion = "Auvergne-Rhône-Alpes"
        Case "33", "24", "40", "47", "64", "16", "17", "19", "23", "79", "86", "87": TrouverRegion = "Nouvelle-Aquitaine"
        Case "31", "34", "30", "11", "66", "09", "12", "46", "48", "65", "81", "82", "32": TrouverRegion = "Occitanie"
        Case "44", "85", "49", "72", "53": TrouverRegion = "Pays de la Loire"
        Case "59", "62", "02", "60", "80": TrouverRegion = "Hauts-de-France"
        Case "67", "68", "08", "10", "51", "52", "54", "55", "57", "88": TrouverRegion = "Grand Est"
        Case "14", "27", "50", "61", "76": TrouverRegion = "Normandie"
        Case "22", "29", "35", "56": TrouverRegion = "Bretagne"
        Case "18", "28", "36", "37", "41", "45": TrouverRegion = "Centre-Val de Loire"
        Case "21", "25", "39", "58", "70", "71", "89", "90": TrouverRegion = "Bourgogne-Franche-Comté"
        Case "20": TrouverRegion = "Corse"
        Case "97": TrouverRegion = "Outre-mer" ' Simplifié pour DOM-TOM ; ajustez si besoin pour régions spécifiques
        Case Else: TrouverRegion = "Autres Régions"
    End Select
End Function

Private Sub L_bilanstock_Click()

End Sub

Private Sub LB_listemoyencate_Click()

End Sub

Private Sub TB_coutmoyen_Change()
    
End Sub

Private Sub UserForm_Click()
   
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler
  
    ' Remplir CB_recherregion avec toutes les régions françaises (13 métropolitaines + 5 ultramarines)
    With CB_recherregion
        .Clear
        .AddItem "Auvergne-Rhône-Alpes"
        .AddItem "Bourgogne-Franche-Comté"
        .AddItem "Bretagne"
        .AddItem "Centre-Val de Loire"
        .AddItem "Corse"
        .AddItem "Grand Est"
        .AddItem "Hauts-de-France"
        .AddItem "Île-de-France"
        .AddItem "Normandie"
        .AddItem "Nouvelle-Aquitaine"
        .AddItem "Occitanie"
        .AddItem "Pays de la Loire"
        .AddItem "Provence-Alpes-Côte d'Azur"
        .AddItem "Guadeloupe"
        .AddItem "Guyane"
        .AddItem "Martinique"
        .AddItem "Mayotte"
        .AddItem "La Réunion"
        .ListIndex = -1 ' Pas de sélection par défaut (ou 0 pour première)
    End With
  
    ' Initialiser CB_rechercate vide (sera remplie au changement de région)
    CB_rechercate.Clear
    CB_rechercate.ListIndex = -1
  
    ' Initialiser TB_coutmoyen vide
    TB_coutmoyen.Value = ""
  
    ' Autres initialisations si besoin (ex. ListBox, etc.)
    Debug.Print "ComboBox CB_recherregion remplie avec 18 régions françaises."
  
    Exit Sub
  
ErrHandler:
    MsgBox "Erreur lors de l'initialisation : " & Err.Number & " - " & Err.Description, vbCritical
End Sub
