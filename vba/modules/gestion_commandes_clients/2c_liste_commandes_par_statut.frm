VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   9372.001
   ClientLeft      =   84
   ClientTop       =   360
   ClientWidth     =   13932
   OleObjectBlob   =   "2c.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ' Titre et instructions
    Me.Caption = "Recherche de commandes par statut"
    
    Dim wsStat As Worksheet
    Set wsStat = ThisWorkbook.Sheets("STATUT_COMMANDE")
    
    Dim lastRowStat As Long
    lastRowStat = wsStat.Cells(wsStat.Rows.Count, 1).End(xlUp).Row
    
    ' Configuration de la ListBox des statuts
    With lstStatuts
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "0pt;150pt"   ' Cacher l'ID, afficher le libellé
        .BoundColumn = 1              ' Valeur retournée = colonne 0 (ID)
        
        Dim i As Long
        For i = 2 To lastRowStat
            .AddItem wsStat.Cells(i, 1).Value                  ' Colonne 0 : ID
            .List(.ListCount - 1, 1) = wsStat.Cells(i, 2).Value ' Colonne 1 : Libellé
        Next i
    End With
    
    ' Configuration de la ListBox des résultats
    With lstResultats
        .Clear
        .ColumnCount = 9
        .ColumnWidths = "80pt;100pt;80pt;150pt;90pt;90pt;90pt;90pt;120pt"
        ' Colonnes : Commande, Statut, Montant, Client, Création, Début prévue, Début réelle, Fin réelle, Artisan
    End With
End Sub

Private Sub cmdAfficher_Click()
    lstResultats.Clear
    
    ' --- Récupérer les statuts sélectionnés ---
    Dim selectedStatuts As Collection
    Set selectedStatuts = New Collection
    
    Dim j As Long
    For j = 0 To lstStatuts.ListCount - 1
        If lstStatuts.Selected(j) Then
            selectedStatuts.Add lstStatuts.List(j, 0)  ' ID du statut
        End If
    Next j
    
    If selectedStatuts.Count = 0 Then
        MsgBox "Veuillez sélectionner au moins un statut de commande.", vbExclamation, "Aucun statut"
        Exit Sub
    End If
    
    ' --- Références feuilles ---
    Dim wsCmd As Worksheet: Set wsCmd = ThisWorkbook.Sheets("COMMANDE")
    Dim wsCli As Worksheet: Set wsCli = ThisWorkbook.Sheets("CLIENT")
    Dim wsArt As Worksheet: Set wsArt = ThisWorkbook.Sheets("ARTISAN")
    Dim wsStat As Worksheet: Set wsStat = ThisWorkbook.Sheets("STATUT_COMMANDE")
    
    Dim lastRowCmd As Long
    lastRowCmd = wsCmd.Cells(wsCmd.Rows.Count, 1).End(xlUp).Row
    
    ' --- Dictionnaires pour accès rapide ---
    Dim dictCli As Object: Set dictCli = CreateObject("Scripting.Dictionary")
    Dim dictArt As Object: Set dictArt = CreateObject("Scripting.Dictionary")
    Dim dictStat As Object: Set dictStat = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 2 To wsCli.Cells(wsCli.Rows.Count, 1).End(xlUp).Row
        dictCli(wsCli.Cells(i, 1).Value) = wsCli.Cells(i, 2).Value & " " & wsCli.Cells(i, 3).Value
    Next i
    
    For i = 2 To wsArt.Cells(wsArt.Rows.Count, 1).End(xlUp).Row
        dictArt(wsArt.Cells(i, 1).Value) = wsArt.Cells(i, 2).Value
    Next i
    
    For i = 2 To wsStat.Cells(wsStat.Rows.Count, 1).End(xlUp).Row
        dictStat(wsStat.Cells(i, 1).Value) = wsStat.Cells(i, 2).Value
    Next i
    
    ' --- Parcours des commandes ---
    Dim ligne As Long: ligne = 0  ' Index de ligne dans lstResultats
    
    For i = 2 To lastRowCmd
        Dim idStat As String
        idStat = Trim(wsCmd.Cells(i, 13).Value)
        
        ' Vérifier si le statut est dans les sélectionnés
        Dim found As Boolean: found = False
        Dim s As Variant
        For Each s In selectedStatuts
            If idStat = CStr(s) Then
                found = True
                Exit For
            End If
        Next s
        
        If found Then
            Dim idCmd As String: idCmd = wsCmd.Cells(i, 1).Value
            Dim montant As Double: montant = val(wsCmd.Cells(i, 7).Value)
            Dim idCli As String: idCli = wsCmd.Cells(i, 8).Value
            Dim idArt As String: idArt = wsCmd.Cells(i, 9).Value
            
            Dim nomCli As String
            nomCli = IIf(dictCli.Exists(idCli), dictCli(idCli), "Client inconnu")
            
            Dim nomArt As String
            nomArt = IIf(dictArt.Exists(idArt), dictArt(idArt), "Artisan inconnu")
            
            Dim libStat As String
            libStat = IIf(dictStat.Exists(idStat), dictStat(idStat), "Statut inconnu")
            
            ' Ajouter la ligne
            lstResultats.AddItem
            lstResultats.List(ligne, 0) = idCmd
            lstResultats.List(ligne, 1) = libStat
            lstResultats.List(ligne, 2) = Format(montant, "#,##0 €")
            lstResultats.List(ligne, 3) = nomCli
            lstResultats.List(ligne, 4) = FormatDate(wsCmd.Cells(i, 2).Value)  ' Date création
            lstResultats.List(ligne, 5) = FormatDate(wsCmd.Cells(i, 3).Value)  ' Début prévue
            lstResultats.List(ligne, 6) = FormatDate(wsCmd.Cells(i, 5).Value)  ' Début réelle
            lstResultats.List(ligne, 7) = FormatDate(wsCmd.Cells(i, 6).Value)  ' Fin réelle
            lstResultats.List(ligne, 8) = nomArt
            
            ligne = ligne + 1
        End If
    Next i
    
    If ligne = 0 Then
        MsgBox "Aucune commande trouvée pour les statuts sélectionnés.", vbInformation, "Résultat vide"
    End If
End Sub

' Fonction helper pour formater les dates Excel en toute sécurité
Private Function FormatDate(val As Variant) As String
    If IsEmpty(val) Or val = "" Or val = "NULL" Then
        FormatDate = ""
    ElseIf IsNumeric(val) Then
        If val > 0 Then
            FormatDate = Format(CDate(val), "dd/mm/yyyy")
        Else
            FormatDate = ""
        End If
    Else
        FormatDate = CStr(val)
    End If
End Function

Private Sub cmdFermer_Click()
    Unload Me
End Sub

Sub OuvrirListeCommandesParStatut()
    frmListeCommandes.Show
End Sub
