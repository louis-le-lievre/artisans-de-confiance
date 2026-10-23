VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   8316.001
   ClientLeft      =   84
   ClientTop       =   360
   ClientWidth     =   14520
   OleObjectBlob   =   "2g.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Me.Caption = "Commandes à terminer ce mois-ci"
    
    ' Configuration de la ListBox
    With Me.ListBox1
        .Clear
        .ColumnCount = 4
        .ColumnWidths = "80;130;130;90"  ' Ajusté pour mieux lire les noms
        
        ' En-têtes
        .AddItem
        .List(0, 0) = "Commande"
        .List(0, 1) = "Artisan"
        .List(0, 2) = "Client"
        .List(0, 3) = "Fin prévue"
    End With
    
    ' Chargement des données
    Call ChargerCommandesMoisCourant
End Sub

Private Sub ChargerCommandesMoisCourant()
    Dim wsCmd As Worksheet, wsCli As Worksheet, wsArt As Worksheet
    Set wsCmd = ThisWorkbook.Sheets("COMMANDE")
    Set wsCli = ThisWorkbook.Sheets("CLIENT")
    Set wsArt = ThisWorkbook.Sheets("ARTISAN")
    
    ' Dictionnaires pour accès rapide aux noms
    Dim dictCli As Object: Set dictCli = CreateObject("Scripting.Dictionary")
    Dim dictArt As Object: Set dictArt = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    
    ' === Remplir dictionnaire Clients : id ? "Nom Prénom" ===
    For i = 2 To wsCli.Cells(wsCli.Rows.count, 1).End(xlUp).Row
        Dim idCli As String: idCli = Trim(wsCli.Cells(i, 1).Value)
        If Len(idCli) > 0 Then
            dictCli(idCli) = Trim(wsCli.Cells(i, 2).Value) & " " & Trim(wsCli.Cells(i, 3).Value)
        End If
    Next i
    
    ' === Remplir dictionnaire Artisans : id ? Nom ===
    For i = 2 To wsArt.Cells(wsArt.Rows.count, 1).End(xlUp).Row
        Dim idArt As String: idArt = Trim(wsArt.Cells(i, 1).Value)
        If Len(idArt) > 0 Then
            dictArt(idArt) = Trim(wsArt.Cells(i, 2).Value)
        End If
    Next i
    
    ' Mois et année courants (au 13 décembre 2025 ? mois 12, année 2025)
    Dim moisCourant As Integer: moisCourant = Month(Date)
    Dim anneeCourante As Integer: anneeCourante = Year(Date)
    
    Dim count As Long: count = 0
    
    ' Nettoyer la ListBox (sauf en-tête)
    Do While Me.ListBox1.ListCount > 1
        Me.ListBox1.RemoveItem (Me.ListBox1.ListCount - 1)
    Loop
    
    ' Parcours des commandes
    Dim lastRow As Long
    lastRow = wsCmd.Cells(wsCmd.Rows.count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        If Trim(wsCmd.Cells(i, 13).Value) = "STAT03" Then  ' EN_COURS
            
            Dim dateFinPrev As Variant: dateFinPrev = wsCmd.Cells(i, 4).Value  ' Colonne D
            
            If IsDate(dateFinPrev) Then
                If Month(dateFinPrev) = moisCourant And Year(dateFinPrev) = anneeCourante Then
                    
                    count = count + 1
                    
                    Dim idCmd As String: idCmd = wsCmd.Cells(i, 1).Value
                    Dim idClient As String: idClient = Trim(wsCmd.Cells(i, 8).Value)
                    Dim idArtisan As String: idArtisan = Trim(wsCmd.Cells(i, 9).Value)
                    
                    Dim nomClient As String
                    nomClient = IIf(dictCli.Exists(idClient), dictCli(idClient), "Client inconnu")
                    
                    Dim nomArtisan As String
                    nomArtisan = IIf(dictArt.Exists(idArtisan), dictArt(idArtisan), "Artisan inconnu")
                    
                    ' Ajouter à la ListBox
                    With Me.ListBox1
                        .AddItem idCmd
                        .List(.ListCount - 1, 1) = nomArtisan
                        .List(.ListCount - 1, 2) = nomClient
                        .List(.ListCount - 1, 3) = Format(dateFinPrev, "dd/mm/yyyy")
                    End With
                End If
            End If
        End If
    Next i
    
    ' === Affichage du résultat ===
    With Me.LabelResultat
        .Caption = "Nombre de commandes en cours à terminer en " & _
                   Format(Date, "mmmm yyyy") & " : " & count
        .Font.Bold = True
        .Font.Size = 12
        .ForeColor = IIf(count = 0, vbRed, vbBlue)
    End With
    
    ' Message si aucune commande
    If count = 0 Then
        Me.ListBox1.AddItem ""
        Me.ListBox1.List(1, 0) = "Aucune commande à terminer ce mois-ci"
        Me.ListBox1.List(1, 1) = ""
        Me.ListBox1.List(1, 2) = ""
        Me.ListBox1.List(1, 3) = ""
    End If
End Sub

Private Sub btnFermer_Click()
    Unload Me
End Sub

' Pour ouvrir ce formulaire depuis un bouton du tableau de bord
Public Sub OuvrirCommandesMoisCourant()
    frmCommandesMois.Show  ' ? Adapte le nom de ton UserForm ici
End Sub

