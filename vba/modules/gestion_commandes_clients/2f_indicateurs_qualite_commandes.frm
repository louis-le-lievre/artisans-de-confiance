VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   9888.001
   ClientLeft      =   84
   ClientTop       =   360
   ClientWidth     =   16104
   OleObjectBlob   =   "2f.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Me.Caption = "Avis clients et satisfaction"
    
    With lstAvis
        .Clear
        .ColumnCount = 5
        .ColumnWidths = "80;150;60;150;80"
        
        ' En-tête
        .AddItem
        .List(0, 0) = "Commande"
        .List(0, 1) = "Client"
        .List(0, 2) = "Note /5"
        .List(0, 3) = "Commentaire"
        .List(0, 4) = "Date avis"
    End With
    
    Call cmdActualiser_Click
End Sub

Private Sub cmdActualiser_Click()
    Dim wsCmd As Worksheet, wsAvis As Worksheet, wsCli As Worksheet
    Set wsCmd = ThisWorkbook.Sheets("COMMANDE")
    Set wsAvis = ThisWorkbook.Sheets("AVIS")
    Set wsCli = ThisWorkbook.Sheets("CLIENT")
    
    Dim dictCli As Object: Set dictCli = CreateObject("Scripting.Dictionary")
    Dim dictCmdClient As Object: Set dictCmdClient = CreateObject("Scripting.Dictionary")
    
    ' === DÉCLARATIONS UNIQUES DE TOUTES LES VARIABLES ===
    Dim totalDelais As Long
    Dim ponctuelles As Long
    Dim nbAvis As Long
    Dim sommeNote As Double
    Dim sommeFinition As Double
    Dim sommePonctualite As Double
    
    Dim i As Long
    Dim idCmd As String, idCliAvis As String, nomCli As String
    Dim note As Double, noteFinition As Double, notePonct As Double
    Dim commentaire As String
    Dim dateAvis As Variant
    
    ' === Remplir dictionnaire clients ===
    For i = 2 To wsCli.Cells(wsCli.Rows.Count, 1).End(xlUp).Row
        Dim idCli As String: idCli = Trim(wsCli.Cells(i, 1).Value)
        If Len(idCli) > 0 Then
            dictCli(idCli) = wsCli.Cells(i, 3).Value & " " & wsCli.Cells(i, 2).Value
        End If
    Next i
    
    ' === Remplir dictionnaire commande ? client + calcul ponctualité ===
    For i = 2 To wsCmd.Cells(wsCmd.Rows.Count, 1).End(xlUp).Row
        idCmd = Trim(wsCmd.Cells(i, 1).Value)
        Dim idCliCmd As String: idCliCmd = Trim(wsCmd.Cells(i, 8).Value)
        
        If Len(idCmd) > 0 And Len(idCliCmd) > 0 Then
            dictCmdClient(idCmd) = idCliCmd
        End If
        
        ' Calcul ponctualité seulement pour commandes TERMINEE
        If Trim(wsCmd.Cells(i, 13).Value) = "STAT04" Then
            Dim dateFinPrev As Variant: dateFinPrev = wsCmd.Cells(i, 4).Value
            Dim dateFinReel As Variant: dateFinReel = wsCmd.Cells(i, 6).Value
            
            If IsDate(dateFinPrev) And IsDate(dateFinReel) Then
                totalDelais = totalDelais + 1
                If dateFinReel <= dateFinPrev Then
                    ponctuelles = ponctuelles + 1
                End If
            End If
        End If
    Next i
    
    ' === Nettoyage ListBox (sauf en-tête) ===
    Do While lstAvis.ListCount > 1
        lstAvis.RemoveItem (lstAvis.ListCount - 1)
    Loop
    
    ' === Parcours des avis ===
    Dim lastRowAvis As Long
    lastRowAvis = wsAvis.Cells(wsAvis.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRowAvis
        idCmd = Trim(wsAvis.Cells(i, 7).Value)
        If Len(idCmd) = 0 Then GoTo NextAvis
        
        note = Val(wsAvis.Cells(i, 2).Value)
        commentaire = Trim(wsAvis.Cells(i, 3).Value & "")
        dateAvis = wsAvis.Cells(i, 4).Value
        noteFinition = Val(wsAvis.Cells(i, 5).Value)
        notePonct = Val(wsAvis.Cells(i, 6).Value)
        
        nbAvis = nbAvis + 1
        sommeNote = sommeNote + note
        sommeFinition = sommeFinition + noteFinition
        sommePonctualite = sommePonctualite + notePonct
        
        ' Nom du client
        If dictCmdClient.Exists(idCmd) Then
            idCliAvis = dictCmdClient(idCmd)
            nomCli = IIf(dictCli.Exists(idCliAvis), dictCli(idCliAvis), "Client inconnu")
        Else
            nomCli = "Commande inconnue"
        End If
        
        ' Ajout dans la ListBox
        With lstAvis
            .AddItem idCmd
            .List(.ListCount - 1, 1) = nomCli
            .List(.ListCount - 1, 2) = note & "/5"
            If Len(commentaire) > 60 Then
                .List(.ListCount - 1, 3) = Left(commentaire, 57) & "..."
            Else
                .List(.ListCount - 1, 3) = commentaire
            End If
            If IsDate(dateAvis) Then
                .List(.ListCount - 1, 4) = Format(dateAvis, "dd/mm/yyyy")
            Else
                .List(.ListCount - 1, 4) = "-"
            End If
        End With
        
NextAvis:
    Next i
    
    ' === Affichage ponctualité ===
    If totalDelais > 0 Then
        Dim taux As Double: taux = ponctuelles / totalDelais
        lblPonctualite.Caption = "Taux de ponctualité : " & Format(taux, "0.0%") & _
                                " (" & ponctuelles & " / " & totalDelais & " commandes terminées)"
        lblPonctualite.ForeColor = IIf(taux >= 0.9, vbGreen, IIf(taux >= 0.7, RGB(255, 140, 0), vbRed))
    Else
        lblPonctualite.Caption = "Taux de ponctualité : Aucune donnée disponible"
        lblPonctualite.ForeColor = vbBlack
    End If
    
    ' === Affichage notes moyennes ===
    If nbAvis > 0 Then
        Dim moyGenerale As Double: moyGenerale = sommeNote / nbAvis
        lblNoteGenerale.Caption = "Note moyenne globale : " & Format(moyGenerale, "0.0") & " / 5  (" & nbAvis & " avis)"
        lblNoteFinition.Caption = "• Note moyenne finition : " & Format(sommeFinition / nbAvis, "0.0") & " / 5"
        lblNotePonctualite.Caption = "• Note moyenne ponctualité : " & Format(sommePonctualite / nbAvis, "0.0") & " / 5"
        
        lblNoteGenerale.ForeColor = IIf(moyGenerale >= 4.5, vbGreen, IIf(moyGenerale >= 3.5, RGB(255, 140, 0), vbRed))
    Else
        lblNoteGenerale.Caption = "Note moyenne globale : Aucun avis client"
        lblNoteFinition.Caption = ""
        lblNotePonctualite.Caption = ""
        lblNoteGenerale.ForeColor = vbBlack
    End If
    
    ' === Aucun avis ===
    If lstAvis.ListCount = 1 Then
        lstAvis.AddItem ""
        lstAvis.List(1, 0) = "Aucun avis client pour le moment"
        lstAvis.List(1, 1) = ""
        lstAvis.List(1, 2) = ""
        lstAvis.List(1, 3) = ""
        lstAvis.List(1, 4) = ""
    End If
End Sub

Private Sub cmdFermer_Click()
    Unload Me
End Sub

