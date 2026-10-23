VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   10596
   ClientLeft      =   84
   ClientTop       =   360
   ClientWidth     =   16944
   OleObjectBlob   =   "2e.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Me.Caption = "Indicateurs de qualité des délais"
    
    lblTitre.Caption = "Indicateurs de qualité des délais" & vbCrLf & _
                       "Basés sur les commandes validées, en cours ou terminées"
    lblTitre.Font.Size = 12
    lblTitre.Font.Bold = True
    ' Correction : TextAlign n'existe pas pour Label en UserForm Excel ? on centre manuellement ou on supprime
    ' On peut centrer le texte en utilisant WordWrap + AutoSize, ou simplement accepter le défaut à gauche.
    ' Alternative : on laisse comme ça, ou on utilise un Label plus large centré visuellement.
    
    With lstRetards
        .ColumnCount = 5
        .ColumnWidths = "70;120;100;100;150"
        .Clear
        
        ' En-tête
        .AddItem
        .List(0, 0) = "Commande"
        .List(0, 1) = "Client"
        .List(0, 2) = "Fin prévue"
        .List(0, 3) = "Fin réelle"
        .List(0, 4) = "Motif retard"
    End With
    
    ' Actualisation automatique à l'ouverture
    cmdActualiser_Click
End Sub

Private Sub cmdActualiser_Click()
    Dim wsCmd As Worksheet, wsCli As Worksheet
    Dim wsAPour As Worksheet, wsMotif As Worksheet
    
    Set wsCmd = ThisWorkbook.Sheets("COMMANDE")
    Set wsCli = ThisWorkbook.Sheets("CLIENT")
    Set wsAPour = ThisWorkbook.Sheets("A_POUR_MOTIF")
    Set wsMotif = ThisWorkbook.Sheets("MOTIF_RETARD")
    
    Dim dictCli As Object: Set dictCli = CreateObject("Scripting.Dictionary")
    Dim dictMotif As Object: Set dictMotif = CreateObject("Scripting.Dictionary")
    
    ' === Variables déclarées UNE SEULE FOIS ici ===
    Dim i As Long
    Dim idCli As String
    Dim idCmd As String
    Dim idMot As String
    
    Dim dateFinPrev As Variant
    Dim dateFinReel As Variant
    Dim dateDebPrev As Variant
    Dim dateDebReel As Variant
    Dim statut As String
    
    Dim nomCli As String
    Dim motif As String
    Dim idx As Long
    
    ' === Chargement des clients ===
    For i = 2 To wsCli.Cells(wsCli.Rows.Count, 1).End(xlUp).Row
        idCli = Trim(wsCli.Cells(i, 1).Value)
        If idCli <> "" Then
            dictCli(idCli) = Trim(wsCli.Cells(i, 3).Value) & " " & Trim(wsCli.Cells(i, 2).Value)
        End If
    Next i
    
    ' === Chargement des motifs ===
    For i = 2 To wsAPour.Cells(wsAPour.Rows.Count, 1).End(xlUp).Row
        idCmd = Trim(wsAPour.Cells(i, 2).Value)
        idMot = Trim(wsAPour.Cells(i, 1).Value)
        If idCmd <> "" And idMot <> "" Then
            Dim libMot As Variant
            libMot = Application.VLookup(idMot, wsMotif.Range("A:B"), 2, False)
            If Not IsError(libMot) Then
                If dictMotif.Exists(idCmd) Then
                    dictMotif(idCmd) = dictMotif(idCmd) & ", " & libMot
                Else
                    dictMotif(idCmd) = libMot
                End If
            End If
        End If
    Next i
    
    ' === Nettoyage ListBox ===
    Do While lstRetards.ListCount > 1
        lstRetards.RemoveItem (lstRetards.ListCount - 1)
    Loop
    
    ' === Compteurs ===
    Dim totalTerminees As Long
    Dim dansDelais As Long
    Dim avecRetard As Long
    Dim sommeRetard As Double
    Dim totalDebut As Long
    Dim debutRespect As Long
    
    ' === Parcours des commandes ===
    For i = 2 To wsCmd.Cells(wsCmd.Rows.Count, 1).End(xlUp).Row
        statut = Trim(wsCmd.Cells(i, 13).Value)
        
        If statut = "STAT03" Or statut = "STAT04" Then
            
            dateFinPrev = wsCmd.Cells(i, 4).Value
            dateFinReel = wsCmd.Cells(i, 6).Value
            dateDebPrev = wsCmd.Cells(i, 3).Value
            dateDebReel = wsCmd.Cells(i, 5).Value
            idCmd = Trim(wsCmd.Cells(i, 1).Value)
            idCli = Trim(wsCmd.Cells(i, 8).Value)
            
            ' --- Fin des travaux ---
            If IsDate(dateFinPrev) And IsDate(dateFinReel) Then
                totalTerminees = totalTerminees + 1
                
                If dateFinReel <= dateFinPrev Then
                    dansDelais = dansDelais + 1
                Else
                    avecRetard = avecRetard + 1
                    sommeRetard = sommeRetard + (dateFinReel - dateFinPrev)
                    
                    nomCli = "Inconnu"
                    If dictCli.Exists(idCli) Then nomCli = dictCli(idCli)
                    
                    motif = "Non renseigné"
                    If dictMotif.Exists(idCmd) Then motif = dictMotif(idCmd)
                    
                    lstRetards.AddItem
                    idx = lstRetards.ListCount - 1
                    lstRetards.List(idx, 0) = idCmd
                    lstRetards.List(idx, 1) = nomCli
                    lstRetards.List(idx, 2) = Format(dateFinPrev, "dd/mm/yyyy")
                    lstRetards.List(idx, 3) = Format(dateFinReel, "dd/mm/yyyy")
                    lstRetards.List(idx, 4) = motif
                End If
            End If
            
            ' --- Début des travaux ---
            If IsDate(dateDebPrev) And IsDate(dateDebReel) Then
                totalDebut = totalDebut + 1
                If dateDebReel <= dateDebPrev + 3 Then
                    debutRespect = debutRespect + 1
                End If
            End If
        End If
    Next i
    
    ' === Mise à jour des labels (inchangée) ===
    If totalTerminees > 0 Then
        lblRespectDelais.Caption = "Respect des délais de fin : " & _
            Format(dansDelais / totalTerminees, "0.0%") & " (" & dansDelais & "/" & totalTerminees & ")"
        
        lblRetard.Caption = "Commandes en retard (fin) : " & _
            Format(avecRetard / totalTerminees, "0.0%") & " (" & avecRetard & " commande" & _
            IIf(avecRetard <> 1, "s", "") & ")"
        
        lblRetardMoyen.Caption = "Retard moyen : " & _
            IIf(avecRetard > 0, Round(sommeRetard / avecRetard, 1) & " jour" & _
            IIf(sommeRetard / avecRetard >= 2, "s", ""), "0 jour")
    Else
        lblRespectDelais.Caption = "Respect des délais de fin : Aucune donnée"
        lblRetard.Caption = "Commandes en retard (fin) : Aucune donnée"
        lblRetardMoyen.Caption = "Retard moyen : Aucune donnée"
    End If
    
    If totalDebut > 0 Then
        lblDebutRespect.Caption = "Début dans les 3 jours prévus : " & _
            Format(debutRespect / totalDebut, "0.0%") & " (" & debutRespect & "/" & totalDebut & ")"
    Else
        lblDebutRespect.Caption = "Respect début prévu : Aucune donnée"
    End If
    
    ' === Aucun retard ===
    If lstRetards.ListCount = 1 Then
        lstRetards.AddItem
        lstRetards.List(1, 0) = "Aucun retard détecté"
    End If
End Sub

Private Sub cmdFermer1_Click()
    Unload Me
End Sub

Public Sub OuvrirIndicateursDelais()
    frmIndicateursDelais.Show
End Sub

