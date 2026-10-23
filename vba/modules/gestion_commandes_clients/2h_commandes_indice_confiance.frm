VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   7572
   ClientLeft      =   84
   ClientTop       =   360
   ClientWidth     =   18960
   OleObjectBlob   =   "2hh.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Dim wsCmd As Worksheet, wsClient As Worksheet, wsArt As Worksheet
    Dim wsAvis As Worksheet, wsMotif As Worksheet, wsAPour As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set wsCmd = ThisWorkbook.Sheets("COMMANDE")
    Set wsClient = ThisWorkbook.Sheets("CLIENT")
    Set wsArt = ThisWorkbook.Sheets("ARTISAN")
    Set wsAvis = ThisWorkbook.Sheets("AVIS")
    Set wsMotif = ThisWorkbook.Sheets("MOTIF_RETARD")
    Set wsAPour = ThisWorkbook.Sheets("A_POUR_MOTIF")
    
    With Me.lstCommandes
        .Clear
        .ColumnCount = 8
        .ColumnWidths = "80;100;100;80;80;60;60;120"
        
        ' En-tête
        .AddItem
        .List(0, 0) = "ID Commande"
        .List(0, 1) = "Client"
        .List(0, 2) = "Artisan"
        .List(0, 3) = "Début prévu"
        .List(0, 4) = "Fin prévue"
        .List(0, 5) = "Conf. Délais"
        .List(0, 6) = "Conf. Qualité"
        .List(0, 7) = "Motif retard"
    End With
    
    Dim today As Date
    today = DateSerial(2025, 12, 13)
    
    Dim lastRow As Long
    lastRow = wsCmd.Cells(wsCmd.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        If Trim(wsCmd.Cells(i, 13).Value) = "STAT03" Then  ' EN_COURS
            Dim idCmd As String: idCmd = Trim(wsCmd.Cells(i, 1).Value)
            Dim debutPrev As Date: debutPrev = wsCmd.Cells(i, 3).Value
            Dim finPrev As Variant: finPrev = wsCmd.Cells(i, 4).Value
            
            Dim idClient As String: idClient = Trim(wsCmd.Cells(i, 8).Value)
            Dim idArtisan As String: idArtisan = Trim(wsCmd.Cells(i, 9).Value)
            
            Dim nomClient As String: nomClient = GetNomClient(wsClient, idClient)
            Dim nomArtisan As String: nomArtisan = GetNomArtisan(wsArt, idArtisan)
            
            ' --- Confiance Délais ---
            Dim confDelais As Double
            If IsDate(finPrev) And finPrev <> 0 Then
                If today <= debutPrev Then
                    confDelais = 100
                ElseIf today >= finPrev Then
                    confDelais = 0
                Else
                    Dim joursTotal As Double: joursTotal = finPrev - debutPrev
                    If joursTotal > 0 Then
                        Dim joursEcoules As Double: joursEcoules = today - debutPrev
                        confDelais = Round(100 * (1 - joursEcoules / joursTotal), 0)
                    Else
                        confDelais = 0
                    End If
                End If
            Else
                confDelais = 100
            End If
            
            If HasMotifRetard(wsAPour, idCmd) Then
                confDelais = Application.Max(0, confDelais - 30)
            End If
            
            ' --- Confiance Qualité (fonction simplifiée et sécurisée) ---
            Dim confQualite As Double
            confQualite = GetConfianceQualiteArtisan(wsAvis, wsCmd, idArtisan)
            
            ' --- Motif retard ---
            Dim motif As String
            motif = GetMotifLibelle(wsAPour, wsMotif, idCmd)
            
            ' Ajout ligne
            With Me.lstCommandes
                .AddItem
                Dim idx As Long: idx = .ListCount - 1
                .List(idx, 0) = idCmd
                .List(idx, 1) = nomClient
                .List(idx, 2) = nomArtisan
                .List(idx, 3) = Format(debutPrev, "dd/mm/yyyy")
                .List(idx, 4) = IIf(IsDate(finPrev), Format(finPrev, "dd/mm/yyyy"), "-")
                .List(idx, 5) = Format(confDelais, "0") & " %"
                .List(idx, 6) = Format(confQualite, "0") & " %"
                .List(idx, 7) = motif
            End With
        End If
    Next i
    
    Me.lblTitre.Caption = "Commandes en cours au " & Format(today, "dd/mm/yyyy") & _
                          " (" & (Me.lstCommandes.ListCount - 1) & " commande(s))"
    Exit Sub

ErrorHandler:
    MsgBox "Erreur " & Err.Number & " : " & Err.Description & vbCrLf & _
           "Ligne concernée : probablement une recherche Find", vbCritical
End Sub


Private Function GetNomClient(ws As Worksheet, idClient As String) As String
    Dim rng As Range
    Set rng = ws.Columns(1).Find(What:=idClient, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rng Is Nothing Then
        GetNomClient = Trim(ws.Cells(rng.Row, 2).Value & " " & ws.Cells(rng.Row, 3).Value)
    Else
        GetNomClient = "Inconnu"
    End If
End Function

Private Function GetNomArtisan(ws As Worksheet, idArtisan As String) As String
    Dim rng As Range
    Set rng = ws.Columns(1).Find(What:=idArtisan, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rng Is Nothing Then
        GetNomArtisan = Trim(ws.Cells(rng.Row, 2).Value)
    Else
        GetNomArtisan = "Inconnu"
    End If
End Function

Private Function HasMotifRetard(wsAPour As Worksheet, idCmd As String) As Boolean
    Dim rng As Range
    Set rng = wsAPour.Columns(2).Find(What:=idCmd, LookIn:=xlValues, LookAt:=xlWhole)
    HasMotifRetard = Not rng Is Nothing
End Function

Private Function GetMotifLibelle(wsAPour As Worksheet, wsMotif As Worksheet, idCmd As String) As String
    Dim rng As Range
    Set rng = wsAPour.Columns(2).Find(What:=idCmd, LookIn:=xlValues, LookAt:=xlWhole)
    If rng Is Nothing Then
        GetMotifLibelle = "-"
        Exit Function
    End If
    
    Dim idMotif As String: idMotif = Trim(wsAPour.Cells(rng.Row, 1).Value)
    Dim rngMotif As Range
    Set rngMotif = wsMotif.Columns(1).Find(What:=idMotif, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rngMotif Is Nothing Then
        GetMotifLibelle = Trim(wsMotif.Cells(rngMotif.Row, 2).Value)
    Else
        GetMotifLibelle = "-"
    End If
End Function

Private Function GetConfianceQualiteArtisan(wsAvis As Worksheet, wsCmd As Worksheet, idArtisan As String) As Double
    Dim totalNotes As Double, nbNotes As Long
    totalNotes = 0: nbNotes = 0
    
    Dim i As Long
    For i = 2 To wsAvis.Cells(wsAvis.Rows.Count, 1).End(xlUp).Row
        If Trim(wsAvis.Cells(i, 8).Value) = idArtisan Then
            Dim idCmdAvis As String: idCmdAvis = Trim(wsAvis.Cells(i, 7).Value)
            If idCmdAvis <> "" Then
                Dim rowCmd As Long
                rowCmd = TrouverLigneCommande(wsCmd, idCmdAvis)
                If rowCmd > 0 Then
                    If Trim(wsCmd.Cells(rowCmd, 13).Value) = "STAT04" Then  ' TERMINEE seulement
                        totalNotes = totalNotes + Nz(wsAvis.Cells(i, 2).Value, 0)  ' note générale
                        totalNotes = totalNotes + Nz(wsAvis.Cells(i, 5).Value, 0)  ' finition
                        totalNotes = totalNotes + Nz(wsAvis.Cells(i, 6).Value, 0)  ' ponctualité
                        nbNotes = nbNotes + 3
                    End If
                End If
            End If
        End If
    Next i
    
    If nbNotes = 0 Then
        GetConfianceQualiteArtisan = 100
    Else
        GetConfianceQualiteArtisan = Round((totalNotes / nbNotes) * 20, 0)
    End If
End Function

Private Function TrouverLigneCommande(ws As Worksheet, idCmd As String) As Long
    Dim rng As Range
    Set rng = ws.Columns(1).Find(What:=idCmd, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rng Is Nothing Then TrouverLigneCommande = rng.Row Else TrouverLigneCommande = 0
End Function

Private Function Nz(v As Variant, default As Double) As Double
    If IsEmpty(v) Or IsNull(v) Or Not IsNumeric(v) Then
        Nz = default
    Else
        Nz = CDbl(v)
    End If
End Function

Private Sub btnFermer_Click()
    Unload Me
End Sub

