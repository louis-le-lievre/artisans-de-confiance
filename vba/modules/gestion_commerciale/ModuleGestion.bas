Attribute VB_Name = "ModuleGestion"
Option Explicit

Const SH_CMD As String = "COMMANDE"
Const SH_ART As String = "ARTISAN"
Const SH_CAT As String = "CATEGORIE_TRAVAUX"
Const SH_STATUT As String = "STATUT_COMMANDE"
Const SH_LIEN_MOTIF As String = "A_POUR_MOTIF"
Const SH_MOTIF As String = "MOTIF_RETARD"
Const SH_AVIS As String = "AVIS"
Const COL_DATE_CREA As Integer = 2
Const COL_FIN_PREVUE As Integer = 4
Const COL_FIN_REELLE As Integer = 6
Const COL_MONTANT As Integer = 7
Const COL_ID_ARTISAN As Integer = 10
Const COL_ID_CATEGORIE As Integer = 12
Const COL_STATUT As Integer = 14

Sub OuvrirMenuGestion()
    usfGestionCom.Show
End Sub

' 3.a

Sub BilanVentesAnnuelles_V2(Annee As Integer)
    Dim wsCmd As Worksheet: Set wsCmd = Sheets(SH_CMD)
    Dim i As Long
    Dim CA_N As Double, CA_N_1 As Double, Evolution As Double
    
    usfGestionCom.lstResultats.Clear
    
    For i = 2 To wsCmd.Cells(Rows.Count, 1).End(xlUp).Row
        If IsDate(wsCmd.Cells(i, COL_DATE_CREA)) Then
            If Year(wsCmd.Cells(i, COL_DATE_CREA)) = Annee Then
                CA_N = CA_N + wsCmd.Cells(i, COL_MONTANT).Value
            ElseIf Year(wsCmd.Cells(i, COL_DATE_CREA)) = Annee - 1 Then
                CA_N_1 = CA_N_1 + wsCmd.Cells(i, COL_MONTANT).Value
            End If
        End If
    Next i
    
    With usfGestionCom.lstResultats
        .AddItem: .List(0, 0) = "INDICATEUR": .List(0, 1) = "VALEUR"
        .AddItem: .List(1, 0) = "CA " & Annee: .List(1, 1) = Format(CA_N, "# ##0 €")
        .AddItem: .List(2, 0) = "CA " & (Annee - 1): .List(2, 1) = Format(CA_N_1, "# ##0 €")
        .AddItem: .List(3, 0) = "Évolution"
        If CA_N_1 > 0 Then .List(3, 1) = Format((CA_N - CA_N_1) / CA_N_1, "+0.00%;-0.00%") Else .List(3, 1) = "-"
    End With
End Sub


' 3.b

Sub BilanParCategorie_V2(Annee As Integer, FiltreCategorie As String)
    Dim wsCmd As Worksheet: Set wsCmd = Sheets(SH_CMD)
    Dim wsCat As Worksheet: Set wsCat = Sheets(SH_CAT)
    Dim dico As Object: Set dico = CreateObject("Scripting.Dictionary")
    Dim dicoNoms As Object: Set dicoNoms = CreateObject("Scripting.Dictionary")
    Dim i As Long, idCat As String, nomCat As String, mnt As Double
    
    usfGestionCom.lstResultats.Clear
    For i = 2 To wsCat.Cells(Rows.Count, 1).End(xlUp).Row
        dicoNoms(wsCat.Cells(i, 1).Value) = wsCat.Cells(i, 2).Value
    Next i
    
    For i = 2 To wsCmd.Cells(Rows.Count, 1).End(xlUp).Row
        If IsDate(wsCmd.Cells(i, COL_DATE_CREA)) Then
            If Annee = 0 Or Year(wsCmd.Cells(i, COL_DATE_CREA)) = Annee Then
                idCat = wsCmd.Cells(i, COL_ID_CATEGORIE).Value
                mnt = wsCmd.Cells(i, COL_MONTANT).Value
                If dicoNoms.Exists(idCat) Then
                    nomCat = dicoNoms(idCat)
                    If FiltreCategorie = "TOUTES LES CATEGORIES" Or FiltreCategorie = nomCat Then
                        If dico.Exists(nomCat) Then dico(nomCat) = dico(nomCat) + mnt Else dico.Add nomCat, mnt
                    End If
                End If
            End If
        End If
    Next i
    If dico.Count > 0 Then
        usfGestionCom.lstResultats.AddItem
        usfGestionCom.lstResultats.List(0, 0) = "CATEGORIE"
        usfGestionCom.lstResultats.List(0, 1) = "CA GLOBAL"
        Dim k As Variant, l As Integer: l = 1
        For Each k In dico.Keys
            usfGestionCom.lstResultats.AddItem
            usfGestionCom.lstResultats.List(l, 0) = k
            usfGestionCom.lstResultats.List(l, 1) = Format(dico(k), "# ##0 €")
            l = l + 1
        Next k
    Else
        usfGestionCom.lstResultats.AddItem "Aucune vente trouvée."
    End If
End Sub

' 3.c
Sub CommandesParArtisan(Annee As Integer, dept As String)
    Dim wsCmd As Worksheet: Set wsCmd = Sheets(SH_CMD)
    Dim wsArt As Worksheet: Set wsArt = Sheets(SH_ART)
    Dim dico As Object: Set dico = CreateObject("Scripting.Dictionary")
    Dim dicoArtValides As Object: Set dicoArtValides = CreateObject("Scripting.Dictionary")
    Dim i As Long, idArt As String, adresse As String, nomArt As String
    
    For i = 2 To wsArt.Cells(Rows.Count, 1).End(xlUp).Row
        adresse = wsArt.Cells(i, 3).Value
        If dept = "TOUS" Or InStr(1, adresse, dept) > 0 Then
            dicoArtValides(wsArt.Cells(i, 1).Value) = wsArt.Cells(i, 2).Value
        End If
    Next i
    For i = 2 To wsCmd.Cells(Rows.Count, 1).End(xlUp).Row
        If IsDate(wsCmd.Cells(i, COL_DATE_CREA)) Then
            If Year(wsCmd.Cells(i, COL_DATE_CREA)) = Annee Then
                idArt = wsCmd.Cells(i, COL_ID_ARTISAN).Value
                If dicoArtValides.Exists(idArt) Then
                    If wsCmd.Cells(i, COL_STATUT).Value = "STAT04" Then
                        nomArt = dicoArtValides(idArt)
                        If dico.Exists(nomArt) Then dico(nomArt) = dico(nomArt) + 1 Else dico.Add nomArt, 1
                    End If
                End If
            End If
        End If
    Next i
    usfGestionCom.lstResultats.Clear
    If dico.Count > 0 Then
        usfGestionCom.lstResultats.AddItem
        usfGestionCom.lstResultats.List(0, 0) = "ARTISAN"
        usfGestionCom.lstResultats.List(0, 1) = "NB COMMANDES"
        Dim ligne As Integer: ligne = 1
        Dim k As Variant
        For Each k In dico.Keys
            usfGestionCom.lstResultats.AddItem
            usfGestionCom.lstResultats.List(ligne, 0) = k
            usfGestionCom.lstResultats.List(ligne, 1) = dico(k)
            ligne = ligne + 1
        Next k
    Else
        usfGestionCom.lstResultats.AddItem "Aucune commande terminée (STAT04)."
    End If
End Sub


' 3.d
Sub BilanUrgence(Annee As Integer)
    Dim wsCmd As Worksheet: Set wsCmd = Sheets(SH_CMD)
    Dim i As Long, totalUrg As Double, totalGlobal As Double
    Dim dateCrea As Date, dateDebut As Date
    
    usfGestionCom.lstResultats.Clear
    totalUrg = 0: totalGlobal = 0
    
    For i = 2 To wsCmd.Cells(Rows.Count, 1).End(xlUp).Row
        If IsDate(wsCmd.Cells(i, COL_DATE_CREA)) And IsDate(wsCmd.Cells(i, 3)) Then
            If Annee = 0 Or Year(wsCmd.Cells(i, COL_DATE_CREA)) = Annee Then
                totalGlobal = totalGlobal + wsCmd.Cells(i, COL_MONTANT).Value
                dateCrea = wsCmd.Cells(i, COL_DATE_CREA)
                dateDebut = wsCmd.Cells(i, 3)
                If (dateDebut - dateCrea) <= 3 Then totalUrg = totalUrg + wsCmd.Cells(i, COL_MONTANT).Value
            End If
        End If
    Next i
    
    With usfGestionCom.lstResultats
        .AddItem: .List(0, 0) = "INDICATEUR": .List(0, 1) = "VALEUR"
        .AddItem: .List(1, 0) = "CA Travaux d'Urgence": .List(1, 1) = Format(totalUrg, "# ##0 €")
        .AddItem: .List(2, 0) = "CA Total (Global)": .List(2, 1) = Format(totalGlobal, "# ##0 €")
        .AddItem: .List(3, 0) = "Part Commerciale"
        If totalGlobal > 0 Then .List(3, 1) = Format(totalUrg / totalGlobal, "0.00%") Else .List(3, 1) = "0%"
    End With
End Sub


' 3.e
Sub AfficherRetards(Annee As Integer)
    Dim wsCmd As Worksheet: Set wsCmd = Sheets(SH_CMD)
    Dim wsArt As Worksheet: Set wsArt = Sheets(SH_ART)
    Dim wsCli As Worksheet: Set wsCli = Sheets("CLIENT")
    Dim i As Long, finReelle As Date, finPrevue As Date
    Dim idArt As String, idCli As String, nomArt As String, nomCli As String
    Dim nbJours As Long, ligne As Integer
    
    usfGestionCom.lstResultats.Clear
    With usfGestionCom.lstResultats
        .ColumnCount = 3: .ColumnWidths = "110;90;50"
        .AddItem: .List(0, 0) = "COMMANDE / CLIENT": .List(0, 1) = "ARTISAN": .List(0, 2) = "RETARD"
    End With
    
    ligne = 1
    For i = 2 To wsCmd.Cells(Rows.Count, 1).End(xlUp).Row
        If IsDate(wsCmd.Cells(i, COL_FIN_PREVUE)) And IsDate(wsCmd.Cells(i, COL_FIN_REELLE)) Then
            If Annee = 0 Or Year(wsCmd.Cells(i, COL_FIN_PREVUE)) = Annee Then
                finPrevue = wsCmd.Cells(i, COL_FIN_PREVUE)
                finReelle = wsCmd.Cells(i, COL_FIN_REELLE)
                If finReelle > finPrevue Then
                    nbJours = DateDiff("d", finPrevue, finReelle)
                    idArt = wsCmd.Cells(i, COL_ID_ARTISAN).Value
                    idCli = wsCmd.Cells(i, 9).Value
                    On Error Resume Next
                    nomArt = Application.VLookup(idArt, wsArt.Range("A:B"), 2, False)
                    nomCli = Application.VLookup(idCli, wsCli.Range("A:B"), 2, False)
                    On Error GoTo 0
                    usfGestionCom.lstResultats.AddItem
                    usfGestionCom.lstResultats.List(ligne, 0) = wsCmd.Cells(i, 1).Value & " - " & nomCli
                    usfGestionCom.lstResultats.List(ligne, 1) = nomArt
                    usfGestionCom.lstResultats.List(ligne, 2) = nbJours & " j"
                    ligne = ligne + 1
                End If
            End If
        End If
    Next i
    If ligne = 1 Then usfGestionCom.lstResultats.AddItem "Aucun retard détecté."
End Sub

' 3.f

Sub StatsQualite(Annee As Integer)
    Dim wsAvis As Worksheet: Set wsAvis = Sheets(SH_AVIS)
    Dim wsMotifLien As Worksheet: Set wsMotifLien = Sheets(SH_LIEN_MOTIF)
    Dim wsMotifDesc As Worksheet: Set wsMotifDesc = Sheets(SH_MOTIF)
    Dim i As Long, totalNote As Double, nbNotes As Long, Moyenne As Double
    Dim dicoMotifs As Object: Set dicoMotifs = CreateObject("Scripting.Dictionary")
    Dim idMotif As String, libelle As String, c As Range
    
    usfGestionCom.lstResultats.Clear
    For i = 2 To wsAvis.Cells(Rows.Count, 1).End(xlUp).Row
        If IsNumeric(wsAvis.Cells(i, 2).Value) Then
            totalNote = totalNote + wsAvis.Cells(i, 2).Value: nbNotes = nbNotes + 1
        End If
    Next i
    
    With usfGestionCom.lstResultats
        .AddItem: .List(0, 0) = "--- SATISFACTION ---": .List(0, 1) = ""
        .AddItem: .List(1, 0) = "Note Moyenne": .List(1, 1) = IIf(nbNotes > 0, Format(totalNote / nbNotes, "0.00") & " / 5", "N/A")
        .AddItem: .List(2, 0) = "Nombre d'avis": .List(2, 1) = nbNotes
        .AddItem ""
    End With
    
    For i = 2 To wsMotifLien.Cells(Rows.Count, 1).End(xlUp).Row
        idMotif = wsMotifLien.Cells(i, 1).Value
        If idMotif <> "" Then
            If dicoMotifs.Exists(idMotif) Then dicoMotifs(idMotif) = dicoMotifs(idMotif) + 1 Else dicoMotifs.Add idMotif, 1
        End If
    Next i
    
    With usfGestionCom.lstResultats
        .AddItem: .List(.ListCount - 1, 0) = "--- CAUSES RETARDS ---"
        If dicoMotifs.Count > 0 Then
            Dim k As Variant
            For Each k In dicoMotifs.Keys
                Set c = wsMotifDesc.Range("A:A").Find(k, LookIn:=xlValues)
                If Not c Is Nothing Then libelle = c.Offset(0, 1).Value Else libelle = k
                .AddItem: .List(.ListCount - 1, 0) = libelle: .List(.ListCount - 1, 1) = dicoMotifs(k) & " fois"
            Next k
        Else
            .AddItem: .List(.ListCount - 1, 0) = "Aucun incident."
        End If
    End With
End Sub
SAUVEGARDE
