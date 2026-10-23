VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   8748.001
   ClientLeft      =   84
   ClientTop       =   360
   ClientWidth     =   14640
   OleObjectBlob   =   "2d.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub OuvrirReduction10()
    frmReduc10.Show
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "Commandes éligibles à la réduction de 10 %"
    
    With lstReduc
        .Clear
        .ColumnCount = 5
        .ColumnWidths = "70pt;150pt;90pt;70pt;90pt"
        ' Colonnes : Commande | Client | Montant HT | Fidélité | Réduction
    End With
    
    ' Chargement initial des données
    Call cmdRafraichir_Click
End Sub
Private Sub cmdRafraichir_Click()
    ' --- Vider complètement la ListBox ---
    lstReduc.Clear
    
    ' --- Références aux feuilles ---
    Dim wsCmd As Worksheet
    Dim wsCli As Worksheet
    Set wsCmd = ThisWorkbook.Sheets("COMMANDE")
    Set wsCli = ThisWorkbook.Sheets("CLIENT")
    
    ' --- Dictionnaire clients : id_client ? Array("Prénom Nom", fidelite) ---
    Dim dictCli As Object
    Set dictCli = CreateObject("Scripting.Dictionary")
    dictCli.CompareMode = vbTextCompare  ' Insensible à la casse (sécurité)
    
    Dim dernierLigneCli As Long
    dernierLigneCli = wsCli.Cells(wsCli.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long
    Dim idCli As String
    Dim prenom As String, nom As String
    Dim fidelite As Variant   ' On garde Variant pour éviter les erreurs de conversion
    
    For i = 2 To dernierLigneCli
        idCli = Trim(wsCli.Cells(i, 1).Value & "")   ' & "" force String et évite Null
        If idCli <> "" Then
            prenom = Trim(wsCli.Cells(i, 3).Value & "")
            nom = Trim(wsCli.Cells(i, 2).Value & "")
            fidelite = wsCli.Cells(i, 8).Value           ' True/False ou texte
            
            ' Conversion sûre en Boolean
            If VarType(fidelite) = vbBoolean Then
                ' déjà Boolean
            ElseIf UCase(Trim(fidelite & "")) = "TRUE" Or fidelite = 1 Then
                fidelite = True
            Else
                fidelite = False
            End If
            
            dictCli(idCli) = Array(prenom & " " & nom, fidelite)
        End If
    Next i
    
    ' --- Parcours des commandes ---
    Dim dernierLigneCmd As Long
    dernierLigneCmd = wsCmd.Cells(wsCmd.Rows.Count, 1).End(xlUp).Row
    
    Dim ligne As Long
    ligne = 0
    
    Dim idCmd As String, idClient As String, montant As Double
    
    For i = 2 To dernierLigneCmd
        idCmd = Trim(wsCmd.Cells(i, 1).Value & "")
        idClient = Trim(wsCmd.Cells(i, 8).Value & "")
        montant = Val(wsCmd.Cells(i, 7).Value & "0")  ' Val gère bien les vides
        
        If dictCli.Exists(idClient) Then
            Dim info As Variant
            info = dictCli(idClient)
            
            Dim nomComplet As String
            nomComplet = info(0)
            
            Dim estFidele As Boolean
            estFidele = info(1)
            
            Dim seuil As Double
            seuil = IIf(estFidele, 5000, 10000)
            
            If montant > seuil Then
                lstReduc.AddItem
                lstReduc.List(ligne, 0) = idCmd
                lstReduc.List(ligne, 1) = nomComplet
                lstReduc.List(ligne, 2) = Format(montant, "#,##0 €")
                lstReduc.List(ligne, 3) = IIf(estFidele, "Oui", "Non")
                lstReduc.List(ligne, 4) = "-10 %"
                ligne = ligne + 1
            End If
        End If
    Next i
    
    ' --- Si aucune commande éligible ---
    If ligne = 0 Then
        lstReduc.AddItem
        lstReduc.List(0, 0) = "Aucune commande éligible pour le moment"
        ' Les autres colonnes restent vides pour un affichage propre
    End If
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub
