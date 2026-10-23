VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_FiltreCommandes 
   Caption         =   "UserForm1"
   ClientHeight    =   8628.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8904.001
   OleObjectBlob   =   "UserForm_FiltreCommandes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_FiltreCommandes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FrameListeCmd_Click()
End Sub
Private Sub UserForm_Initialize()

    lstCmd.ColumnCount = 7
    lstCmd.ColumnWidths = "60;150;110;100;100;100;80"

End Sub

Private Sub BtnAdd_Click()
    Dim i As Long
    For i = lstDispo.ListCount - 1 To 0 Step -1
        If lstDispo.Selected(i) Then
            lstSelect.AddItem lstDispo.List(i)
            lstDispo.RemoveItem i
        End If
    Next i
End Sub
Private Sub BtnRemove_Click()
    Dim i As Long
    For i = lstSelect.ListCount - 1 To 0 Step -1
        If lstSelect.Selected(i) Then
            lstDispo.AddItem lstSelect.List(i)
            lstSelect.RemoveItem i
        End If
    Next i
End Sub
Private Sub btnFiltrer_Click()

    Dim wsCmd As Worksheet
    Dim wsStatut As Worksheet
    Dim wsClient As Worksheet
    Dim i As Long, lastCmd As Long, lastStatut As Long, lastClient As Long
    Dim statutID As Variant, clientID As Variant
    Dim statutLib As String, clientNom As String
    Dim filtre As Variant, ok As Boolean
    
    Dim dictStatuts As Object
    Dim dictClients As Object

    ' --- Charger les feuilles ---
    On Error Resume Next
    Set wsCmd = ThisWorkbook.Sheets("COMMANDE")
    Set wsStatut = ThisWorkbook.Sheets("STATUT_COMMANDE")
    Set wsClient = ThisWorkbook.Sheets("CLIENT")
    On Error GoTo 0

    If wsCmd Is Nothing Or wsStatut Is Nothing Or wsClient Is Nothing Then
        MsgBox "Feuille COMMANDE, STATUT_COMMANDE ou CLIENT manquante.", vbCritical
        Exit Sub
    End If

    ' --- Dictionnaire des statuts ---
    Set dictStatuts = CreateObject("Scripting.Dictionary")

    lastStatut = wsStatut.Cells(wsStatut.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastStatut
        dictStatuts(wsStatut.Cells(i, "A").Value) = wsStatut.Cells(i, "B").Value
    Next i

    ' --- Dictionnaire des clients ---
    Set dictClients = CreateObject("Scripting.Dictionary")

    lastClient = wsClient.Cells(wsClient.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastClient
        dictClients(wsClient.Cells(i, "A").Value) = wsClient.Cells(i, "B").Value & " " & wsClient.Cells(i, "C").Value
    Next i

    ' --- Vérification sélection ---
    If lstSelect.ListCount = 0 Then
        MsgBox "Veuillez sélectionner au moins un statut.", vbExclamation
        Exit Sub
    End If

    lstCmd.Clear

    ' --- Lecture de COMMANDE ---
    lastCmd = wsCmd.Cells(wsCmd.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastCmd
        
        ' --- Récupération du statut (colonne N) ---
        statutID = wsCmd.Cells(i, "N").Value
        If dictStatuts.exists(statutID) Then
            statutLib = dictStatuts(statutID)
        Else
            statutLib = "Inconnu"
        End If

        ' --- Vérification du filtre (par libellé) ---
        ok = False
        For Each filtre In lstSelect.List
            If UCase(Trim(filtre)) = UCase(Trim(statutLib)) Then
                ok = True
                Exit For
            End If
        Next filtre

        If ok Then
        
            ' --- Récupération du client (colonne I) ---
            clientID = wsCmd.Cells(i, "I").Value
            
            If dictClients.exists(clientID) Then
                clientNom = dictClients(clientID)
            Else
                clientNom = "Client inconnu"
            End If

            ' --- Ajout dans lstCmd ---
            lstCmd.AddItem wsCmd.Cells(i, "A").Value                         ' ID commande
            lstCmd.List(lstCmd.ListCount - 1, 1) = clientNom                ' Nom + Prénom
            lstCmd.List(lstCmd.ListCount - 1, 2) = statutLib                ' Libellé statut
            lstCmd.List(lstCmd.ListCount - 1, 3) = wsCmd.Cells(i, "B").Value ' Date création
            lstCmd.List(lstCmd.ListCount - 1, 4) = wsCmd.Cells(i, "C").Value ' Date début
            lstCmd.List(lstCmd.ListCount - 1, 5) = wsCmd.Cells(i, "F").Value ' Date fin
            lstCmd.List(lstCmd.ListCount - 1, 6) = wsCmd.Cells(i, "G").Value ' Montant

        End If

    Next i

End Sub

