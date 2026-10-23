VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfGestionCom 
   Caption         =   "Menu gestion commerciale "
   ClientHeight    =   9012.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5160
   OleObjectBlob   =   "usfGestionCom.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfGestionCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    RemplirAnnees
    RemplirRegionsDisponibles
    RemplirCategories
    CacherDetails
    ModeMenuPrincipal
End Sub

Private Sub btnRetour_Click()
    ModeMenuPrincipal
End Sub

Private Sub ModeMenuPrincipal()
    Me.btn3a.Visible = True: Me.btn3b.Visible = True
    Me.btn3c.Visible = True: Me.btn3d.Visible = True
    Me.btn3e.Visible = True: Me.btn3f.Visible = True
    Me.btnRetour.Visible = False
    CacherDetails
End Sub

Private Sub ModeFocus()
    Me.btn3a.Visible = False: Me.btn3b.Visible = False
    Me.btn3c.Visible = False: Me.btn3d.Visible = False
    Me.btn3e.Visible = False: Me.btn3f.Visible = False
    Me.btnRetour.Visible = True
    Me.lstResultats.Clear
End Sub

Private Sub CacherDetails()
    Me.lblAnnee.Visible = False: Me.cmbAnnee.Visible = False
    Me.lblRegion.Visible = False: Me.cmbRegion.Visible = False
    Me.lblDepartement.Visible = False: Me.cmbDepartement.Visible = False
    Me.cmbCategorie.Visible = False
    Me.lstResultats.Visible = False
End Sub


' 3.a
Private Sub btn3a_Click()
    ModeFocus
    Me.lblAnnee.Visible = True: Me.cmbAnnee.Visible = True
    Me.lstResultats.Visible = True
    Me.lstResultats.ColumnCount = 2: Me.lstResultats.ColumnWidths = "140;80"
    If ValiderSaisie Then Call ModuleGestion.BilanVentesAnnuelles_V2(CInt(Me.cmbAnnee.Value))
End Sub

' 3.b
Private Sub btn3b_Click()
    ModeFocus
    Me.cmbCategorie.Visible = True
    Me.lstResultats.Visible = True
    Me.lstResultats.ColumnCount = 2: Me.lstResultats.ColumnWidths = "140;80"
    Me.cmbCategorie.Value = "TOUTES LES CATEGORIES"
    Call ModuleGestion.BilanParCategorie_V2(0, "TOUTES LES CATEGORIES")
End Sub

' 3.c
Private Sub btn3c_Click()
    ModeFocus
    Me.lblAnnee.Visible = True: Me.cmbAnnee.Visible = True
    Me.lblRegion.Visible = True: Me.cmbRegion.Visible = True
    Me.lblDepartement.Visible = True: Me.cmbDepartement.Visible = True
    Me.lstResultats.Visible = True
    Me.cmbRegion.Value = "TOUTES LES REGIONS"
    Me.cmbDepartement.Value = ""
    Me.lstResultats.ColumnCount = 2: Me.lstResultats.ColumnWidths = "140;80"
    If ValiderSaisie Then Call ModuleGestion.CommandesParArtisan(CInt(Me.cmbAnnee.Value), "TOUS")
End Sub

' 3.d
Private Sub btn3d_Click()
    ModeFocus
    Me.lstResultats.Visible = True
    Me.lstResultats.ColumnCount = 2: Me.lstResultats.ColumnWidths = "140;80"
    Call ModuleGestion.BilanUrgence(0)
End Sub

' 3.e
Private Sub btn3e_Click()
    ModeFocus
    Me.lstResultats.Visible = True
    Me.lstResultats.ColumnCount = 3: Me.lstResultats.ColumnWidths = "110;90;50"
    Call ModuleGestion.AfficherRetards(0)
End Sub

' 3.f
Private Sub btn3f_Click()
    ModeFocus
    Me.lstResultats.Visible = True
    Me.lstResultats.ColumnCount = 2: Me.lstResultats.ColumnWidths = "140;80"
    Call ModuleGestion.StatsQualite(0)
End Sub


Private Sub cmbRegion_Change()
    If Me.cmbRegion.Visible = True Then
        If Me.cmbRegion.Value = "TOUTES LES REGIONS" Then
            RemplirDepartementsFiltres ""
            Me.cmbDepartement.Value = ""
            If ValiderSaisie Then Call ModuleGestion.CommandesParArtisan(CInt(Me.cmbAnnee.Value), "TOUS")
        Else
            RemplirDepartementsFiltres Me.cmbRegion.Value
            Me.lstResultats.Clear
        End If
    End If
End Sub

Private Sub cmbDepartement_Change()
    If Me.cmbDepartement.Visible = True And Me.cmbDepartement.Value <> "" Then
        If ValiderSaisie Then
            Call ModuleGestion.CommandesParArtisan(CInt(Me.cmbAnnee.Value), Me.cmbDepartement.Value)
        End If
    End If
End Sub

Private Sub cmbCategorie_Change()
    If Me.cmbCategorie.Visible = True Then
        Call ModuleGestion.BilanParCategorie_V2(0, Me.cmbCategorie.Value)
    End If
End Sub

Private Sub cmbAnnee_Change()
    If Me.lstResultats.Visible = True And ValiderSaisie Then
        If Me.cmbDepartement.Visible = True Then
            If Me.cmbRegion.Value = "TOUTES LES REGIONS" Then
                Call ModuleGestion.CommandesParArtisan(CInt(Me.cmbAnnee.Value), "TOUS")
            ElseIf Me.cmbDepartement.Value <> "" Then
                Call ModuleGestion.CommandesParArtisan(CInt(Me.cmbAnnee.Value), Me.cmbDepartement.Value)
            End If
        ElseIf Me.cmbCategorie.Visible = False And Me.lblAnnee.Visible = True Then
             Call ModuleGestion.BilanVentesAnnuelles_V2(CInt(Me.cmbAnnee.Value))
        End If
    End If
End Sub


Private Sub RemplirAnnees()
    Dim i As Long, dico As Object, wsCmd As Worksheet
    Set dico = CreateObject("Scripting.Dictionary")
    Set wsCmd = ThisWorkbook.Sheets("COMMANDE")
    Me.cmbAnnee.Clear
    For i = 2 To wsCmd.Cells(wsCmd.Rows.Count, 1).End(xlUp).Row
        If IsDate(wsCmd.Cells(i, 2).Value) Then
            Dim an As Integer: an = Year(wsCmd.Cells(i, 2).Value)
            If Not dico.Exists(an) Then dico.Add an, an
        End If
    Next i
    If Not dico.Exists(2024) Then dico.Add 2024, 2024
    If Not dico.Exists(2025) Then dico.Add 2025, 2025
    Dim t() As Variant: t = dico.Keys: TrierTableau t
    For i = LBound(t) To UBound(t): Me.cmbAnnee.AddItem t(i): Next i
    Me.cmbAnnee.Value = "2025"
End Sub

Private Sub RemplirRegionsDisponibles()
    Dim i As Long, dicoReg As Object, valCell As String, dept As String, nomRegion As String
    Dim wsArt As Worksheet, tabSplit() As String
    Set dicoReg = CreateObject("Scripting.Dictionary")
    Set wsArt = ThisWorkbook.Sheets("ARTISAN")
    Me.cmbRegion.Clear: Me.cmbRegion.AddItem "TOUTES LES REGIONS"
    For i = 2 To wsArt.Cells(wsArt.Rows.Count, 1).End(xlUp).Row
        valCell = wsArt.Cells(i, 3).Value
        If InStr(valCell, ",") > 0 Then
            tabSplit = Split(valCell, ",")
            valCell = Trim(tabSplit(1))
            If Len(valCell) >= 2 Then
                dept = Left(valCell, 2)
                If IsNumeric(dept) Then
                    nomRegion = TrouverRegion(dept)
                    If Not dicoReg.Exists(nomRegion) Then dicoReg.Add nomRegion, nomRegion
                End If
            End If
        End If
    Next i
    Dim k As Variant
    For Each k In dicoReg.Keys: Me.cmbRegion.AddItem k: Next k
End Sub

Private Sub RemplirDepartementsFiltres(RegionFiltre As String)
    Dim i As Long, dicoDept As Object, valCell As String, dept As String, regionDuDept As String
    Dim wsArt As Worksheet, tabSplit() As String
    Set dicoDept = CreateObject("Scripting.Dictionary")
    Set wsArt = ThisWorkbook.Sheets("ARTISAN")
    Me.cmbDepartement.Clear
    
    For i = 2 To wsArt.Cells(wsArt.Rows.Count, 1).End(xlUp).Row
        valCell = wsArt.Cells(i, 3).Value
        If InStr(valCell, ",") > 0 Then
            tabSplit = Split(valCell, ",")
            valCell = Trim(tabSplit(1))
            If Len(valCell) >= 2 Then
                dept = Left(valCell, 2)
                If IsNumeric(dept) Then
                    regionDuDept = TrouverRegion(dept)
                    If RegionFiltre = "" Or RegionFiltre = regionDuDept Then
                        If Not dicoDept.Exists(dept) Then dicoDept.Add dept, dept
                    End If
                End If
            End If
        End If
    Next i
    
    If dicoDept.Count > 0 Then
        Dim tabD() As Variant: tabD = dicoDept.Keys: TrierTableau tabD
        For i = LBound(tabD) To UBound(tabD): Me.cmbDepartement.AddItem tabD(i): Next i
    End If
End Sub

Private Function TrouverRegion(CodeDept As String) As String
    Select Case CodeDept
        Case "75", "77", "78", "91", "92", "93", "94", "95": TrouverRegion = "Ile-de-France"
        Case "13", "83", "06", "84", "04", "05": TrouverRegion = "PACA"
        Case "69", "01", "38", "73", "74", "42", "26", "07": TrouverRegion = "Auvergne-Rhone-Alpes"
        Case "33", "24", "40", "47", "64": TrouverRegion = "Nouvelle-Aquitaine"
        Case "31", "34", "30", "11", "66": TrouverRegion = "Occitanie"
        Case "44", "85", "49", "72", "53": TrouverRegion = "Pays de la Loire"
        Case "59", "62": TrouverRegion = "Hauts-de-France"
        Case "67", "68": TrouverRegion = "Grand Est"
        Case Else: TrouverRegion = "Autres Régions"
    End Select
End Function

Private Sub RemplirCategories()
    Dim i As Long, wsCat As Worksheet
    Set wsCat = ThisWorkbook.Sheets("CATEGORIE_TRAVAUX")
    Me.cmbCategorie.Clear
    On Error Resume Next
    For i = 2 To wsCat.Cells(wsCat.Rows.Count, 1).End(xlUp).Row
        Me.cmbCategorie.AddItem wsCat.Cells(i, 2).Value
    Next i
    On Error GoTo 0
    Me.cmbCategorie.AddItem "TOUTES LES CATEGORIES"
End Sub

Private Sub TrierTableau(ByRef Arr As Variant)
    Dim i As Long, j As Long, temp As Variant
    If UBound(Arr) < LBound(Arr) Then Exit Sub
    For i = LBound(Arr) To UBound(Arr) - 1
        For j = i + 1 To UBound(Arr)
            If Arr(i) > Arr(j) Then
                temp = Arr(i): Arr(i) = Arr(j): Arr(j) = temp
            End If
        Next j
    Next i
End Sub

Private Function ValiderSaisie() As Boolean
    ValiderSaisie = IsNumeric(Me.cmbAnnee.Value)
End Function
