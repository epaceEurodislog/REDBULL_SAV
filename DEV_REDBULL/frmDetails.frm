VERSION 5.00
Begin VB.Form frmDetails 
   Caption         =   "Form2"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
VERSION 5#
Begin VB.Form frmDetails
   Caption = "D�tails de l'�quipement"
   ClientHeight = 7695
   ClientLeft = 45
   ClientTop = 435
   ClientWidth = 8175
   LinkTopic = "Form1"
   ScaleHeight = 7695
   ScaleWidth = 8175
   StartUpPosition = 1    'CenterOwner
End
Attribute VB_Name = "frmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' === D�CLARATIONS DES CONTR�LES (EN HAUT) ===
Dim cmdSauvegarder As CommandButton
Dim cmdFermer As CommandButton
Dim Frame1 As Frame
Dim Frame2 As Frame
Dim Frame3 As Frame
Dim lblTitreEquipement As Label
Dim lblStatutIndicateur As Label
Dim lblID As Label
Dim txtIDDetails As TextBox
Dim lblType As Label
Dim txtTypeDetails As TextBox
Dim lblModele As Label
Dim txtModeleDetails As TextBox
Dim lblStatut As Label
Dim cmbStatutDetails As ComboBox
Dim lblDestination As Label
Dim txtDestinationDetails As TextBox
Dim lblDate As Label
Dim txtDateDetails As TextBox
Dim lblRemarques As Label
Dim txtRemarquesDetails As TextBox
Dim lblDiagnostic As Label
Dim txtDiagnosticDetails As TextBox
Dim lblPriorite As Label
Dim cmbPrioriteDetails As ComboBox
Dim lblTechnicien As Label
Dim txtTechnicien As TextBox
Dim lblHistoriqueLabel As Label
Dim txtHistorique As TextBox

Private Sub Form_Load()
    ' Configurer le formulaire
    Me.BorderStyle = 3 ' Fixed Dialog
    Me.ControlBox = False
    Me.MaxButton = False
    Me.MinButton = False
    Me.ShowInTaskbar = False
    
    ' Cr�er les contr�les
    CreerControles
    
    ' Initialiser les donn�es
    InitialiserFormulaire
End Sub

Private Sub CreerControles()
    ' === BOUTONS DU HAUT ===
    Set cmdSauvegarder = Me.Controls.Add("VB.CommandButton", "cmdSauvegarder")
    With cmdSauvegarder
        .Left = 5760
        .Top = 120
        .Width = 1095
        .Height = 375
        .Caption = "Sauvegarder"
        .Visible = True
    End With
    
    Set cmdFermer = Me.Controls.Add("VB.CommandButton", "cmdFermer")
    With cmdFermer
        .Left = 6960
        .Top = 120
        .Width = 1095
        .Height = 375
        .Caption = "Fermer"
        .Visible = True
    End With
    
    ' === FRAME INFORMATIONS G�N�RALES ===
    Set Frame1 = Me.Controls.Add("VB.Frame", "Frame1")
    With Frame1
        .Left = 120
        .Top = 120
        .Width = 3855
        .Height = 4935
        .Caption = "Informations G�n�rales"
        .Visible = True
    End With
    
    Set lblTitreEquipement = Frame1.Controls.Add("VB.Label", "lblTitreEquipement")
    With lblTitreEquipement
        .Left = 120
        .Top = 360
        .Width = 3615
        .Height = 375
        .Caption = "�quipement #XXX"
        .Alignment = 2 ' Center
        .Font.Size = 12
        .Font.Bold = True
        .Visible = True
    End With
    
    Set lblStatutIndicateur = Frame1.Controls.Add("VB.Label", "lblStatutIndicateur")
    With lblStatutIndicateur
        .Left = 120
        .Top = 840
        .Width = 3615
        .Height = 375
        .Caption = "EN STOCK"
        .Alignment = 2 ' Center
        .BackColor = &H80FF80
        .ForeColor = &HFFFFFF
        .BorderStyle = 1 ' Fixed Single
        .Font.Size = 9.75
        .Font.Bold = True
        .Visible = True
    End With
    
    ' Champs dans Frame1
    Set lblID = Frame1.Controls.Add("VB.Label", "lblID")
    With lblID
        .Left = 120
        .Top = 1500
        .Width = 255
        .Height = 255
        .Caption = "ID:"
        .Visible = True
    End With
    
    Set txtIDDetails = Frame1.Controls.Add("VB.TextBox", "txtIDDetails")
    With txtIDDetails
        .Left = 1200
        .Top = 1440
        .Width = 1215
        .Height = 315
        .Locked = True
        .Visible = True
    End With
    
    Set lblType = Frame1.Controls.Add("VB.Label", "lblType")
    With lblType
        .Left = 120
        .Top = 1860
        .Width = 375
        .Height = 255
        .Caption = "Type:"
        .Visible = True
    End With
    
    Set txtTypeDetails = Frame1.Controls.Add("VB.TextBox", "txtTypeDetails")
    With txtTypeDetails
        .Left = 1200
        .Top = 1800
        .Width = 2535
        .Height = 315
        .Visible = True
    End With
    
    Set lblModele = Frame1.Controls.Add("VB.Label", "lblModele")
    With lblModele
        .Left = 120
        .Top = 2220
        .Width = 615
        .Height = 255
        .Caption = "Mod�le:"
        .Visible = True
    End With
    
    Set txtModeleDetails = Frame1.Controls.Add("VB.TextBox", "txtModeleDetails")
    With txtModeleDetails
        .Left = 1200
        .Top = 2160
        .Width = 2535
        .Height = 315
        .Visible = True
    End With
    
    Set lblStatut = Frame1.Controls.Add("VB.Label", "lblStatut")
    With lblStatut
        .Left = 120
        .Top = 2580
        .Width = 495
        .Height = 255
        .Caption = "Statut:"
        .Visible = True
    End With
    
    Set cmbStatutDetails = Frame1.Controls.Add("VB.ComboBox", "cmbStatutDetails")
    With cmbStatutDetails
        .Left = 1200
        .Top = 2520
        .Width = 2055
        .Height = 315
        .Visible = True
    End With
    
    Set lblDestination = Frame1.Controls.Add("VB.Label", "lblDestination")
    With lblDestination
        .Left = 120
        .Top = 2940
        .Width = 975
        .Height = 255
        .Caption = "Destination:"
        .Visible = True
    End With
    
    Set txtDestinationDetails = Frame1.Controls.Add("VB.TextBox", "txtDestinationDetails")
    With txtDestinationDetails
        .Left = 1200
        .Top = 2880
        .Width = 2535
        .Height = 315
        .Visible = True
    End With
    
    Set lblDate = Frame1.Controls.Add("VB.Label", "lblDate")
    With lblDate
        .Left = 120
        .Top = 3300
        .Width = 495
        .Height = 255
        .Caption = "Date:"
        .Visible = True
    End With
    
    Set txtDateDetails = Frame1.Controls.Add("VB.TextBox", "txtDateDetails")
    With txtDateDetails
        .Left = 1200
        .Top = 3240
        .Width = 2535
        .Height = 315
        .Visible = True
    End With
    
    Set lblRemarques = Frame1.Controls.Add("VB.Label", "lblRemarques")
    With lblRemarques
        .Left = 120
        .Top = 3780
        .Width = 975
        .Height = 255
        .Caption = "Remarques:"
        .Visible = True
    End With
    
    Set txtRemarquesDetails = Frame1.Controls.Add("VB.TextBox", "txtRemarquesDetails")
    With txtRemarquesDetails
        .Left = 1200
        .Top = 3720
        .Width = 2535
        .Height = 1215
        .MultiLine = True
        .ScrollBars = 2 ' Vertical
        .Visible = True
    End With
    
    ' === FRAME INFORMATIONS TECHNIQUES ===
    Set Frame2 = Me.Controls.Add("VB.Frame", "Frame2")
    With Frame2
        .Left = 4080
        .Top = 3000
        .Width = 3975
        .Height = 2175
        .Caption = "Informations Techniques"
        .Visible = True
    End With
    
    Set lblDiagnostic = Frame2.Controls.Add("VB.Label", "lblDiagnostic")
    With lblDiagnostic
        .Left = 120
        .Top = 420
        .Width = 975
        .Height = 255
        .Caption = "Diagnostic:"
        .Visible = True
    End With
    
    Set txtDiagnosticDetails = Frame2.Controls.Add("VB.TextBox", "txtDiagnosticDetails")
    With txtDiagnosticDetails
        .Left = 1200
        .Top = 360
        .Width = 2655
        .Height = 855
        .MultiLine = True
        .ScrollBars = 2 ' Vertical
        .Visible = True
    End With
    
    Set lblPriorite = Frame2.Controls.Add("VB.Label", "lblPriorite")
    With lblPriorite
        .Left = 120
        .Top = 1380
        .Width = 615
        .Height = 255
        .Caption = "Priorit�:"
        .Visible = True
    End With
    
    Set cmbPrioriteDetails = Frame2.Controls.Add("VB.ComboBox", "cmbPrioriteDetails")
    With cmbPrioriteDetails
        .Left = 1200
        .Top = 1320
        .Width = 1455
        .Height = 315
        .Visible = True
    End With
    
    Set lblTechnicien = Frame2.Controls.Add("VB.Label", "lblTechnicien")
    With lblTechnicien
        .Left = 120
        .Top = 1740
        .Width = 975
        .Height = 255
        .Caption = "Technicien:"
        .Visible = True
    End With
    
    Set txtTechnicien = Frame2.Controls.Add("VB.TextBox", "txtTechnicien")
    With txtTechnicien
        .Left = 1200
        .Top = 1680
        .Width = 2655
        .Height = 315
        .Visible = True
    End With
    
    ' === FRAME HISTORIQUE ===
    Set Frame3 = Me.Controls.Add("VB.Frame", "Frame3")
    With Frame3
        .Left = 120
        .Top = 5280
        .Width = 7935
        .Height = 2055
        .Caption = "Historique et Suivi"
        .Visible = True
    End With
    
    Set lblHistoriqueLabel = Frame3.Controls.Add("VB.Label", "lblHistoriqueLabel")
    With lblHistoriqueLabel
        .Left = 120
        .Top = 240
        .Width = 1935
        .Height = 255
        .Caption = "Historique des op�rations:"
        .Visible = True
    End With
    
    Set txtHistorique = Frame3.Controls.Add("VB.TextBox", "txtHistorique")
    With txtHistorique
        .Left = 120
        .Top = 360
        .Width = 7695
        .Height = 1575
        .MultiLine = True
        .ScrollBars = 2 ' Vertical
        .Locked = True
        .Visible = True
    End With
End Sub

Private Sub InitialiserFormulaire()
    ' Valeurs d'exemple
    txtIDDetails.Text = "1"
    txtTypeDetails.Text = "Frigo"
    txtModeleDetails.Text = "RB-2024-001"
    txtDestinationDetails.Text = "Magasin Paris"
    txtDateDetails.Text = Format(Date, "dd/mm/yyyy")
    txtRemarquesDetails.Text = "�quipement en bon �tat"
    
    lblTitreEquipement.Caption = "�quipement #" & txtIDDetails.Text
    
    ' Remplir les combobox
    ' Statuts
    cmbStatutDetails.AddItem "R�ception"
    cmbStatutDetails.AddItem "Stock"
    cmbStatutDetails.AddItem "Pr�paration"
    cmbStatutDetails.AddItem "Exp�dition"
    cmbStatutDetails.AddItem "Diagnostic"
    cmbStatutDetails.AddItem "R�parable"
    cmbStatutDetails.AddItem "Atelier"
    cmbStatutDetails.Text = "Stock"
    
    ' Priorit�s
    cmbPrioriteDetails.AddItem "Haute"
    cmbPrioriteDetails.AddItem "Normale"
    cmbPrioriteDetails.AddItem "Basse"
    cmbPrioriteDetails.Text = "Normale"
    
    GenererHistorique
End Sub

Private Sub GenererHistorique()
    Dim historique As String
    
    historique = "=== HISTORIQUE DE L'�QUIPEMENT ===" & vbCrLf & vbCrLf
    historique = historique & Format(Date, "dd/mm/yyyy") & " - �quipement en stock" & vbCrLf
    historique = historique & "� Destination: " & txtDestinationDetails.Text & vbCrLf & vbCrLf
    
    historique = historique & "=== �TAPES PR�C�DENTES ===" & vbCrLf
    historique = historique & Format(DateAdd("d", -3, Date), "dd/mm/yyyy") & " - R�ception de l'�quipement" & vbCrLf
    historique = historique & Format(DateAdd("d", -2, Date), "dd/mm/yyyy") & " - Contr�le qualit� effectu�" & vbCrLf
    historique = historique & Format(DateAdd("d", -1, Date), "dd/mm/yyyy") & " - Mise en stock" & vbCrLf
    
    txtHistorique.Text = historique
End Sub

' === �V�NEMENTS ===
Private Sub cmdSauvegarder_Click()
    ' Validation des champs
    If Trim(txtModeleDetails.Text) = "" Then
        MsgBox "Le mod�le ne peut pas �tre vide.", vbExclamation
        txtModeleDetails.SetFocus
        Exit Sub
    End If
    
    If Trim(txtDestinationDetails.Text) = "" Then
        MsgBox "La destination ne peut pas �tre vide.", vbExclamation
        txtDestinationDetails.SetFocus
        Exit Sub
    End If
    
    MsgBox "�quipement sauvegard� !" & vbCrLf & _
           "Type: " & txtTypeDetails.Text & vbCrLf & _
           "Mod�le: " & txtModeleDetails.Text & vbCrLf & _
           "Statut: " & cmbStatutDetails.Text, vbInformation
    
    ' Mettre � jour l'affichage
    MettreAJourIndicateurStatut cmbStatutDetails.Text
    GenererHistorique
End Sub

Private Sub cmdFermer_Click()
    Unload Me
End Sub

Private Sub cmbStatutDetails_Click()
    ' Mettre � jour l'indicateur en temps r�el
    MettreAJourIndicateurStatut cmbStatutDetails.Text
    
    ' Adapter les champs selon le statut
    Select Case cmbStatutDetails.Text
        Case "Diagnostic", "Attente Pi�ces", "R�parable", "Donneur Pi�ces", "Atelier", "Stock Pr�t"
            ' Activer les champs de r�paration
            txtTechnicien.Enabled = True
            cmbPrioriteDetails.Enabled = True
            txtDiagnosticDetails.Enabled = True
            txtDestinationDetails.Text = "Service R�paration"
        Case Else
            ' D�sactiver les champs de r�paration pour les autres statuts
            txtTechnicien.Enabled = False
            cmbPrioriteDetails.Enabled = False
            txtDiagnosticDetails.Enabled = False
    End Select
End Sub

Private Sub MettreAJourIndicateurStatut(statut As String)
    lblStatutIndicateur.Caption = UCase(statut)
    
    ' Changer la couleur selon le statut
    Select Case statut
        Case "R�ception"
            lblStatutIndicateur.BackColor = &H80FF80   ' Vert clair
        Case "Stock"
            lblStatutIndicateur.BackColor = &H8080FF   ' Bleu
        Case "Pr�paration"
            lblStatutIndicateur.BackColor = &HFFFF80   ' Jaune
            lblStatutIndicateur.ForeColor = &H0&       ' Texte noir pour lisibilit�
        Case "Exp�dition"
            lblStatutIndicateur.BackColor = &HFF8080   ' Rouge clair
        Case "Diagnostic"
            lblStatutIndicateur.BackColor = &HFF8000   ' Orange
        Case "Attente Pi�ces"
            lblStatutIndicateur.BackColor = &H8000FF   ' Violet
        Case "R�parable"
            lblStatutIndicateur.BackColor = &H80FFFF   ' Cyan
            lblStatutIndicateur.ForeColor = &H0&       ' Texte noir
        Case "Donneur Pi�ces"
            lblStatutIndicateur.BackColor = &H400040   ' Violet fonc�
        Case "Atelier"
            lblStatutIndicateur.BackColor = &H4080FF   ' Bleu orange
        Case "Stock Pr�t"
            lblStatutIndicateur.BackColor = &H40FF40   ' Vert
        Case Else
            lblStatutIndicateur.BackColor = &H808080   ' Gris
    End Select
    
    ' R�initialiser la couleur du texte si n�cessaire
    If lblStatutIndicateur.ForeColor <> &H0& Then
        lblStatutIndicateur.ForeColor = &HFFFFFF     ' Blanc
    End If
End Sub

