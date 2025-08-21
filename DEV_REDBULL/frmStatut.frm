VERSION 5.00
Begin VB.Form frmStatut 
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
Attribute VB_Name = "frmStatut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
VERSION 5#
Begin VB.Form frmStatut
   Caption = "Changer le Statut"
   ClientHeight = 5535
   ClientLeft = 45
   ClientTop = 435
   ClientWidth = 6015
   LinkTopic = "Form1"
   ScaleHeight = 5535
   ScaleWidth = 6015
   StartUpPosition = 1    'CenterOwner
End
Attribute VB_Name = "frmStatut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' === D�CLARATIONS DES CONTR�LES (EN HAUT) ===
Dim Frame1 As Frame
Dim Frame2 As Frame
Dim lblIDLabel As Label
Dim lblIDStatut As Label
Dim lblTypeLabel As Label
Dim lblTypeStatut As Label
Dim lblModeleLabel As Label
Dim lblModeleStatut As Label
Dim lblStatutActuelLabel As Label
Dim lblStatutActuel As Label
Dim lblNouveauStatutLabel As Label
Dim cmbNouveauStatut As ComboBox
Dim lblDateLabel As Label
Dim txtDateStatut As TextBox
Dim lblNotesLabel As Label
Dim txtNotesStatut As TextBox
Dim lblPrioriteLabel As Label
Dim cmbPrioriteStatut As ComboBox
Dim lblTechnicienLabel As Label
Dim cmbTechnicienStatut As ComboBox
Dim cmdConfirmerStatut As CommandButton
Dim cmdAnnulerStatut As CommandButton

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
    ' === FRAME INFORMATIONS ACTUELLES ===
    Set Frame1 = Me.Controls.Add("VB.Frame", "Frame1")
    With Frame1
        .Left = 120
        .Top = 120
        .Width = 5775
        .Height = 2535
        .Caption = "Informations Actuelles"
        .Visible = True
    End With
    
    ' Labels dans Frame1
    Set lblIDLabel = Frame1.Controls.Add("VB.Label", "lblIDLabel")
    With lblIDLabel
        .Left = 240
        .Top = 420
        .Width = 1215
        .Height = 255
        .Caption = "ID �quipement:"
        .Visible = True
    End With
    
    Set lblIDStatut = Frame1.Controls.Add("VB.Label", "lblIDStatut")
    With lblIDStatut
        .Left = 1680
        .Top = 360
        .Width = 1215
        .Height = 255
        .Caption = "1"
        .Visible = True
    End With
    
    Set lblTypeLabel = Frame1.Controls.Add("VB.Label", "lblTypeLabel")
    With lblTypeLabel
        .Left = 240
        .Top = 780
        .Width = 375
        .Height = 255
        .Caption = "Type:"
        .Visible = True
    End With
    
    Set lblTypeStatut = Frame1.Controls.Add("VB.Label", "lblTypeStatut")
    With lblTypeStatut
        .Left = 1680
        .Top = 720
        .Width = 2535
        .Height = 255
        .Caption = "Frigo"
        .Visible = True
    End With
    
    Set lblModeleLabel = Frame1.Controls.Add("VB.Label", "lblModeleLabel")
    With lblModeleLabel
        .Left = 240
        .Top = 1140
        .Width = 615
        .Height = 255
        .Caption = "Mod�le:"
        .Visible = True
    End With
    
    Set lblModeleStatut = Frame1.Controls.Add("VB.Label", "lblModeleStatut")
    With lblModeleStatut
        .Left = 1680
        .Top = 1080
        .Width = 2535
        .Height = 255
        .Caption = "RB-2024-001"
        .Visible = True
    End With
    
    Set lblStatutActuelLabel = Frame1.Controls.Add("VB.Label", "lblStatutActuelLabel")
    With lblStatutActuelLabel
        .Left = 240
        .Top = 1500
        .Width = 1095
        .Height = 255
        .Caption = "Statut actuel:"
        .Visible = True
    End With
    
    Set lblStatutActuel = Frame1.Controls.Add("VB.Label", "lblStatutActuel")
    With lblStatutActuel
        .Left = 1680
        .Top = 1440
        .Width = 2535
        .Height = 255
        .Caption = "Stock"
        .Font.Size = 9.75
        .Font.Bold = True
        .ForeColor = &HFF0000
        .Visible = True
    End With
    
    Set lblNouveauStatutLabel = Frame1.Controls.Add("VB.Label", "lblNouveauStatutLabel")
    With lblNouveauStatutLabel
        .Left = 240
        .Top = 1860
        .Width = 1335
        .Height = 255
        .Caption = "Nouveau statut:"
        .Visible = True
    End With
    
    Set cmbNouveauStatut = Frame1.Controls.Add("VB.ComboBox", "cmbNouveauStatut")
    With cmbNouveauStatut
        .Left = 1680
        .Top = 1800
        .Width = 2535
        .Height = 315
        .Visible = True
    End With
    
    ' === FRAME INFORMATIONS ADDITIONNELLES ===
    Set Frame2 = Me.Controls.Add("VB.Frame", "Frame2")
    With Frame2
        .Left = 120
        .Top = 2760
        .Width = 5775
        .Height = 2175
        .Caption = "Informations Additionnelles"
        .Visible = True
    End With
    
    Set lblDateLabel = Frame2.Controls.Add("VB.Label", "lblDateLabel")
    With lblDateLabel
        .Left = 240
        .Top = 300
        .Width = 375
        .Height = 255
        .Caption = "Date:"
        .Visible = True
    End With
    
    Set txtDateStatut = Frame2.Controls.Add("VB.TextBox", "txtDateStatut")
    With txtDateStatut
        .Left = 1320
        .Top = 240
        .Width = 2055
        .Height = 315
        .Visible = True
    End With
    
    Set lblNotesLabel = Frame2.Controls.Add("VB.Label", "lblNotesLabel")
    With lblNotesLabel
        .Left = 240
        .Top = 540
        .Width = 495
        .Height = 255
        .Caption = "Notes:"
        .Visible = True
    End With
    
    Set txtNotesStatut = Frame2.Controls.Add("VB.TextBox", "txtNotesStatut")
    With txtNotesStatut
        .Left = 1320
        .Top = 480
        .Width = 4335
        .Height = 1095
        .MultiLine = True
        .ScrollBars = 2 ' Vertical
        .Visible = True
    End With
    
    Set lblPrioriteLabel = Frame2.Controls.Add("VB.Label", "lblPrioriteLabel")
    With lblPrioriteLabel
        .Left = 3600
        .Top = 1440
        .Width = 615
        .Height = 255
        .Caption = "Priorit�:"
        .Visible = True
    End With
    
    Set cmbPrioriteStatut = Frame2.Controls.Add("VB.ComboBox", "cmbPrioriteStatut")
    With cmbPrioriteStatut
        .Left = 3600
        .Top = 1680
        .Width = 1575
        .Height = 315
        .Visible = True
    End With
    
    Set lblTechnicienLabel = Frame2.Controls.Add("VB.Label", "lblTechnicienLabel")
    With lblTechnicienLabel
        .Left = 240
        .Top = 1740
        .Width = 975
        .Height = 255
        .Caption = "Technicien:"
        .Visible = True
    End With
    
    Set cmbTechnicienStatut = Frame2.Controls.Add("VB.ComboBox", "cmbTechnicienStatut")
    With cmbTechnicienStatut
        .Left = 1320
        .Top = 1680
        .Width = 2055
        .Height = 315
        .Visible = True
    End With
    
    ' === BOUTONS ===
    Set cmdConfirmerStatut = Me.Controls.Add("VB.CommandButton", "cmdConfirmerStatut")
    With cmdConfirmerStatut
        .Left = 3360
        .Top = 5040
        .Width = 1215
        .Height = 375
        .Caption = "Confirmer"
        .Visible = True
    End With
    
    Set cmdAnnulerStatut = Me.Controls.Add("VB.CommandButton", "cmdAnnulerStatut")
    With cmdAnnulerStatut
        .Left = 4680
        .Top = 5040
        .Width = 1215
        .Height = 375
        .Caption = "Annuler"
        .Visible = True
    End With
End Sub

Private Sub InitialiserFormulaire()
    ' Date par d�faut
    txtDateStatut.Text = Format(Date, "dd/mm/yyyy")
    
    ' Remplir les combobox
    ' Statuts
    cmbNouveauStatut.AddItem "R�ception"
    cmbNouveauStatut.AddItem "Stock"
    cmbNouveauStatut.AddItem "Pr�paration"
    cmbNouveauStatut.AddItem "Exp�dition"
    cmbNouveauStatut.AddItem "Retour"
    cmbNouveauStatut.AddItem "Diagnostic"
    cmbNouveauStatut.AddItem "Attente Pi�ces"
    cmbNouveauStatut.AddItem "R�parable"
    cmbNouveauStatut.AddItem "Donneur Pi�ces"
    cmbNouveauStatut.AddItem "Atelier"
    cmbNouveauStatut.AddItem "Stock Pr�t"
    cmbNouveauStatut.Text = "Stock"
    
    ' Priorit�s
    cmbPrioriteStatut.AddItem "Haute"
    cmbPrioriteStatut.AddItem "Normale"
    cmbPrioriteStatut.AddItem "Basse"
    cmbPrioriteStatut.Text = "Normale"
    
    ' Techniciens
    cmbTechnicienStatut.AddItem "Martin L."
    cmbTechnicienStatut.AddItem "Sophie M."
    cmbTechnicienStatut.AddItem "Jean-Paul D."
    cmbTechnicienStatut.AddItem "Marie C."
    cmbTechnicienStatut.AddItem "Pierre R."
End Sub

' === �V�NEMENTS ===
Private Sub cmdConfirmerStatut_Click()
    ' Validation
    If cmbNouveauStatut.Text = "" Then
        MsgBox "Veuillez s�lectionner un nouveau statut.", vbExclamation
        cmbNouveauStatut.SetFocus
        Exit Sub
    End If
    
    ' V�rifications sp�cifiques selon le statut
    Select Case cmbNouveauStatut.Text
        Case "Diagnostic", "Attente Pi�ces", "R�parable", "Donneur Pi�ces", "Atelier", "Stock Pr�t"
            If cmbTechnicienStatut.Text = "" Then
                MsgBox "Veuillez assigner un technicien pour les op�rations de r�paration.", vbExclamation
                cmbTechnicienStatut.SetFocus
                Exit Sub
            End If
    End Select
    
    ' Confirmer le changement
    Dim message As String
    message = "Confirmer le changement de statut ?" & vbCrLf & vbCrLf
    message = message & "De: " & lblStatutActuel.Caption & vbCrLf
    message = message & "Vers: " & cmbNouveauStatut.Text & vbCrLf & vbCrLf
    message = message & "Date: " & txtDateStatut.Text
    
    If MsgBox(message, vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    MsgBox "Statut mis � jour avec succ�s!" & vbCrLf & _
           "Nouveau statut: " & cmbNouveauStatut.Text & vbCrLf & _
           "Date: " & txtDateStatut.Text, vbInformation
    
    ' Fermer le formulaire
    Unload Me
End Sub

Private Sub cmdAnnulerStatut_Click()
    If MsgBox("Annuler les modifications?", vbQuestion + vbYesNo) = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmbNouveauStatut_Click()
    ' Adapter les champs selon le statut s�lectionn�
    Select Case cmbNouveauStatut.Text
        Case "Diagnostic", "Attente Pi�ces", "R�parable", "Donneur Pi�ces", "Atelier", "Stock Pr�t"
            ' Statuts de r�paration - activer les champs techniques
            cmbTechnicienStatut.Enabled = True
            cmbPrioriteStatut.Enabled = True
            
            ' Valeurs par d�faut
            If cmbTechnicienStatut.Text = "" Then
                cmbTechnicienStatut.ListIndex = 0
            End If
            
        Case Else
            ' Statuts normaux - d�sactiver les champs techniques
            cmbTechnicienStatut.Enabled = False
            cmbPrioriteStatut.Enabled = False
    End Select
    
    ' Suggestions de notes selon le statut
    Select Case cmbNouveauStatut.Text
        Case "R�ception"
            txtNotesStatut.Text = "�quipement re�u et v�rifi�"
        Case "Stock"
            txtNotesStatut.Text = "�quipement contr�l� et mis en stock"
        Case "Pr�paration"
            txtNotesStatut.Text = "Pr�paration pour exp�dition"
        Case "Exp�dition"
            txtNotesStatut.Text = "Exp�dition en cours"
        Case "Diagnostic"
            txtNotesStatut.Text = "Diagnostic technique en cours"
        Case "Attente Pi�ces"
            txtNotesStatut.Text = "En attente de pi�ces d�tach�es"
        Case "R�parable"
            txtNotesStatut.Text = "�quipement identifi� comme r�parable"
        Case "Donneur Pi�ces"
            txtNotesStatut.Text = "�quipement utilis� comme donneur de pi�ces"
        Case "Atelier"
            txtNotesStatut.Text = "R�paration en cours dans l'atelier"
        Case "Stock Pr�t"
            txtNotesStatut.Text = "R�paration termin�e - Pr�t � exp�dier"
    End Select
End Sub

