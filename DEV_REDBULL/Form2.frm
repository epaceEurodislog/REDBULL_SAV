VERSION 5.00
Begin VB.Form frmReparation 
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
Attribute VB_Name = "frmReparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
VERSION 5#
Begin VB.Form frmReparation
   Caption = "Demande de Réparation"
   ClientHeight = 6030
   ClientLeft = 45
   ClientTop = 435
   ClientWidth = 6735
   LinkTopic = "Form1"
   ScaleHeight = 6030
   ScaleWidth = 6735
   StartUpPosition = 1    'CenterOwner
End
Attribute VB_Name = "frmReparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' === DÉCLARATIONS DES CONTRÔLES (EN HAUT) ===
Dim Frame1 As Frame
Dim Frame2 As Frame
Dim lblReference As Label
Dim txtReference As TextBox
Dim lblType As Label
Dim cmbTypeRep As ComboBox
Dim lblModele As Label
Dim cmbModeleRep As ComboBox
Dim cmdNouveauModele As CommandButton
Dim lblProbleme As Label
Dim txtProbleme As TextBox
Dim lblDateEntree As Label
Dim txtDateEntree As TextBox
Dim lblStatutReparation As Label
Dim cmbStatutReparation As ComboBox
Dim lblDiagnostic As Label
Dim txtDiagnostic As TextBox
Dim lblPriorite As Label
Dim cmbPriorite As ComboBox
Dim lblTechnicien As Label
Dim cmbTechnicien As ComboBox
Dim cmdValider As CommandButton
Dim cmdAnnuler As CommandButton
Dim lblInfo As Label

Private Sub Form_Load()
    ' Configurer le formulaire
    Me.BorderStyle = 3 ' Fixed Dialog
    Me.ControlBox = False
    Me.MaxButton = False
    Me.MinButton = False
    Me.ShowInTaskbar = False
    
    ' Créer les contrôles
    CreerControles
    
    ' Initialiser les données
    InitialiserFormulaire
End Sub

Private Sub CreerControles()
    ' === FRAME INFORMATIONS ÉQUIPEMENT ===
    Set Frame1 = Me.Controls.Add("VB.Frame", "Frame1")
    With Frame1
        .Left = 120
        .Top = 120
        .Width = 6495
        .Height = 2535
        .Caption = "Informations Équipement"
        .Visible = True
    End With
    
    Set lblReference = Frame1.Controls.Add("VB.Label", "lblReference")
    With lblReference
        .Left = 240
        .Top = 300
        .Width = 855
        .Height = 255
        .Caption = "Référence:"
        .Visible = True
    End With
    
    Set txtReference = Frame1.Controls.Add("VB.TextBox", "txtReference")
    With txtReference
        .Left = 1320
        .Top = 240
        .Width = 2055
        .Height = 315
        .Visible = True
    End With
    
    Set lblType = Frame1.Controls.Add("VB.Label", "lblType")
    With lblType
        .Left = 240
        .Top = 660
        .Width = 375
        .Height = 255
        .Caption = "Type:"
        .Visible = True
    End With
    
    Set cmbTypeRep = Frame1.Controls.Add("VB.ComboBox", "cmbTypeRep")
    With cmbTypeRep
        .Left = 1320
        .Top = 600
        .Width = 2055
        .Height = 315
        .Text = "Frigo"
        .Visible = True
    End With
    
    Set lblModele = Frame1.Controls.Add("VB.Label", "lblModele")
    With lblModele
        .Left = 240
        .Top = 1020
        .Width = 615
        .Height = 255
        .Caption = "Modèle:"
        .Visible = True
    End With
    
    Set cmbModeleRep = Frame1.Controls.Add("VB.ComboBox", "cmbModeleRep")
    With cmbModeleRep
        .Left = 1320
        .Top = 960
        .Width = 2055
        .Height = 315
        .Visible = True
    End With
    
    Set cmdNouveauModele = Frame1.Controls.Add("VB.CommandButton", "cmdNouveauModele")
    With cmdNouveauModele
        .Left = 3480
        .Top = 960
        .Width = 855
        .Height = 315
        .Caption = "Nouveau"
        .Visible = True
    End With
    
    Set lblProbleme = Frame1.Controls.Add("VB.Label", "lblProbleme")
    With lblProbleme
        .Left = 240
        .Top = 1380
        .Width = 735
        .Height = 255
        .Caption = "Problème:"
        .Visible = True
    End With
    
    Set txtProbleme = Frame1.Controls.Add("VB.TextBox", "txtProbleme")
    With txtProbleme
        .Left = 1320
        .Top = 1320
        .Width = 4935
        .Height = 615
        .MultiLine = True
        .ScrollBars = 2 ' Vertical
        .Visible = True
    End With
    
    Set lblDateEntree = Frame1.Controls.Add("VB.Label", "lblDateEntree")
    With lblDateEntree
        .Left = 240
        .Top = 2100
        .Width = 975
        .Height = 255
        .Caption = "Date entrée:"
        .Visible = True
    End With
    
    Set txtDateEntree = Frame1.Controls.Add("VB.TextBox", "txtDateEntree")
    With txtDateEntree
        .Left = 1320
        .Top = 2040
        .Width = 2055
        .Height = 315
        .Visible = True
    End With
    
    ' === FRAME INFORMATIONS TECHNIQUES ===
    Set Frame2 = Me.Controls.Add("VB.Frame", "Frame2")
    With Frame2
        .Left = 120
        .Top = 2760
        .Width = 6495
        .Height = 2175
        .Caption = "Informations Techniques"
        .Visible = True
    End With
    
    Set lblStatutReparation = Frame2.Controls.Add("VB.Label", "lblStatutReparation")
    With lblStatutReparation
        .Left = 240
        .Top = 420
        .Width = 495
        .Height = 255
        .Caption = "Statut:"
        .Visible = True
    End With
    
    Set cmbStatutReparation = Frame2.Controls.Add("VB.ComboBox", "cmbStatutReparation")
    With cmbStatutReparation
        .Left = 1320
        .Top = 360
        .Width = 2055
        .Height = 315
        .Text = "Diagnostic"
        .Visible = True
    End With
    
    Set lblDiagnostic = Frame2.Controls.Add("VB.Label", "lblDiagnostic")
    With lblDiagnostic
        .Left = 240
        .Top = 780
        .Width = 975
        .Height = 255
        .Caption = "Diagnostic:"
        .Visible = True
    End With
    
    Set txtDiagnostic = Frame2.Controls.Add("VB.TextBox", "txtDiagnostic")
    With txtDiagnostic
        .Left = 1320
        .Top = 720
        .Width = 4935
        .Height = 855
        .MultiLine = True
        .ScrollBars = 2 ' Vertical
        .Visible = True
    End With
    
    Set lblPriorite = Frame2.Controls.Add("VB.Label", "lblPriorite")
    With lblPriorite
        .Left = 3600
        .Top = 1740
        .Width = 615
        .Height = 255
        .Caption = "Priorité:"
        .Visible = True
    End With
    
    Set cmbPriorite = Frame2.Controls.Add("VB.ComboBox", "cmbPriorite")
    With cmbPriorite
        .Left = 4320
        .Top = 1680
        .Width = 1575
        .Height = 315
        .Text = "Normale"
        .Visible = True
    End With
    
    Set lblTechnicien = Frame2.Controls.Add("VB.Label", "lblTechnicien")
    With lblTechnicien
        .Left = 240
        .Top = 1740
        .Width = 975
        .Height = 255
        .Caption = "Technicien:"
        .Visible = True
    End With
    
    Set cmbTechnicien = Frame2.Controls.Add("VB.ComboBox", "cmbTechnicien")
    With cmbTechnicien
        .Left = 1320
        .Top = 1680
        .Width = 2055
        .Height = 315
        .Text = "Martin L."
        .Visible = True
    End With
    
    ' === BOUTONS ===
    Set cmdValider = Me.Controls.Add("VB.CommandButton", "cmdValider")
    With cmdValider
        .Left = 4080
        .Top = 5520
        .Width = 1215
        .Height = 375
        .Caption = "Valider"
        .Visible = True
    End With
    
    Set cmdAnnuler = Me.Controls.Add("VB.CommandButton", "cmdAnnuler")
    With cmdAnnuler
        .Left = 5400
        .Top = 5520
        .Width = 1215
        .Height = 375
        .Caption = "Annuler"
        .Visible = True
    End With
    
    ' === LABEL INFO ===
    Set lblInfo = Me.Controls.Add("VB.Label", "lblInfo")
    With lblInfo
        .Left = 120
        .Top = 5160
        .Width = 6495
        .Height = 255
        .Caption = "Saisissez les informations de l'équipement à envoyer en réparation."
        .Visible = True
    End With
End Sub

Private Sub InitialiserFormulaire()
    ' Date par défaut
    txtDateEntree.Text = Format(Date, "dd/mm/yyyy")
    
    ' Remplir les combobox
    ' Types d'équipements
    cmbTypeRep.AddItem "Frigo"
    cmbTypeRep.AddItem "Distributeur"
    cmbTypeRep.AddItem "Présentoir"
    
    ' Statuts de réparation
    cmbStatutReparation.AddItem "Diagnostic"
    cmbStatutReparation.AddItem "Attente Pièces"
    cmbStatutReparation.AddItem "Réparable"
    cmbStatutReparation.AddItem "Donneur Pièces"
    cmbStatutReparation.AddItem "Atelier"
    cmbStatutReparation.AddItem "Stock Prêt"
    
    ' Priorités
    cmbPriorite.AddItem "Haute"
    cmbPriorite.AddItem "Normale"
    cmbPriorite.AddItem "Basse"
    
    ' Techniciens
    cmbTechnicien.AddItem "Martin L."
    cmbTechnicien.AddItem "Sophie M."
    cmbTechnicien.AddItem "Jean-Paul D."
    cmbTechnicien.AddItem "Marie C."
    cmbTechnicien.AddItem "Pierre R."
    
    ' Remplir les modèles selon le type par défaut
    RemplirModeles
End Sub

Private Sub RemplirModeles()
    ' Modèles selon le type sélectionné
    cmbModeleRep.Clear
    
    Select Case cmbTypeRep.Text
        Case "Frigo"
            cmbModeleRep.AddItem "RB-2024-001"
            cmbModeleRep.AddItem "RB-2024-002"
            cmbModeleRep.AddItem "RB-2024-003"
            cmbModeleRep.AddItem "RB-2023-089"
            cmbModeleRep.AddItem "RB-2023-045"
            cmbModeleRep.AddItem "RB-2022-156"
        Case "Distributeur"
            cmbModeleRep.AddItem "RB-DIST-042"
            cmbModeleRep.AddItem "RB-DIST-028"
            cmbModeleRep.AddItem "RB-DIST-033"
            cmbModeleRep.AddItem "RB-DIST-051"
        Case "Présentoir"
            cmbModeleRep.AddItem "RB-PRES-15"
            cmbModeleRep.AddItem "RB-PRES-33"
            cmbModeleRep.AddItem "RB-PRES-27"
            cmbModeleRep.AddItem "RB-PRES-41"
        Case Else
            ' Tous les modèles
            cmbModeleRep.AddItem "RB-2024-001"
            cmbModeleRep.AddItem "RB-2024-002"
            cmbModeleRep.AddItem "RB-DIST-042"
            cmbModeleRep.AddItem "RB-PRES-15"
    End Select
End Sub

' === ÉVÉNEMENTS ===
Private Sub cmbTypeRep_Click()
    ' Filtrer les modèles selon le type sélectionné
    RemplirModeles
End Sub

Private Sub cmdNouveauModele_Click()
    Dim nouveauModele As String
    nouveauModele = InputBox("Saisissez le nouveau modèle:", "Nouveau Modèle")
    
    If Trim(nouveauModele) <> "" Then
        cmbModeleRep.AddItem nouveauModele
        cmbModeleRep.Text = nouveauModele
    End If
End Sub

Private Sub cmdValider_Click()
    ' Validation des champs obligatoires
    If Trim(txtReference.Text) = "" Then
        MsgBox "Veuillez saisir une référence.", vbExclamation
        txtReference.SetFocus
        Exit Sub
    End If
    
    If cmbTypeRep.Text = "" Then
        MsgBox "Veuillez sélectionner un type.", vbExclamation
        cmbTypeRep.SetFocus
        Exit Sub
    End If
    
    If cmbModeleRep.Text = "" Then
        MsgBox "Veuillez sélectionner ou saisir un modèle.", vbExclamation
        cmbModeleRep.SetFocus
        Exit Sub
    End If
    
    If Trim(txtProbleme.Text) = "" Then
        MsgBox "Veuillez décrire le problème.", vbExclamation
        txtProbleme.SetFocus
        Exit Sub
    End If
    
    If cmbTechnicien.Text = "" Then
        MsgBox "Veuillez assigner un technicien.", vbExclamation
        cmbTechnicien.SetFocus
        Exit Sub
    End If
    
    MsgBox "Demande de réparation créée avec succès!" & vbCrLf & _
           "Référence: " & txtReference.Text & vbCrLf & _
           "Type: " & cmbTypeRep.Text & vbCrLf & _
           "Modèle: " & cmbModeleRep.Text & vbCrLf & _
           "Technicien assigné: " & cmbTechnicien.Text, vbInformation
    
    ' Fermer le formulaire
    Unload Me
End Sub

Private Sub cmdAnnuler_Click()
    ' Confirmer l'annulation
    If MsgBox("Êtes-vous sûr de vouloir annuler?", vbQuestion + vbYesNo) = vbYes Then
        Unload Me
    End If
End Sub

Private Sub txtReference_KeyPress(KeyAscii As Integer)
    ' Permettre seulement les caractères alphanumériques et tirets
    If KeyAscii < 32 Then Exit Sub ' Touches de contrôle
    
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or _
            (KeyAscii >= 65 And KeyAscii <= 90) Or _
            (KeyAscii >= 97 And KeyAscii <= 122) Or _
            KeyAscii = 45) Then
        KeyAscii = 0
    End If
End Sub

