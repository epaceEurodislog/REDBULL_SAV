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
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReparation
   BorderStyle = 3        'Fixed Dialog
   Caption = "Demande de Réparation"
   ClientHeight = 6030
   ClientLeft = 45
   ClientTop = 435
   ClientWidth = 6735
   ControlBox = 0          'False
   LinkTopic = "Form1"
   MaxButton = 0           'False
   MinButton = 0           'False
   ScaleHeight = 6030
   ScaleWidth = 6735
   ShowInTaskbar = 0       'False
   StartUpPosition = 1    'CenterOwner
   Begin VB.Frame Frame2
      Caption = "Informations Techniques"
      Height = 2175
      Left = 120
      TabIndex = 12
      Top = 2760
      Width = 6495
      Begin VB.ComboBox cmbTechnicien
         Height = 315
         Left = 1320
         TabIndex = 18
         Text = "Martin L."
         Top = 1680
         Width = 2055
      End
      Begin VB.ComboBox cmbPriorite
         Height = 315
         ItemData        =   "frmReparation.frx":0000
         Left = 4320
         List            =   "frmReparation.frx":000D
         TabIndex = 16
         Text = "Normale"
         Top = 1680
         Width = 1575
      End
      Begin VB.TextBox txtDiagnostic
         Height = 855
         Left = 1320
         MultiLine = -1          'True
         ScrollBars = 2         'Vertical
         TabIndex = 14
         Top = 720
         Width = 4935
      End
      Begin VB.ComboBox cmbStatutReparation
         Height = 315
         ItemData        =   "frmReparation.frx":0027
         Left = 1320
         List            =   "frmReparation.frx":003D
         TabIndex = 13
         Text = "Diagnostic"
         Top = 360
         Width = 2055
      End
      Begin VB.Label Label9
         Caption = "Technicien:"
         Height = 255
         Left = 240
         TabIndex = 19
         Top = 1740
         Width = 975
      End
      Begin VB.Label Label8
         Caption = "Priorité:"
         Height = 255
         Left = 3600
         TabIndex = 17
         Top = 1740
         Width = 615
      End
      Begin VB.Label Label7
         Caption = "Diagnostic:"
         Height = 255
         Left = 240
         TabIndex = 15
         Top = 780
         Width = 975
      End
      Begin VB.Label Label6
         Caption = "Statut:"
         Height = 255
         Left = 240
         TabIndex = 20
         Top = 420
         Width = 495
      End
   End
   Begin VB.Frame Frame1
      Caption = "Informations Équipement"
      Height = 2535
      Left = 120
      TabIndex = 3
      Top = 120
      Width = 6495
      Begin MSComCtl2.DTPicker dtpDateEntree
         Height = 315
         Left = 1320
         TabIndex = 11
         Top = 2040
         Width = 2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Format = 133234689
         CurrentDate = 45529
      End
      Begin VB.TextBox txtProbleme
         Height = 615
         Left = 1320
         MultiLine = -1          'True
         ScrollBars = 2         'Vertical
         TabIndex = 9
         Top = 1320
         Width = 4935
      End
      Begin VB.ComboBox cmbModeleRep
         Height = 315
         Left = 1320
         TabIndex = 7
         Top = 960
         Width = 2055
      End
      Begin VB.ComboBox cmbTypeRep
         Height = 315
         ItemData        =   "frmReparation.frx":0088
         Left = 1320
         List            =   "frmReparation.frx":0098
         TabIndex = 5
         Text = "Frigo"
         Top = 600
         Width = 2055
      End
      Begin VB.TextBox txtReference
         Height = 315
         Left = 1320
         TabIndex = 4
         Top = 240
         Width = 2055
      End
      Begin VB.Label Label5
         Caption = "Date entrée:"
         Height = 255
         Left = 240
         TabIndex = 10
         Top = 2100
         Width = 975
      End
      Begin VB.Label Label4
         Caption = "Problème:"
         Height = 255
         Left = 240
         TabIndex = 8
         Top = 1380
         Width = 735
      End
      Begin VB.Label Label3
         Caption = "Modèle:"
         Height = 255
         Left = 240
         TabIndex = 23
         Top = 1020
         Width = 615
      End
      Begin VB.Label Label2
         Caption = "Type:"
         Height = 255
         Left = 240
         TabIndex = 22
         Top = 660
         Width = 375
      End
      Begin VB.Label Label1
         Caption = "Référence:"
         Height = 255
         Left = 240
         TabIndex = 21
         Top = 300
         Width = 855
      End
   End
   Begin VB.CommandButton cmdAnnuler
      Caption = "Annuler"
      Height = 375
      Left = 5400
      TabIndex = 2
      Top = 5520
      Width = 1215
   End
   Begin VB.CommandButton cmdValider
      Caption = "Valider"
      Height = 375
      Left = 4080
      TabIndex = 1
      Top = 5520
      Width = 1215
   End
   Begin VB.CommandButton cmdNouveauModele
      Caption = "Nouveau"
      Height = 315
      Left = 3480
      TabIndex = 6
      Top = 960
      Width = 855
   End
   Begin VB.Label lblInfo
      Caption = "Saisissez les informations de l'équipement à envoyer en réparation."
      Height = 255
      Left = 120
      TabIndex = 0
      Top = 5160
      Width = 6495
   End
End
Attribute VB_Name = "frmReparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    ' Initialiser les données
    InitialiserFormulaire
    
    ' Remplir les listes déroulantes
    RemplirModeles
    RemplirTechniciens
    
    ' Date par défaut
    dtpDateEntree.Value = Date
End Sub

Private Sub InitialiserFormulaire()
    ' Types d'équipements déjà dans le designer
    
    ' Statuts de réparation déjà dans le designer
    
    ' Priorités déjà dans le designer
End Sub

Private Sub RemplirModeles()
    ' Modèles de frigos disponibles
    cmbModeleRep.AddItem "RB-2024-001"
    cmbModeleRep.AddItem "RB-2024-002"
    cmbModeleRep.AddItem "RB-2024-003"
    cmbModeleRep.AddItem "RB-2023-089"
    cmbModeleRep.AddItem "RB-2023-045"
    cmbModeleRep.AddItem "RB-2022-156"
    cmbModeleRep.AddItem "RB-DIST-042"
    cmbModeleRep.AddItem "RB-DIST-028"
    cmbModeleRep.AddItem "RB-PRES-15"
    cmbModeleRep.AddItem "RB-PRES-33"
End Sub

Private Sub RemplirTechniciens()
    cmbTechnicien.AddItem "Martin L."
    cmbTechnicien.AddItem "Sophie M."
    cmbTechnicien.AddItem "Jean-Paul D."
    cmbTechnicien.AddItem "Marie C."
    cmbTechnicien.AddItem "Pierre R."
End Sub

Private Sub cmbTypeRep_Click()
    ' Filtrer les modèles selon le type sélectionné
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
    End Select
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
    
    ' Créer un nouvel équipement en réparation
    Dim NouvelleReparation As Equipement
    
    With NouvelleReparation
        .TypeEq = cmbTypeRep.Text
        .Modele = cmbModeleRep.Text
        .statut = cmbStatutReparation.Text
        .DateOperation = dtpDateEntree.Value
        .Destination = "Service Réparation"
        .Remarques = txtProbleme.Text
        If Trim(txtDiagnostic.Text) <> "" Then
            .Remarques = .Remarques & " - Diagnostic: " & txtDiagnostic.Text
        End If
        .Technicien = cmbTechnicien.Text
        .Priorite = cmbPriorite.Text
    End With
    
    ' Ajouter à la liste principale via Form1
    Form1.AjouterEquipement NouvelleReparation
    
    MsgBox "Demande de réparation créée avec succès!" & vbCrLf & _
           "Référence: " & txtReference.Text & vbCrLf & _
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
    
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or _  ' 0-9
            (KeyAscii >= 65 And KeyAscii <= 90) Or _  ' A-Z
            (KeyAscii >= 97 And KeyAscii <= 122) Or _ ' a-z
            KeyAscii = 45) Then                      ' -
        KeyAscii = 0
    End If
End Sub

