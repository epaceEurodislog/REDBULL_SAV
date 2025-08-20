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
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStatut
   BorderStyle = 3        'Fixed Dialog
   Caption = "Changer le Statut"
   ClientHeight = 5535
   ClientLeft = 45
   ClientTop = 435
   ClientWidth = 6015
   ControlBox = 0          'False
   LinkTopic = "Form1"
   MaxButton = 0           'False
   MinButton = 0           'False
   ScaleHeight = 5535
   ScaleWidth = 6015
   ShowInTaskbar = 0       'False
   StartUpPosition = 1    'CenterOwner
   Begin VB.Frame Frame2
      Caption = "Informations Additionnelles"
      Height = 2175
      Left = 120
      TabIndex = 8
      Top = 2760
      Width = 5775
      Begin VB.ComboBox cmbTechnicienStatut
         Height = 315
         Left = 1320
         TabIndex = 14
         Top = 1680
         Width = 2055
      End
      Begin VB.ComboBox cmbPrioriteStatut
         Height = 315
         ItemData        =   "frmStatut.frx":0000
         Left = 3600
         List            =   "frmStatut.frx":000D
         TabIndex = 12
         Top = 1680
         Width = 1575
      End
      Begin VB.TextBox txtNotesStatut
         Height = 1095
         Left = 1320
         MultiLine = -1          'True
         ScrollBars = 2         'Vertical
         TabIndex = 10
         Top = 480
         Width = 4335
      End
      Begin MSComCtl2.DTPicker dtpDateStatut
         Height = 315
         Left = 1320
         TabIndex = 13
         Top = 240
         Width = 2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Format = 133234689
         CurrentDate = 45529
      End
      Begin VB.Label Label6
         Caption = "Technicien:"
         Height = 255
         Left = 240
         TabIndex = 15
         Top = 1740
         Width = 975
      End
      Begin VB.Label Label5
         Caption = "Priorité:"
         Height = 255
         Left = 3600
         TabIndex = 16
         Top = 1440
         Width = 615
      End
      Begin VB.Label Label4
         Caption = "Notes:"
         Height = 255
         Left = 240
         TabIndex = 11
         Top = 540
         Width = 495
      End
      Begin VB.Label Label3
         Caption = "Date:"
         Height = 255
         Left = 240
         TabIndex = 9
         Top = 300
         Width = 375
      End
   End
   Begin VB.Frame Frame1
      Caption = "Informations Actuelles"
      Height = 2535
      Left = 120
      TabIndex = 2
      Top = 120
      Width = 5775
      Begin VB.ComboBox cmbNouveauStatut
         Height = 315
         ItemData        =   "frmStatut.frx":0027
         Left = 1680
         List            =   "frmStatut.frx":0052
         TabIndex = 6
         Top = 1800
         Width = 2535
      End
      Begin VB.Label lblStatutActuel
         Caption = "Stock"
         BeginProperty Font
            Name = "MS Sans Serif"
            Size = 9.75
            Charset = 0
            Weight = 700
            Underline = 0           'False
            Italic = 0              'False
            Strikethrough = 0       'False
         EndProperty
         ForeColor = &HFF0000
         Height = 255
         Left = 1680
         TabIndex = 7
         Top = 1440
         Width = 2535
      End
      Begin VB.Label lblModeleStatut
         Caption = "RB-2024-001"
         Height = 255
         Left = 1680
         TabIndex = 5
         Top = 1080
         Width = 2535
      End
      Begin VB.Label lblTypeStatut
         Caption = "Frigo"
         Height = 255
         Left = 1680
         TabIndex = 4
         Top = 720
         Width = 2535
      End
      Begin VB.Label lblIDStatut
         Caption = "1"
         Height = 255
         Left = 1680
         TabIndex = 3
         Top = 360
         Width = 1215
      End
      Begin VB.Label Label9
         Caption = "Nouveau statut:"
         Height = 255
         Left = 240
         TabIndex = 17
         Top = 1860
         Width = 1335
      End
      Begin VB.Label Label8
         Caption = "Statut actuel:"
         Height = 255
         Left = 240
         TabIndex = 18
         Top = 1500
         Width = 1095
      End
      Begin VB.Label Label7
         Caption = "Modèle:"
         Height = 255
         Left = 240
         TabIndex = 19
         Top = 1140
         Width = 615
      End
      Begin VB.Label Label2
         Caption = "Type:"
         Height = 255
         Left = 240
         TabIndex = 20
         Top = 780
         Width = 375
      End
      Begin VB.Label Label1
         Caption = "ID Équipement:"
         Height = 255
         Left = 240
         TabIndex = 21
         Top = 420
         Width = 1215
      End
   End
   Begin VB.CommandButton cmdAnnulerStatut
      Caption = "Annuler"
      Height = 375
      Left = 4680
      TabIndex = 1
      Top = 5040
      Width = 1215
   End
   Begin VB.CommandButton cmdConfirmerStatut
      Caption = "Confirmer"
      Height = 375
      Left = 3360
      TabIndex = 0
      Top = 5040
      Width = 1215
   End
End
Attribute VB_Name = "frmStatut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public EquipementID As Long
Private EquipementCourant As Equipement

Private Sub Form_Load()
    ' Charger les données de l'équipement
    If EquipementID > 0 Then
        ChargerEquipement
    End If
    
    ' Initialiser les listes déroulantes
    InitialiserListes
    
    ' Date par défaut
    dtpDateStatut.Value = Date
End Sub

Private Sub ChargerEquipement()
    ' Récupérer les données depuis Form1
    EquipementCourant = Form1.GetEquipement(EquipementID)
    
    ' Afficher les informations actuelles
    With EquipementCourant
        lblIDStatut.Caption = CStr(.ID)
        lblTypeStatut.Caption = .TypeEq
        lblModeleStatut.Caption = .Modele
        lblStatutActuel.Caption = .statut
        
        ' Sélectionner le statut actuel dans la liste
        cmbNouveauStatut.Text = .statut
        
        ' Pré-remplir les champs si des données existent
        If .Technicien <> "" Then
            cmbTechnicienStatut.Text = .Technicien
        End If
        
        If .Priorite <> "" Then
            cmbPrioriteStatut.Text = .Priorite
        End If
        
        txtNotesStatut.Text = .Remarques
    End With
End Sub

Private Sub InitialiserListes()
    ' Techniciens
    cmbTechnicienStatut.AddItem "Martin L."
    cmbTechnicienStatut.AddItem "Sophie M."
    cmbTechnicienStatut.AddItem "Jean-Paul D."
    cmbTechnicienStatut.AddItem "Marie C."
    cmbTechnicienStatut.AddItem "Pierre R."
End Sub

Private Sub cmbNouveauStatut_Click()
    ' Adapter les champs selon le nouveau statut sélectionné
    Select Case cmbNouveauStatut.Text
        Case "Diagnostic", "Attente Pièces", "Réparable", "Donneur Pièces", "Atelier", "Stock Prêt"
            ' Statuts de réparation - activer les champs techniques
            cmbTechnicienStatut.Enabled = True
            cmbPrioriteStatut.Enabled = True
            
            ' Valeurs par défaut
            If cmbTechnicienStatut.Text = "" Then
                cmbTechnicienStatut.ListIndex = 0
            End If
            If cmbPrioriteStatut.Text = "" Then
                cmbPrioriteStatut.Text = "Normale"
            End If
            
        Case Else
            ' Statuts normaux - désactiver les champs techniques
            cmbTechnicienStatut.Enabled = False
            cmbPrioriteStatut.Enabled = False
    End Select
    
    ' Suggestions de notes selon le statut
    Select Case cmbNouveauStatut.Text
        Case "Réception"
            If txtNotesStatut.Text = "" Then
                txtNotesStatut.Text = "Équipement reçu et vérifié"
            End If
        Case "Stock"
            If txtNotesStatut.Text = "" Then
                txtNotesStatut.Text = "Équipement contrôlé et mis en stock"
            End If
        Case "Préparation"
            If txtNotesStatut.Text = "" Then
                txtNotesStatut.Text = "Préparation pour expédition"
            End If
        Case "Expédition"
            If txtNotesStatut.Text = "" Then
                txtNotesStatut.Text = "Expédition en cours"
            End If
        Case "Diagnostic"
            If txtNotesStatut.Text = "" Then
                txtNotesStatut.Text = "Diagnostic technique en cours"
            End If
        Case "Attente Pièces"
            If txtNotesStatut.Text = "" Then
                txtNotesStatut.Text = "En attente de pièces détachées"
            End If
        Case "Réparable"
            If txtNotesStatut.Text = "" Then
                txtNotesStatut.Text = "Équipement identifié comme réparable"
            End If
        Case "Donneur Pièces"
            If txtNotesStatut.Text = "" Then
                txtNotesStatut.Text = "Équipement utilisé comme donneur de pièces"
            End If
        Case "Atelier"
            If txtNotesStatut.Text = "" Then
                txtNotesStatut.Text = "Réparation en cours dans l'atelier"
            End If
        Case "Stock Prêt"
            If txtNotesStatut.Text = "" Then
                txtNotesStatut.Text = "Réparation terminée - Prêt à expédier"
            End If
    End Select
End Sub

Private Sub cmdConfirmerStatut_Click()
    ' Validation
    If cmbNouveauStatut.Text = "" Then
        MsgBox "Veuillez sélectionner un nouveau statut.", vbExclamation
        cmbNouveauStatut.SetFocus
        Exit Sub
    End If
    
    ' Vérifications spécifiques selon le statut
    Select Case cmbNouveauStatut.Text
        Case "Diagnostic", "Attente Pièces", "Réparable", "Donneur Pièces", "Atelier", "Stock Prêt"
            If cmbTechnicienStatut.Text = "" Then
                MsgBox "Veuillez assigner un technicien pour les opérations de réparation.", vbExclamation
                cmbTechnicienStatut.SetFocus
                Exit Sub
            End If
    End Select
    
    ' Confirmer le changement
    Dim message As String
    message = "Confirmer le changement de statut ?" & vbCrLf & vbCrLf
    message = message & "De: " & EquipementCourant.statut & vbCrLf
    message = message & "Vers: " & cmbNouveauStatut.Text & vbCrLf & vbCrLf
    message = message & "Date: " & Format(dtpDateStatut.Value, "dd/mm/yyyy")
    
    If MsgBox(message, vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    ' Mettre à jour l'équipement
    With EquipementCourant
        .statut = cmbNouveauStatut.Text
        .DateOperation = dtpDateStatut.Value
        .Remarques = txtNotesStatut.Text
        
        ' Mettre à jour les informations techniques si nécessaire
        If cmbTechnicienStatut.Enabled And cmbTechnicienStatut.Text <> "" Then
            .Technicien = cmbTechnicienStatut.Text
        End If
        
        If cmbPrioriteStatut.Enabled And cmbPrioriteStatut.Text <> "" Then
            .Priorite = cmbPrioriteStatut.Text
        End If
        
        ' Ajuster la destination selon le statut
        Select Case .statut
            Case "Diagnostic", "Attente Pièces", "Réparable", "Donneur Pièces", "Atelier", "Stock Prêt"
                .Destination = "Service Réparation"
        End Select
    End With
    
    ' Sauvegarder via Form1
    Form1.ModifierEquipement EquipementID, EquipementCourant
    
    MsgBox "Statut mis à jour avec succès!", vbInformation
    
    ' Fermer le formulaire
    Unload Me
End Sub

Private Sub cmdAnnulerStatut_Click()
    If MsgBox("Annuler les modifications?", vbQuestion + vbYesNo) = vbYes Then
        Unload Me
    End If
End Sub
