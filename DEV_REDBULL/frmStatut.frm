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
         Caption = "Priorit�:"
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
         Caption = "Mod�le:"
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
         Caption = "ID �quipement:"
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
    ' Charger les donn�es de l'�quipement
    If EquipementID > 0 Then
        ChargerEquipement
    End If
    
    ' Initialiser les listes d�roulantes
    InitialiserListes
    
    ' Date par d�faut
    dtpDateStatut.Value = Date
End Sub

Private Sub ChargerEquipement()
    ' R�cup�rer les donn�es depuis Form1
    EquipementCourant = Form1.GetEquipement(EquipementID)
    
    ' Afficher les informations actuelles
    With EquipementCourant
        lblIDStatut.Caption = CStr(.ID)
        lblTypeStatut.Caption = .TypeEq
        lblModeleStatut.Caption = .Modele
        lblStatutActuel.Caption = .statut
        
        ' S�lectionner le statut actuel dans la liste
        cmbNouveauStatut.Text = .statut
        
        ' Pr�-remplir les champs si des donn�es existent
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
    ' Adapter les champs selon le nouveau statut s�lectionn�
    Select Case cmbNouveauStatut.Text
        Case "Diagnostic", "Attente Pi�ces", "R�parable", "Donneur Pi�ces", "Atelier", "Stock Pr�t"
            ' Statuts de r�paration - activer les champs techniques
            cmbTechnicienStatut.Enabled = True
            cmbPrioriteStatut.Enabled = True
            
            ' Valeurs par d�faut
            If cmbTechnicienStatut.Text = "" Then
                cmbTechnicienStatut.ListIndex = 0
            End If
            If cmbPrioriteStatut.Text = "" Then
                cmbPrioriteStatut.Text = "Normale"
            End If
            
        Case Else
            ' Statuts normaux - d�sactiver les champs techniques
            cmbTechnicienStatut.Enabled = False
            cmbPrioriteStatut.Enabled = False
    End Select
    
    ' Suggestions de notes selon le statut
    Select Case cmbNouveauStatut.Text
        Case "R�ception"
            If txtNotesStatut.Text = "" Then
                txtNotesStatut.Text = "�quipement re�u et v�rifi�"
            End If
        Case "Stock"
            If txtNotesStatut.Text = "" Then
                txtNotesStatut.Text = "�quipement contr�l� et mis en stock"
            End If
        Case "Pr�paration"
            If txtNotesStatut.Text = "" Then
                txtNotesStatut.Text = "Pr�paration pour exp�dition"
            End If
        Case "Exp�dition"
            If txtNotesStatut.Text = "" Then
                txtNotesStatut.Text = "Exp�dition en cours"
            End If
        Case "Diagnostic"
            If txtNotesStatut.Text = "" Then
                txtNotesStatut.Text = "Diagnostic technique en cours"
            End If
        Case "Attente Pi�ces"
            If txtNotesStatut.Text = "" Then
                txtNotesStatut.Text = "En attente de pi�ces d�tach�es"
            End If
        Case "R�parable"
            If txtNotesStatut.Text = "" Then
                txtNotesStatut.Text = "�quipement identifi� comme r�parable"
            End If
        Case "Donneur Pi�ces"
            If txtNotesStatut.Text = "" Then
                txtNotesStatut.Text = "�quipement utilis� comme donneur de pi�ces"
            End If
        Case "Atelier"
            If txtNotesStatut.Text = "" Then
                txtNotesStatut.Text = "R�paration en cours dans l'atelier"
            End If
        Case "Stock Pr�t"
            If txtNotesStatut.Text = "" Then
                txtNotesStatut.Text = "R�paration termin�e - Pr�t � exp�dier"
            End If
    End Select
End Sub

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
    message = message & "De: " & EquipementCourant.statut & vbCrLf
    message = message & "Vers: " & cmbNouveauStatut.Text & vbCrLf & vbCrLf
    message = message & "Date: " & Format(dtpDateStatut.Value, "dd/mm/yyyy")
    
    If MsgBox(message, vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    ' Mettre � jour l'�quipement
    With EquipementCourant
        .statut = cmbNouveauStatut.Text
        .DateOperation = dtpDateStatut.Value
        .Remarques = txtNotesStatut.Text
        
        ' Mettre � jour les informations techniques si n�cessaire
        If cmbTechnicienStatut.Enabled And cmbTechnicienStatut.Text <> "" Then
            .Technicien = cmbTechnicienStatut.Text
        End If
        
        If cmbPrioriteStatut.Enabled And cmbPrioriteStatut.Text <> "" Then
            .Priorite = cmbPrioriteStatut.Text
        End If
        
        ' Ajuster la destination selon le statut
        Select Case .statut
            Case "Diagnostic", "Attente Pi�ces", "R�parable", "Donneur Pi�ces", "Atelier", "Stock Pr�t"
                .Destination = "Service R�paration"
        End Select
    End With
    
    ' Sauvegarder via Form1
    Form1.ModifierEquipement EquipementID, EquipementCourant
    
    MsgBox "Statut mis � jour avec succ�s!", vbInformation
    
    ' Fermer le formulaire
    Unload Me
End Sub

Private Sub cmdAnnulerStatut_Click()
    If MsgBox("Annuler les modifications?", vbQuestion + vbYesNo) = vbYes Then
        Unload Me
    End If
End Sub
