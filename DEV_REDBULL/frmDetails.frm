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
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDetails
   BorderStyle = 3        'Fixed Dialog
   Caption = "D�tails de l'�quipement"
   ClientHeight = 7695
   ClientLeft = 45
   ClientTop = 435
   ClientWidth = 8175
   ControlBox = 0          'False
   LinkTopic = "Form1"
   MaxButton = 0           'False
   MinButton = 0           'False
   ScaleHeight = 7695
   ScaleWidth = 8175
   ShowInTaskbar = 0       'False
   StartUpPosition = 1    'CenterOwner
   Begin VB.Frame Frame3
      Caption = "Historique et Suivi"
      Height = 2055
      Left = 120
      TabIndex = 21
      Top = 5280
      Width = 7935
      Begin VB.TextBox txtHistorique
         Height = 1575
         Left = 120
         Locked = -1             'True
         MultiLine = -1          'True
         ScrollBars = 2         'Vertical
         TabIndex = 22
         Text            =   "frmDetails.frx":0000
         Top = 360
         Width = 7695
      End
      Begin VB.Label Label14
         Caption = "Historique des op�rations:"
         Height = 255
         Left = 120
         TabIndex = 23
         Top = 240
         Width = 1935
      End
   End
   Begin VB.Frame Frame2
      Caption = "Informations Techniques"
      Height = 2175
      Left = 4080
      TabIndex = 11
      Top = 3000
      Width = 3975
      Begin VB.TextBox txtTechnicien
         Height = 315
         Left = 1200
         TabIndex = 19
         Top = 1680
         Width = 2655
      End
      Begin VB.ComboBox cmbPrioriteDetails
         Height = 315
         ItemData        =   "frmDetails.frx":0128
         Left = 1200
         List            =   "frmDetails.frx":0135
         TabIndex = 17
         Top = 1320
         Width = 1455
      End
      Begin VB.TextBox txtDiagnosticDetails
         Height = 855
         Left = 1200
         MultiLine = -1          'True
         ScrollBars = 2         'Vertical
         TabIndex = 14
         Top = 360
         Width = 2655
      End
      Begin VB.Label Label12
         Caption = "Technicien:"
         Height = 255
         Left = 120
         TabIndex = 20
         Top = 1740
         Width = 975
      End
      Begin VB.Label Label11
         Caption = "Priorit�:"
         Height = 255
         Left = 120
         TabIndex = 18
         Top = 1380
         Width = 615
      End
      Begin VB.Label Label10
         Caption = "Diagnostic:"
         Height = 255
         Left = 120
         TabIndex = 15
         Top = 420
         Width = 975
      End
   End
   Begin VB.Frame Frame1
      Caption = "Informations G�n�rales"
      Height = 4935
      Left = 120
      TabIndex = 2
      Top = 120
      Width = 3855
      Begin MSComCtl2.DTPicker dtpDateDetails
         Height = 315
         Left = 1200
         TabIndex = 12
         Top = 3240
         Width = 2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Format = 133234689
         CurrentDate = 45529
      End
      Begin VB.TextBox txtRemarquesDetails
         Height = 1215
         Left = 1200
         MultiLine = -1          'True
         ScrollBars = 2         'Vertical
         TabIndex = 10
         Top = 3720
         Width = 2535
      End
      Begin VB.TextBox txtDestinationDetails
         Height = 315
         Left = 1200
         TabIndex = 8
         Top = 2880
         Width = 2535
      End
      Begin VB.ComboBox cmbStatutDetails
         Height = 315
         ItemData        =   "frmDetails.frx":014F
         Left = 1200
         List            =   "frmDetails.frx":0171
         TabIndex = 6
         Top = 2520
         Width = 2055
      End
      Begin VB.TextBox txtModeleDetails
         Height = 315
         Left = 1200
         TabIndex = 5
         Top = 2160
         Width = 2535
      End
      Begin VB.TextBox txtTypeDetails
         Height = 315
         Left = 1200
         TabIndex = 4
         Top = 1800
         Width = 2535
      End
      Begin VB.TextBox txtIDDetails
         Height = 315
         Left = 1200
         Locked = -1             'True
         TabIndex = 3
         Top = 1440
         Width = 1215
      End
      Begin VB.Label Label9
         Caption = "Remarques:"
         Height = 255
         Left = 120
         TabIndex = 24
         Top = 3780
         Width = 975
      End
      Begin VB.Label Label8
         Caption = "Date:"
         Height = 255
         Left = 120
         TabIndex = 25
         Top = 3300
         Width = 495
      End
      Begin VB.Label Label7
         Caption = "Destination:"
         Height = 255
         Left = 120
         TabIndex = 26
         Top = 2940
         Width = 975
      End
      Begin VB.Label Label6
         Caption = "Statut:"
         Height = 255
         Left = 120
         TabIndex = 27
         Top = 2580
         Width = 495
      End
      Begin VB.Label Label5
         Caption = "Mod�le:"
         Height = 255
         Left = 120
         TabIndex = 28
         Top = 2220
         Width = 615
      End
      Begin VB.Label Label4
         Caption = "Type:"
         Height = 255
         Left = 120
         TabIndex = 29
         Top = 1860
         Width = 375
      End
      Begin VB.Label Label3
         Caption = "ID:"
         Height = 255
         Left = 120
         TabIndex = 30
         Top = 1500
         Width = 255
      End
      Begin VB.Label lblTitreEquipement
         Alignment = 2          'Center
         Caption = "�quipement #XXX"
         BeginProperty Font
            Name = "MS Sans Serif"
            Size = 12
            Charset = 0
            Weight = 700
            Underline = 0           'False
            Italic = 0              'False
            Strikethrough = 0       'False
         EndProperty
         Height = 375
         Left = 120
         TabIndex = 13
         Top = 360
         Width = 3615
      End
      Begin VB.Label lblStatutIndicateur
         Alignment = 2          'Center
         BackColor = &H80FF80
         BorderStyle = 1        'Fixed Single
         Caption = "EN STOCK"
         BeginProperty Font
            Name = "MS Sans Serif"
            Size = 9.75
            Charset = 0
            Weight = 700
            Underline = 0           'False
            Italic = 0              'False
            Strikethrough = 0       'False
         EndProperty
         ForeColor = &HFFFFFF
         Height = 375
         Left = 120
         TabIndex = 16
         Top = 840
         Width = 3615
      End
   End
   Begin VB.CommandButton cmdFermer
      Caption = "Fermer"
      Height = 375
      Left = 6960
      TabIndex = 1
      Top = 120
      Width = 1095
   End
   Begin VB.CommandButton cmdSauvegarder
      Caption = "Sauvegarder"
      Height = 375
      Left = 5760
      TabIndex = 0
      Top = 120
      Width = 1095
   End
End
Attribute VB_Name = "frmDetails"
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
End Sub

Private Sub ChargerEquipement()
    ' R�cup�rer les donn�es depuis Form1
    EquipementCourant = Form1.GetEquipement(EquipementID)
    
    ' Remplir les champs
    With EquipementCourant
        lblTitreEquipement.Caption = "�quipement #" & .ID
        txtIDDetails.Text = CStr(.ID)
        txtTypeDetails.Text = .TypeEq
        txtModeleDetails.Text = .Modele
        cmbStatutDetails.Text = .statut
        dtpDateDetails.Value = .DateOperation
        txtDestinationDetails.Text = .Destination
        txtRemarquesDetails.Text = .Remarques
        
        ' Informations techniques (si disponibles)
        If .Technicien <> "" Then
            txtTechnicien.Text = .Technicien
        End If
        
        If .Priorite <> "" Then
            cmbPrioriteDetails.Text = .Priorite
        End If
        
        ' Diagnostic dans les remarques techniques
        txtDiagnosticDetails.Text = ExtraireTexteApres(.Remarques, "Diagnostic: ")
        
        ' Mettre � jour l'indicateur de statut
        MettreAJourIndicateurStatut .statut
        
        ' G�n�rer l'historique
        GenererHistorique
    End With
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

Private Sub GenererHistorique()
    Dim historique As String
    
    historique = "=== HISTORIQUE DE L'�QUIPEMENT ===" & vbCrLf & vbCrLf
    
    ' Informations de base
    historique = historique & Format(EquipementCourant.DateOperation, "dd/mm/yyyy") & " - "
    
    Select Case EquipementCourant.statut
        Case "R�ception"
            historique = historique & "R�ception de l'�quipement" & vbCrLf
            historique = historique & "� Destination pr�vue: " & EquipementCourant.Destination & vbCrLf
        Case "Stock"
            historique = historique & "Mise en stock" & vbCrLf
            historique = historique & "� �quipement v�rifi� et stock�" & vbCrLf
        Case "Pr�paration"
            historique = historique & "Pr�paration pour exp�dition" & vbCrLf
            historique = historique & "� Destination: " & EquipementCourant.Destination & vbCrLf
        Case "Exp�dition"
            historique = historique & "Exp�dition en cours" & vbCrLf
            historique = historique & "� Livraison vers: " & EquipementCourant.Destination & vbCrLf
        Case "Diagnostic"
            historique = historique & "Diagnostic en cours" & vbCrLf
            historique = historique & "� Technicien: " & EquipementCourant.Technicien & vbCrLf
            historique = historique & "� Priorit�: " & EquipementCourant.Priorite & vbCrLf
        Case "Attente Pi�ces"
            historique = historique & "En attente de pi�ces d�tach�es" & vbCrLf
            historique = historique & "� Technicien: " & EquipementCourant.Technicien & vbCrLf
        Case "R�parable"
            historique = historique & "�quipement r�parable identifi�" & vbCrLf
            historique = historique & "� Technicien: " & EquipementCourant.Technicien & vbCrLf
        Case "Donneur Pi�ces"
            historique = historique & "D�sign� comme donneur de pi�ces" & vbCrLf
            historique = historique & "� D�montage en cours" & vbCrLf
        Case "Atelier"
            historique = historique & "R�paration en atelier" & vbCrLf
            historique = historique & "� Technicien: " & EquipementCourant.Technicien & vbCrLf
        Case "Stock Pr�t"
            historique = historique & "R�paration termin�e - Pr�t � exp�dier" & vbCrLf
            historique = historique & "� Destination: " & EquipementCourant.Destination & vbCrLf
    End Select
    
    historique = historique & vbCrLf
    
    ' Remarques
    If Trim(EquipementCourant.Remarques) <> "" Then
        historique = historique & "REMARQUES:" & vbCrLf
        historique = historique & EquipementCourant.Remarques & vbCrLf & vbCrLf
    End If
    
    ' Informations de suivi simul�es
    historique = historique & "=== �TAPES PR�C�DENTES ===" & vbCrLf
    historique = historique & Format(DateAdd("d", -5, EquipementCourant.DateOperation), "dd/mm/yyyy") & " - Cr�ation de la fiche �quipement" & vbCrLf
    historique = historique & Format(DateAdd("d", -3, EquipementCourant.DateOperation), "dd/mm/yyyy") & " - Contr�le qualit� effectu�" & vbCrLf
    historique = historique & Format(DateAdd("d", -1, EquipementCourant.DateOperation), "dd/mm/yyyy") & " - Mise � jour du statut" & vbCrLf
    
    txtHistorique.Text = historique
End Sub

Private Function ExtraireTexteApres(texte As String, motif As String) As String
    Dim position As Integer
    position = InStr(texte, motif)
    
    If position > 0 Then
        ExtraireTexteApres = Mid(texte, position + Len(motif))
        ' Nettoyer jusqu'au premier retour � la ligne ou fin
        position = InStr(ExtraireTexteApres, vbCrLf)
        If position > 0 Then
            ExtraireTexteApres = Left(ExtraireTexteApres, position - 1)
        End If
    Else
        ExtraireTexteApres = ""
    End If
End Function

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
    
    ' Mettre � jour l'�quipement
    With EquipementCourant
        .TypeEq = txtTypeDetails.Text
        .Modele = txtModeleDetails.Text
        .statut = cmbStatutDetails.Text
        .DateOperation = dtpDateDetails.Value
        .Destination = txtDestinationDetails.Text
        .Remarques = txtRemarquesDetails.Text
        
        ' Ajouter le diagnostic aux remarques si pr�sent
        If Trim(txtDiagnosticDetails.Text) <> "" Then
            .Remarques = .Remarques & " - Diagnostic: " & txtDiagnosticDetails.Text
        End If
        
        .Technicien = txtTechnicien.Text
        .Priorite = cmbPrioriteDetails.Text
    End With
    
    ' Sauvegarder via Form1
    Form1.ModifierEquipement EquipementID, EquipementCourant
    
    ' Mettre � jour l'affichage
    MettreAJourIndicateurStatut EquipementCourant.statut
    GenererHistorique
    
    MsgBox "�quipement mis � jour avec succ�s!", vbInformation
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
