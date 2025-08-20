VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
VERSION 5#
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrincipal
   BackColor = &HF0F0F0
   BorderStyle = 1        'Fixed Single
   Caption = "SAV Red Bull Scanner Pro - v2.1 - [frmPrincipal.frm]"
   ClientHeight = 9030
   ClientLeft = 45
   ClientTop = 435
   ClientWidth = 11415
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic = "Form1"
   MaxButton = 0           'False
   ScaleHeight = 9030
   ScaleWidth = 11415
   StartUpPosition = 2    'CenterScreen
   Begin VB.Frame frameInfoGenerales
      BackColor = &HE0E0E0
      Caption = "?? Informations générales"
      BeginProperty Font
         Name = "MS Sans Serif"
         Size = 8.25
         Charset = 0
         Weight = 700
         Underline = 0           'False
         Italic = 0              'False
         Strikethrough = 0       'False
      EndProperty
      ForeColor = &H4080&
      Height = 2535
      Left = 240
      TabIndex = 11
      Top = 1800
      Width = 10935
      Begin MSComCtl2.DTPicker dtpDate
         Height = 315
         Left = 1680
         TabIndex = 19
         Top = 1440
         Width = 2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Format = 133234689
         CurrentDate = 45529
      End
      Begin VB.TextBox txtReferenceProduit
         Height = 315
         Left = 1680
         TabIndex = 18
         Top = 1920
         Width = 8775
      End
      Begin VB.TextBox txtReception
         Height = 315
         Left = 1680
         TabIndex = 16
         Top = 960
         Width = 2415
      End
      Begin VB.TextBox txtEnlevement
         Height = 315
         Left = 1680
         TabIndex = 14
         Top = 480
         Width = 2415
      End
      Begin VB.Label Label4
         BackColor = &HE0E0E0
         Caption = "Référence produit:"
         Height = 255
         Left = 240
         TabIndex = 17
         Top = 1980
         Width = 1335
      End
      Begin VB.Label Label3
         BackColor = &HE0E0E0
         Caption = "Date:"
         Height = 255
         Left = 240
         TabIndex = 20
         Top = 1500
         Width = 375
      End
      Begin VB.Label Label2
         BackColor = &HE0E0E0
         Caption = "N° Réception:"
         Height = 255
         Left = 240
         TabIndex = 15
         Top = 1020
         Width = 1095
      End
      Begin VB.Label Label1
         BackColor = &HE0E0E0
         Caption = "N° Enlèvement:"
         Height = 255
         Left = 240
         TabIndex = 13
         Top = 540
         Width = 1215
      End
   End
   Begin VB.Frame frameDiagnostic
      BackColor = &HE0E0E0
      Caption = "?? Diagnostic technique"
      BeginProperty Font
         Name = "MS Sans Serif"
         Size = 8.25
         Charset = 0
         Weight = 700
         Underline = 0           'False
         Italic = 0              'False
         Strikethrough = 0       'False
      EndProperty
      ForeColor = &H4080&
      Height = 2415
      Left = 240
      TabIndex = 6
      Top = 6360
      Width = 10935
      Begin VB.CheckBox chkRayures
         BackColor = &HF0F0F0
         Caption = "??? RAYURES TROP IMPORTANTES"
         Height = 255
         Left = 240
         TabIndex = 10
         Top = 1920
         Value = 1              'Checked
         Width = 10455
      End
      Begin VB.CheckBox chkTechnique
         BackColor = &HF0F0F0
         Caption = "? TECHNIQUE ? LUMIÈRE // FROID // MOTEUR // VITRE BRISÉE"
         Height = 255
         Left = 240
         TabIndex = 9
         Top = 1440
         Width = 10455
      End
      Begin VB.CheckBox chkPieceManquante
         BackColor = &HF0F0F0
         Caption = "?? PIÈCE MANQUANTE // PROBLÈME CAPOT DU BAS DU FRIGO"
         Height = 255
         Left = 240
         TabIndex = 8
         Top = 960
         Width = 10455
      End
      Begin VB.Label lblDiagnosticTitle
         BackColor = &HE0E0E0
         Caption = "Sélectionnez les problèmes identifiés:"
         Height = 255
         Left = 240
         TabIndex = 7
         Top = 600
         Width = 2655
      End
   End
   Begin VB.Frame frameMotifRetour
      BackColor = &HE0E0E0
      Caption = "?? Motif du retour"
      BeginProperty Font
         Name = "MS Sans Serif"
         Size = 8.25
         Charset = 0
         Weight = 700
         Underline = 0           'False
         Italic = 0              'False
         Strikethrough = 0       'False
      EndProperty
      ForeColor = &H4080&
      Height = 2175
      Left = 240
      TabIndex = 0
      Top = 4440
      Width = 10935
      Begin VB.Frame frameCoherence
         BackColor = &HE0E0E0
         Caption = "Cohérence avec la boutique:"
         Height = 855
         Left = 5520
         TabIndex = 3
         Top = 480
         Width = 5175
         Begin VB.OptionButton optNon
            BackColor = &HF0F0F0
            Caption = "? NON"
            Height = 255
            Left = 240
            TabIndex = 5
            Top = 480
            Width = 4695
         End
         Begin VB.OptionButton optOui
            BackColor = &HC0FFC0
            Caption = "? OUI"
            Height = 255
            Left = 240
            TabIndex = 4
            Top = 240
            Value = -1              'True
            Width = 4695
         End
      End
      Begin VB.OptionButton optEsthetique
         BackColor = &HF0F0F0
         Caption = "?? ESTHÉTIQUE"
         Height = 255
         Left = 240
         TabIndex = 2
         Top = 840
         Width = 4935
      End
      Begin VB.OptionButton optMecanique
         BackColor = &H4080FF
         Caption = "?? MÉCANIQUE"
         BeginProperty Font
            Name = "MS Sans Serif"
            Size = 8.25
            Charset = 0
            Weight = 700
            Underline = 0           'False
            Italic = 0              'False
            Strikethrough = 0       'False
         EndProperty
         ForeColor = &HFFFFFF
         Height = 255
         Left = 240
         TabIndex = 1
         Top = 480
         Value = -1              'True
         Width = 4935
      End
   End
   Begin VB.CommandButton cmdHistorique
      BackColor = &HE0E0E0
      Caption = "?? Historique"
      Height = 735
      Left = 7800
      TabIndex = 24
      Top = 840
      Width = 1575
   End
   Begin VB.CommandButton cmdFormulaire
      BackColor = &H4080FF
      Caption = "?? Formulaire"
      BeginProperty Font
         Name = "MS Sans Serif"
         Size = 8.25
         Charset = 0
         Weight = 700
         Underline = 0           'False
         Italic = 0              'False
         Strikethrough = 0       'False
      EndProperty
      ForeColor = &HFFFFFF
      Height = 735
      Left = 6120
      TabIndex = 23
      Top = 840
      Width = 1575
   End
   Begin VB.CommandButton cmdScanner
      BackColor = &HE0E0E0
      Caption = "?? Scanner"
      Height = 735
      Left = 4440
      TabIndex = 22
      Top = 840
      Width = 1575
   End
   Begin VB.CommandButton cmdScannerPro
      BackColor = &H80FF80
      Caption = "Scanner Pro"
      BeginProperty Font
         Name = "MS Sans Serif"
         Size = 8.25
         Charset = 0
         Weight = 700
         Underline = 0           'False
         Italic = 0              'False
         Strikethrough = 0       'False
      EndProperty
      ForeColor = &HFFFFFF
      Height = 375
      Left = 9600
      TabIndex = 21
      Top = 120
      Width = 1575
   End
   Begin VB.Label lblTitre
      Alignment = 2          'Center
      BackColor = &H4080FF
      Caption = "?? FICHE RETOUR SAV RED BULL ??"
      BeginProperty Font
         Name = "MS Sans Serif"
         Size = 12
         Charset = 0
         Weight = 700
         Underline = 0           'False
         Italic = 0              'False
         Strikethrough = 0       'False
      EndProperty
      ForeColor = &HFFFFFF
      Height = 375
      Left = 240
      TabIndex = 12
      Top = 240
      Width = 9255
   End
   Begin VB.Label lblSousTitre
      Alignment = 2          'Center
      BackColor = &H4080FF
      Caption = "Système de Gestion des Réfrigérateurs - Interface Frigoriste"
      BeginProperty Font
         Name = "MS Sans Serif"
         Size = 8.25
         Charset = 0
         Weight = 700
         Underline = 0           'False
         Italic = 0              'False
         Strikethrough = 0       'False
      EndProperty
      ForeColor = &HFFFFFF
      Height = 255
      Left = 240
      TabIndex = 25
      Top = 600
      Width = 9255
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    ' Initialiser les valeurs par défaut
    InitialiserFormulaire
End Sub

Private Sub InitialiserFormulaire()
    ' Valeurs par défaut basées sur l'image
    txtEnlevement.Text = "69113"
    txtReception.Text = "19108"
    dtpDate.Value = DateSerial(2025, 6, 5) ' 05/06/25
    txtReferenceProduit.Text = "VCZ286 52000-1"
    
    ' Sélections par défaut
    optMecanique.Value = True
    optOui.Value = True
    chkRayures.Value = vbChecked
    
    ' Mettre le focus sur le premier champ
    txtEnlevement.SetFocus
End Sub

Private Sub cmdScanner_Click()
    MsgBox "Fonction Scanner à implémenter" & vbCrLf & _
           "Cette fonction permettra de scanner des codes-barres ou QR codes.", _
           vbInformation, "Scanner"
End Sub

Private Sub cmdFormulaire_Click()
    ' Afficher ou traiter le formulaire actuel
    If Not ValiderFormulaire() Then
        Exit Sub
    End If
    
    MsgBox "Formulaire validé avec succès!" & vbCrLf & _
           "N° Enlèvement: " & txtEnlevement.Text & vbCrLf & _
           "N° Réception: " & txtReception.Text & vbCrLf & _
           "Date: " & Format(dtpDate.Value, "dd/mm/yyyy") & vbCrLf & _
           "Référence: " & txtReferenceProduit.Text, _
           vbInformation, "Formulaire"
End Sub

Private Sub cmdHistorique_Click()
    ' Ouvrir l'historique
    MsgBox "Fonction Historique à implémenter" & vbCrLf & _
           "Cette fonction affichera l'historique des interventions.", _
           vbInformation, "Historique"
End Sub

Private Sub cmdScannerPro_Click()
    ' Fonction Scanner Pro
    MsgBox "Scanner Pro activé" & vbCrLf & _
           "Mode avancé de scan avec fonctionnalités étendues.", _
           vbInformation, "Scanner Pro"
End Sub

Private Function ValiderFormulaire() As Boolean
    ' Validation des champs obligatoires
    If Trim(txtEnlevement.Text) = "" Then
        MsgBox "Le numéro d'enlèvement est obligatoire.", vbExclamation
        txtEnlevement.SetFocus
        ValiderFormulaire = False
        Exit Function
    End If
    
    If Trim(txtReception.Text) = "" Then
        MsgBox "Le numéro de réception est obligatoire.", vbExclamation
        txtReception.SetFocus
        ValiderFormulaire = False
        Exit Function
    End If
    
    If Trim(txtReferenceProduit.Text) = "" Then
        MsgBox "La référence produit est obligatoire.", vbExclamation
        txtReferenceProduit.SetFocus
        ValiderFormulaire = False
        Exit Function
    End If
    
    ' Vérifier qu'un motif de retour est sélectionné
    If Not optMecanique.Value And Not optEsthetique.Value Then
        MsgBox "Veuillez sélectionner un motif de retour.", vbExclamation
        ValiderFormulaire = False
        Exit Function
    End If
    
    ' Vérifier la cohérence avec la boutique
    If Not optOui.Value And Not optNon.Value Then
        MsgBox "Veuillez indiquer la cohérence avec la boutique.", vbExclamation
        ValiderFormulaire = False
        Exit Function
    End If
    
    ValiderFormulaire = True
End Function

Private Sub optMecanique_Click()
    ' Mettre à jour l'apparence quand MÉCANIQUE est sélectionné
    If optMecanique.Value Then
        optMecanique.BackColor = &H4080FF
        optEsthetique.BackColor = &HF0F0F0
    End If
End Sub

Private Sub optEsthetique_Click()
    ' Mettre à jour l'apparence quand ESTHÉTIQUE est sélectionné
    If optEsthetique.Value Then
        optEsthetique.BackColor = &HFFC080
        optMecanique.BackColor = &HF0F0F0
    End If
End Sub

Private Sub optOui_Click()
    ' Mettre à jour l'apparence pour OUI
    If optOui.Value Then
        optOui.BackColor = &HC0FFC0
        optNon.BackColor = &HF0F0F0
    End If
End Sub

Private Sub optNon_Click()
    ' Mettre à jour l'apparence pour NON
    If optNon.Value Then
        optNon.BackColor = &HC0C0FF
        optOui.BackColor = &HF0F0F0
    End If
End Sub

' Fonction pour sauvegarder les données (simulation)
Public Sub SauvegarderDonnees()
    Dim rapport As String
    
    rapport = "=== FICHE RETOUR SAV RED BULL ===" & vbCrLf & vbCrLf
    rapport = rapport & "INFORMATIONS GÉNÉRALES:" & vbCrLf
    rapport = rapport & "N° Enlèvement: " & txtEnlevement.Text & vbCrLf
    rapport = rapport & "N° Réception: " & txtReception.Text & vbCrLf
    rapport = rapport & "Date: " & Format(dtpDate.Value, "dd/mm/yyyy") & vbCrLf
    rapport = rapport & "Référence produit: " & txtReferenceProduit.Text & vbCrLf & vbCrLf
    
    rapport = rapport & "MOTIF DU RETOUR:" & vbCrLf
    If optMecanique.Value Then
        rapport = rapport & "- MÉCANIQUE" & vbCrLf
    End If
    If optEsthetique.Value Then
        rapport = rapport & "- ESTHÉTIQUE" & vbCrLf
    End If
    
    rapport = rapport & vbCrLf & "COHÉRENCE AVEC LA BOUTIQUE:" & vbCrLf
    If optOui.Value Then
        rapport = rapport & "- OUI" & vbCrLf
    Else
        rapport = rapport & "- NON" & vbCrLf
    End If
    
    rapport = rapport & vbCrLf & "DIAGNOSTIC TECHNIQUE:" & vbCrLf
    If chkPieceManquante.Value = vbChecked Then
        rapport = rapport & "- PIÈCE MANQUANTE // PROBLÈME CAPOT DU BAS DU FRIGO" & vbCrLf
    End If
    If chkTechnique.Value = vbChecked Then
        rapport = rapport & "- TECHNIQUE ? LUMIÈRE // FROID // MOTEUR // VITRE BRISÉE" & vbCrLf
    End If
    If chkRayures.Value = vbChecked Then
        rapport = rapport & "- RAYURES TROP IMPORTANTES" & vbCrLf
    End If
    
    ' Simuler la sauvegarde (dans une vraie application, ceci irait en base de données)
    MsgBox rapport, vbInformation, "Données sauvegardées"
End Sub

Private Sub txtEnlevement_KeyPress(KeyAscii As Integer)
    ' Permettre seulement les chiffres
    If KeyAscii < 32 Then Exit Sub ' Touches de contrôle
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtReception_KeyPress(KeyAscii As Integer)
    ' Permettre seulement les chiffres
    If KeyAscii < 32 Then Exit Sub ' Touches de contrôle
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Raccourcis clavier
    Select Case KeyCode
        Case vbKeyF1
            cmdScanner_Click
        Case vbKeyF2
            cmdFormulaire_Click
        Case vbKeyF3
            cmdHistorique_Click
        Case vbKeyEscape
            If MsgBox("Voulez-vous vraiment quitter?", vbQuestion + vbYesNo) = vbYes Then
                End
            End If
    End Select
End Sub

