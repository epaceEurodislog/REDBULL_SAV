VERSION 5.00
Begin VB.Form frmPrincipal 
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
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ' Configurer le formulaire
    Me.Caption = "SAV Red Bull Scanner Pro - v2.1 - [frmPrincipal.frm]"
    Me.BorderStyle = 1 ' Fixed Single
    Me.MaxButton = False
    Me.StartUpPosition = 2 ' CenterScreen
    Me.BackColor = &HF0F0F0
    Me.Width = 11415
    Me.Height = 9030
    
    ' Créer les contrôles
    CreerControles
    
    ' Initialiser les valeurs
    InitialiserFormulaire
End Sub

Private Sub CreerControles()
    ' Cette fonction crée tous les contrôles dynamiquement
    
    ' === TITRE ===
    Set lblTitre = Me.Controls.Add("VB.Label", "lblTitre")
    With lblTitre
        .Left = 240
        .Top = 240
        .Width = 9255
        .Height = 375
        .Caption = "FICHE RETOUR SAV RED BULL"
        .Alignment = 2 ' Center
        .BackColor = &H4080FF
        .ForeColor = &HFFFFFF
        .Font.Size = 12
        .Font.Bold = True
        .Visible = True
    End With
    
    ' === SOUS-TITRE ===
    Set lblSousTitre = Me.Controls.Add("VB.Label", "lblSousTitre")
    With lblSousTitre
        .Left = 240
        .Top = 600
        .Width = 9255
        .Height = 255
        .Caption = "Système de Gestion des Réfrigérateurs - Interface Frigoriste"
        .Alignment = 2 ' Center
        .BackColor = &H4080FF
        .ForeColor = &HFFFFFF
        .Font.Bold = True
        .Visible = True
    End With
    
    ' === BOUTONS ===
    Set cmdScanner = Me.Controls.Add("VB.CommandButton", "cmdScanner")
    With cmdScanner
        .Left = 4440
        .Top = 840
        .Width = 1575
        .Height = 735
        .Caption = "Scanner"
        .Visible = True
    End With
    
    Set cmdFormulaire = Me.Controls.Add("VB.CommandButton", "cmdFormulaire")
    With cmdFormulaire
        .Left = 6120
        .Top = 840
        .Width = 1575
        .Height = 735
        .Caption = "Formulaire"
        .BackColor = &H4080FF
        .Font.Bold = True
        .Visible = True
    End With
    
    Set cmdHistorique = Me.Controls.Add("VB.CommandButton", "cmdHistorique")
    With cmdHistorique
        .Left = 7800
        .Top = 840
        .Width = 1575
        .Height = 735
        .Caption = "Historique"
        .Visible = True
    End With
    
    Set cmdScannerPro = Me.Controls.Add("VB.CommandButton", "cmdScannerPro")
    With cmdScannerPro
        .Left = 9600
        .Top = 120
        .Width = 1575
        .Height = 375
        .Caption = "Scanner Pro"
        .BackColor = &H80FF80
        .Font.Bold = True
        .Visible = True
    End With
    
    ' === FRAME INFORMATIONS ===
    Set frameInfos = Me.Controls.Add("VB.Frame", "frameInfos")
    With frameInfos
        .Left = 240
        .Top = 1800
        .Width = 10935
        .Height = 2535
        .Caption = "Informations générales"
        .Font.Bold = True
        .Visible = True
    End With
    
    ' === CHAMPS TEXTE ===
    Set txtEnlevement = frameInfos.Controls.Add("VB.TextBox", "txtEnlevement")
    With txtEnlevement
        .Left = 1680
        .Top = 480
        .Width = 2415
        .Height = 315
        .Visible = True
    End With
    
    Set lblEnlevement = frameInfos.Controls.Add("VB.Label", "lblEnlevement")
    With lblEnlevement
        .Left = 240
        .Top = 540
        .Width = 1215
        .Height = 255
        .Caption = "N° Enlèvement:"
        .Visible = True
    End With
    
    Set txtReception = frameInfos.Controls.Add("VB.TextBox", "txtReception")
    With txtReception
        .Left = 1680
        .Top = 960
        .Width = 2415
        .Height = 315
        .Visible = True
    End With
    
    Set lblReception = frameInfos.Controls.Add("VB.Label", "lblReception")
    With lblReception
        .Left = 240
        .Top = 1020
        .Width = 1095
        .Height = 255
        .Caption = "N° Réception:"
        .Visible = True
    End With
    
    Set txtDate = frameInfos.Controls.Add("VB.TextBox", "txtDate")
    With txtDate
        .Left = 1680
        .Top = 1440
        .Width = 2415
        .Height = 315
        .Visible = True
    End With
    
    Set lblDate = frameInfos.Controls.Add("VB.Label", "lblDate")
    With lblDate
        .Left = 240
        .Top = 1500
        .Width = 375
        .Height = 255
        .Caption = "Date:"
        .Visible = True
    End With
    
    Set txtReference = frameInfos.Controls.Add("VB.TextBox", "txtReference")
    With txtReference
        .Left = 1680
        .Top = 1920
        .Width = 8775
        .Height = 315
        .Visible = True
    End With
    
    Set lblReference = frameInfos.Controls.Add("VB.Label", "lblReference")
    With lblReference
        .Left = 240
        .Top = 1980
        .Width = 1335
        .Height = 255
        .Caption = "Référence produit:"
        .Visible = True
    End With
    
    ' === FRAME MOTIF ===
    Set frameMotif = Me.Controls.Add("VB.Frame", "frameMotif")
    With frameMotif
        .Left = 240
        .Top = 4440
        .Width = 10935
        .Height = 2175
        .Caption = "Motif du retour"
        .Font.Bold = True
        .Visible = True
    End With
    
    Set optMecanique = frameMotif.Controls.Add("VB.OptionButton", "optMecanique")
    With optMecanique
        .Left = 240
        .Top = 480
        .Width = 4935
        .Height = 255
        .Caption = "MECANIQUE"
        .Value = True
        .Visible = True
    End With
    
    Set optEsthetique = frameMotif.Controls.Add("VB.OptionButton", "optEsthetique")
    With optEsthetique
        .Left = 240
        .Top = 840
        .Width = 4935
        .Height = 255
        .Caption = "ESTHETIQUE"
        .Visible = True
    End With
    
    ' === FRAME COHERENCE ===
    Set frameCoherence = frameMotif.Controls.Add("VB.Frame", "frameCoherence")
    With frameCoherence
        .Left = 5520
        .Top = 480
        .Width = 5175
        .Height = 855
        .Caption = "Cohérence avec la boutique:"
        .Visible = True
    End With
    
    Set optOui = frameCoherence.Controls.Add("VB.OptionButton", "optOui")
    With optOui
        .Left = 240
        .Top = 240
        .Width = 4695
        .Height = 255
        .Caption = "OUI"
        .Value = True
        .Visible = True
    End With
    
    Set optNon = frameCoherence.Controls.Add("VB.OptionButton", "optNon")
    With optNon
        .Left = 240
        .Top = 480
        .Width = 4695
        .Height = 255
        .Caption = "NON"
        .Visible = True
    End With
    
    ' === FRAME DIAGNOSTIC ===
    Set frameDiagnostic = Me.Controls.Add("VB.Frame", "frameDiagnostic")
    With frameDiagnostic
        .Left = 240
        .Top = 6360
        .Width = 10935
        .Height = 2415
        .Caption = "Diagnostic technique"
        .Font.Bold = True
        .Visible = True
    End With
    
    Set chkPiece = frameDiagnostic.Controls.Add("VB.CheckBox", "chkPiece")
    With chkPiece
        .Left = 240
        .Top = 960
        .Width = 10455
        .Height = 255
        .Caption = "PIECE MANQUANTE / PROBLEME CAPOT DU BAS DU FRIGO"
        .Visible = True
    End With
    
    Set chkTechnique = frameDiagnostic.Controls.Add("VB.CheckBox", "chkTechnique")
    With chkTechnique
        .Left = 240
        .Top = 1440
        .Width = 10455
        .Height = 255
        .Caption = "TECHNIQUE : LUMIERE / FROID / MOTEUR / VITRE BRISEE"
        .Visible = True
    End With
    
    Set chkRayures = frameDiagnostic.Controls.Add("VB.CheckBox", "chkRayures")
    With chkRayures
        .Left = 240
        .Top = 1920
        .Width = 10455
        .Height = 255
        .Caption = "RAYURES TROP IMPORTANTES"
        .Value = 1 ' Checked
        .Visible = True
    End With
End Sub

Private Sub InitialiserFormulaire()
    txtEnlevement.Text = "69113"
    txtReception.Text = "19108"
    txtDate.Text = "05/06/25"
    txtReference.Text = "VCZ286 52000-1"
End Sub

' === DÉCLARATIONS DES CONTRÔLES ===
Dim lblTitre As Label
Dim lblSousTitre As Label
Dim cmdScanner As CommandButton
Dim cmdFormulaire As CommandButton
Dim cmdHistorique As CommandButton
Dim cmdScannerPro As CommandButton
Dim frameInfos As Frame
Dim txtEnlevement As TextBox
Dim lblEnlevement As Label
Dim txtReception As TextBox
Dim lblReception As Label
Dim txtDate As TextBox
Dim lblDate As Label
Dim txtReference As TextBox
Dim lblReference As Label
Dim frameMotif As Frame
Dim optMecanique As OptionButton
Dim optEsthetique As OptionButton
Dim frameCoherence As Frame
Dim optOui As OptionButton
Dim optNon As OptionButton
Dim frameDiagnostic As Frame
Dim chkPiece As CheckBox
Dim chkTechnique As CheckBox
Dim chkRayures As CheckBox

' === ÉVÉNEMENTS ===
Private Sub cmdScanner_Click()
    MsgBox "Fonction Scanner", vbInformation
End Sub

Private Sub cmdFormulaire_Click()
    MsgBox "Formulaire validé!" & vbCrLf & _
           "N° Enlèvement: " & txtEnlevement.Text & vbCrLf & _
           "N° Réception: " & txtReception.Text & vbCrLf & _
           "Date: " & txtDate.Text & vbCrLf & _
           "Référence: " & txtReference.Text, vbInformation
End Sub

Private Sub cmdHistorique_Click()
    MsgBox "Fonction Historique", vbInformation
End Sub

Private Sub cmdScannerPro_Click()
    MsgBox "Scanner Pro activé", vbInformation
End Sub

