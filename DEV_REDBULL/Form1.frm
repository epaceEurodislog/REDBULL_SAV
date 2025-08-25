VERSION 5.00
Begin VB.Form frmPrincipal 
   Caption         =   "Form1"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ' Configuration de l'apparence du formulaire
    Me.BackColor = RGB(240, 240, 240)
    Me.Caption = "SAV Red Bull"
    Me.Width = 12000
    Me.Height = 9000
    
    ' Centrer le formulaire
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    ' Cr�er l'interface
    CreerInterface
End Sub

Private Sub CreerInterface()
    Dim ctrl As Object
    
    ' Titre principal
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitre")
    ctrl.Left = 240
    ctrl.Top = 120
    ctrl.Width = 8295
    ctrl.Height = 375
    ctrl.Caption = "FICHE RETOUR SAV RED BULL"
    ctrl.BackColor = RGB(51, 102, 255)
    ctrl.ForeColor = RGB(255, 255, 255)
    ctrl.Font.Size = 14
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Sous-titre
    Set ctrl = Me.Controls.Add("VB.Label", "lblSousTitre")
    ctrl.Left = 240
    ctrl.Top = 480
    ctrl.Width = 8295
    ctrl.Height = 255
    ctrl.Caption = "Syst�me de Gestion des R�frig�rateurs - Interface Frigoriste"
    ctrl.BackColor = RGB(51, 102, 255)
    ctrl.ForeColor = RGB(255, 255, 255)
    ctrl.Font.Size = 10
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Bouton Scanner
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdScanner")
    ctrl.Left = 480
    ctrl.Top = 840
    ctrl.Width = 1575
    ctrl.Height = 495
    ctrl.Caption = "Scanner"
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    ' Bouton Formulaire
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdFormulaire")
    ctrl.Left = 3240
    ctrl.Top = 840
    ctrl.Width = 1575
    ctrl.Height = 495
    ctrl.Caption = "Formulaire"
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    ' Bouton Historique
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdHistorique")
    ctrl.Left = 6000
    ctrl.Top = 840
    ctrl.Width = 1575
    ctrl.Height = 495
    ctrl.Caption = "Historique"
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    ' === SECTION INFORMATIONS G�N�RALES ===
    
    ' Titre section Informations
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitreInfos")
    ctrl.Left = 240
    ctrl.Top = 1440
    ctrl.Width = 8295
    ctrl.Height = 300
    ctrl.Caption = "=== INFORMATIONS G�N�RALES ==="
    ctrl.BackColor = RGB(200, 200, 200)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Label N� Enl�vement
    Set ctrl = Me.Controls.Add("VB.Label", "lblEnlevement")
    ctrl.Left = 480
    ctrl.Top = 1800
    ctrl.Width = 1335
    ctrl.Height = 255
    ctrl.Caption = "N� Enl�vement:"
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    ' TextBox N� Enl�vement
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtEnlevement")
    ctrl.Left = 480
    ctrl.Top = 2040
    ctrl.Width = 7055
    ctrl.Height = 285
    ctrl.Text = "69113"
    ctrl.Visible = True
    
    ' Label N� R�ception
    Set ctrl = Me.Controls.Add("VB.Label", "lblReception")
    ctrl.Left = 480
    ctrl.Top = 2400
    ctrl.Width = 1215
    ctrl.Height = 255
    ctrl.Caption = "N� R�ception:"
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    ' TextBox N� R�ception
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtReception")
    ctrl.Left = 480
    ctrl.Top = 2640
    ctrl.Width = 7055
    ctrl.Height = 285
    ctrl.Text = "19108"
    ctrl.Visible = True
    
    ' Label Date
    Set ctrl = Me.Controls.Add("VB.Label", "lblDate")
    ctrl.Left = 480
    ctrl.Top = 3000
    ctrl.Width = 495
    ctrl.Height = 255
    ctrl.Caption = "Date:"
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    ' TextBox Date
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtDate")
    ctrl.Left = 480
    ctrl.Top = 3240
    ctrl.Width = 7055
    ctrl.Height = 285
    ctrl.Text = "05/06/25"
    ctrl.Visible = True
    
    ' Label R�f�rence
    Set ctrl = Me.Controls.Add("VB.Label", "lblReference")
    ctrl.Left = 480
    ctrl.Top = 3600
    ctrl.Width = 1455
    ctrl.Height = 255
    ctrl.Caption = "R�f�rence produit:"
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    ' TextBox R�f�rence
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtReference")
    ctrl.Left = 480
    ctrl.Top = 3840
    ctrl.Width = 7055
    ctrl.Height = 285
    ctrl.Text = "VC2286 52000-1"
    ctrl.Visible = True
    
    ' === SECTION MOTIF DU RETOUR ===
    
    ' Titre section Motif
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitreMotif")
    ctrl.Left = 240
    ctrl.Top = 4200
    ctrl.Width = 8295
    ctrl.Height = 300
    ctrl.Caption = "=== MOTIF DU RETOUR ==="
    ctrl.BackColor = RGB(200, 200, 200)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Option M�canique
    Set ctrl = Me.Controls.Add("VB.OptionButton", "optMecanique")
    ctrl.Left = 480
    ctrl.Top = 4560
    ctrl.Width = 1575
    ctrl.Height = 255
    ctrl.Caption = "M�CANIQUE"
    ctrl.Value = True
    ctrl.BackColor = RGB(255, 128, 128)
    ctrl.Visible = True
    
    ' Option Esth�tique
    Set ctrl = Me.Controls.Add("VB.OptionButton", "optEsthetique")
    ctrl.Left = 480
    ctrl.Top = 4840
    ctrl.Width = 1575
    ctrl.Height = 255
    ctrl.Caption = "ESTH�TIQUE"
    ctrl.Visible = True
    
    ' === SECTION COH�RENCE ===
    
    ' Titre section Coh�rence
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitreCoherence")
    ctrl.Left = 240
    ctrl.Top = 5160
    ctrl.Width = 8295
    ctrl.Height = 300
    ctrl.Caption = "=== COH�RENCE AVEC LA BOUTIQUE ==="
    ctrl.BackColor = RGB(200, 200, 200)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Option OUI
    Set ctrl = Me.Controls.Add("VB.OptionButton", "optOui")
    ctrl.Left = 480
    ctrl.Top = 5520
    ctrl.Width = 855
    ctrl.Height = 255
    ctrl.Caption = "OUI"
    ctrl.Value = True
    ctrl.BackColor = RGB(128, 255, 128)
    ctrl.Visible = True
    
    ' Option NON
    Set ctrl = Me.Controls.Add("VB.OptionButton", "optNon")
    ctrl.Left = 1680
    ctrl.Top = 5520
    ctrl.Width = 855
    ctrl.Height = 255
    ctrl.Caption = "NON"
    ctrl.Visible = True
    
    ' === SECTION DIAGNOSTIC ===
    
    ' Titre section Diagnostic
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitreDiagnostic")
    ctrl.Left = 240
    ctrl.Top = 5880
    ctrl.Width = 8295
    ctrl.Height = 300
    ctrl.Caption = "=== DIAGNOSTIC TECHNIQUE ==="
    ctrl.BackColor = RGB(200, 200, 200)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' CheckBox Pi�ce
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkPiece")
    ctrl.Left = 480
    ctrl.Top = 6240
    ctrl.Width = 5295
    ctrl.Height = 255
    ctrl.Caption = "PI�CE MANQUANTE // PROBL�ME CAPOT OU BAS DU FRIGO"
    ctrl.Visible = True
    
    ' CheckBox Technique
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkTechnique")
    ctrl.Left = 480
    ctrl.Top = 6520
    ctrl.Width = 5775
    ctrl.Height = 255
    ctrl.Caption = "TECHNIQUE � LUMI�RE // FROID // MOTEUR // VITRE BRIS�E"
    ctrl.Visible = True
    
    ' CheckBox Rayures
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkRayures")
    ctrl.Left = 480
    ctrl.Top = 6800
    ctrl.Width = 3375
    ctrl.Height = 255
    ctrl.Caption = "RAYURES TROP IMPORTANTES"
    ctrl.Value = 1
    ctrl.Visible = True
    
    ' === BOUTONS D'ACTION ===
    
    ' Bouton Sauvegarder
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdSauvegarder")
    ctrl.Left = 9000
    ctrl.Top = 2000
    ctrl.Width = 1200
    ctrl.Height = 400
    ctrl.Caption = "Sauvegarder"
    ctrl.Font.Bold = True
    ctrl.BackColor = RGB(128, 255, 128)
    ctrl.Visible = True
    
    ' Bouton Nouveau
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdNouveau")
    ctrl.Left = 9000
    ctrl.Top = 2500
    ctrl.Width = 1200
    ctrl.Height = 400
    ctrl.Caption = "Nouveau"
    ctrl.Font.Bold = True
    ctrl.BackColor = RGB(255, 255, 128)
    ctrl.Visible = True
    
    ' Bouton Imprimer
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdImprimer")
    ctrl.Left = 9000
    ctrl.Top = 3000
    ctrl.Width = 1200
    ctrl.Height = 400
    ctrl.Caption = "Imprimer"
    ctrl.Font.Bold = True
    ctrl.BackColor = RGB(200, 200, 255)
    ctrl.Visible = True
End Sub

Private Sub cmdScanner_Click()
    MsgBox "Fonction Scanner activ�e", vbInformation, "Scanner"
End Sub
    MsgBox "Ouverture du formulaire", vbInformation, "Formulaire"

Private Sub cmdFormulaire_Click()
End Sub

Private Sub cmdHistorique_Click()
    MsgBox "Affichage de l'historique", vbInformation, "Historique"
End Sub

Private Sub cmdSauvegarder_Click()
    SauvegarderDonnees
End Sub

Private Sub cmdNouveau_Click()
    ReinitialiserFormulaire
End Sub

Private Sub cmdImprimer_Click()
    ImprimerFiche
End Sub

Private Sub SauvegarderDonnees()
    Dim fichier As String
    Dim numeroFichier As Integer
    
    ' Validation des donn�es
    If Len(Trim(Me.Controls("txtEnlevement").Text)) = 0 Then
        MsgBox "Le num�ro d'enl�vement est obligatoire !", vbExclamation
        Me.Controls("txtEnlevement").SetFocus
        Exit Sub
    End If
    
    fichier = App.Path & "\SAV_" & Me.Controls("txtEnlevement").Text & ".txt"
    numeroFichier = FreeFile
    
    Open fichier For Output As #numeroFichier
    Print #numeroFichier, "=== FICHE RETOUR SAV RED BULL ==="
    Print #numeroFichier, "Date de cr�ation: " & Now
    Print #numeroFichier, String(50, "=")
    Print #numeroFichier, ""
    Print #numeroFichier, "INFORMATIONS G�N�RALES:"
    Print #numeroFichier, "N� Enl�vement: " & Me.Controls("txtEnlevement").Text
    Print #numeroFichier, "N� R�ception: " & Me.Controls("txtReception").Text
    Print #numeroFichier, "Date: " & Me.Controls("txtDate").Text
    Print #numeroFichier, "R�f�rence produit: " & Me.Controls("txtReference").Text
    Print #numeroFichier, ""
    Print #numeroFichier, "MOTIF DU RETOUR:"
    If Me.Controls("optMecanique").Value Then Print #numeroFichier, "- M�CANIQUE"
    If Me.Controls("optEsthetique").Value Then Print #numeroFichier, "- ESTH�TIQUE"
    Print #numeroFichier, ""
    Print #numeroFichier, "COH�RENCE AVEC LA BOUTIQUE:"
    If Me.Controls("optOui").Value Then Print #numeroFichier, "- OUI"
    If Me.Controls("optNon").Value Then Print #numeroFichier, "- NON"
    Print #numeroFichier, ""
    Print #numeroFichier, "DIAGNOSTIC TECHNIQUE:"
    If Me.Controls("chkPiece").Value = 1 Then Print #numeroFichier, "- PI�CE MANQUANTE // PROBL�ME CAPOT OU BAS DU FRIGO"
    If Me.Controls("chkTechnique").Value = 1 Then Print #numeroFichier, "- TECHNIQUE � LUMI�RE // FROID // MOTEUR // VITRE BRIS�E"
    If Me.Controls("chkRayures").Value = 1 Then Print #numeroFichier, "- RAYURES TROP IMPORTANTES"
    Print #numeroFichier, ""
    Print #numeroFichier, String(50, "=")
    Close #numeroFichier
    
    MsgBox "Donn�es sauvegard�es avec succ�s !" & vbCrLf & "Fichier: " & fichier, vbInformation, "Sauvegarde"
End Sub

Private Sub ReinitialiserFormulaire()
    Me.Controls("txtEnlevement").Text = ""
    Me.Controls("txtReception").Text = ""
    Me.Controls("txtDate").Text = Format(Date, "dd/mm/yy")
    Me.Controls("txtReference").Text = ""
    
    Me.Controls("optMecanique").Value = True
    Me.Controls("optOui").Value = True
    
    Me.Controls("chkPiece").Value = 0
    Me.Controls("chkTechnique").Value = 0
    Me.Controls("chkRayures").Value = 0
    
    Me.Controls("txtEnlevement").SetFocus
End Sub

Private Sub ImprimerFiche()
    ' Fonction d'impression simple
    Printer.Font.Name = "Arial"
    Printer.Font.Size = 12
    Printer.Font.Bold = True
    
    Printer.Print "=== FICHE RETOUR SAV RED BULL ==="
    Printer.Print "Date: " & Now
    Printer.Print ""
    Printer.Font.Bold = False
    Printer.Print "N� Enl�vement: " & Me.Controls("txtEnlevement").Text
    Printer.Print "N� R�ception: " & Me.Controls("txtReception").Text
    Printer.Print "Date: " & Me.Controls("txtDate").Text
    Printer.Print "R�f�rence: " & Me.Controls("txtReference").Text
    
    Printer.EndDoc
    MsgBox "Impression lanc�e !", vbInformation
End Sub
