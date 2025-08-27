VERSION 5.00
Begin VB.Form frmFicheRetour 
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
Attribute VB_Name = "frmFicheRetour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private referenceFrigo As String
Private WithEvents cmdValider As CommandButton
Attribute cmdValider.VB_VarHelpID = -1
Private WithEvents cmdAnnuler As CommandButton
Attribute cmdAnnuler.VB_VarHelpID = -1
Private WithEvents optMecanique As OptionButton
Attribute optMecanique.VB_VarHelpID = -1
Private WithEvents optEsthetique As OptionButton
Attribute optEsthetique.VB_VarHelpID = -1
Private WithEvents optCoherenceOui As OptionButton
Attribute optCoherenceOui.VB_VarHelpID = -1
Private WithEvents optCoherenceNon As OptionButton
Attribute optCoherenceNon.VB_VarHelpID = -1
Private WithEvents optStandard As OptionButton
Attribute optStandard.VB_VarHelpID = -1
Private WithEvents optHS As OptionButton
Attribute optHS.VB_VarHelpID = -1

Private Sub Form_Load()
    Me.BackColor = RGB(240, 240, 240)
    Me.Caption = "FICHE RETOUR - RED BULL"
    Me.Width = 12000
    Me.Height = 9000
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    CreerInterfaceFiche
End Sub

Public Sub InitialiserAvecReference(reference As String)
    referenceFrigo = reference
    On Error Resume Next
    Me.Controls("txtReference").Text = referenceFrigo
    On Error GoTo 0
End Sub

Private Sub CreerInterfaceFiche()
    Dim ctrl As Object
    
    ' TITRE FICHE RETOUR
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitre")
    ctrl.Left = 1000
    ctrl.Top = 200
    ctrl.Width = 8000
    ctrl.Height = 400
    ctrl.Caption = "FICHE RETOUR"
    ctrl.BackColor = RGB(200, 200, 200)
    ctrl.BorderStyle = 1
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' N° ENLEVEMENT
    Set ctrl = Me.Controls.Add("VB.Label", "lblEnlevement")
    ctrl.Left = 500
    ctrl.Top = 800
    ctrl.Width = 1800
    ctrl.Caption = "N° ENLEVEMENT :"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtEnlevement")
    ctrl.Left = 2400
    ctrl.Top = 800
    ctrl.Width = 3000
    ctrl.Height = 300
    ctrl.Visible = True
    
    ' N° RECEPTION
    Set ctrl = Me.Controls.Add("VB.Label", "lblReception")
    ctrl.Left = 500
    ctrl.Top = 1200
    ctrl.Width = 1800
    ctrl.Caption = "N° RECEPTION :"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtReception")
    ctrl.Left = 2400
    ctrl.Top = 1200
    ctrl.Width = 3000
    ctrl.Height = 300
    ctrl.Visible = True
    
    ' REFERENCE
    Set ctrl = Me.Controls.Add("VB.Label", "lblReference")
    ctrl.Left = 500
    ctrl.Top = 1600
    ctrl.Width = 1800
    ctrl.Caption = "REFERENCE :"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtReference")
    ctrl.Left = 2400
    ctrl.Top = 1600
    ctrl.Width = 3000
    ctrl.Height = 300
    ctrl.Text = referenceFrigo
    ctrl.Enabled = False
    ctrl.BackColor = RGB(240, 240, 240)
    ctrl.Visible = True
    
    ' MOTIF DU RETOUR
    Set ctrl = Me.Controls.Add("VB.Label", "lblMotif")
    ctrl.Left = 500
    ctrl.Top = 2100
    ctrl.Width = 2000
    ctrl.Caption = "MOTIF DU RETOUR :"
    ctrl.Visible = True
    
    Set optMecanique = Me.Controls.Add("VB.OptionButton", "optMecanique")
    optMecanique.Left = 2600
    optMecanique.Top = 2100
    optMecanique.Width = 1500
    optMecanique.Caption = "MECANIQUE"
    optMecanique.Value = True
    optMecanique.Visible = True
    
    Set optEsthetique = Me.Controls.Add("VB.OptionButton", "optEsthetique")
    optEsthetique.Left = 4200
    optEsthetique.Top = 2100
    optEsthetique.Width = 1500
    optEsthetique.Caption = "ESTHETIQUE"
    optEsthetique.Visible = True
    
    ' COHERENCE AVEC LA BOUTIQUE
    Set ctrl = Me.Controls.Add("VB.Label", "lblCoherence")
    ctrl.Left = 500
    ctrl.Top = 2500
    ctrl.Width = 2500
    ctrl.Caption = "COHERENCE AVEC LA BOUTIQUE :"
    ctrl.Visible = True
    
    Set optCoherenceOui = Me.Controls.Add("VB.OptionButton", "optCoherenceOui")
    optCoherenceOui.Left = 3100
    optCoherenceOui.Top = 2500
    optCoherenceOui.Width = 600
    optCoherenceOui.Caption = "OUI"
    optCoherenceOui.Value = True
    optCoherenceOui.Visible = True
    
    Set optCoherenceNon = Me.Controls.Add("VB.OptionButton", "optCoherenceNon")
    optCoherenceNon.Left = 3800
    optCoherenceNon.Top = 2500
    optCoherenceNon.Width = 600
    optCoherenceNon.Caption = "NON"
    optCoherenceNon.Visible = True
    
    ' DIAGNOSTIC
    Set ctrl = Me.Controls.Add("VB.Label", "lblDiagnostic")
    ctrl.Left = 500
    ctrl.Top = 2900
    ctrl.Width = 1500
    ctrl.Caption = "DIAGNOSTIC :"
    ctrl.Visible = True
    
    ' Cases à cocher diagnostic
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkPieceManquante")
    ctrl.Left = 500
    ctrl.Top = 3300
    ctrl.Width = 4000
    ctrl.Caption = "PIECE MANQUANTE // PROBLEME CAPOT OU BAS DU FRIGO"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkTechnique")
    ctrl.Left = 500
    ctrl.Top = 3600
    ctrl.Width = 4000
    ctrl.Caption = "TECHNIQUE -> LUMIERE // FROID // MOTEUR // VITRE BRISEE"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkRayures")
    ctrl.Left = 500
    ctrl.Top = 3900
    ctrl.Width = 2000
    ctrl.Caption = "RAYURES TROP IMPORTANTES"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkLogoDegradé")
    ctrl.Left = 500
    ctrl.Top = 4200
    ctrl.Width = 2000
    ctrl.Caption = "LOGO DEGRADE"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkObsolete")
    ctrl.Left = 500
    ctrl.Top = 4500
    ctrl.Width = 2000
    ctrl.Caption = "OBSOLETE"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkBonEtat")
    ctrl.Left = 500
    ctrl.Top = 4800
    ctrl.Width = 3000
    ctrl.Caption = "BON ETAT -> REMIS DANS LE CIRCUIT"
    ctrl.Visible = True
    
    ' REPARE / RECUPERE
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkRepare")
    ctrl.Left = 500
    ctrl.Top = 5100
    ctrl.Width = 1500
    ctrl.Caption = "REPARE :"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtTempsRepare")
    ctrl.Left = 2500
    ctrl.Top = 5100
    ctrl.Width = 1000
    ctrl.Height = 300
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.Label", "lblTempsPasseRepare")
    ctrl.Left = 3600
    ctrl.Top = 5100
    ctrl.Width = 1200
    ctrl.Caption = "/ TEMPS PASSE :"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkRecupere")
    ctrl.Left = 500
    ctrl.Top = 5500
    ctrl.Width = 1500
    ctrl.Caption = "RECUPERE :"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtTempsRecupere")
    ctrl.Left = 2500
    ctrl.Top = 5500
    ctrl.Width = 1000
    ctrl.Height = 300
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.Label", "lblTempsPasseRecupere")
    ctrl.Left = 3600
    ctrl.Top = 5500
    ctrl.Width = 1200
    ctrl.Caption = "/ TEMPS PASSE :"
    ctrl.Visible = True
    
    ' N° SERIE
    Set ctrl = Me.Controls.Add("VB.Label", "lblSerie")
    ctrl.Left = 500
    ctrl.Top = 5900
    ctrl.Width = 1200
    ctrl.Caption = "N° SERIE :"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtSerie")
    ctrl.Left = 1800
    ctrl.Top = 5900
    ctrl.Width = 2000
    ctrl.Height = 300
    ctrl.Visible = True
    
    ' COMMENTAIRE
    Set ctrl = Me.Controls.Add("VB.Label", "lblCommentaire")
    ctrl.Left = 500
    ctrl.Top = 6300
    ctrl.Width = 1500
    ctrl.Caption = "COMMENTAIRE :"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtCommentaire")
    ctrl.Left = 500
    ctrl.Top = 6600
    ctrl.Width = 6000
    ctrl.Height = 300
    ctrl.Visible = True
    
    ' QUALITE
    Set ctrl = Me.Controls.Add("VB.Label", "lblQualite")
    ctrl.Left = 500
    ctrl.Top = 7000
    ctrl.Width = 1200
    ctrl.Caption = "QUALITE :"
    ctrl.Visible = True
    
    Set optStandard = Me.Controls.Add("VB.OptionButton", "optStandard")
    optStandard.Left = 1800
    optStandard.Top = 7000
    optStandard.Width = 1500
    optStandard.Caption = "STANDARD"
    optStandard.Value = True
    optStandard.Visible = True
    
    Set optHS = Me.Controls.Add("VB.OptionButton", "optHS")
    optHS.Left = 3500
    optHS.Top = 7000
    optHS.Width = 1000
    optHS.Caption = "HS"
    optHS.Visible = True
    
    ' Boutons
    Set cmdValider = Me.Controls.Add("VB.CommandButton", "cmdValider")
    cmdValider.Left = 2000
    cmdValider.Top = 7600
    cmdValider.Width = 1800
    cmdValider.Height = 400
    cmdValider.Caption = "VALIDER FICHE"
    cmdValider.BackColor = RGB(128, 255, 128)
    cmdValider.Visible = True
    
    Set cmdAnnuler = Me.Controls.Add("VB.CommandButton", "cmdAnnuler")
    cmdAnnuler.Left = 4000
    cmdAnnuler.Top = 7600
    cmdAnnuler.Width = 1800
    cmdAnnuler.Height = 400
    cmdAnnuler.Caption = "ANNULER"
    cmdAnnuler.BackColor = RGB(255, 128, 128)
    cmdAnnuler.Visible = True
End Sub

Private Sub cmdValider_Click()
    If Not ValiderFormulaire() Then Exit Sub
    
    Dim statut As String
    If optHS.Value = True Then
        statut = "HS"
    Else
        statut = "STANDARD"
    End If
    
    SauvegarderFiche statut
    
    If statut = "HS" Then
        MsgBox "Fiche sauvegardée - Frigo marqué HS" & vbCrLf & "Processus de récupération des pièces à implémenter", vbInformation
    Else
        MsgBox "Fiche sauvegardée - Frigo marqué STANDARD" & vbCrLf & "Frigo remis en circuit", vbInformation
    End If
    
    Me.Hide
End Sub

Private Function ValiderFormulaire() As Boolean
    If Len(Trim(Me.Controls("txtEnlevement").Text)) = 0 Then
        MsgBox "Veuillez saisir le numéro d'enlèvement !", vbExclamation
        ValiderFormulaire = False
        Exit Function
    End If
    
    If Len(Trim(Me.Controls("txtReception").Text)) = 0 Then
        MsgBox "Veuillez saisir le numéro de réception !", vbExclamation
        ValiderFormulaire = False
        Exit Function
    End If
    
    ValiderFormulaire = True
End Function

Private Sub SauvegarderFiche(statut As String)
    On Error GoTo GestionErreur
    
    If Dir(App.Path & "\Fiches", vbDirectory) = "" Then
        MkDir App.Path & "\Fiches"
    End If
    
    Dim fichier As String
    Dim numeroFichier As Integer
    
    fichier = App.Path & "\Fiches\Fiche_" & referenceFrigo & "_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    numeroFichier = FreeFile
    
    Open fichier For Output As #numeroFichier
    Print #numeroFichier, "=== FICHE RETOUR RED BULL ==="
    Print #numeroFichier, "N° ENLEVEMENT: " & Me.Controls("txtEnlevement").Text
    Print #numeroFichier, "N° RECEPTION: " & Me.Controls("txtReception").Text
    Print #numeroFichier, "REFERENCE: " & Me.Controls("txtReference").Text
    Print #numeroFichier, ""
    Print #numeroFichier, "MOTIF DU RETOUR:"
    If optMecanique.Value Then Print #numeroFichier, "- MECANIQUE"
    If optEsthetique.Value Then Print #numeroFichier, "- ESTHETIQUE"
    Print #numeroFichier, ""
    Print #numeroFichier, "COHERENCE AVEC LA BOUTIQUE:"
    If optCoherenceOui.Value Then Print #numeroFichier, "- OUI"
    If optCoherenceNon.Value Then Print #numeroFichier, "- NON"
    Print #numeroFichier, ""
    Print #numeroFichier, "DIAGNOSTIC:"
    If Me.Controls("chkPieceManquante").Value = 1 Then Print #numeroFichier, "- PIECE MANQUANTE"
    If Me.Controls("chkTechnique").Value = 1 Then Print #numeroFichier, "- TECHNIQUE"
    If Me.Controls("chkRayures").Value = 1 Then Print #numeroFichier, "- RAYURES"
    If Me.Controls("chkLogoDegradé").Value = 1 Then Print #numeroFichier, "- LOGO DEGRADE"
    If Me.Controls("chkObsolete").Value = 1 Then Print #numeroFichier, "- OBSOLETE"
    If Me.Controls("chkBonEtat").Value = 1 Then Print #numeroFichier, "- BON ETAT"
    Print #numeroFichier, ""
    If Me.Controls("chkRepare").Value = 1 Then Print #numeroFichier, "REPARE - Temps: " & Me.Controls("txtTempsRepare").Text
    If Me.Controls("chkRecupere").Value = 1 Then Print #numeroFichier, "RECUPERE - Temps: " & Me.Controls("txtTempsRecupere").Text
    Print #numeroFichier, "N° SERIE: " & Me.Controls("txtSerie").Text
    Print #numeroFichier, "COMMENTAIRE: " & Me.Controls("txtCommentaire").Text
    Print #numeroFichier, ""
    Print #numeroFichier, "QUALITE: " & statut
    Print #numeroFichier, "Date création: " & Now
    Close #numeroFichier
    
    Exit Sub
    
GestionErreur:
    MsgBox "Erreur lors de la sauvegarde: " & Err.Description, vbCritical
End Sub

Private Sub cmdAnnuler_Click()
    If MsgBox("Etes-vous sûr de vouloir annuler cette fiche ?", vbYesNo + vbQuestion) = vbYes Then
        Me.Hide
    End If
End Sub
