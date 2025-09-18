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
' === FORM1.FRM - VERSION PRODUCTION SAV RED BULL ===

Option Explicit

Private referenceValidee As String
Private numeroSerieValide As String
Private WithEvents cmdValider As CommandButton
Attribute cmdValider.VB_VarHelpID = -1
Private WithEvents cmdOuvrirFiche As CommandButton
Attribute cmdOuvrirFiche.VB_VarHelpID = -1
Private WithEvents tmrVerifBDD As Timer
Attribute tmrVerifBDD.VB_VarHelpID = -1

' === CONSTANTES COULEURS RED BULL ===
Private Const ROUGE_REDBULL = &H1414DC
Private Const BLEU_REDBULL = &HCC6600
Private Const JAUNE_REDBULL = &HFFFF00
Private Const ARGENT_REDBULL = &HC0C0C0

Private Sub Form_Load()
    Me.Caption = "SAV Red Bull"
    Me.Width = 8000
    Me.Height = 6000
    referenceValidee = ""
    numeroSerieValide = ""
    
    ' Initialiser l'application avec BDD
    InitialiserApplication
    
    ' Créer les contrôles dynamiquement
    CreerControles
    
    ' Démarrer le timer de vérification BDD
    DemarrerTimerBDD
End Sub

Private Sub CreerControles()
    Dim ctrl As Object
    
    ' === TITRE PRINCIPAL ===
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitre")
    ctrl.Left = 600
    ctrl.Top = 200
    ctrl.Width = 6800
    ctrl.Height = 500
    ctrl.Caption = "SAV RED BULL"
    ctrl.BackColor = ROUGE_REDBULL
    ctrl.ForeColor = RGB(255, 255, 255)
    ctrl.Alignment = 2
    ctrl.Font.Size = 18
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    ' === INDICATEUR STATUT BDD ===
    Set ctrl = Me.Controls.Add("VB.Label", "lblStatutBDD")
    ctrl.Left = 600
    ctrl.Top = 750
    ctrl.Width = 6800
    ctrl.Height = 300
    ActualiserStatutBDD
    ctrl.Alignment = 2
    ctrl.Font.Bold = True
    ctrl.Font.Size = 10
    ctrl.Visible = True
    
    ' === ZONE DE SAISIE ===
    Set ctrl = Me.Controls.Add("VB.Label", "lblRef")
    ctrl.Left = 600
    ctrl.Top = 1200
    ctrl.Width = 2000
    ctrl.Caption = "Numéro de série frigo :"
    ctrl.ForeColor = ROUGE_REDBULL
    ctrl.Font.Size = 11
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtReference")
    ctrl.Left = 2700
    ctrl.Top = 1200
    ctrl.Width = 2800
    ctrl.Height = 400
    ctrl.Font.Size = 12
    ctrl.BackColor = RGB(255, 255, 255)
    ctrl.Visible = True
    
    Set cmdValider = Me.Controls.Add("VB.CommandButton", "cmdValider")
    cmdValider.Left = 5600
    cmdValider.Top = 1200
    cmdValider.Width = 1400
    cmdValider.Height = 450
    cmdValider.Caption = "SCANNER"
    cmdValider.Font.Bold = True
    cmdValider.Font.Size = 11
    cmdValider.BackColor = BLEU_REDBULL
    cmdValider.Visible = True
    
    ' === ZONE D'INFORMATIONS ===
    Set ctrl = Me.Controls.Add("VB.Label", "lblInfo")
    ctrl.Left = 600
    ctrl.Top = 1800
    ctrl.Width = 6800
    ctrl.Height = 2000
    ctrl.Caption = "INSTRUCTIONS D'UTILISATION :" & vbCrLf & vbCrLf & _
                   "1. Scannez le code-barres du frigo Red Bull" & vbCrLf & _
                   "2. Cliquez SCANNER pour vérification" & vbCrLf & _
                   "3. Si l'équipement est reconnu, vous pourrez ouvrir la fiche SAV" & vbCrLf & vbCrLf & _
                   "ÉQUIPEMENTS PRIS EN CHARGE :" & vbCrLf & _
                   "• Frigos vitrine Red Bull" & vbCrLf & _
                   "• Distributeurs de boissons Red Bull" & vbCrLf & _
                   "• Équipements de réfrigération Red Bull" & vbCrLf & vbCrLf & _
                   "En cas de problème, contactez votre superviseur."
    ctrl.BackColor = RGB(248, 249, 250)
    ctrl.BorderStyle = 1
    ctrl.Alignment = 0
    ctrl.Font.Size = 10
    ctrl.Visible = True
    
    ' === BOUTON PRINCIPAL ===
    Set cmdOuvrirFiche = Me.Controls.Add("VB.CommandButton", "cmdOuvrirFiche")
    cmdOuvrirFiche.Left = 2000
    cmdOuvrirFiche.Top = 4000
    cmdOuvrirFiche.Width = 4000
    cmdOuvrirFiche.Height = 600
    cmdOuvrirFiche.Caption = "OUVRIR FICHE RETOUR SAV"
    cmdOuvrirFiche.Enabled = False
    cmdOuvrirFiche.BackColor = RGB(150, 150, 150)
    cmdOuvrirFiche.Font.Bold = True
    cmdOuvrirFiche.Font.Size = 12
    cmdOuvrirFiche.Visible = True
    
    ' === ZONE DE STATUT ===
    Set ctrl = Me.Controls.Add("VB.Label", "lblStatut")
    ctrl.Left = 600
    ctrl.Top = 4700
    ctrl.Width = 6800
    ctrl.Height = 300
    ctrl.Caption = "Prêt - En attente de scan | " & ObtenirDateTimeFormatee()
    ctrl.BackColor = RGB(240, 240, 240)
    ctrl.BorderStyle = 1
    ctrl.Alignment = 2
    ctrl.Font.Size = 9
    ctrl.Visible = True
End Sub

' === ÉVÉNEMENTS PRINCIPAUX ===

Private Sub cmdValider_Click()
    Dim numeroSerie As String
    numeroSerie = Trim(UCase(Me.Controls("txtReference").Text))
    
    Me.Controls("lblStatut").Caption = "Validation en cours... | " & ObtenirDateTimeFormatee()
    Me.Refresh
    
    If Len(numeroSerie) = 0 Then
        MsgBox "Veuillez saisir un numéro de série !", vbExclamation, "Erreur de saisie"
        Me.Controls("lblStatut").Caption = "Erreur: Numéro de série manquant | " & ObtenirDateTimeFormatee()
        Exit Sub
    End If
    
    ' Validation du format de base
    If Not ValiderFormatNumeroSerie(numeroSerie) Then
        AfficherErreurValidation "Format de numéro de série invalide", numeroSerie
        Exit Sub
    End If
    
    ' Vérification dans la base de données
    If Not VerifierConnexionBDD() Then
        AfficherErreurConnexion numeroSerie
        Exit Sub
    End If
    
    ' Rechercher dans NSE_DAT et ART_PAR
    Dim resultatsValidation As TypeValidationBDD
    resultatsValidation = ValiderNumeroSerieBDD(numeroSerie)
    
    If resultatsValidation.existe Then
        ' Validation réussie
        referenceValidee = numeroSerie
        numeroSerieValide = numeroSerie
        AfficherValidationReussie resultatsValidation
        
        ' Enregistrer dans l'historique
        EcrireHistoriqueScan numeroSerie, resultatsValidation.modeleArticle
    Else
        ' Numéro de série non autorisé
        AfficherErreurValidation resultatsValidation.statut, numeroSerie
    End If
End Sub

Private Sub cmdOuvrirFiche_Click()
    If Len(referenceValidee) = 0 Or Len(numeroSerieValide) = 0 Then
        MsgBox "Aucun numéro de série validé !", vbExclamation
        Exit Sub
    End If
    
    Me.Controls("lblStatut").Caption = "Ouverture fiche retour... | " & ObtenirDateTimeFormatee()
    Me.Refresh
    
    ' Récupérer l'ART_CODE correspondant au numéro de série
    Dim validationComplete As TypeValidationBDD
    validationComplete = ValiderNumeroSerieBDD(numeroSerieValide)
    
    If validationComplete.existe Then
        ' Ouvrir la fiche retour
        Load frmFicheRetour
        frmFicheRetour.InitialiserAvecReference validationComplete.codeArticle, numeroSerieValide
        frmFicheRetour.Show vbModal
        
        ' Log de l'action
        EcrireHistoriqueScan numeroSerieValide, "Fiche SAV créée pour " & validationComplete.modeleArticle
    Else
        MsgBox "Erreur lors de la récupération des données. Veuillez rescanner.", vbExclamation
    End If
    
    ' Reset après fermeture de la fiche
    ResetFormulaire
End Sub

' === GESTION DES RACCOURCIS CLAVIER ===

Private Sub txtReference_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ' Touche Entrée
        KeyAscii = 0
        cmdValider_Click
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 ' F1 = Aide
            AfficherAideComplete
        Case vbKeyEscape ' Échap = Reset
            ResetFormulaire
    End Select
End Sub

' === FONCTIONS D'AFFICHAGE ===

Private Sub AfficherValidationReussie(resultats As TypeValidationBDD)
    Dim info As String
    
    info = "ÉQUIPEMENT RECONNU" & vbCrLf & vbCrLf
    info = info & "INFORMATIONS :" & vbCrLf
    info = info & "• Numéro de série : " & resultats.numeroSerie & vbCrLf
    info = info & "• Référence produit : " & resultats.codeArticle & vbCrLf
    info = info & "• Modèle : " & resultats.modeleArticle & vbCrLf
    info = info & vbCrLf & "Vous pouvez maintenant ouvrir la fiche SAV"
    
    Me.Controls("lblInfo").Caption = info
    Me.Controls("lblInfo").BackColor = RGB(212, 237, 218)
    
    Me.Controls("cmdOuvrirFiche").Enabled = True
    Me.Controls("cmdOuvrirFiche").BackColor = ROUGE_REDBULL
    
    Me.Controls("lblStatut").Caption = "Équipement reconnu : " & resultats.numeroSerie & " | " & ObtenirDateTimeFormatee()
End Sub

Private Sub AfficherErreurValidation(messageErreur As String, numeroSerie As String)
    Dim info As String
    
    info = "ÉQUIPEMENT NON RECONNU" & vbCrLf & vbCrLf
    info = info & "Numéro saisi : " & numeroSerie & vbCrLf & vbCrLf
    info = info & "VÉRIFICATIONS :" & vbCrLf
    info = info & "• Le numéro de série est-il correct ?" & vbCrLf
    info = info & "• S'agit-il bien d'un équipement Red Bull ?" & vbCrLf & vbCrLf
    info = info & "SOLUTIONS :" & vbCrLf
    info = info & "• Vérifiez la saisie" & vbCrLf
    info = info & "• Rescannez le code-barres" & vbCrLf
    info = info & "• Contactez le support si le problème persiste"
    
    Me.Controls("lblInfo").Caption = info
    Me.Controls("lblInfo").BackColor = RGB(248, 215, 218)
    
    Me.Controls("cmdOuvrirFiche").Enabled = False
    Me.Controls("cmdOuvrirFiche").BackColor = RGB(150, 150, 150)
    
    EcrireHistoriqueScan numeroSerie, "ÉQUIPEMENT NON RECONNU"
    
    Me.Controls("lblStatut").Caption = "Équipement non reconnu | " & ObtenirDateTimeFormatee()
End Sub

Private Sub AfficherErreurConnexion(numeroSerie As String)
    Dim info As String
    
    info = "PROBLÈME DE CONNEXION" & vbCrLf & vbCrLf
    info = info & "La vérification de l'équipement n'est pas possible actuellement." & vbCrLf & vbCrLf
    info = info & "ACTIONS POSSIBLES :" & vbCrLf
    info = info & "• Vérifiez votre connexion réseau" & vbCrLf
    info = info & "• Redémarrez l'application" & vbCrLf
    info = info & "• Contactez le support technique" & vbCrLf & vbCrLf
    info = info & "Mode dégradé : fonctions limitées"
    
    Me.Controls("lblInfo").Caption = info
    Me.Controls("lblInfo").BackColor = RGB(255, 243, 205)
    
    Me.Controls("lblStatut").Caption = "Problème de connexion | " & ObtenirDateTimeFormatee()
End Sub

' === FONCTIONS UTILITAIRES ===

Private Function ValiderFormatNumeroSerie(numeroSerie As String) As Boolean
    ValiderFormatNumeroSerie = ValiderFormatNumeroSerieRB(numeroSerie)
End Function

Private Sub ActualiserStatutBDD()
    Dim ctrl As Object
    Set ctrl = Me.Controls("lblStatutBDD")
    
    If VerifierConnexionBDD() Then
        ctrl.Caption = "SYSTÈME CONNECTÉ - PRÊT À UTILISER"
        ctrl.BackColor = RGB(212, 237, 218)
    Else
        ctrl.Caption = "PROBLÈME DE CONNEXION - FONCTIONS LIMITÉES"
        ctrl.BackColor = RGB(248, 215, 218)
    End If
End Sub

Private Sub DemarrerTimerBDD()
    Set tmrVerifBDD = Me.Controls.Add("VB.Timer", "tmrVerifBDD")
    tmrVerifBDD.Interval = 30000 ' 30 secondes
    tmrVerifBDD.Enabled = True
End Sub

Private Sub tmrVerifBDD_Timer()
    ActualiserStatutBDD
    
    If Not VerifierConnexionBDD() Then
        ConnecterBDD
        ActualiserStatutBDD
    End If
End Sub

Private Sub ResetFormulaire()
    Me.Controls("txtReference").Text = ""
    Me.Controls("lblInfo").Caption = "FICHE TRAITÉE AVEC SUCCÈS" & vbCrLf & vbCrLf & _
                                     "Vous pouvez scanner un nouvel équipement Red Bull" & vbCrLf & _
                                     "Le système vérifiera automatiquement sa validité" & vbCrLf & vbCrLf & _
                                     "Appuyez sur F1 pour l'aide"
    Me.Controls("lblInfo").BackColor = RGB(248, 249, 250)
    Me.Controls("cmdOuvrirFiche").Enabled = False
    Me.Controls("cmdOuvrirFiche").BackColor = RGB(150, 150, 150)
    referenceValidee = ""
    numeroSerieValide = ""
    
    Me.Controls("txtReference").SetFocus
    Me.Controls("lblStatut").Caption = "Prêt - Système de validation actif | " & ObtenirDateTimeFormatee()
End Sub

Private Sub AfficherAideComplete()
    Dim aide As String
    
    aide = "=== AIDE SAV RED BULL ===" & vbCrLf & vbCrLf
    aide = aide & "OBJECTIF :" & vbCrLf
    aide = aide & "Scanner et valider les numéros de série des frigos Red Bull" & vbCrLf
    aide = aide & "pour créer des fiches retour SAV." & vbCrLf & vbCrLf
    
    aide = aide & "UTILISATION :" & vbCrLf
    aide = aide & "1. Scannez ou saisissez le numéro de série" & vbCrLf
    aide = aide & "2. Cliquez SCANNER (ou appuyez Entrée)" & vbCrLf
    aide = aide & "3. Si validé, cliquez OUVRIR FICHE RETOUR" & vbCrLf & vbCrLf
    
    aide = aide & "RACCOURCIS CLAVIER :" & vbCrLf
    aide = aide & "F1 = Cette aide | Entrée = Scanner | Échap = Reset" & vbCrLf & vbCrLf
    
    aide = aide & "SUPPORT :" & vbCrLf
    aide = aide & "En cas de problème, contactez votre superviseur."
    
    MsgBox aide, vbInformation, "Aide SAV Red Bull"
End Sub

' === GESTION DE LA FERMETURE ===

Private Sub Form_Unload(Cancel As Integer)
    If Not tmrVerifBDD Is Nothing Then
        tmrVerifBDD.Enabled = False
        Set tmrVerifBDD = Nothing
    End If
    
    FermerBDD
    
    Me.Controls("lblStatut").Caption = "Sauvegarde en cours... | " & ObtenirDateTimeFormatee()
    Me.Refresh
    SauvegardeAutomatique
    
    EcrireHistoriqueScan "SYSTEM", "Application fermée"
    
    Set cmdValider = Nothing
    Set cmdOuvrirFiche = Nothing
    Set tmrVerifBDD = Nothing
End Sub


