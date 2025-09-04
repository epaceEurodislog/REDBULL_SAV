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
' === FORM1.FRM - FORMULAIRE PRINCIPAL SAV RED BULL AVEC BDD COMPLÈTE ===

Option Explicit

Private referenceValidee As String
Private numeroSerieValide As String
Private WithEvents cmdValider As CommandButton
Attribute cmdValider.VB_VarHelpID = -1
Private WithEvents cmdOuvrirFiche As CommandButton
Attribute cmdOuvrirFiche.VB_VarHelpID = -1
Private WithEvents cmdTestBDD As CommandButton
Attribute cmdTestBDD.VB_VarHelpID = -1
Private WithEvents tmrVerifBDD As Timer
Attribute tmrVerifBDD.VB_VarHelpID = -1

Private Sub Form_Load()
    Me.Caption = "SAV Red Bull"
    Me.Width = 8000
    Me.Height = 6500
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
    
    ' Label titre principal
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitre")
    ctrl.Left = 600
    ctrl.Top = 200
    ctrl.Width = 6800
    ctrl.Height = 500
    ctrl.Caption = "SAV RED BULL"
    ctrl.BackColor = RGB(220, 20, 60) ' Rouge Red Bull
    'ctrl.ForeColor = RGB(255, 255, 255)
    ctrl.Alignment = 2
    ctrl.Font.Size = 18
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    ' Indicateur de statut BDD
    Set ctrl = Me.Controls.Add("VB.Label", "lblStatutBDD")
    ctrl.Left = 600
    ctrl.Top = 750
    ctrl.Width = 6800
    ctrl.Height = 250
    ActualiserStatutBDD ' Sera défini plus bas
    ctrl.Alignment = 2
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    ' Label "Numéro de série"
    Set ctrl = Me.Controls.Add("VB.Label", "lblRef")
    ctrl.Left = 600
    ctrl.Top = 1100
    ctrl.Width = 1800
    ctrl.Caption = "Numéro de série frigo:"
    ctrl.Font.Size = 10
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    ' TextBox pour saisie du numéro de série
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtReference")
    ctrl.Left = 2500
    ctrl.Top = 1100
    ctrl.Width = 2800
    ctrl.Height = 350
    ctrl.Font.Size = 12
    ctrl.Visible = True
    
    ' Bouton Valider
    Set cmdValider = Me.Controls.Add("VB.CommandButton", "cmdValider")
    cmdValider.Left = 5400
    cmdValider.Top = 1100
    cmdValider.Width = 1000
    cmdValider.Height = 350
    cmdValider.Caption = "SCANNER"
    cmdValider.Font.Bold = True
    cmdValider.BackColor = RGB(0, 123, 255)
    'cmdValider.ForeColor = RGB(255, 255, 255)
    cmdValider.Visible = True
    
    ' Zone d'information détaillée
    Set ctrl = Me.Controls.Add("VB.Label", "lblInfo")
    ctrl.Left = 600
    ctrl.Top = 1550
    ctrl.Width = 6800
    ctrl.Height = 1800
    ctrl.Caption = "INSTRUCTIONS - VALIDATION FILTRÉE :" & vbCrLf & vbCrLf & _
                   "1. Scannez ou saisissez le numéro de série du frigo Red Bull" & vbCrLf & _
                   "2. Cliquez SCANNER pour vérifier dans la liste des 92 codes autorisés" & vbCrLf & _
                   "3. Si validé, vous pourrez ouvrir la fiche retour SAV" & vbCrLf & vbCrLf & _
                   "SÉCURITÉ - CODES FILTRÉS :" & vbCrLf & _
                   "• Seulement 92 codes articles Red Bull autorisés" & vbCrLf & _
                   "• Validation stricte : INNER JOIN + DISTINCT + filtres" & vbCrLf & _
                   "• Élimination des valeurs NULL et vides" & vbCrLf & _
                   "• Seuls les équipements de la liste peuvent être traités" & vbCrLf & vbCrLf & _
                   "BASE : ART_PAR + NSE_DAT avec act_code = 'RB'"
    ctrl.BackColor = RGB(248, 249, 250)
    ctrl.BorderStyle = 1
    ctrl.Alignment = 0 ' Alignement à gauche
    ctrl.Font.Size = 9
    ctrl.Visible = True
    
    ' Bouton Ouvrir Fiche (désactivé au début)
    Set cmdOuvrirFiche = Me.Controls.Add("VB.CommandButton", "cmdOuvrirFiche")
    cmdOuvrirFiche.Left = 2000
    cmdOuvrirFiche.Top = 3500
    cmdOuvrirFiche.Width = 2800
    cmdOuvrirFiche.Height = 450
    cmdOuvrirFiche.Caption = "?? OUVRIR FICHE RETOUR SAV"
    cmdOuvrirFiche.Enabled = False
    cmdOuvrirFiche.BackColor = RGB(150, 150, 150)
    cmdOuvrirFiche.Font.Bold = True
    cmdOuvrirFiche.Font.Size = 11
    cmdOuvrirFiche.Visible = True
    
    ' Bouton Test BDD
    Set cmdTestBDD = Me.Controls.Add("VB.CommandButton", "cmdTestBDD")
    cmdTestBDD.Left = 600
    cmdTestBDD.Top = 4100
    cmdTestBDD.Width = 1800
    cmdTestBDD.Height = 350
    cmdTestBDD.Caption = "?? TESTER BDD"
    cmdTestBDD.BackColor = RGB(108, 117, 125)
    'cmdTestBDD.ForeColor = RGB(255, 255, 255)
    cmdTestBDD.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdTestFiltrage")
    ctrl.Left = 2500
    ctrl.Top = 4100
    ctrl.Width = 1800
    ctrl.Height = 350
    ctrl.Caption = "?? TEST FILTRAGE"
    ctrl.BackColor = RGB(255, 193, 7)
    ctrl.Visible = True
    
    ' Bouton Historique
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdHistorique")
    ctrl.Left = 4400
    ctrl.Top = 4100
    ctrl.Width = 1800
    ctrl.Height = 350
    ctrl.Caption = "?? HISTORIQUE"
    ctrl.BackColor = RGB(40, 167, 69)
    'ctrl.ForeColor = RGB(255, 255, 255)
    ctrl.Visible = True
    
    ' Zone de statut en bas
    Set ctrl = Me.Controls.Add("VB.Label", "lblStatut")
    ctrl.Left = 600
    ctrl.Top = 4600
    ctrl.Width = 6800
    ctrl.Height = 300
    ctrl.Caption = "Prêt - En attente de scan | " & ObtenirDateTimeFormatee()
    ctrl.BackColor = RGB(220, 220, 220)
    ctrl.BorderStyle = 1
    ctrl.Alignment = 2
    ctrl.Font.Size = 8
    ctrl.Visible = True
End Sub

Private Sub cmdValider_Click()
    Dim numeroSerie As String
    numeroSerie = Trim(UCase(Me.Controls("txtReference").Text))
    
    Me.Controls("lblStatut").Caption = "Validation filtrée en cours... | " & ObtenirDateTimeFormatee()
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
        ' Validation réussie dans la liste filtrée
        referenceValidee = numeroSerie
        numeroSerieValide = numeroSerie
        AfficherValidationReussie resultatsValidation
        
        ' Enregistrer dans l'historique avec mention du filtrage
        EcrireHistoriqueScan numeroSerie, resultatsValidation.modeleArticle & " [FILTRÉ-92]"
    Else
        ' Numéro de série non autorisé dans la liste filtrée
        AfficherErreurValidation resultatsValidation.statut, numeroSerie
    End If
End Sub


' Fonction pour afficher les informations système détaillées
Private Sub AfficherInfosSystemeDetaillees()
    Dim infos As String
    
    infos = ObtenirInfosSysteme() & vbCrLf
    
    ' Ajouter des statistiques de session
    infos = infos & vbCrLf & "=== STATISTIQUES SESSION ===" & vbCrLf
    infos = infos & "Référence validée actuelle: " & IIf(Len(referenceValidee) > 0, referenceValidee, "Aucune") & vbCrLf
    infos = infos & "Numéro de série validé: " & IIf(Len(numeroSerieValide) > 0, numeroSerieValide, "Aucun") & vbCrLf
    
    ' Ajouter l'état des fichiers
    infos = infos & vbCrLf & "=== ÉTAT DES FICHIERS ===" & vbCrLf
    infos = infos & "Historique: " & IIf(Dir(App.Path & FICHIER_HISTORIQUE) <> "", "? Présent", "? Absent") & vbCrLf
    infos = infos & "Stock pièces: " & IIf(Dir(App.Path & FICHIER_STOCK_PIECES) <> "", "? Présent", "? Absent") & vbCrLf
    infos = infos & "Stock réparable: " & IIf(Dir(App.Path & FICHIER_STOCK_REPARABLE) <> "", "? Présent", "? Absent") & vbCrLf
    
    MsgBox infos, vbInformation, "Informations Système Détaillées"
End Sub

' === FONCTIONS DE COMMUNICATION AVEC SCANNER EXTERNE ===

' Fonction pour traiter les données reçues du port série (scanner code-barres)
Private Sub MSComm1_OnComm()
    On Error GoTo ErrorHandler
    
    Dim donnees As String
    
    Select Case MSComm1.CommEvent
        Case comEvReceive
            donnees = MSComm1.Input
            If Len(donnees) > 0 Then
                ' Nettoyer les données reçues du scanner
                donnees = Replace(donnees, vbCr, "")
                donnees = Replace(donnees, vbLf, "")
                donnees = Replace(donnees, Chr(0), "") ' Supprimer les caractères null
                donnees = Trim(UCase(donnees))
                
                If Len(donnees) >= 8 Then
                    ' Traitement automatique du scan
                    TraiterScanAutomatique donnees
                End If
            End If
            
        Case comEvError
            MsgBox "Erreur de communication avec le scanner !", vbExclamation
            Me.Controls("lblStatut").Caption = "? Erreur scanner | " & ObtenirDateTimeFormatee()
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur lors du traitement du scan : " & Err.description, vbCritical
End Sub

' Fonction pour traiter automatiquement un scan
Private Sub TraiterScanAutomatique(numeroSerie As String)
    ' Remplir automatiquement la zone de texte
    Me.Controls("txtReference").Text = numeroSerie
    
    ' Ajouter un effet visuel pour indiquer le scan automatique
    Me.Controls("txtReference").BackColor = RGB(255, 255, 0) ' Jaune temporaire
    Me.Refresh
    
    ' Petite pause pour l'effet visuel
    Sleep 300 ' Nécessite API Windows
    
    Me.Controls("txtReference").BackColor = RGB(255, 255, 255) ' Retour au blanc
    
    ' Déclencher automatiquement la validation
    Me.Controls("lblStatut").Caption = "?? Scan automatique détecté... | " & ObtenirDateTimeFormatee()
    cmdValider_Click
End Sub

' === FONCTIONS AVANCÉES DE VALIDATION BDD ===

' Fonction pour vérifier les dépendances et contraintes
Private Function VerifierContraintesBDD(numeroSerie As String) As String
    On Error GoTo ErrorHandler
    
    Dim sql As String
    Dim rsContraintes As ADODB.Recordset
    Dim resultats As String
    
    ' Vérifier s'il y a des fiches SAV existantes pour ce numéro
    sql = "SELECT COUNT(*) as nb_fiches FROM sav_historique WHERE nse_nums = '" & numeroSerie & "'"
    Set rsContraintes = ExecuterRequete(sql)
    
    If Not rsContraintes Is Nothing And Not rsContraintes.EOF Then
        If rsContraintes!nb_fiches > 0 Then
            resultats = resultats & "?? " & rsContraintes!nb_fiches & " fiche(s) SAV existante(s)" & vbCrLf
        End If
        rsContraintes.Close
    End If
    
    ' Vérifier le statut de l'équipement
    sql = "SELECT nse_statut FROM nse_dat WHERE nse_nums = '" & numeroSerie & "'"
    Set rsContraintes = ExecuterRequete(sql)
    
    If Not rsContraintes Is Nothing And Not rsContraintes.EOF Then
        If Not IsNull(rsContraintes!nse_statut) Then
            resultats = resultats & "?? Statut actuel: " & rsContraintes!nse_statut & vbCrLf
        End If
        rsContraintes.Close
    End If
    
    Set rsContraintes = Nothing
    VerifierContraintesBDD = resultats
    Exit Function
    
ErrorHandler:
    VerifierContraintesBDD = "Erreur vérification contraintes: " & Err.description
End Function

' === GESTION DES ÉVÉNEMENTS DE FERMETURE ===

' Événement de fermeture du formulaire avec nettoyage complet
Private Sub Form_Unload(Cancel As Integer)
    ' Arrêter le timer
    If Not tmrVerifBDD Is Nothing Then
        tmrVerifBDD.Enabled = False
        Set tmrVerifBDD = Nothing
    End If
    
    ' Fermer proprement la connexion BDD
    FermerBDD
    
    ' Créer une sauvegarde automatique avant fermeture
    Me.Controls("lblStatut").Caption = "Sauvegarde en cours... | " & ObtenirDateTimeFormatee()
    Me.Refresh
    SauvegardeAutomatique
    
    ' Log de fermeture
    EcrireHistoriqueScan "SYSTEM", "Application fermée"
End Sub

' === FONCTIONS D'INTERFACE UTILISATEUR AVANCÉES ===

' Fonction pour gérer le clic droit (menu contextuel)
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then ' Clic droit
        AfficherMenuContextuel
    End If
End Sub

' Fonction pour afficher un menu contextuel avancé
Private Sub AfficherMenuContextuel()
    Dim reponse As Integer
    Dim menu As String
    
    menu = "?? MENU CONTEXTUEL SAV RED BULL" & vbCrLf & vbCrLf
    menu = menu & "Choisissez une action :" & vbCrLf
    menu = menu & "• OUI = ?? Voir l'historique des scans" & vbCrLf
    menu = menu & "• NON = ?? Informations système détaillées" & vbCrLf
    menu = menu & "• ANNULER = ?? Effectuer une maintenance rapide"
    
    reponse = MsgBox(menu, vbYesNoCancel + vbQuestion, "Menu Actions")
    
    Select Case reponse
        Case vbYes
            cmdHistorique_Click
        Case vbNo
            AfficherInfosSystemeDetaillees
        Case vbCancel
            cmdMaintenance_Click
    End Select
End Sub

' === GESTION DES ERREURS ET LOGGING ===

' Fonction pour logger les erreurs système
Private Sub LoggerErreur(source As String, description As String)
    On Error Resume Next
    
    Dim fichierLog As String
    Dim numeroFichier As Integer
    
    fichierLog = App.Path & "\Logs\Erreurs_" & Format(Now, "yyyymmdd") & ".txt"
    
    ' Créer le répertoire Logs s'il n'existe pas
    If Dir(App.Path & "\Logs", vbDirectory) = "" Then
        MkDir App.Path & "\Logs"
    End If
    
    numeroFichier = FreeFile
    Open fichierLog For Append As #numeroFichier
    Print #numeroFichier, Format(Now, "dd/mm/yyyy hh:nn:ss") & " - [" & source & "] " & description
    Close #numeroFichier
End Sub

' === FONCTIONS DE DIAGNOSTIC RÉSEAU ===

' Fonction pour tester la connectivité réseau vers le serveur BDD
Private Function TesterConnectiviteReseau() As Boolean
    On Error GoTo ErrorHandler
    
    ' Test simple de ping (simulation)
    ' Dans un environnement réel, vous pourriez utiliser une API Windows pour ping
    TesterConnectiviteReseau = True
    Exit Function
    
ErrorHandler:
    TesterConnectiviteReseau = False
End Function

' === SAUVEGARDE ET RÉCUPÉRATION D'ÉTAT ===

' Fonction pour sauvegarder l'état actuel de la session
Private Sub SauvegarderEtatSession()
    On Error Resume Next
    
    Dim fichierEtat As String
    Dim numeroFichier As Integer
    
    fichierEtat = App.Path & "\Session_" & Format(Now, "yyyymmdd") & ".tmp"
    numeroFichier = FreeFile
    
    Open fichierEtat For Output As #numeroFichier
    Print #numeroFichier, "DERNIERE_REFERENCE=" & referenceValidee
    Print #numeroFichier, "DERNIER_NUMERO_SERIE=" & numeroSerieValide
    Print #numeroFichier, "TIMESTAMP=" & Format(Now, "dd/mm/yyyy hh:nn:ss")
    Print #numeroFichier, "STATUT_BDD=" & IIf(VerifierConnexionBDD(), "CONNECTE", "DECONNECTE")
    Close #numeroFichier
End Sub

Private Sub cmdTestFiltrage_Click()
    Me.Controls("lblStatut").Caption = "Test du filtrage 92 codes... | " & ObtenirDateTimeFormatee()
    Me.Refresh
    
    ' Appeler la nouvelle fonction de test
    TesterRequeteFiltree92Codes
    
    Me.Controls("lblStatut").Caption = "Test filtrage terminé | " & ObtenirDateTimeFormatee()
End Sub

' Fonction pour récupérer l'état de la dernière session
Private Sub RecupererEtatSession()
    On Error GoTo ErrorHandler
    
    Dim fichierEtat As String
    Dim numeroFichier As Integer
    Dim ligne As String
    
    fichierEtat = App.Path & "\Session_" & Format(Now, "yyyymmdd") & ".tmp"
    
    If Dir(fichierEtat) <> "" Then
        numeroFichier = FreeFile
        Open fichierEtat For Input As #numeroFichier
        
        Do While Not EOF(numeroFichier)
            Line Input #numeroFichier, ligne
            
            If InStr(ligne, "DERNIERE_REFERENCE=") > 0 Then
                ' Récupérer la dernière référence si besoin
            ElseIf InStr(ligne, "DERNIER_NUMERO_SERIE=") > 0 Then
                ' Récupérer le dernier numéro de série si besoin
            End If
        Loop
        
        Close #numeroFichier
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Erreur non critique, continuer normalement
End Sub

' Sauvegarde automatique de l'état toutes les 5 minutes
Private Sub tmrSauvegardeAuto_Timer()
    SauvegarderEtatSession
End Sub
' Fonction pour valider le format du numéro de série (appel au module)
Private Function ValiderFormatNumeroSerie(numeroSerie As String) As Boolean
    ValiderFormatNumeroSerie = ValiderFormatNumeroSerieRB(numeroSerie)
End Function

' Fonction pour obtenir des informations complémentaires (appel au module)
Private Function ObtenirInfosComplementaires(codeArticle As String) As String
    ObtenirInfosComplementaires = ObtenirInfosComplementairesArticle(codeArticle)
End Function

' Fonction pour afficher une validation réussie
Private Sub AfficherValidationReussie(resultats As TypeValidationBDD)
    Dim info As String
    
    info = "? NUMÉRO DE SÉRIE VALIDÉ - LISTE AUTORISÉE" & vbCrLf & vbCrLf
    info = info & "?? DÉTAILS PRODUIT:" & vbCrLf
    info = info & "• Numéro de série: " & resultats.numeroSerie & vbCrLf
    info = info & "• Code article: " & resultats.codeArticle & vbCrLf
    info = info & "• Modèle: " & resultats.modeleArticle & vbCrLf
    info = info & "• Date création: " & resultats.dateCreation & vbCrLf
    
    If resultats.prixCatalogue > 0 Then
        info = info & "• Prix catalogue: " & Format(resultats.prixCatalogue, "0.00") & "€" & vbCrLf
    End If
    
    info = info & vbCrLf & "?? VALIDATION SÉCURISÉE:" & vbCrLf
    info = info & "• Serveur: " & SERVER_NAME & vbCrLf
    info = info & "• Base: " & DATABASE_NAME & vbCrLf
    info = info & "• Statut: " & resultats.statut & vbCrLf
    info = info & "• Liste: 92 codes articles Red Bull autorisés" & vbCrLf
    
    If Len(resultats.informationsComplementaires) > 0 Then
        info = info & "• " & resultats.informationsComplementaires & vbCrLf
    End If
    
    info = info & vbCrLf & "? Vous pouvez maintenant ouvrir la fiche retour SAV"
    
    Me.Controls("lblInfo").Caption = info
    Me.Controls("lblInfo").BackColor = RGB(212, 237, 218) ' Vert clair
    
    ' Activer le bouton fiche retour
    Me.Controls("cmdOuvrirFiche").Enabled = True
    Me.Controls("cmdOuvrirFiche").BackColor = RGB(40, 167, 69) ' Vert
    
    Me.Controls("lblStatut").Caption = "? Validé (liste 92 codes): " & resultats.numeroSerie & " | " & ObtenirDateTimeFormatee()
End Sub

' Fonction pour afficher une erreur de validation
Private Sub AfficherErreurValidation(messageErreur As String, numeroSerie As String)
    Dim info As String
    
    info = "? VALIDATION ÉCHOUÉE - HORS LISTE AUTORISÉE" & vbCrLf & vbCrLf
    info = info & "?? ERREUR: " & messageErreur & vbCrLf
    info = info & "?? Numéro saisi: " & numeroSerie & vbCrLf & vbCrLf
    
    info = info & "?? VÉRIFICATIONS EFFECTUÉES:" & vbCrLf
    info = info & "• Format du numéro de série" & vbCrLf
    info = info & "• Existence dans les 92 codes autorisés" & vbCrLf
    info = info & "• Validation act_code = 'RB'" & vbCrLf
    info = info & "• Filtrage DISTINCT avec conditions strictes" & vbCrLf & vbCrLf
    
    info = info & "?? CONSEILS:" & vbCrLf
    info = info & "• Vérifiez la saisie du numéro de série" & vbCrLf
    info = info & "• Ce numéro doit appartenir aux 92 codes autorisés" & vbCrLf
    info = info & "• Contactez l'administrateur si l'équipement devrait être autorisé"
    
    Me.Controls("lblInfo").Caption = info
    Me.Controls("lblInfo").BackColor = RGB(248, 215, 218) ' Rouge clair
    
    ' Désactiver le bouton fiche
    Me.Controls("cmdOuvrirFiche").Enabled = False
    Me.Controls("cmdOuvrirFiche").BackColor = RGB(150, 150, 150)
    
    ' Enregistrer l'erreur dans l'historique avec mention du filtrage
    EcrireHistoriqueScan numeroSerie, "ERREUR FILTRAGE: " & messageErreur
    
    Me.Controls("lblStatut").Caption = "? Hors liste autorisée | " & ObtenirDateTimeFormatee()
End Sub

' Fonction pour afficher une erreur de connexion
Private Sub AfficherErreurConnexion(numeroSerie As String)
    Dim info As String
    
    info = "?? PROBLÈME DE CONNEXION BASE DE DONNÉES" & vbCrLf & vbCrLf
    info = info & "?? STATUT: Connexion à la base de données interrompue" & vbCrLf
    info = info & "??? Serveur: " & SERVER_NAME & vbCrLf
    info = info & "??? Base: " & DATABASE_NAME & vbCrLf & vbCrLf
    
    info = info & "? ACTIONS DISPONIBLES:" & vbCrLf
    info = info & "• Cliquez 'TESTER BDD' pour retenter la connexion" & vbCrLf
    info = info & "• Vérifiez votre connexion réseau" & vbCrLf
    info = info & "• Contactez l'administrateur système" & vbCrLf & vbCrLf
    
    info = info & "?? Mode dégradé activé - fonctions limitées"
    
    Me.Controls("lblInfo").Caption = info
    Me.Controls("lblInfo").BackColor = RGB(255, 243, 205) ' Orange clair
    
    Me.Controls("lblStatut").Caption = "?? Connexion BDD perdue | " & ObtenirDateTimeFormatee()
End Sub

Private Sub cmdOuvrirFiche_Click()
    If Len(referenceValidee) = 0 Or Len(numeroSerieValide) = 0 Then
        MsgBox "Aucun numéro de série validé !", vbExclamation
        Exit Sub
    End If
    
    Me.Controls("lblStatut").Caption = "Ouverture fiche retour... | " & ObtenirDateTimeFormatee()
    Me.Refresh
    
    ' Ouvrir la fiche retour avec toutes les données validées
    Load frmFicheRetour
    frmFicheRetour.InitialiserAvecReference referenceValidee, numeroSerieValide
    frmFicheRetour.Show vbModal
    
    ' Reset après fermeture de la fiche
    ResetFormulaire
End Sub

Private Sub cmdTestBDD_Click()
    Me.Controls("lblStatut").Caption = "Test système en cours... | " & ObtenirDateTimeFormatee()
    Me.Refresh
    
    ' Tester complètement le système
    TesterSystemeComplet
    
    ' Mettre à jour l'indicateur de statut BDD
    ActualiserStatutBDD
    
    Me.Controls("lblStatut").Caption = "Test terminé | " & ObtenirDateTimeFormatee()
End Sub

' Fonction pour actualiser le statut BDD dans l'interface
Private Sub ActualiserStatutBDD()
    Dim ctrl As Object
    Set ctrl = Me.Controls("lblStatutBDD")
    
    If VerifierConnexionBDD() Then
        ctrl.Caption = "? BASE CONNECTÉE - FILTRAGE 92 CODES ACTIF - " & SERVER_NAME & " (" & DATABASE_NAME & ")"
        ctrl.BackColor = RGB(212, 237, 218) ' Vert clair
    Else
        ctrl.Caption = "? BASE DÉCONNECTÉE - MODE DÉGRADÉ"
        ctrl.BackColor = RGB(248, 215, 218) ' Rouge clair
    End If
End Sub

' Fonction pour démarrer le timer de vérification BDD
Private Sub DemarrerTimerBDD()
    Set tmrVerifBDD = Me.Controls.Add("VB.Timer", "tmrVerifBDD")
    tmrVerifBDD.Interval = 30000 ' 30 secondes
    tmrVerifBDD.Enabled = True
End Sub

' Timer pour vérification périodique de la connexion BDD
Private Sub tmrVerifBDD_Timer()
    ' Vérifier la connexion et actualiser le statut
    ActualiserStatutBDD
    
    ' Si déconnecté, essayer de reconnecter automatiquement
    If Not VerifierConnexionBDD() Then
        ConnecterBDD
        ActualiserStatutBDD
    End If
End Sub

' Fonction pour remettre à zéro le formulaire
Private Sub ResetFormulaire()
    Me.Controls("txtReference").Text = ""
    Me.Controls("lblInfo").Caption = "? FICHE TRAITÉE AVEC SUCCÈS" & vbCrLf & vbCrLf & _
                                     "?? Vous pouvez scanner un nouveau frigo Red Bull" & vbCrLf & _
                                     "autorisé dans la liste des 92 codes" & vbCrLf & vbCrLf & _
                                     "? Appuyez sur F1 pour l'aide complète"
    Me.Controls("lblInfo").BackColor = RGB(248, 249, 250)
    Me.Controls("cmdOuvrirFiche").Enabled = False
    Me.Controls("cmdOuvrirFiche").BackColor = RGB(150, 150, 150)
    referenceValidee = ""
    numeroSerieValide = ""
    
    ' Donner le focus à la zone de saisie
    Me.Controls("txtReference").SetFocus
    
    Me.Controls("lblStatut").Caption = "Prêt - Filtrage 92 codes actif | " & ObtenirDateTimeFormatee()
End Sub

' Gestion de la touche Entrée dans la zone de texte
Private Sub txtReference_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ' Touche Entrée
        KeyAscii = 0 ' Annuler le bip
        cmdValider_Click ' Déclencher la validation
    End If
End Sub

' Gestion des raccourcis clavier globaux
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 ' F1 = Aide complète
            AfficherAideComplete
            
        Case vbKeyF2 ' F2 = Test BDD
            cmdTestBDD_Click
            
        Case vbKeyF3 ' F3 = Historique
            cmdHistorique_Click
            
        Case vbKeyF5 ' F5 = Maintenance
            cmdMaintenance_Click
            
        Case vbKeyF12 ' F12 = Infos système détaillées
            AfficherInfosSystemeDetaillees
            
        Case vbKeyEscape ' Échap = Reset
            ResetFormulaire
    End Select
End Sub

' Fonction pour afficher l'aide complète
Private Sub AfficherAideComplete()
    Dim aide As String
    
    aide = "=== AIDE SAV RED BULL SCANNER PRO v2.1 ===" & vbCrLf & vbCrLf
    aide = aide & "?? OBJECTIF:" & vbCrLf
    aide = aide & "Scanner et valider les numéros de série des frigos Red Bull" & vbCrLf
    aide = aide & "pour créer des fiches retour SAV." & vbCrLf & vbCrLf
    
    aide = aide & "?? UTILISATION:" & vbCrLf
    aide = aide & "1. Scannez ou saisissez le numéro de série" & vbCrLf
    aide = aide & "2. Cliquez SCANNER (ou appuyez Entrée)" & vbCrLf
    aide = aide & "3. Si validé, cliquez OUVRIR FICHE RETOUR" & vbCrLf & vbCrLf
    
    aide = aide & "??? BASE DE DONNÉES:" & vbCrLf
    aide = aide & "• Table NSE_DAT: Numéros de série" & vbCrLf
    aide = aide & "• Table ART_PAR: Articles et modèles" & vbCrLf
    aide = aide & "• Filtre ACT_CODE = 'RB' (Red Bull uniquement)" & vbCrLf & vbCrLf
    
    aide = aide & "?? RACCOURCIS CLAVIER:" & vbCrLf
    aide = aide & "F1 = Cette aide | F2 = Test BDD | F3 = Historique" & vbCrLf
    aide = aide & "F5 = Maintenance | F12 = Infos système | Échap = Reset" & vbCrLf & vbCrLf
    
    aide = aide & "?? MAINTENANCE:" & vbCrLf
    aide = aide & "Le système effectue automatiquement:" & vbCrLf
    aide = aide & "• Vérification connexion BDD toutes les 30s" & vbCrLf
    aide = aide & "• Sauvegarde automatique à la fermeture" & vbCrLf
    aide = aide & "• Nettoyage des fichiers temporaires"
    
    MsgBox aide, vbInformation, "Aide SAV Red Bull Scanner Pro"
End Sub

' Événements pour les autres boutons
Private Sub cmdMaintenance_Click()
    Dim reponse As Integer
    reponse = MsgBox("Lancer la maintenance complète du système ?" & vbCrLf & vbCrLf & _
                     "Cette opération va :" & vbCrLf & _
                     "• Vérifier l'intégrité des fichiers" & vbCrLf & _
                     "• Nettoyer les fichiers temporaires" & vbCrLf & _
                     "• Tester la connexion BDD" & vbCrLf & _
                     "• Créer une sauvegarde complète" & vbCrLf & _
                     "• Synchroniser les données", vbYesNo + vbQuestion, "Maintenance")
    
    If reponse = vbYes Then
        MaintenanceRapide
    End If
End Sub

Private Sub cmdHistorique_Click()
    AfficherHistoriqueDetaille
End Sub

' Fonction pour afficher l'historique détaillé
Private Sub AfficherHistoriqueDetaille()
    Dim historique As String
    historique = LireHistoriqueScan()
    
    If Len(historique) = 0 Then
        MsgBox "Aucun historique disponible pour le moment.", vbInformation, "Historique"
        Exit Sub
    End If
    
    ' Compter les entrées
    Dim lignes() As String
    lignes = Split(historique, vbCrLf)
    Dim nbEntrees As Integer
    nbEntrees = 0
    
    ' Compter les lignes non vides
    Dim i As Integer
    For i = 0 To UBound(lignes)
        If Len(Trim(lignes(i))) > 0 Then
            nbEntrees = nbEntrees + 1
        End If
    Next i
    
    Dim titre As String
    titre = "Historique des scans (" & nbEntrees & " entrées)"
    
    ' Limiter l'affichage pour éviter les messages trop longs
    If Len(historique) > 2000 Then
        historique = Left(historique, 2000) & vbCrLf & vbCrLf & "... [Historique tronqué - voir fichier complet]"
    End If
    
    ' Afficher dans une MessageBox pour simplicité
    ' Dans une version avancée, vous pourriez créer une form dédiée
    MsgBox titre & vbCrLf & String(50, "=") & vbCrLf & vbCrLf & historique, vbInformation, titre
End Sub

' Fonction de nettoyage final
Private Sub NettoyageFinal()
    On Error Resume Next
    
    ' Sauvegarder l'état final
    SauvegarderEtatSession
    
    ' Logger la fermeture
    EcrireHistoriqueScan "SYSTEM", "Application fermée normalement"
    
    ' Libérer les ressources
    Set cmdValider = Nothing
    Set cmdOuvrirFiche = Nothing
    Set cmdTestBDD = Nothing
    Set tmrVerifBDD = Nothing
End Sub
