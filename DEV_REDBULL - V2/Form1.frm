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
' === FORM1.FRM - FORMULAIRE PRINCIPAL SAV RED BULL AVEC BDD COMPL�TE ===

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
    
    ' Cr�er les contr�les dynamiquement
    CreerControles
    
    ' D�marrer le timer de v�rification BDD
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
    ActualiserStatutBDD ' Sera d�fini plus bas
    ctrl.Alignment = 2
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    ' Label "Num�ro de s�rie"
    Set ctrl = Me.Controls.Add("VB.Label", "lblRef")
    ctrl.Left = 600
    ctrl.Top = 1100
    ctrl.Width = 1800
    ctrl.Caption = "Num�ro de s�rie frigo:"
    ctrl.Font.Size = 10
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    ' TextBox pour saisie du num�ro de s�rie
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
    
    ' Zone d'information d�taill�e
    Set ctrl = Me.Controls.Add("VB.Label", "lblInfo")
    ctrl.Left = 600
    ctrl.Top = 1550
    ctrl.Width = 6800
    ctrl.Height = 1800
    ctrl.Caption = "INSTRUCTIONS - VALIDATION FILTR�E :" & vbCrLf & vbCrLf & _
                   "1. Scannez ou saisissez le num�ro de s�rie du frigo Red Bull" & vbCrLf & _
                   "2. Cliquez SCANNER pour v�rifier dans la liste des 92 codes autoris�s" & vbCrLf & _
                   "3. Si valid�, vous pourrez ouvrir la fiche retour SAV" & vbCrLf & vbCrLf & _
                   "S�CURIT� - CODES FILTR�S :" & vbCrLf & _
                   "� Seulement 92 codes articles Red Bull autoris�s" & vbCrLf & _
                   "� Validation stricte : INNER JOIN + DISTINCT + filtres" & vbCrLf & _
                   "� �limination des valeurs NULL et vides" & vbCrLf & _
                   "� Seuls les �quipements de la liste peuvent �tre trait�s" & vbCrLf & vbCrLf & _
                   "BASE : ART_PAR + NSE_DAT avec act_code = 'RB'"
    ctrl.BackColor = RGB(248, 249, 250)
    ctrl.BorderStyle = 1
    ctrl.Alignment = 0 ' Alignement � gauche
    ctrl.Font.Size = 9
    ctrl.Visible = True
    
    ' Bouton Ouvrir Fiche (d�sactiv� au d�but)
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
    ctrl.Caption = "Pr�t - En attente de scan | " & ObtenirDateTimeFormatee()
    ctrl.BackColor = RGB(220, 220, 220)
    ctrl.BorderStyle = 1
    ctrl.Alignment = 2
    ctrl.Font.Size = 8
    ctrl.Visible = True
End Sub

Private Sub cmdValider_Click()
    Dim numeroSerie As String
    numeroSerie = Trim(UCase(Me.Controls("txtReference").Text))
    
    Me.Controls("lblStatut").Caption = "Validation filtr�e en cours... | " & ObtenirDateTimeFormatee()
    Me.Refresh
    
    If Len(numeroSerie) = 0 Then
        MsgBox "Veuillez saisir un num�ro de s�rie !", vbExclamation, "Erreur de saisie"
        Me.Controls("lblStatut").Caption = "Erreur: Num�ro de s�rie manquant | " & ObtenirDateTimeFormatee()
        Exit Sub
    End If
    
    ' Validation du format de base
    If Not ValiderFormatNumeroSerie(numeroSerie) Then
        AfficherErreurValidation "Format de num�ro de s�rie invalide", numeroSerie
        Exit Sub
    End If
    
    ' V�rification dans la base de donn�es
    If Not VerifierConnexionBDD() Then
        AfficherErreurConnexion numeroSerie
        Exit Sub
    End If
    
    ' Rechercher dans NSE_DAT et ART_PAR
    Dim resultatsValidation As TypeValidationBDD
    resultatsValidation = ValiderNumeroSerieBDD(numeroSerie)
    
    If resultatsValidation.existe Then
        ' Validation r�ussie dans la liste filtr�e
        referenceValidee = numeroSerie
        numeroSerieValide = numeroSerie
        AfficherValidationReussie resultatsValidation
        
        ' Enregistrer dans l'historique avec mention du filtrage
        EcrireHistoriqueScan numeroSerie, resultatsValidation.modeleArticle & " [FILTR�-92]"
    Else
        ' Num�ro de s�rie non autoris� dans la liste filtr�e
        AfficherErreurValidation resultatsValidation.statut, numeroSerie
    End If
End Sub


' Fonction pour afficher les informations syst�me d�taill�es
Private Sub AfficherInfosSystemeDetaillees()
    Dim infos As String
    
    infos = ObtenirInfosSysteme() & vbCrLf
    
    ' Ajouter des statistiques de session
    infos = infos & vbCrLf & "=== STATISTIQUES SESSION ===" & vbCrLf
    infos = infos & "R�f�rence valid�e actuelle: " & IIf(Len(referenceValidee) > 0, referenceValidee, "Aucune") & vbCrLf
    infos = infos & "Num�ro de s�rie valid�: " & IIf(Len(numeroSerieValide) > 0, numeroSerieValide, "Aucun") & vbCrLf
    
    ' Ajouter l'�tat des fichiers
    infos = infos & vbCrLf & "=== �TAT DES FICHIERS ===" & vbCrLf
    infos = infos & "Historique: " & IIf(Dir(App.Path & FICHIER_HISTORIQUE) <> "", "? Pr�sent", "? Absent") & vbCrLf
    infos = infos & "Stock pi�ces: " & IIf(Dir(App.Path & FICHIER_STOCK_PIECES) <> "", "? Pr�sent", "? Absent") & vbCrLf
    infos = infos & "Stock r�parable: " & IIf(Dir(App.Path & FICHIER_STOCK_REPARABLE) <> "", "? Pr�sent", "? Absent") & vbCrLf
    
    MsgBox infos, vbInformation, "Informations Syst�me D�taill�es"
End Sub

' === FONCTIONS DE COMMUNICATION AVEC SCANNER EXTERNE ===

' Fonction pour traiter les donn�es re�ues du port s�rie (scanner code-barres)
Private Sub MSComm1_OnComm()
    On Error GoTo ErrorHandler
    
    Dim donnees As String
    
    Select Case MSComm1.CommEvent
        Case comEvReceive
            donnees = MSComm1.Input
            If Len(donnees) > 0 Then
                ' Nettoyer les donn�es re�ues du scanner
                donnees = Replace(donnees, vbCr, "")
                donnees = Replace(donnees, vbLf, "")
                donnees = Replace(donnees, Chr(0), "") ' Supprimer les caract�res null
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
    Sleep 300 ' N�cessite API Windows
    
    Me.Controls("txtReference").BackColor = RGB(255, 255, 255) ' Retour au blanc
    
    ' D�clencher automatiquement la validation
    Me.Controls("lblStatut").Caption = "?? Scan automatique d�tect�... | " & ObtenirDateTimeFormatee()
    cmdValider_Click
End Sub

' === FONCTIONS AVANC�ES DE VALIDATION BDD ===

' Fonction pour v�rifier les d�pendances et contraintes
Private Function VerifierContraintesBDD(numeroSerie As String) As String
    On Error GoTo ErrorHandler
    
    Dim sql As String
    Dim rsContraintes As ADODB.Recordset
    Dim resultats As String
    
    ' V�rifier s'il y a des fiches SAV existantes pour ce num�ro
    sql = "SELECT COUNT(*) as nb_fiches FROM sav_historique WHERE nse_nums = '" & numeroSerie & "'"
    Set rsContraintes = ExecuterRequete(sql)
    
    If Not rsContraintes Is Nothing And Not rsContraintes.EOF Then
        If rsContraintes!nb_fiches > 0 Then
            resultats = resultats & "?? " & rsContraintes!nb_fiches & " fiche(s) SAV existante(s)" & vbCrLf
        End If
        rsContraintes.Close
    End If
    
    ' V�rifier le statut de l'�quipement
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
    VerifierContraintesBDD = "Erreur v�rification contraintes: " & Err.description
End Function

' === GESTION DES �V�NEMENTS DE FERMETURE ===

' �v�nement de fermeture du formulaire avec nettoyage complet
Private Sub Form_Unload(Cancel As Integer)
    ' Arr�ter le timer
    If Not tmrVerifBDD Is Nothing Then
        tmrVerifBDD.Enabled = False
        Set tmrVerifBDD = Nothing
    End If
    
    ' Fermer proprement la connexion BDD
    FermerBDD
    
    ' Cr�er une sauvegarde automatique avant fermeture
    Me.Controls("lblStatut").Caption = "Sauvegarde en cours... | " & ObtenirDateTimeFormatee()
    Me.Refresh
    SauvegardeAutomatique
    
    ' Log de fermeture
    EcrireHistoriqueScan "SYSTEM", "Application ferm�e"
End Sub

' === FONCTIONS D'INTERFACE UTILISATEUR AVANC�ES ===

' Fonction pour g�rer le clic droit (menu contextuel)
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then ' Clic droit
        AfficherMenuContextuel
    End If
End Sub

' Fonction pour afficher un menu contextuel avanc�
Private Sub AfficherMenuContextuel()
    Dim reponse As Integer
    Dim menu As String
    
    menu = "?? MENU CONTEXTUEL SAV RED BULL" & vbCrLf & vbCrLf
    menu = menu & "Choisissez une action :" & vbCrLf
    menu = menu & "� OUI = ?? Voir l'historique des scans" & vbCrLf
    menu = menu & "� NON = ?? Informations syst�me d�taill�es" & vbCrLf
    menu = menu & "� ANNULER = ?? Effectuer une maintenance rapide"
    
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

' Fonction pour logger les erreurs syst�me
Private Sub LoggerErreur(source As String, description As String)
    On Error Resume Next
    
    Dim fichierLog As String
    Dim numeroFichier As Integer
    
    fichierLog = App.Path & "\Logs\Erreurs_" & Format(Now, "yyyymmdd") & ".txt"
    
    ' Cr�er le r�pertoire Logs s'il n'existe pas
    If Dir(App.Path & "\Logs", vbDirectory) = "" Then
        MkDir App.Path & "\Logs"
    End If
    
    numeroFichier = FreeFile
    Open fichierLog For Append As #numeroFichier
    Print #numeroFichier, Format(Now, "dd/mm/yyyy hh:nn:ss") & " - [" & source & "] " & description
    Close #numeroFichier
End Sub

' === FONCTIONS DE DIAGNOSTIC R�SEAU ===

' Fonction pour tester la connectivit� r�seau vers le serveur BDD
Private Function TesterConnectiviteReseau() As Boolean
    On Error GoTo ErrorHandler
    
    ' Test simple de ping (simulation)
    ' Dans un environnement r�el, vous pourriez utiliser une API Windows pour ping
    TesterConnectiviteReseau = True
    Exit Function
    
ErrorHandler:
    TesterConnectiviteReseau = False
End Function

' === SAUVEGARDE ET R�CUP�RATION D'�TAT ===

' Fonction pour sauvegarder l'�tat actuel de la session
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
    
    Me.Controls("lblStatut").Caption = "Test filtrage termin� | " & ObtenirDateTimeFormatee()
End Sub

' Fonction pour r�cup�rer l'�tat de la derni�re session
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
                ' R�cup�rer la derni�re r�f�rence si besoin
            ElseIf InStr(ligne, "DERNIER_NUMERO_SERIE=") > 0 Then
                ' R�cup�rer le dernier num�ro de s�rie si besoin
            End If
        Loop
        
        Close #numeroFichier
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Erreur non critique, continuer normalement
End Sub

' Sauvegarde automatique de l'�tat toutes les 5 minutes
Private Sub tmrSauvegardeAuto_Timer()
    SauvegarderEtatSession
End Sub
' Fonction pour valider le format du num�ro de s�rie (appel au module)
Private Function ValiderFormatNumeroSerie(numeroSerie As String) As Boolean
    ValiderFormatNumeroSerie = ValiderFormatNumeroSerieRB(numeroSerie)
End Function

' Fonction pour obtenir des informations compl�mentaires (appel au module)
Private Function ObtenirInfosComplementaires(codeArticle As String) As String
    ObtenirInfosComplementaires = ObtenirInfosComplementairesArticle(codeArticle)
End Function

' Fonction pour afficher une validation r�ussie
Private Sub AfficherValidationReussie(resultats As TypeValidationBDD)
    Dim info As String
    
    info = "? NUM�RO DE S�RIE VALID� - LISTE AUTORIS�E" & vbCrLf & vbCrLf
    info = info & "?? D�TAILS PRODUIT:" & vbCrLf
    info = info & "� Num�ro de s�rie: " & resultats.numeroSerie & vbCrLf
    info = info & "� Code article: " & resultats.codeArticle & vbCrLf
    info = info & "� Mod�le: " & resultats.modeleArticle & vbCrLf
    info = info & "� Date cr�ation: " & resultats.dateCreation & vbCrLf
    
    If resultats.prixCatalogue > 0 Then
        info = info & "� Prix catalogue: " & Format(resultats.prixCatalogue, "0.00") & "�" & vbCrLf
    End If
    
    info = info & vbCrLf & "?? VALIDATION S�CURIS�E:" & vbCrLf
    info = info & "� Serveur: " & SERVER_NAME & vbCrLf
    info = info & "� Base: " & DATABASE_NAME & vbCrLf
    info = info & "� Statut: " & resultats.statut & vbCrLf
    info = info & "� Liste: 92 codes articles Red Bull autoris�s" & vbCrLf
    
    If Len(resultats.informationsComplementaires) > 0 Then
        info = info & "� " & resultats.informationsComplementaires & vbCrLf
    End If
    
    info = info & vbCrLf & "? Vous pouvez maintenant ouvrir la fiche retour SAV"
    
    Me.Controls("lblInfo").Caption = info
    Me.Controls("lblInfo").BackColor = RGB(212, 237, 218) ' Vert clair
    
    ' Activer le bouton fiche retour
    Me.Controls("cmdOuvrirFiche").Enabled = True
    Me.Controls("cmdOuvrirFiche").BackColor = RGB(40, 167, 69) ' Vert
    
    Me.Controls("lblStatut").Caption = "? Valid� (liste 92 codes): " & resultats.numeroSerie & " | " & ObtenirDateTimeFormatee()
End Sub

' Fonction pour afficher une erreur de validation
Private Sub AfficherErreurValidation(messageErreur As String, numeroSerie As String)
    Dim info As String
    
    info = "? VALIDATION �CHOU�E - HORS LISTE AUTORIS�E" & vbCrLf & vbCrLf
    info = info & "?? ERREUR: " & messageErreur & vbCrLf
    info = info & "?? Num�ro saisi: " & numeroSerie & vbCrLf & vbCrLf
    
    info = info & "?? V�RIFICATIONS EFFECTU�ES:" & vbCrLf
    info = info & "� Format du num�ro de s�rie" & vbCrLf
    info = info & "� Existence dans les 92 codes autoris�s" & vbCrLf
    info = info & "� Validation act_code = 'RB'" & vbCrLf
    info = info & "� Filtrage DISTINCT avec conditions strictes" & vbCrLf & vbCrLf
    
    info = info & "?? CONSEILS:" & vbCrLf
    info = info & "� V�rifiez la saisie du num�ro de s�rie" & vbCrLf
    info = info & "� Ce num�ro doit appartenir aux 92 codes autoris�s" & vbCrLf
    info = info & "� Contactez l'administrateur si l'�quipement devrait �tre autoris�"
    
    Me.Controls("lblInfo").Caption = info
    Me.Controls("lblInfo").BackColor = RGB(248, 215, 218) ' Rouge clair
    
    ' D�sactiver le bouton fiche
    Me.Controls("cmdOuvrirFiche").Enabled = False
    Me.Controls("cmdOuvrirFiche").BackColor = RGB(150, 150, 150)
    
    ' Enregistrer l'erreur dans l'historique avec mention du filtrage
    EcrireHistoriqueScan numeroSerie, "ERREUR FILTRAGE: " & messageErreur
    
    Me.Controls("lblStatut").Caption = "? Hors liste autoris�e | " & ObtenirDateTimeFormatee()
End Sub

' Fonction pour afficher une erreur de connexion
Private Sub AfficherErreurConnexion(numeroSerie As String)
    Dim info As String
    
    info = "?? PROBL�ME DE CONNEXION BASE DE DONN�ES" & vbCrLf & vbCrLf
    info = info & "?? STATUT: Connexion � la base de donn�es interrompue" & vbCrLf
    info = info & "??? Serveur: " & SERVER_NAME & vbCrLf
    info = info & "??? Base: " & DATABASE_NAME & vbCrLf & vbCrLf
    
    info = info & "? ACTIONS DISPONIBLES:" & vbCrLf
    info = info & "� Cliquez 'TESTER BDD' pour retenter la connexion" & vbCrLf
    info = info & "� V�rifiez votre connexion r�seau" & vbCrLf
    info = info & "� Contactez l'administrateur syst�me" & vbCrLf & vbCrLf
    
    info = info & "?? Mode d�grad� activ� - fonctions limit�es"
    
    Me.Controls("lblInfo").Caption = info
    Me.Controls("lblInfo").BackColor = RGB(255, 243, 205) ' Orange clair
    
    Me.Controls("lblStatut").Caption = "?? Connexion BDD perdue | " & ObtenirDateTimeFormatee()
End Sub

Private Sub cmdOuvrirFiche_Click()
    If Len(referenceValidee) = 0 Or Len(numeroSerieValide) = 0 Then
        MsgBox "Aucun num�ro de s�rie valid� !", vbExclamation
        Exit Sub
    End If
    
    Me.Controls("lblStatut").Caption = "Ouverture fiche retour... | " & ObtenirDateTimeFormatee()
    Me.Refresh
    
    ' Ouvrir la fiche retour avec toutes les donn�es valid�es
    Load frmFicheRetour
    frmFicheRetour.InitialiserAvecReference referenceValidee, numeroSerieValide
    frmFicheRetour.Show vbModal
    
    ' Reset apr�s fermeture de la fiche
    ResetFormulaire
End Sub

Private Sub cmdTestBDD_Click()
    Me.Controls("lblStatut").Caption = "Test syst�me en cours... | " & ObtenirDateTimeFormatee()
    Me.Refresh
    
    ' Tester compl�tement le syst�me
    TesterSystemeComplet
    
    ' Mettre � jour l'indicateur de statut BDD
    ActualiserStatutBDD
    
    Me.Controls("lblStatut").Caption = "Test termin� | " & ObtenirDateTimeFormatee()
End Sub

' Fonction pour actualiser le statut BDD dans l'interface
Private Sub ActualiserStatutBDD()
    Dim ctrl As Object
    Set ctrl = Me.Controls("lblStatutBDD")
    
    If VerifierConnexionBDD() Then
        ctrl.Caption = "? BASE CONNECT�E - FILTRAGE 92 CODES ACTIF - " & SERVER_NAME & " (" & DATABASE_NAME & ")"
        ctrl.BackColor = RGB(212, 237, 218) ' Vert clair
    Else
        ctrl.Caption = "? BASE D�CONNECT�E - MODE D�GRAD�"
        ctrl.BackColor = RGB(248, 215, 218) ' Rouge clair
    End If
End Sub

' Fonction pour d�marrer le timer de v�rification BDD
Private Sub DemarrerTimerBDD()
    Set tmrVerifBDD = Me.Controls.Add("VB.Timer", "tmrVerifBDD")
    tmrVerifBDD.Interval = 30000 ' 30 secondes
    tmrVerifBDD.Enabled = True
End Sub

' Timer pour v�rification p�riodique de la connexion BDD
Private Sub tmrVerifBDD_Timer()
    ' V�rifier la connexion et actualiser le statut
    ActualiserStatutBDD
    
    ' Si d�connect�, essayer de reconnecter automatiquement
    If Not VerifierConnexionBDD() Then
        ConnecterBDD
        ActualiserStatutBDD
    End If
End Sub

' Fonction pour remettre � z�ro le formulaire
Private Sub ResetFormulaire()
    Me.Controls("txtReference").Text = ""
    Me.Controls("lblInfo").Caption = "? FICHE TRAIT�E AVEC SUCC�S" & vbCrLf & vbCrLf & _
                                     "?? Vous pouvez scanner un nouveau frigo Red Bull" & vbCrLf & _
                                     "autoris� dans la liste des 92 codes" & vbCrLf & vbCrLf & _
                                     "? Appuyez sur F1 pour l'aide compl�te"
    Me.Controls("lblInfo").BackColor = RGB(248, 249, 250)
    Me.Controls("cmdOuvrirFiche").Enabled = False
    Me.Controls("cmdOuvrirFiche").BackColor = RGB(150, 150, 150)
    referenceValidee = ""
    numeroSerieValide = ""
    
    ' Donner le focus � la zone de saisie
    Me.Controls("txtReference").SetFocus
    
    Me.Controls("lblStatut").Caption = "Pr�t - Filtrage 92 codes actif | " & ObtenirDateTimeFormatee()
End Sub

' Gestion de la touche Entr�e dans la zone de texte
Private Sub txtReference_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ' Touche Entr�e
        KeyAscii = 0 ' Annuler le bip
        cmdValider_Click ' D�clencher la validation
    End If
End Sub

' Gestion des raccourcis clavier globaux
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 ' F1 = Aide compl�te
            AfficherAideComplete
            
        Case vbKeyF2 ' F2 = Test BDD
            cmdTestBDD_Click
            
        Case vbKeyF3 ' F3 = Historique
            cmdHistorique_Click
            
        Case vbKeyF5 ' F5 = Maintenance
            cmdMaintenance_Click
            
        Case vbKeyF12 ' F12 = Infos syst�me d�taill�es
            AfficherInfosSystemeDetaillees
            
        Case vbKeyEscape ' �chap = Reset
            ResetFormulaire
    End Select
End Sub

' Fonction pour afficher l'aide compl�te
Private Sub AfficherAideComplete()
    Dim aide As String
    
    aide = "=== AIDE SAV RED BULL SCANNER PRO v2.1 ===" & vbCrLf & vbCrLf
    aide = aide & "?? OBJECTIF:" & vbCrLf
    aide = aide & "Scanner et valider les num�ros de s�rie des frigos Red Bull" & vbCrLf
    aide = aide & "pour cr�er des fiches retour SAV." & vbCrLf & vbCrLf
    
    aide = aide & "?? UTILISATION:" & vbCrLf
    aide = aide & "1. Scannez ou saisissez le num�ro de s�rie" & vbCrLf
    aide = aide & "2. Cliquez SCANNER (ou appuyez Entr�e)" & vbCrLf
    aide = aide & "3. Si valid�, cliquez OUVRIR FICHE RETOUR" & vbCrLf & vbCrLf
    
    aide = aide & "??? BASE DE DONN�ES:" & vbCrLf
    aide = aide & "� Table NSE_DAT: Num�ros de s�rie" & vbCrLf
    aide = aide & "� Table ART_PAR: Articles et mod�les" & vbCrLf
    aide = aide & "� Filtre ACT_CODE = 'RB' (Red Bull uniquement)" & vbCrLf & vbCrLf
    
    aide = aide & "?? RACCOURCIS CLAVIER:" & vbCrLf
    aide = aide & "F1 = Cette aide | F2 = Test BDD | F3 = Historique" & vbCrLf
    aide = aide & "F5 = Maintenance | F12 = Infos syst�me | �chap = Reset" & vbCrLf & vbCrLf
    
    aide = aide & "?? MAINTENANCE:" & vbCrLf
    aide = aide & "Le syst�me effectue automatiquement:" & vbCrLf
    aide = aide & "� V�rification connexion BDD toutes les 30s" & vbCrLf
    aide = aide & "� Sauvegarde automatique � la fermeture" & vbCrLf
    aide = aide & "� Nettoyage des fichiers temporaires"
    
    MsgBox aide, vbInformation, "Aide SAV Red Bull Scanner Pro"
End Sub

' �v�nements pour les autres boutons
Private Sub cmdMaintenance_Click()
    Dim reponse As Integer
    reponse = MsgBox("Lancer la maintenance compl�te du syst�me ?" & vbCrLf & vbCrLf & _
                     "Cette op�ration va :" & vbCrLf & _
                     "� V�rifier l'int�grit� des fichiers" & vbCrLf & _
                     "� Nettoyer les fichiers temporaires" & vbCrLf & _
                     "� Tester la connexion BDD" & vbCrLf & _
                     "� Cr�er une sauvegarde compl�te" & vbCrLf & _
                     "� Synchroniser les donn�es", vbYesNo + vbQuestion, "Maintenance")
    
    If reponse = vbYes Then
        MaintenanceRapide
    End If
End Sub

Private Sub cmdHistorique_Click()
    AfficherHistoriqueDetaille
End Sub

' Fonction pour afficher l'historique d�taill�
Private Sub AfficherHistoriqueDetaille()
    Dim historique As String
    historique = LireHistoriqueScan()
    
    If Len(historique) = 0 Then
        MsgBox "Aucun historique disponible pour le moment.", vbInformation, "Historique"
        Exit Sub
    End If
    
    ' Compter les entr�es
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
    titre = "Historique des scans (" & nbEntrees & " entr�es)"
    
    ' Limiter l'affichage pour �viter les messages trop longs
    If Len(historique) > 2000 Then
        historique = Left(historique, 2000) & vbCrLf & vbCrLf & "... [Historique tronqu� - voir fichier complet]"
    End If
    
    ' Afficher dans une MessageBox pour simplicit�
    ' Dans une version avanc�e, vous pourriez cr�er une form d�di�e
    MsgBox titre & vbCrLf & String(50, "=") & vbCrLf & vbCrLf & historique, vbInformation, titre
End Sub

' Fonction de nettoyage final
Private Sub NettoyageFinal()
    On Error Resume Next
    
    ' Sauvegarder l'�tat final
    SauvegarderEtatSession
    
    ' Logger la fermeture
    EcrireHistoriqueScan "SYSTEM", "Application ferm�e normalement"
    
    ' Lib�rer les ressources
    Set cmdValider = Nothing
    Set cmdOuvrirFiche = Nothing
    Set cmdTestBDD = Nothing
    Set tmrVerifBDD = Nothing
End Sub
