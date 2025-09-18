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
Private numeroSerieFrigo As String

Private WithEvents cmdValider As CommandButton
Attribute cmdValider.VB_VarHelpID = -1
Private WithEvents cmdAnnuler As CommandButton
Attribute cmdAnnuler.VB_VarHelpID = -1

' Remplacement des OptionButton par des CheckBox
Private WithEvents chkMecanique As CheckBox
Attribute chkMecanique.VB_VarHelpID = -1
Private WithEvents chkEsthetique As CheckBox
Attribute chkEsthetique.VB_VarHelpID = -1
Private WithEvents chkCoherenceOui As CheckBox
Attribute chkCoherenceOui.VB_VarHelpID = -1
Private WithEvents chkCoherenceNon As CheckBox
Attribute chkCoherenceNon.VB_VarHelpID = -1
Private WithEvents chkReparable As CheckBox
Attribute chkReparable.VB_VarHelpID = -1
Private WithEvents chkHS As CheckBox
Attribute chkHS.VB_VarHelpID = -1

' Variable pour éviter les boucles infinies lors de la création
Private creationEnCours As Boolean

Private Sub Form_Load()
    Me.BackColor = RGB(240, 240, 240)
    Me.Caption = "FICHE RETOUR - RED BULL"
    Me.Width = 13000
    Me.Height = 12000
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    creationEnCours = True
    CreerInterfaceFiche
    creationEnCours = False
End Sub

Public Sub InitialiserAvecReference(reference As String, numeroSerie As String)
    ' Stocker les deux valeurs
    referenceFrigo = reference
    numeroSerieFrigo = numeroSerie
    
    ' Remplir les contrôles correspondants
    On Error Resume Next
    Me.Controls("txtReference").Text = referenceFrigo
    Me.Controls("txtSerie").Text = numeroSerieFrigo
    
    ' Mettre à jour le titre du formulaire
    Me.Caption = "FICHE RETOUR - RED BULL - " & numeroSerieFrigo
    
    On Error GoTo 0
    
    ' Récupérer automatiquement le numéro de réception REE_Nore
    RecupererNumeroReceptionREE numeroSerieFrigo
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
    ctrl.Top = 900
    ctrl.Width = 1800
    ctrl.Caption = "N° ENLEVEMENT :"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtEnlevement")
    ctrl.Left = 2400
    ctrl.Top = 900
    ctrl.Width = 3000
    ctrl.Height = 300
    ctrl.Visible = True
    
    ' N° RECEPTION
    Set ctrl = Me.Controls.Add("VB.Label", "lblReception")
    ctrl.Left = 500
    ctrl.Top = 1400
    ctrl.Width = 1800
    ctrl.Caption = "N° RECEPTION :"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtReception")
    ctrl.Left = 2400
    ctrl.Top = 1400
    ctrl.Width = 3000
    ctrl.Height = 300
    ctrl.Visible = True
    
    ' REFERENCE
    Set ctrl = Me.Controls.Add("VB.Label", "lblReference")
    ctrl.Left = 500
    ctrl.Top = 1900
    ctrl.Width = 1800
    ctrl.Caption = "REFERENCE :"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtReference")
    ctrl.Left = 2400
    ctrl.Top = 1900
    ctrl.Width = 3000
    ctrl.Height = 300
    ctrl.Text = referenceFrigo
    ctrl.Enabled = False
    ctrl.BackColor = RGB(240, 240, 240)
    ctrl.Visible = True
    
    ' MOTIF DU RETOUR - TITRE
    Set ctrl = Me.Controls.Add("VB.Label", "lblMotifTitre")
    ctrl.Left = 500
    ctrl.Top = 2500
    ctrl.Width = 2000
    ctrl.Caption = "MOTIF DU RETOUR :"
    ctrl.BackColor = RGB(220, 220, 220)
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' GROUPE MOTIF - Remplacé par des CheckBox
    Set chkMecanique = Me.Controls.Add("VB.CheckBox", "chkMecanique")
    chkMecanique.Left = 2600
    chkMecanique.Top = 2500
    chkMecanique.Width = 1500
    chkMecanique.Caption = "MECANIQUE"
    chkMecanique.Visible = True
    
    Set chkEsthetique = Me.Controls.Add("VB.CheckBox", "chkEsthetique")
    chkEsthetique.Left = 4200
    chkEsthetique.Top = 2500
    chkEsthetique.Width = 1500
    chkEsthetique.Caption = "ESTHETIQUE"
    chkEsthetique.Visible = True
    
    ' COHERENCE AVEC LA BOUTIQUE - TITRE
    Set ctrl = Me.Controls.Add("VB.Label", "lblCoherenceTitre")
    ctrl.Left = 500
    ctrl.Top = 3000
    ctrl.Width = 2500
    ctrl.Caption = "COHERENCE AVEC LA BOUTIQUE :"
    ctrl.BackColor = RGB(220, 220, 220)
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' GROUPE COHERENCE - Remplacé par des CheckBox
    Set chkCoherenceOui = Me.Controls.Add("VB.CheckBox", "chkCoherenceOui")
    chkCoherenceOui.Left = 3100
    chkCoherenceOui.Top = 3000
    chkCoherenceOui.Width = 600
    chkCoherenceOui.Caption = "OUI"
    chkCoherenceOui.Visible = True
    
    Set chkCoherenceNon = Me.Controls.Add("VB.CheckBox", "chkCoherenceNon")
    chkCoherenceNon.Left = 4000
    chkCoherenceNon.Top = 3000
    chkCoherenceNon.Width = 800
    chkCoherenceNon.Caption = "NON"
    chkCoherenceNon.Visible = True
    
    ' DIAGNOSTIC
    Set ctrl = Me.Controls.Add("VB.Label", "lblDiagnostic")
    ctrl.Left = 500
    ctrl.Top = 3600
    ctrl.Width = 1500
    ctrl.Caption = "DIAGNOSTIC :"
    ctrl.Visible = True
    
    ' Cases à cocher diagnostic
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkPieceManquante")
    ctrl.Left = 500
    ctrl.Top = 4100
    ctrl.Width = 4000
    ctrl.Caption = "PIECE MANQUANTE // PROBLEME CAPOT OU BAS DU FRIGO"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkTechnique")
    ctrl.Left = 500
    ctrl.Top = 4550
    ctrl.Width = 4000
    ctrl.Caption = "TECHNIQUE -> LUMIERE // FROID // MOTEUR // VITRE BRISEE"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkRayures")
    ctrl.Left = 500
    ctrl.Top = 4900
    ctrl.Width = 3500
    ctrl.Caption = "RAYURES TROP IMPORTANTES"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkLogoDegradé")
    ctrl.Left = 500
    ctrl.Top = 5300
    ctrl.Width = 2000
    ctrl.Caption = "LOGO DEGRADE"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkObsolete")
    ctrl.Left = 500
    ctrl.Top = 5700
    ctrl.Width = 2000
    ctrl.Caption = "OBSOLETE"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkBonEtat")
    ctrl.Left = 500
    ctrl.Top = 6150
    ctrl.Width = 3500
    ctrl.Caption = "BON ETAT -> REMIS DANS LE CIRCUIT"
    ctrl.Visible = True
    
    ' N° SERIE - REPOSITIONNÉ
    Set ctrl = Me.Controls.Add("VB.Label", "lblSerie")
    ctrl.Left = 500
    ctrl.Top = 6600
    ctrl.Width = 1200
    ctrl.Caption = "N° SERIE :"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtSerie")
    ctrl.Left = 1800
    ctrl.Top = 6600
    ctrl.Width = 2000
    ctrl.Height = 300
    ctrl.Visible = True
    
    ' COMMENTAIRE - REPOSITIONNÉ
    Set ctrl = Me.Controls.Add("VB.Label", "lblCommentaire")
    ctrl.Left = 500
    ctrl.Top = 7100
    ctrl.Width = 1500
    ctrl.Caption = "COMMENTAIRE :"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtCommentaire")
    ctrl.Left = 500
    ctrl.Top = 7500
    ctrl.Width = 6000
    ctrl.Height = 300
    ctrl.Visible = True
    
    ' QUALITE - TITRE - REPOSITIONNÉ
    Set ctrl = Me.Controls.Add("VB.Label", "lblQualiteTitre")
    ctrl.Left = 500
    ctrl.Top = 8000
    ctrl.Width = 1200
    ctrl.Caption = "QUALITE :"
    ctrl.BackColor = RGB(220, 220, 220)
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' GROUPE QUALITE - Remplacé par des CheckBox - REPOSITIONNÉ
    Set chkReparable = Me.Controls.Add("VB.CheckBox", "chkReparable")
    chkReparable.Left = 1800
    chkReparable.Top = 8000
    chkReparable.Width = 1500
    chkReparable.Caption = "REPARABLE"
    chkReparable.Visible = True
    
    Set chkHS = Me.Controls.Add("VB.CheckBox", "chkHS")
    chkHS.Left = 3500
    chkHS.Top = 8000
    chkHS.Width = 1000
    chkHS.Caption = "HS"
    chkHS.Visible = True
    
    ' Boutons - REPOSITIONNÉS
    Set cmdValider = Me.Controls.Add("VB.CommandButton", "cmdValider")
    cmdValider.Left = 2000
    cmdValider.Top = 8700
    cmdValider.Width = 1800
    cmdValider.Height = 400
    cmdValider.Caption = "VALIDER FICHE"
    cmdValider.BackColor = RGB(128, 255, 128)
    cmdValider.Visible = True
    
    Set cmdAnnuler = Me.Controls.Add("VB.CommandButton", "cmdAnnuler")
    cmdAnnuler.Left = 4000
    cmdAnnuler.Top = 8700
    cmdAnnuler.Width = 1800
    cmdAnnuler.Height = 400
    cmdAnnuler.Caption = "ANNULER"
    cmdAnnuler.BackColor = RGB(255, 128, 128)
    cmdAnnuler.Visible = True
End Sub

' GESTION DES GROUPES EXCLUSIFS - GROUPE MOTIF
Private Sub chkMecanique_Click()
    If creationEnCours Then Exit Sub
    If chkMecanique.Value = 1 Then
        If Not chkEsthetique Is Nothing Then chkEsthetique.Value = 0
    End If
End Sub

Private Sub chkEsthetique_Click()
    If creationEnCours Then Exit Sub
    If chkEsthetique.Value = 1 Then
        If Not chkMecanique Is Nothing Then chkMecanique.Value = 0
    End If
End Sub

' GESTION DES GROUPES EXCLUSIFS - GROUPE COHERENCE
Private Sub chkCoherenceOui_Click()
    If creationEnCours Then Exit Sub
    If chkCoherenceOui.Value = 1 Then
        If Not chkCoherenceNon Is Nothing Then chkCoherenceNon.Value = 0
    End If
End Sub

Private Sub chkCoherenceNon_Click()
    If creationEnCours Then Exit Sub
    If chkCoherenceNon.Value = 1 Then
        If Not chkCoherenceOui Is Nothing Then chkCoherenceOui.Value = 0
    End If
End Sub

' GESTION DES GROUPES EXCLUSIFS - GROUPE QUALITE
Private Sub chkReparable_Click()
    If creationEnCours Then Exit Sub
    If chkReparable.Value = 1 Then
        If Not chkHS Is Nothing Then chkHS.Value = 0
    End If
End Sub

Private Sub chkHS_Click()
    If creationEnCours Then Exit Sub
    If chkHS.Value = 1 Then
        If Not chkReparable Is Nothing Then chkReparable.Value = 0
    End If
End Sub

' === MODIFICATION DANS frmFicheRetour.frm ===
' Remplacer la méthode cmdValider_Click() existante par cette version :

Private Sub cmdValider_Click()
    If Not ValiderFormulaire() Then Exit Sub
    
    Dim statut As String
    If chkHS.Value = 1 Then
        statut = "HS"
    ElseIf chkReparable.Value = 1 Then
        statut = "REPARABLE"
    Else
        ' Aucune qualité sélectionnée - ne devrait pas arriver avec la validation
        MsgBox "Erreur : Aucune qualité sélectionnée !", vbCritical
        Exit Sub
    End If
    
    ' Sauvegarder d'abord la fiche
    SauvegarderFiche statut
    
    If statut = "HS" Then
        MsgBox "Fiche sauvegardée - Frigo marqué HS" & vbCrLf & "Ouverture du processus de récupération des pièces", vbInformation
        
        ' Ouvrir le formulaire de récupération des pièces
        Load frmRecuperationPieces
        frmRecuperationPieces.InitialiserAvecFrigo referenceFrigo, "Nom_Frigoriste"
        frmRecuperationPieces.Show vbModal
        
    ElseIf statut = "REPARABLE" Then
        MsgBox "Fiche sauvegardée - Frigo marqué REPARABLE" & vbCrLf & "Ouverture de l'affectation des pièces", vbInformation
        
        ' CORRECTION : Nom correct du formulaire et ordre des opérations
        Load frmAffectationPieces
        frmAffectationPieces.InitialiserAvecFrigo referenceFrigo, numeroSerieFrigo, "Nom_Frigoriste"
        frmAffectationPieces.Show vbModal
        
        ' Une fois l'affectation terminée, confirmer
        MsgBox "Processus d'affectation terminé." & vbCrLf & "Le frigo est maintenant en cours de réparation.", vbInformation
    End If
    
    Me.Hide
End Sub

Private Function ValiderFormulaire() As Boolean
    ' Validation des champs obligatoires
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
    
    ' VALIDATION DES GROUPES EXCLUSIFS
    
    ' 1. Vérification MOTIF : Mécanique ET Esthétique ne peuvent pas être cochés ensemble
    If chkMecanique.Value = 1 And chkEsthetique.Value = 1 Then
        MsgBox "ERREUR : Vous ne pouvez pas sélectionner MECANIQUE et ESTHETIQUE en même temps !" & vbCrLf & _
               "Veuillez ne choisir qu'un seul motif de retour.", vbExclamation + vbCritical, "Sélection invalide"
        ValiderFormulaire = False
        Exit Function
    End If
    
    ' 2. Vérification COHERENCE : OUI ET NON ne peuvent pas être cochés ensemble
    If chkCoherenceOui.Value = 1 And chkCoherenceNon.Value = 1 Then
        MsgBox "ERREUR : Vous ne pouvez pas sélectionner OUI et NON en même temps pour la cohérence !" & vbCrLf & _
               "Veuillez choisir une seule option.", vbExclamation + vbCritical, "Sélection invalide"
        ValiderFormulaire = False
        Exit Function
    End If
    
    ' 3. Vérification QUALITE : REPARABLE ET HS ne peuvent pas être cochés ensemble
    If chkReparable.Value = 1 And chkHS.Value = 1 Then
        MsgBox "ERREUR : Vous ne pouvez pas sélectionner REPARABLE et HS en même temps !" & vbCrLf & _
               "Veuillez choisir une seule option de qualité.", vbExclamation + vbCritical, "Sélection invalide"
        ValiderFormulaire = False
        Exit Function
    End If
    
    ' 4. Vérification qu'au moins une option est sélectionnée pour chaque groupe obligatoire
    If chkMecanique.Value = 0 And chkEsthetique.Value = 0 Then
        MsgBox "ATTENTION : Veuillez sélectionner un motif de retour (MECANIQUE ou ESTHETIQUE).", vbExclamation
        ValiderFormulaire = False
        Exit Function
    End If
    
    If chkCoherenceOui.Value = 0 And chkCoherenceNon.Value = 0 Then
        MsgBox "ATTENTION : Veuillez indiquer la cohérence avec la boutique (OUI ou NON).", vbExclamation
        ValiderFormulaire = False
        Exit Function
    End If
    
    If chkReparable.Value = 0 And chkHS.Value = 0 Then
        MsgBox "ATTENTION : Veuillez indiquer la qualité du frigo (REPARABLE ou HS).", vbExclamation
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
    If chkMecanique.Value = 1 Then Print #numeroFichier, "- MECANIQUE"
    If chkEsthetique.Value = 1 Then Print #numeroFichier, "- ESTHETIQUE"
    Print #numeroFichier, ""
    Print #numeroFichier, "COHERENCE AVEC LA BOUTIQUE:"
    If chkCoherenceOui.Value = 1 Then Print #numeroFichier, "- OUI"
    If chkCoherenceNon.Value = 1 Then Print #numeroFichier, "- NON"
    Print #numeroFichier, ""
    Print #numeroFichier, "DIAGNOSTIC:"
    If Me.Controls("chkPieceManquante").Value = 1 Then Print #numeroFichier, "- PIECE MANQUANTE"
    If Me.Controls("chkTechnique").Value = 1 Then Print #numeroFichier, "- TECHNIQUE"
    If Me.Controls("chkRayures").Value = 1 Then Print #numeroFichier, "- RAYURES"
    If Me.Controls("chkLogoDegradé").Value = 1 Then Print #numeroFichier, "- LOGO DEGRADE"
    If Me.Controls("chkObsolete").Value = 1 Then Print #numeroFichier, "- OBSOLETE"
    If Me.Controls("chkBonEtat").Value = 1 Then Print #numeroFichier, "- BON ETAT"
    Print #numeroFichier, ""
    Print #numeroFichier, "N° SERIE: " & Me.Controls("txtSerie").Text
    Print #numeroFichier, "COMMENTAIRE: " & Me.Controls("txtCommentaire").Text
    Print #numeroFichier, ""
    Print #numeroFichier, "QUALITE: " & statut
    Print #numeroFichier, "Date création: " & Now
    Print #numeroFichier, ""
    Print #numeroFichier, "NOTE: Les temps de réparation/récupération seront saisis dans les formulaires dédiés."
    Close #numeroFichier
    
    Exit Sub
    
GestionErreur:
    MsgBox "Erreur lors de la sauvegarde: " & Err.description, vbCritical
End Sub

Private Sub cmdAnnuler_Click()
    If MsgBox("Etes-vous sûr de vouloir annuler cette fiche ?", vbYesNo + vbQuestion) = vbYes Then
        Me.Hide
    End If
End Sub

' Récupération automatique du numéro de réception (VERSION CORRIGÉE)
Private Sub RecupererNumeroReceptionREE(numeroSerie As String)
    On Error GoTo ErrorHandler
    
    If Not VerifierConnexionBDD() Then
        MsgBox "Impossible de récupérer le numéro de réception : pas de connexion BDD", vbExclamation
        Exit Sub
    End If
    
    ' UTILISER LA NOUVELLE FONCTION CORRIGÉE
    Dim donneesREE As TypeDonneesREE
    donneesREE = RecupererNumeroReceptionCorrect(numeroSerie)
    
    If donneesREE.trouve Then
        ' Numéro de réception trouvé
        On Error Resume Next
        Me.Controls("txtReception").Text = donneesREE.numeroReception
        Me.Controls("txtReception").Enabled = False
        Me.Controls("txtReception").BackColor = RGB(240, 240, 240)
        On Error GoTo 0
        
        MsgBox "Numéro de réception récupéré automatiquement :" & vbCrLf & _
               "• N° Réception : " & donneesREE.numeroReception, _
               vbInformation, "Données BDD récupérées"
    Else
        ' Permettre saisie manuelle
        On Error Resume Next
        Me.Controls("txtReception").Text = ""
        Me.Controls("txtReception").Enabled = True
        Me.Controls("txtReception").BackColor = RGB(255, 255, 255)
        On Error GoTo 0
        
        ' AJOUTER UN DIAGNOSTIC AUTOMATIQUE
        Debug.Print "=== DIAGNOSTIC AUTOMATIQUE ==="
        DiagnostiquerProblemeREE numeroSerie
        
        MsgBox "Aucun numéro de réception trouvé." & vbCrLf & _
               "Erreur: " & donneesREE.messageErreur & vbCrLf & _
               "Saisie manuelle requise." & vbCrLf & vbCrLf & _
               "Vérifiez la fenêtre Debug pour plus de détails.", _
               vbExclamation, "Saisie manuelle requise"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur lors de la récupération du numéro de réception :" & vbCrLf & _
           Err.description, vbCritical
End Sub

