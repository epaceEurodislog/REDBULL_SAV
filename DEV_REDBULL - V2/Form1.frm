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
' === FORM1.FRM - INTERFACE PROFESSIONNELLE SAV RED BULL ===

' Variables globales
Private referenceScannee As String
Private informationsFrigo As String

Private Sub Form_Load()
    ' Configuration du formulaire
    Me.BackColor = RGB(245, 245, 245)
    Me.Caption = "SAV Red Bull Scanner Pro - v2.1"
    Me.Width = 14000
    Me.Height = 10000
    Me.WindowState = 0
    
    ' Centrer le formulaire
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    ' Initialiser
    referenceScannee = ""
    informationsFrigo = ""
    
    ' Configurer l'interface
    ConfigurerInterface
End Sub

Private Sub ConfigurerInterface()
    ' Cette procédure configure les contrôles que VOUS devez créer dans le designer VB6
    
    ' Vérifier que les contrôles existent avant de les configurer
    On Error Resume Next
    
    ' Configuration titre
    If Not lblTitre Is Nothing Then
        lblTitre.Caption = "?? SAV RED BULL - SCANNER PROFESSIONNEL ??"
        lblTitre.BackColor = RGB(30, 144, 255)
        lblTitre.ForeColor = RGB(255, 255, 255)
        lblTitre.Font.Size = 16
        lblTitre.Font.Bold = True
        lblTitre.Alignment = 2
    End If
    
    ' Configuration zone de scan
    If Not txtCodeBarre Is Nothing Then
        txtCodeBarre.Font.Size = 14
        txtCodeBarre.Font.Bold = True
        txtCodeBarre.Text = ""
    End If
    
    ' Configuration des boutons
    If Not cmdScanner Is Nothing Then
        cmdScanner.Caption = "?? SCANNER"
        cmdScanner.BackColor = RGB(46, 204, 113)
        cmdScanner.Font.Bold = True
    End If
    
    If Not cmdTest1 Is Nothing Then
        cmdTest1.Caption = "Test VC2286"
        cmdTest1.BackColor = RGB(230, 230, 230)
    End If
    
    If Not cmdTest2 Is Nothing Then
        cmdTest2.Caption = "Test RB4458"
        cmdTest2.BackColor = RGB(230, 230, 230)
    End If
    
    If Not cmdCreerFiche Is Nothing Then
        cmdCreerFiche.Caption = "?? CRÉER FICHE RETOUR"
        cmdCreerFiche.BackColor = RGB(189, 195, 199)
        cmdCreerFiche.Enabled = False
    End If
    
    If Not cmdStockReparable Is Nothing Then
        cmdStockReparable.Caption = "?? STOCK RÉPARABLE"
        cmdStockReparable.BackColor = RGB(52, 152, 219)
    End If
    
    If Not cmdStockPieces Is Nothing Then
        cmdStockPieces.Caption = "?? STOCK PIÈCES"
        cmdStockPieces.BackColor = RGB(155, 89, 182)
    End If
    
    ' Configuration zone info
    If Not lblInfoFrigo Is Nothing Then
        lblInfoFrigo.Caption = "Aucun réfrigérateur scanné" & vbCrLf & "Veuillez scanner un code-barres pour afficher les informations"
        lblInfoFrigo.BackColor = RGB(248, 249, 250)
    End If
    
    ' Configuration statut
    If Not lblStatut Is Nothing Then
        lblStatut.Caption = "Prêt - En attente de scan..."
        lblStatut.BackColor = RGB(236, 240, 241)
    End If
    
    On Error GoTo 0
End Sub

' === ÉVÉNEMENTS DES CONTRÔLES ===

Private Sub txtCodeBarre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ' Entrée
        TraiterCodeBarre
        KeyAscii = 0
    End If
End Sub

Private Sub cmdScanner_Click()
    TraiterCodeBarre
End Sub

Private Sub cmdTest1_Click()
    txtCodeBarre.Text = "VC2286-52000-1"
    TraiterCodeBarre
End Sub

Private Sub cmdTest2_Click()
    txtCodeBarre.Text = "RB4458-78900-2"
    TraiterCodeBarre
End Sub

Private Sub cmdCreerFiche_Click()
    CreerFicheRetour
End Sub

Private Sub cmdStockReparable_Click()
    OuvrirStockReparable
End Sub

Private Sub cmdStockPieces_Click()
    OuvrirStockPieces
End Sub

Private Sub cmdHistorique_Click()
    AfficherHistorique
End Sub

Private Sub cmdEffacer_Click()
    EffacerHistorique
End Sub

' === LOGIQUE MÉTIER ===

Private Sub TraiterCodeBarre()
    Dim codeBarre As String
    
    codeBarre = Trim(txtCodeBarre.Text)
    
    If Len(codeBarre) = 0 Then
        MsgBox "?? Veuillez saisir ou scanner un code-barres !", vbExclamation, "Code manquant"
        If Not lblStatut Is Nothing Then
            lblStatut.Caption = "Erreur - Code-barres manquant"
            lblStatut.BackColor = RGB(231, 76, 60)
            lblStatut.ForeColor = RGB(255, 255, 255)
        End If
        Exit Sub
    End If
    
    ' Animation traitement
    If Not cmdScanner Is Nothing Then
        cmdScanner.Caption = "? Traitement..."
        cmdScanner.BackColor = RGB(241, 196, 15)
        cmdScanner.Enabled = False
    End If
    
    If Not lblStatut Is Nothing Then
        lblStatut.Caption = "Traitement du code: " & codeBarre
        lblStatut.BackColor = RGB(241, 196, 15)
        lblStatut.ForeColor = RGB(0, 0, 0)
    End If
    
    ' Simuler délai
    Dim i As Long
    For i = 1 To 30000000: Next i
    
    ' Récupérer infos
    referenceScannee = codeBarre
    informationsFrigo = ObtenirInfosFrigo(codeBarre)
    
    ' Afficher résultat
    If Not lblInfoFrigo Is Nothing Then
        lblInfoFrigo.Caption = informationsFrigo
        lblInfoFrigo.BackColor = RGB(255, 255, 240)
    End If
    
    ' Activer bouton fiche
    If Not cmdCreerFiche Is Nothing Then
        cmdCreerFiche.Enabled = True
        cmdCreerFiche.BackColor = RGB(231, 76, 60)
        cmdCreerFiche.ForeColor = RGB(255, 255, 255)
        cmdCreerFiche.Caption = "?? CRÉER FICHE RETOUR ?"
    End If
    
    ' Ajouter à l'historique
    If Not lstHistorique Is Nothing Then
        lstHistorique.AddItem Format(Now, "dd/mm/yy hh:nn:ss") & " - " & codeBarre & " - " & GetModeleFromCode(codeBarre)
        lstHistorique.TopIndex = lstHistorique.ListCount - 1
    End If
    
    ' Restaurer bouton
    If Not cmdScanner Is Nothing Then
        cmdScanner.Caption = "?? SCANNER"
        cmdScanner.BackColor = RGB(46, 204, 113)
        cmdScanner.Enabled = True
    End If
    
    ' Statut success
    If Not lblStatut Is Nothing Then
        lblStatut.Caption = "? Scan réussi - " & codeBarre & " - " & GetModeleFromCode(codeBarre)
        lblStatut.BackColor = RGB(46, 204, 113)
        lblStatut.ForeColor = RGB(255, 255, 255)
    End If
    
    MsgBox "? Code-barres traité avec succès !" & vbCrLf & vbCrLf & _
           "Référence: " & codeBarre & vbCrLf & _
           "Modèle: " & GetModeleFromCode(codeBarre) & vbCrLf & vbCrLf & _
           "Vous pouvez maintenant créer la fiche retour.", vbInformation, "Scan réussi"
End Sub

Private Function ObtenirInfosFrigo(codeBarre As String) As String
    Dim info As String
    
    Select Case Left(codeBarre, 6)
        Case "VC2286"
            info = "??? RÉFRIGÉRATEUR VITRINE - Modèle VC2286" & vbCrLf & vbCrLf
            info = info & "?? Capacité: 250L" & vbCrLf
            info = info & "??? Température: +2°C à +8°C" & vbCrLf
            info = info & "?? Composants: Compresseur, LED, Vitre, Thermostat" & vbCrLf
            info = info & "?? Fabrication: 2023" & vbCrLf
            info = info & "??? Garantie: 24 mois" & vbCrLf
            info = info & "?? Prix neuf: 1,250€" & vbCrLf
            info = info & "? Série: Premium Vitrine"
            
        Case "RB4458"
            info = "??? RÉFRIGÉRATEUR RED BULL - Modèle RB4458" & vbCrLf & vbCrLf
            info = info & "?? Capacité: 180L" & vbCrLf
            info = info & "??? Température: +1°C à +6°C" & vbCrLf
            info = info & "?? Composants: Compresseur, LED, Vitre sécurisée, Écran digital" & vbCrLf
            info = info & "?? Fabrication: 2024" & vbCrLf
            info = info & "??? Garantie: 36 mois" & vbCrLf
            info = info & "?? Prix neuf: 1,580€" & vbCrLf
            info = info & "? Série: Red Bull Edition"
            
        Case Else
            info = "??? RÉFRIGÉRATEUR GÉNÉRIQUE" & vbCrLf & vbCrLf
            info = info & "?? Référence: " & codeBarre & vbCrLf
            info = info & "?? Modèle: Non identifié" & vbCrLf
            info = info & "?? Composants: Standard (à vérifier)" & vbCrLf
            info = info & "?? Fabrication: Inconnue" & vbCrLf
            info = info & "??? Garantie: À déterminer" & vbCrLf
            info = info & "?? Vérification manuelle requise" & vbCrLf
            info = info & "?? Consultez la documentation technique"
    End Select
    
    ObtenirInfosFrigo = info
End Function

Private Function GetModeleFromCode(codeBarre As String) As String
    Select Case Left(codeBarre, 6)
        Case "VC2286": GetModeleFromCode = "Vitrine VC2286"
        Case "RB4458": GetModeleFromCode = "Red Bull RB4458"
        Case Else: GetModeleFromCode = "Modèle générique"
    End Select
End Function

Private Sub CreerFicheRetour()
    If Len(referenceScannee) = 0 Then
        MsgBox "?? SCAN REQUIS" & vbCrLf & vbCrLf & _
               "Veuillez d'abord scanner un réfrigérateur !" & vbCrLf & vbCrLf & _
               "• Utilisez les boutons de test" & vbCrLf & _
               "• Ou saisissez manuellement un code", _
               vbExclamation, "Erreur"
        Exit Sub
    End If
    
    ' Simuler l'ouverture du formulaire de fiche
    Dim message As String
    message = "?? OUVERTURE FICHE RETOUR SAV" & vbCrLf & vbCrLf
    message = message & "??? Référence: " & referenceScannee & vbCrLf
    message = message & "?? Modèle: " & GetModeleFromCode(referenceScannee) & vbCrLf & vbCrLf
    message = message & "Le formulaire de création de fiche retour" & vbCrLf
    message = message & "s'ouvrirait maintenant avec ces informations."
    
    MsgBox message, vbInformation, "Fiche Retour"
    
    ' TODO: Décommenter quand le formulaire sera créé
    ' Load frmFicheRetour
    ' frmFicheRetour.InitialiserAvecReference referenceScannee
    ' frmFicheRetour.Show
End Sub

Private Sub OuvrirStockReparable()
    MsgBox "?? STOCK RÉPARABLE" & vbCrLf & vbCrLf & _
           "Ouverture de la gestion du stock des" & vbCrLf & _
           "réfrigérateurs réparables..." & vbCrLf & vbCrLf & _
           "Cette fonction ouvrira le formulaire de" & vbCrLf & _
           "gestion et d'affectation des pièces.", _
           vbInformation, "Stock Réparable"
    
    ' TODO: Décommenter quand le formulaire sera créé
    ' Load frmStockReparable
    ' frmStockReparable.Show
End Sub

Private Sub OuvrirStockPieces()
    MsgBox "?? STOCK PIÈCES" & vbCrLf & vbCrLf & _
           "Ouverture de l'inventaire des pièces" & vbCrLf & _
           "détachées récupérées..." & vbCrLf & vbCrLf & _
           "Cette fonction ouvrira le formulaire de" & vbCrLf & _
           "visualisation et gestion du stock.", _
           vbInformation, "Stock Pièces"
    
    ' TODO: Décommenter quand le formulaire sera créé
    ' Load frmStockPieces
    ' frmStockPieces.Show
End Sub

Private Sub AfficherHistorique()
    If lstHistorique Is Nothing Then
        MsgBox "Contrôle historique non disponible", vbExclamation
        Exit Sub
    End If
    
    If lstHistorique.ListCount = 0 Then
        MsgBox "?? HISTORIQUE VIDE" & vbCrLf & vbCrLf & _
               "Aucun scan effectué." & vbCrLf & vbCrLf & _
               "Utilisez les boutons de test pour commencer.", _
               vbInformation, "Historique"
    Else
        Dim message As String
        Dim i As Integer
        
        message = "?? HISTORIQUE DES SCANS" & vbCrLf & vbCrLf
        
        For i = 0 To lstHistorique.ListCount - 1
            message = message & "• " & lstHistorique.List(i) & vbCrLf
        Next i
        
        message = message & vbCrLf & "Total: " & lstHistorique.ListCount & " scan(s)"
        
        MsgBox message, vbInformation, "Historique complet"
    End If
End Sub

Private Sub EffacerHistorique()
    If lstHistorique Is Nothing Then Exit Sub
    
    If MsgBox("Effacer tout l'historique ?", vbYesNo + vbQuestion) = vbYes Then
        lstHistorique.Clear
        If Not lblStatut Is Nothing Then
            lblStatut.Caption = "Historique effacé"
            lblStatut.BackColor = RGB(52, 152, 219)
        End If
    End If
End Sub

' === INSTRUCTIONS POUR CRÉER L'INTERFACE ===
'
' CRÉEZ CES CONTRÔLES DANS LE DESIGNER VB6 :
'
' LABELS :
' Name: lblTitre - Caption: "TITRE" - Top: 100, Width: 10000, Height: 600
' Name: lblInfoFrigo - Caption: "INFO" - Top: 3200, Width: 10000, Height: 1500
' Name: lblStatut - Caption: "STATUT" - Top: 8500, Width: 10000, Height: 300
'
' TEXTBOX :
' Name: txtCodeBarre - Top: 1700, Width: 4000, Height: 400
'
' COMMAND BUTTONS :
' Name: cmdScanner - Caption: "SCANNER" - Top: 1700, Width: 1500, Height: 400
' Name: cmdTest1 - Caption: "Test1" - Top: 1700, Width: 800, Height: 200
' Name: cmdTest2 - Caption: "Test2" - Top: 1920, Width: 800, Height: 200
' Name: cmdCreerFiche - Caption: "FICHE" - Top: 4800, Width: 2500, Height: 600
' Name: cmdStockReparable - Caption: "STOCK REP" - Top: 4800, Width: 2500, Height: 600
' Name: cmdStockPieces - Caption: "STOCK PIECES" - Top: 4800, Width: 2500, Height: 600
' Name: cmdHistorique - Caption: "HISTORIQUE" - Top: 5500, Width: 1500, Height: 400
' Name: cmdEffacer - Caption: "EFFACER" - Top: 5500, Width: 1500, Height: 400
'
' LISTBOX :
' Name: lstHistorique - Top: 6000, Width: 10000, Height: 2000
'
' POSITIONNEMENT SUGGÉRÉ :
' - Titre en haut centré
' - Zone scan au milieu avec boutons tests à droite
' - Zone info frigo dessous
' - 3 boutons principaux alignés
' - Historique en bas avec boutons de gestion
' - Barre statut tout en bas

