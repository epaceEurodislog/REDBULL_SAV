VERSION 5.00
Begin VB.Form frmRecuperationPieces 
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
Attribute VB_Name = "frmRecuperationPieces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' === FRMRECUPERATIONPIECES.FRM - VERSION FONCTIONNELLE CORRIGÉE ===

Private referenceFrigo As String
Private nomFrigoriste As String

' DÉCLARATIONS WITHEVENTS POUR LES BOUTONS PRINCIPAUX
Private WithEvents cmdValiderRecuperation As CommandButton
Attribute cmdValiderRecuperation.VB_VarHelpID = -1
Private WithEvents cmdAnnuler As CommandButton
Attribute cmdAnnuler.VB_VarHelpID = -1
Private WithEvents cmdMettreAJourResume As CommandButton
Attribute cmdMettreAJourResume.VB_VarHelpID = -1

' DÉCLARATIONS WITHEVENTS POUR TOUS LES BOUTONS + ET -
Private WithEvents cmdPlus0 As CommandButton
Attribute cmdPlus0.VB_VarHelpID = -1
Private WithEvents cmdMoins0 As CommandButton
Attribute cmdMoins0.VB_VarHelpID = -1
Private WithEvents cmdPlus1 As CommandButton
Attribute cmdPlus1.VB_VarHelpID = -1
Private WithEvents cmdMoins1 As CommandButton
Attribute cmdMoins1.VB_VarHelpID = -1
Private WithEvents cmdPlus2 As CommandButton
Attribute cmdPlus2.VB_VarHelpID = -1
Private WithEvents cmdMoins2 As CommandButton
Attribute cmdMoins2.VB_VarHelpID = -1
Private WithEvents cmdPlus3 As CommandButton
Attribute cmdPlus3.VB_VarHelpID = -1
Private WithEvents cmdMoins3 As CommandButton
Attribute cmdMoins3.VB_VarHelpID = -1
Private WithEvents cmdPlus4 As CommandButton
Attribute cmdPlus4.VB_VarHelpID = -1
Private WithEvents cmdMoins4 As CommandButton
Attribute cmdMoins4.VB_VarHelpID = -1
Private WithEvents cmdPlus5 As CommandButton
Attribute cmdPlus5.VB_VarHelpID = -1
Private WithEvents cmdMoins5 As CommandButton
Attribute cmdMoins5.VB_VarHelpID = -1
Private WithEvents cmdPlus6 As CommandButton
Attribute cmdPlus6.VB_VarHelpID = -1
Private WithEvents cmdMoins6 As CommandButton
Attribute cmdMoins6.VB_VarHelpID = -1
Private WithEvents cmdPlus7 As CommandButton
Attribute cmdPlus7.VB_VarHelpID = -1
Private WithEvents cmdMoins7 As CommandButton
Attribute cmdMoins7.VB_VarHelpID = -1
Private WithEvents cmdPlus8 As CommandButton
Attribute cmdPlus8.VB_VarHelpID = -1
Private WithEvents cmdMoins8 As CommandButton
Attribute cmdMoins8.VB_VarHelpID = -1
Private WithEvents cmdPlus9 As CommandButton
Attribute cmdPlus9.VB_VarHelpID = -1
Private WithEvents cmdMoins9 As CommandButton
Attribute cmdMoins9.VB_VarHelpID = -1
Private WithEvents cmdPlus10 As CommandButton
Attribute cmdPlus10.VB_VarHelpID = -1
Private WithEvents cmdMoins10 As CommandButton
Attribute cmdMoins10.VB_VarHelpID = -1
Private WithEvents cmdPlus11 As CommandButton
Attribute cmdPlus11.VB_VarHelpID = -1
Private WithEvents cmdMoins11 As CommandButton
Attribute cmdMoins11.VB_VarHelpID = -1
Private WithEvents cmdPlus12 As CommandButton
Attribute cmdPlus12.VB_VarHelpID = -1
Private WithEvents cmdMoins12 As CommandButton
Attribute cmdMoins12.VB_VarHelpID = -1

Private Sub Form_Load()
    Me.BackColor = RGB(250, 250, 250)
    Me.Caption = "Récupération Pièces - " & referenceFrigo
    Me.Width = 14000
    Me.Height = 15000
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    CreerInterfaceRecuperation
End Sub

Public Sub InitialiserAvecFrigo(reference As String, frigoriste As String)
    referenceFrigo = reference
    nomFrigoriste = frigoriste
    Me.Caption = "Récupération Pièces - " & referenceFrigo
End Sub

Private Sub CreerInterfaceRecuperation()
    Dim ctrl As Object
    
    ' === TITRE ===
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitre")
    ctrl.Left = 500
    ctrl.Top = 300
    ctrl.Width = 11000
    ctrl.Height = 500
    ctrl.Caption = "RÉCUPÉRATION DES PIÈCES - FRIGO HS"
    ctrl.BackColor = RGB(255, 100, 100)
    ctrl.ForeColor = RGB(255, 255, 255)
    ctrl.Font.Size = 16
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' === INFOS FRIGO ===
    Set ctrl = Me.Controls.Add("VB.Label", "lblInfoFrigo")
    ctrl.Left = 500
    ctrl.Top = 1000
    ctrl.Width = 11000
    ctrl.Height = 400
    ctrl.Caption = "Référence: " & referenceFrigo & " | Frigoriste: " & nomFrigoriste & " | Date: " & Format(Now, "dd/mm/yyyy")
    ctrl.BackColor = RGB(255, 255, 200)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' === INSTRUCTIONS ===
    Set ctrl = Me.Controls.Add("VB.Label", "lblInstructions")
    ctrl.Left = 500
    ctrl.Top = 1600
    ctrl.Width = 11000
    ctrl.Height = 400
    ctrl.Caption = "Cochez les pièces récupérables et ajustez les quantités"
    ctrl.Font.Size = 12
    ctrl.Alignment = 2
    ctrl.BackColor = RGB(240, 248, 255)
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' === LISTE DES PIÈCES (espacées de 500 pixels) ===
    CreerLignePiece 2300, "PORTE", "PORTE", 1, 0
    CreerLignePiece 2800, "JOINT PORTE", "JOINT_PORTE", 2, 1
    CreerLignePiece 3300, "GOND PORTE", "GOND_PORTE", 2, 2
    CreerLignePiece 3800, "ETAGERE", "ETAGERE", 4, 3
    CreerLignePiece 4300, "ATTACHE ETAGERE", "ATTACHE_ETAGERE", 8, 4
    CreerLignePiece 4800, "VENTILATEUR", "VENTILATEUR", 1, 5
    CreerLignePiece 5300, "PRISE", "PRISE", 1, 6
    CreerLignePiece 5800, "CAPOT HAUT", "CAPOT_HAUT", 1, 7
    CreerLignePiece 6300, "CAPOT BAS", "CAPOT_BAS", 1, 8
    CreerLignePiece 6800, "LOGO", "LOGO", 2, 9
    CreerLignePiece 7300, "POMPE", "POMPE", 1, 10
    CreerLignePiece 7800, "VIS DIVERSES", "VIS_DIVERSES", 10, 11
    CreerLignePiece 8300, "GRILLE ARRIERE", "GRILLE_ARRIERE", 1, 12
    
    ' === SECTION TEMPS (plus d'espace) ===
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitreTempsRecup")
    ctrl.Left = 500
    ctrl.Top = 9200
    ctrl.Width = 11000
    ctrl.Height = 400
    ctrl.Caption = "TEMPS DE RÉCUPÉRATION"
    ctrl.BackColor = RGB(100, 150, 255)
    ctrl.ForeColor = RGB(255, 255, 255)
    ctrl.Font.Bold = True
    ctrl.Font.Size = 12
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkRecuperationEffectuee")
    ctrl.Left = 800
    ctrl.Top = 9800
    ctrl.Width = 2500
    ctrl.Height = 350
    ctrl.Caption = "RÉCUPÉRATION EFFECTUÉE"
    ctrl.Font.Bold = True
    ctrl.Font.Size = 11
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.Label", "lblTempsRecuperation")
    ctrl.Left = 3500
    ctrl.Top = 9800
    ctrl.Width = 2000
    ctrl.Height = 350
    ctrl.Caption = "TEMPS PASSÉ (min):"
    ctrl.Font.Bold = True
    ctrl.Font.Size = 11
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtTempsRecuperation")
    ctrl.Left = 5700
    ctrl.Top = 9800
    ctrl.Width = 1200
    ctrl.Height = 300
    ctrl.Alignment = 2
    ctrl.Font.Bold = True
    ctrl.Font.Size = 11
    ctrl.Text = "0"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.Label", "lblCommentaireRecuperation")
    ctrl.Left = 7200
    ctrl.Top = 9800
    ctrl.Width = 1550
    ctrl.Height = 300
    ctrl.Caption = "COMMENTAIRE:"
    ctrl.Font.Bold = True
    ctrl.Font.Size = 11
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtCommentaireRecuperation")
    ctrl.Left = 8900
    ctrl.Top = 9800
    ctrl.Width = 2600
    ctrl.Height = 300
    ctrl.Font.Size = 10
    ctrl.Visible = True

    ' === RÉSUMÉ (plus d'espace) ===
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitreResume")
    ctrl.Left = 500
    ctrl.Top = 11100
    ctrl.Width = 11000
    ctrl.Height = 400
    ctrl.Caption = "RÉSUMÉ DES PIÈCES RÉCUPÉRÉES"
    ctrl.BackColor = RGB(200, 200, 200)
    ctrl.Font.Bold = True
    ctrl.Font.Size = 12
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.Label", "txtResume")
    ctrl.Left = 500
    ctrl.Top = 11700
    ctrl.Width = 11000
    ctrl.Height = 1200
    ctrl.BackColor = RGB(255, 255, 240)
    ctrl.Caption = "Aucune pièce sélectionnée"
    ctrl.BorderStyle = 1
    ctrl.Alignment = 0
    ctrl.Font.Size = 10
    ctrl.Visible = True
    
    ' === BOUTONS (plus espacés) - UTILISER WITHEVENTS ===
    Set cmdMettreAJourResume = Me.Controls.Add("VB.CommandButton", "cmdMettreAJourResume")
    cmdMettreAJourResume.Left = 500
    cmdMettreAJourResume.Top = 13200
    cmdMettreAJourResume.Width = 2500
    cmdMettreAJourResume.Height = 500
    cmdMettreAJourResume.Caption = "ACTUALISER RÉSUMÉ"
    cmdMettreAJourResume.Font.Bold = True
    cmdMettreAJourResume.Font.Size = 11
    cmdMettreAJourResume.BackColor = RGB(200, 230, 255)
    cmdMettreAJourResume.Visible = True
    
    Set cmdValiderRecuperation = Me.Controls.Add("VB.CommandButton", "cmdValiderRecuperation")
    cmdValiderRecuperation.Left = 4500
    cmdValiderRecuperation.Top = 13200
    cmdValiderRecuperation.Width = 3000
    cmdValiderRecuperation.Height = 600
    cmdValiderRecuperation.Caption = "VALIDER RÉCUPÉRATION"
    cmdValiderRecuperation.Font.Bold = True
    cmdValiderRecuperation.Font.Size = 12
    cmdValiderRecuperation.BackColor = RGB(128, 255, 128)
    cmdValiderRecuperation.Visible = True
    
    Set cmdAnnuler = Me.Controls.Add("VB.CommandButton", "cmdAnnuler")
    cmdAnnuler.Left = 9000
    cmdAnnuler.Top = 13200
    cmdAnnuler.Width = 2500
    cmdAnnuler.Height = 500
    cmdAnnuler.Caption = "ANNULER"
    cmdAnnuler.Font.Bold = True
    cmdAnnuler.Font.Size = 11
    cmdAnnuler.BackColor = RGB(255, 150, 150)
    cmdAnnuler.Visible = True
    
    ' Mettre à jour le résumé initial
    MettreAJourResume
End Sub

Private Sub CreerLignePiece(topPosition As Long, nomPiece As String, codePiece As String, quantiteMax As Integer, index As Integer)
    Dim ctrl As Object
    
    ' CheckBox
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chk" & index)
    ctrl.Left = 800
    ctrl.Top = topPosition
    ctrl.Width = 400
    ctrl.Height = 350
    ctrl.Caption = ""
    ctrl.Visible = True
    
    ' Nom de la pièce
    Set ctrl = Me.Controls.Add("VB.Label", "lblPiece" & index)
    ctrl.Left = 1400
    ctrl.Top = topPosition
    ctrl.Width = 3500
    ctrl.Height = 350
    ctrl.Caption = nomPiece
    ctrl.Font.Bold = True
    ctrl.Font.Size = 11
    ctrl.Visible = True
    
    ' Quantité disponible
    Set ctrl = Me.Controls.Add("VB.Label", "lblDispo" & index)
    ctrl.Left = 5200
    ctrl.Top = topPosition
    ctrl.Width = 1800
    ctrl.Height = 350
    ctrl.Caption = "Dispo: " & quantiteMax
    ctrl.Tag = quantiteMax
    ctrl.Font.Size = 10
    ctrl.Visible = True
    
    ' BOUTONS AVEC WITHEVENTS
    Select Case index
        Case 0
            Set cmdMoins0 = Me.Controls.Add("VB.CommandButton", "cmdMoins0")
            Set cmdPlus0 = Me.Controls.Add("VB.CommandButton", "cmdPlus0")
        Case 1
            Set cmdMoins1 = Me.Controls.Add("VB.CommandButton", "cmdMoins1")
            Set cmdPlus1 = Me.Controls.Add("VB.CommandButton", "cmdPlus1")
        Case 2
            Set cmdMoins2 = Me.Controls.Add("VB.CommandButton", "cmdMoins2")
            Set cmdPlus2 = Me.Controls.Add("VB.CommandButton", "cmdPlus2")
        Case 3
            Set cmdMoins3 = Me.Controls.Add("VB.CommandButton", "cmdMoins3")
            Set cmdPlus3 = Me.Controls.Add("VB.CommandButton", "cmdPlus3")
        Case 4
            Set cmdMoins4 = Me.Controls.Add("VB.CommandButton", "cmdMoins4")
            Set cmdPlus4 = Me.Controls.Add("VB.CommandButton", "cmdPlus4")
        Case 5
            Set cmdMoins5 = Me.Controls.Add("VB.CommandButton", "cmdMoins5")
            Set cmdPlus5 = Me.Controls.Add("VB.CommandButton", "cmdPlus5")
        Case 6
            Set cmdMoins6 = Me.Controls.Add("VB.CommandButton", "cmdMoins6")
            Set cmdPlus6 = Me.Controls.Add("VB.CommandButton", "cmdPlus6")
        Case 7
            Set cmdMoins7 = Me.Controls.Add("VB.CommandButton", "cmdMoins7")
            Set cmdPlus7 = Me.Controls.Add("VB.CommandButton", "cmdPlus7")
        Case 8
            Set cmdMoins8 = Me.Controls.Add("VB.CommandButton", "cmdMoins8")
            Set cmdPlus8 = Me.Controls.Add("VB.CommandButton", "cmdPlus8")
        Case 9
            Set cmdMoins9 = Me.Controls.Add("VB.CommandButton", "cmdMoins9")
            Set cmdPlus9 = Me.Controls.Add("VB.CommandButton", "cmdPlus9")
        Case 10
            Set cmdMoins10 = Me.Controls.Add("VB.CommandButton", "cmdMoins10")
            Set cmdPlus10 = Me.Controls.Add("VB.CommandButton", "cmdPlus10")
        Case 11
            Set cmdMoins11 = Me.Controls.Add("VB.CommandButton", "cmdMoins11")
            Set cmdPlus11 = Me.Controls.Add("VB.CommandButton", "cmdPlus11")
        Case 12
            Set cmdMoins12 = Me.Controls.Add("VB.CommandButton", "cmdMoins12")
            Set cmdPlus12 = Me.Controls.Add("VB.CommandButton", "cmdPlus12")
    End Select
    
    ' Configuration du bouton -
    Dim btnMoins As CommandButton
    Set btnMoins = Me.Controls("cmdMoins" & index)
    btnMoins.Left = 7300
    btnMoins.Top = topPosition
    btnMoins.Width = 500
    btnMoins.Height = 350
    btnMoins.Caption = "-"
    btnMoins.Font.Size = 16
    btnMoins.Font.Bold = True
    btnMoins.Visible = True
    
    ' Configuration du bouton +
    Dim btnPlus As CommandButton
    Set btnPlus = Me.Controls("cmdPlus" & index)
    btnPlus.Left = 9000
    btnPlus.Top = topPosition
    btnPlus.Width = 500
    btnPlus.Height = 350
    btnPlus.Caption = "+"
    btnPlus.Font.Size = 16
    btnPlus.Font.Bold = True
    btnPlus.Visible = True
    
    ' Quantité sélectionnée
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtQte" & index)
    ctrl.Left = 8000
    ctrl.Top = topPosition
    ctrl.Width = 800
    ctrl.Height = 350
    ctrl.Text = "0"
    ctrl.Alignment = 2
    ctrl.Font.Bold = True
    ctrl.Font.Size = 12
    ctrl.Visible = True
End Sub

' === ÉVÉNEMENTS DES BOUTONS - CORRIGÉS ===
Private Sub cmdPlus0_Click()
    AjusterQuantite 0, 1
End Sub
Private Sub cmdMoins0_Click()
    AjusterQuantite 0, -1
End Sub
Private Sub cmdPlus1_Click()
    AjusterQuantite 1, 1
End Sub
Private Sub cmdMoins1_Click()
    AjusterQuantite 1, -1
End Sub
Private Sub cmdPlus2_Click()
    AjusterQuantite 2, 1
End Sub
Private Sub cmdMoins2_Click()
    AjusterQuantite 2, -1
End Sub
Private Sub cmdPlus3_Click()
    AjusterQuantite 3, 1
End Sub
Private Sub cmdMoins3_Click()
    AjusterQuantite 3, -1
End Sub
Private Sub cmdPlus4_Click()
    AjusterQuantite 4, 1
End Sub
Private Sub cmdMoins4_Click()
    AjusterQuantite 4, -1
End Sub
Private Sub cmdPlus5_Click()
    AjusterQuantite 5, 1
End Sub
Private Sub cmdMoins5_Click()
    AjusterQuantite 5, -1
End Sub
Private Sub cmdPlus6_Click()
    AjusterQuantite 6, 1
End Sub
Private Sub cmdMoins6_Click()
    AjusterQuantite 6, -1
End Sub
Private Sub cmdPlus7_Click()
    AjusterQuantite 7, 1
End Sub
Private Sub cmdMoins7_Click()
    AjusterQuantite 7, -1
End Sub
Private Sub cmdPlus8_Click()
    AjusterQuantite 8, 1
End Sub
Private Sub cmdMoins8_Click()
    AjusterQuantite 8, -1
End Sub
Private Sub cmdPlus9_Click()
    AjusterQuantite 9, 1
End Sub
Private Sub cmdMoins9_Click()
    AjusterQuantite 9, -1
End Sub
Private Sub cmdPlus10_Click()
    AjusterQuantite 10, 1
End Sub
Private Sub cmdMoins10_Click()
    AjusterQuantite 10, -1
End Sub
Private Sub cmdPlus11_Click()
    AjusterQuantite 11, 1
End Sub
Private Sub cmdMoins11_Click()
    AjusterQuantite 11, -1
End Sub
Private Sub cmdPlus12_Click()
    AjusterQuantite 12, 1
End Sub
Private Sub cmdMoins12_Click()
    AjusterQuantite 12, -1
End Sub

' === FONCTION D'AJUSTEMENT CORRIGÉE ===
Private Sub AjusterQuantite(index As Integer, ajustement As Integer)
    On Error GoTo ErrorHandler
    
    Dim quantiteActuelle As Integer
    Dim nouvelleQuantite As Integer
    Dim quantiteMax As Integer
    
    ' Récupérer la quantité actuelle
    quantiteActuelle = Val(Me.Controls("txtQte" & index).Text)
    nouvelleQuantite = quantiteActuelle + ajustement
    
    ' RÉCUPÉRER LA QUANTITÉ MAX DEPUIS LE TAG DU LABEL DISPO
    quantiteMax = Val(Me.Controls("lblDispo" & index).Tag)
    
    ' SI PAS DE TAG, UTILISER LES VALEURS PAR DÉFAUT
    If quantiteMax = 0 Then
        Select Case index
            Case 0, 5, 6, 7, 8, 10, 12  ' Pièces uniques
                quantiteMax = 1
            Case 1, 2, 9  ' 2 pièces max
                quantiteMax = 2
            Case 3  ' ETAGERE
                quantiteMax = 4
            Case 4  ' ATTACHE ETAGERE
                quantiteMax = 8
            Case 11  ' VIS DIVERSES
                quantiteMax = 10
            Case Else
                quantiteMax = 1
        End Select
    End If
    
    ' Appliquer les limites
    If nouvelleQuantite < 0 Then nouvelleQuantite = 0
    If nouvelleQuantite > quantiteMax Then nouvelleQuantite = quantiteMax
    
    ' Mettre à jour les contrôles
    Me.Controls("txtQte" & index).Text = CStr(nouvelleQuantite)
    Me.Controls("chk" & index).Value = IIf(nouvelleQuantite > 0, 1, 0)
    
    ' Debug pour vérifier
    Debug.Print "Index " & index & ": " & quantiteActuelle & " -> " & nouvelleQuantite & " (Max: " & quantiteMax & ")"
    
    ' Mise à jour automatique du résumé
    MettreAJourResume
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur lors de l'ajustement de la quantité: " & Err.description, vbCritical
End Sub

' === ÉVÉNEMENTS DES BOUTONS PRINCIPAUX ===
Private Sub cmdMettreAJourResume_Click()
    MettreAJourResume
End Sub

Private Sub cmdValiderRecuperation_Click()
    MettreAJourResume
    
    If InStr(Me.Controls("txtResume").Caption, "AUCUNE PIÈCE") > 0 Then
        MsgBox "Veuillez sélectionner au moins une pièce à récupérer !", vbExclamation
        Exit Sub
    End If
    
    If Me.Controls("chkRecuperationEffectuee").Value = 1 Then
        If Val(Me.Controls("txtTempsRecuperation").Text) <= 0 Then
            MsgBox "Veuillez saisir un temps de récupération valide !", vbExclamation
            Exit Sub
        End If
    End If
    
    Dim message As String
    message = "Confirmer la récupération des pièces sélectionnées ?" & vbCrLf & vbCrLf
    message = message & "Frigo: " & referenceFrigo & vbCrLf
    message = message & "Frigoriste: " & nomFrigoriste
    
    If MsgBox(message, vbYesNo + vbQuestion) = vbYes Then
        SauvegarderRecuperation
        AjouterAuStockPieces
        
        MsgBox "Récupération validée avec succès !" & vbCrLf & "Pièces ajoutées au stock.", vbInformation
        Me.Hide
    End If
End Sub

Private Sub cmdAnnuler_Click()
    If MsgBox("Annuler la récupération des pièces ?", vbYesNo + vbQuestion) = vbYes Then
        Me.Hide
    End If
End Sub

Private Sub MettreAJourResume()
    Dim resumeTexte As String
    Dim nbPiecesRecuperees As Integer
    
    resumeTexte = "RÉCUPÉRATION - " & referenceFrigo & vbCrLf
    resumeTexte = resumeTexte & "Frigoriste: " & nomFrigoriste & " | Date: " & Format(Now, "dd/mm/yyyy hh:nn:ss") & vbCrLf & vbCrLf
    
    Dim pieces(12) As String
    pieces(0) = "PORTE"
    pieces(1) = "JOINT PORTE"
    pieces(2) = "GOND PORTE"
    pieces(3) = "ETAGERE"
    pieces(4) = "ATTACHE ETAGERE"
    pieces(5) = "VENTILATEUR"
    pieces(6) = "PRISE"
    pieces(7) = "CAPOT HAUT"
    pieces(8) = "CAPOT BAS"
    pieces(9) = "LOGO"
    pieces(10) = "POMPE"
    pieces(11) = "VIS DIVERSES"
    pieces(12) = "GRILLE ARRIERE"
    
    For i = 0 To 12
        If Me.Controls("chk" & i).Value = 1 And Val(Me.Controls("txtQte" & i).Text) > 0 Then
            Dim qte As Integer
            qte = Val(Me.Controls("txtQte" & i).Text)
            resumeTexte = resumeTexte & "• " & pieces(i) & " : " & qte & vbCrLf
            nbPiecesRecuperees = nbPiecesRecuperees + qte
        End If
    Next i
    
    If nbPiecesRecuperees = 0 Then
        resumeTexte = resumeTexte & "AUCUNE PIÈCE SÉLECTIONNÉE"
    Else
        resumeTexte = resumeTexte & vbCrLf & "TOTAL : " & nbPiecesRecuperees & " pièces récupérées"
    End If
    
    Me.Controls("txtResume").Caption = resumeTexte
End Sub

Private Sub SauvegarderRecuperation()
    Dim fichier As String
    Dim numeroFichier As Integer
    
    If Dir(App.Path & "\Recuperations", vbDirectory) = "" Then
        MkDir App.Path & "\Recuperations"
    End If
    
    fichier = App.Path & "\Recuperations\Recup_" & referenceFrigo & "_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    numeroFichier = FreeFile
    
    Open fichier For Output As #numeroFichier
    Print #numeroFichier, Me.Controls("txtResume").Caption
    Print #numeroFichier, ""
    If Me.Controls("chkRecuperationEffectuee").Value = 1 Then
        Print #numeroFichier, "Temps de récupération: " & Me.Controls("txtTempsRecuperation").Text & " minutes"
        Print #numeroFichier, "Commentaire: " & Me.Controls("txtCommentaireRecuperation").Text
    End If
    Close #numeroFichier
End Sub

Private Sub AjouterAuStockPieces()
    Dim fichier As String
    Dim numeroFichier As Integer
    
    fichier = App.Path & "\StockPieces.txt"
    numeroFichier = FreeFile
    
    If Dir(fichier) = "" Then
        Open fichier For Output As #numeroFichier
        Print #numeroFichier, "CODE|PIECE|QUANTITE|ETAT|ORIGINE|DATE"
        Close #numeroFichier
    End If
    
    Open fichier For Append As #numeroFichier
    
    Dim codes(12) As String
    codes(0) = "PORTE": codes(1) = "JOINT_PORTE": codes(2) = "GOND_PORTE"
    codes(3) = "ETAGERE": codes(4) = "ATTACHE_ETAGERE": codes(5) = "VENTILATEUR"
    codes(6) = "PRISE": codes(7) = "CAPOT_HAUT": codes(8) = "CAPOT_BAS"
    codes(9) = "LOGO": codes(10) = "POMPE": codes(11) = "VIS_DIVERSES"
    codes(12) = "GRILLE_ARRIERE"
    
    Dim pieces(12) As String
    pieces(0) = "Porte": pieces(1) = "Joint porte": pieces(2) = "Gond porte"
    pieces(3) = "Etagere": pieces(4) = "Attache etagere": pieces(5) = "Ventilateur"
    pieces(6) = "Prise": pieces(7) = "Capot haut": pieces(8) = "Capot bas"
    pieces(9) = "Logo": pieces(10) = "Pompe": pieces(11) = "Vis diverses"
    pieces(12) = "Grille arriere"
    
    For i = 0 To 12
        If Me.Controls("chk" & i).Value = 1 And Val(Me.Controls("txtQte" & i).Text) > 0 Then
            Dim ligne As String
            ligne = codes(i) & "|" & pieces(i) & "|" & Me.Controls("txtQte" & i).Text & "|"
            ligne = ligne & "Recupere|" & referenceFrigo & "|" & Format(Now, "dd/mm/yyyy hh:nn:ss")
            Print #numeroFichier, ligne
        End If
    Next i
    
    Close #numeroFichier
End Sub
