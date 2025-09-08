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
' === FRMRECUPERATIONPIECES.FRM - RÉCUPÉRATION DES PIÈCES ===

Private referenceFrigo As String
Private nomFrigoriste As String

Private Sub Form_Load()
    Me.BackColor = RGB(240, 240, 240)
    Me.Caption = "Récupération Pièces - " & referenceFrigo
    Me.Width = 15000
    Me.Height = 14000
    
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
    
    ' Titre
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitre")
    ctrl.Left = 240
    ctrl.Top = 120
    ctrl.Width = 12000
    ctrl.Height = 400
    ctrl.Caption = "RÉCUPÉRATION DES PIÈCES - FRIGO HS"
    ctrl.BackColor = RGB(255, 100, 100)
    ctrl.ForeColor = RGB(255, 255, 255)
    ctrl.Font.Size = 16
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Infos frigo
    Set ctrl = Me.Controls.Add("VB.Label", "lblInfoFrigo")
    ctrl.Left = 240
    ctrl.Top = 600
    ctrl.Width = 12000
    ctrl.Height = 300
    ctrl.Caption = "Référence: " & referenceFrigo & " | Frigoriste: " & nomFrigoriste & " | Date: " & Format(Now, "dd/mm/yyyy")
    ctrl.BackColor = RGB(255, 255, 200)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Instructions
    Set ctrl = Me.Controls.Add("VB.Label", "lblInstructions")
    ctrl.Left = 240
    ctrl.Top = 1000
    ctrl.Width = 12000
    ctrl.Height = 400
    ctrl.Caption = "Sélectionnez les pièces récupérables et ajustez les quantités avec les boutons + et -"
    ctrl.Font.Size = 11
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' === LISTE DES PIÈCES RÉCUPÉRABLES ===
    
    ' Compresseur
    CreerLignePiece 1600, "Compresseur", "COMP", 1, 0
    
    ' Éclairage LED
    CreerLignePiece 2000, "Éclairage LED", "LED", 2, 1
    
    ' Vitre
    CreerLignePiece 2400, "Vitre principale", "VITRE", 1, 2
    
    ' Thermostat
    CreerLignePiece 2800, "Thermostat digital", "THERMO", 1, 3
    
    ' Joints de porte
    CreerLignePiece 3200, "Joints de porte", "JOINT", 4, 4
    
    ' Grilles
    CreerLignePiece 3600, "Grilles métalliques", "GRILLE", 3, 5
    
    ' Ventilateur
    CreerLignePiece 4000, "Ventilateur", "VENTILO", 1, 6
    
    ' Capot arrière
    CreerLignePiece 4400, "Capot arrière", "CAPOT", 1, 7
    
    ' Pieds réglables
    CreerLignePiece 4800, "Pieds réglables", "PIED", 4, 8
    
    ' Câblage électrique
    CreerLignePiece 5200, "Câblage électrique", "CABLE", 1, 9
    
    ' === RÉSUMÉ ET VALIDATION ===
    
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitreResume")
    ctrl.Left = 240
    ctrl.Top = 5800
    ctrl.Width = 12000
    ctrl.Height = 300
    ctrl.Caption = "=== RÉSUMÉ DES PIÈCES RÉCUPÉRÉES ==="
    ctrl.BackColor = RGB(200, 200, 200)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtResume")
    ctrl.Left = 240
    ctrl.Top = 6200
    ctrl.Width = 12000
    ctrl.Height = 1200
    ctrl.BackColor = RGB(255, 255, 240)
    ctrl.Text = "Aucune pièce sélectionnée"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdMettreAJourResume")
    ctrl.Left = 240
    ctrl.Top = 7600
    ctrl.Width = 2000
    ctrl.Height = 400
    ctrl.Caption = "Mettre à jour résumé"
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdValiderRecuperation")
    ctrl.Left = 4000
    ctrl.Top = 7600
    ctrl.Width = 2500
    ctrl.Height = 400
    ctrl.Caption = "VALIDER RÉCUPÉRATION"
    ctrl.Font.Bold = True
    ctrl.Font.Size = 12
    ctrl.BackColor = RGB(128, 255, 128)
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdAnnuler")
    ctrl.Left = 7000
    ctrl.Top = 7600
    ctrl.Width = 2000
    ctrl.Height = 400
    ctrl.Caption = "ANNULER"
    ctrl.Font.Bold = True
    ctrl.BackColor = RGB(255, 128, 128)
    ctrl.Visible = True
    
    ' Mettre à jour le résumé initial
    MettreAJourResume
End Sub

Private Sub CreerLignePiece(topPosition As Long, nomPiece As String, codePiece As String, quantiteMax As Integer, index As Integer)
    Dim ctrl As Object
    
    ' CheckBox pour sélectionner la pièce
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chk" & index)
    ctrl.Left = 480
    ctrl.Top = topPosition
    ctrl.Width = 300
    ctrl.Height = 300
    ctrl.Caption = ""
    ctrl.Visible = True
    
    ' Nom de la pièce
    Set ctrl = Me.Controls.Add("VB.Label", "lblPiece" & index)
    ctrl.Left = 840
    ctrl.Top = topPosition
    ctrl.Width = 3000
    ctrl.Height = 300
    ctrl.Caption = nomPiece & " (" & codePiece & ")"
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    ' Quantité disponible
    Set ctrl = Me.Controls.Add("VB.Label", "lblDispo" & index)
    ctrl.Left = 4000
    ctrl.Top = topPosition
    ctrl.Width = 1500
    ctrl.Height = 300
    ctrl.Caption = "Dispo: " & quantiteMax
    ctrl.Visible = True
    
    ' Bouton -
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdMoins" & index)
    ctrl.Left = 5800
    ctrl.Top = topPosition
    ctrl.Width = 400
    ctrl.Height = 300
    ctrl.Caption = "-"
    ctrl.Font.Size = 14
    ctrl.Tag = index
    ctrl.Visible = True
    
    ' Quantité sélectionnée
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtQte" & index)
    ctrl.Left = 6240
    ctrl.Top = topPosition
    ctrl.Width = 600
    ctrl.Height = 300
    ctrl.Text = "0"
    ctrl.Alignment = 2
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    ' Bouton +
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdPlus" & index)
    ctrl.Left = 6880
    ctrl.Top = topPosition
    ctrl.Width = 400
    ctrl.Height = 300
    ctrl.Caption = "+"
    ctrl.Font.Size = 14
    ctrl.Tag = index
    ctrl.Visible = True
End Sub

Private Function CalculerPrixPiece(codePiece As String, Quantite As Integer) As String
    Dim prixUnitaire As Double
    
    Select Case codePiece
        Case "COMP": prixUnitaire = 450
        Case "LED": prixUnitaire = 35
        Case "VITRE": prixUnitaire = 120
        Case "THERMO": prixUnitaire = 85
        Case "JOINT": prixUnitaire = 25
        Case "GRILLE": prixUnitaire = 40
        Case "VENTILO": prixUnitaire = 95
        Case "CAPOT": prixUnitaire = 75
        Case "PIED": prixUnitaire = 15
        Case "CABLE": prixUnitaire = 65
        Case Else: prixUnitaire = 20
    End Select
    
    CalculerPrixPiece = Format(prixUnitaire * Quantite, "0.00")
End Function


Private Sub cmdPlus0_Click()
    AjusterQuantite 0, 1
    MettreAJourResume
End Sub
Private Sub cmdMoins0_Click()
    AjusterQuantite 0, -1
    MettreAJourResume
End Sub
Private Sub cmdPlus1_Click()
    AjusterQuantite 1, 1
    MettreAJourResume
End Sub
Private Sub cmdMoins1_Click()
    AjusterQuantite 1, -1
    MettreAJourResume
End Sub
Private Sub cmdPlus2_Click()
    AjusterQuantite 2, 1
    MettreAJourResume
End Sub
Private Sub cmdMoins2_Click()
    AjusterQuantite 2, -1
    MettreAJourResume
End Sub
Private Sub cmdPlus3_Click()
    AjusterQuantite 3, 1
    MettreAJourResume
End Sub
Private Sub cmdMoins3_Click()
    AjusterQuantite 3, -1
    MettreAJourResume
End Sub
Private Sub cmdPlus4_Click()
    AjusterQuantite 4, 1
    MettreAJourResume
End Sub
Private Sub cmdMoins4_Click()
    AjusterQuantite 4, -1
    MettreAJourResume
End Sub
Private Sub cmdPlus5_Click()
    AjusterQuantite 5, 1
    MettreAJourResume
End Sub
Private Sub cmdMoins5_Click()
    AjusterQuantite 5, -1
    MettreAJourResume
End Sub
Private Sub cmdPlus6_Click()
    AjusterQuantite 6, 1
    MettreAJourResume
End Sub
Private Sub cmdMoins6_Click()
    AjusterQuantite 6, -1
    MettreAJourResume
End Sub
Private Sub cmdPlus7_Click()
    AjusterQuantite 7, 1
    MettreAJourResume
End Sub
Private Sub cmdMoins7_Click()
    AjusterQuantite 7, -1
    MettreAJourResume
End Sub
Private Sub cmdPlus8_Click()
    AjusterQuantite 8, 1
    MettreAJourResume
End Sub
Private Sub cmdMoins8_Click()
    AjusterQuantite 8, -1
    MettreAJourResume
End Sub
Private Sub cmdPlus9_Click()
    AjusterQuantite 9, 1
    MettreAJourResume
End Sub
Private Sub cmdMoins9_Click()
    AjusterQuantite 9, -1
    MettreAJourResume
End Sub

Private Sub AjusterQuantite(index As Integer, ajustement As Integer)
    Dim quantiteActuelle As Integer
    Dim nouvelleQuantite As Integer
    Dim quantiteMax As Integer
    
    quantiteActuelle = Val(Me.Controls("txtQte" & index).Text)
    nouvelleQuantite = quantiteActuelle + ajustement
    
    ' Définir les limites selon la pièce
    Select Case index
        Case 0, 2, 3, 6, 7, 9: quantiteMax = 1 ' Pièces uniques
        Case 1: quantiteMax = 2 ' LEDs
        Case 5: quantiteMax = 3 ' Grilles
        Case 4, 8: quantiteMax = 4 ' Joints et pieds
    End Select
    
    If nouvelleQuantite < 0 Then nouvelleQuantite = 0
    If nouvelleQuantite > quantiteMax Then nouvelleQuantite = quantiteMax
    
    Me.Controls("txtQte" & index).Text = nouvelleQuantite
    
    ' Cocher automatiquement si quantité > 0
    Me.Controls("chk" & index).Value = IIf(nouvelleQuantite > 0, 1, 0)
    
    ' Mettre à jour le prix
    Dim codePiece As String
    Select Case index
        Case 0: codePiece = "COMP"
        Case 1: codePiece = "LED"
        Case 2: codePiece = "VITRE"
        Case 3: codePiece = "THERMO"
        Case 4: codePiece = "JOINT"
        Case 5: codePiece = "GRILLE"
        Case 6: codePiece = "VENTILO"
        Case 7: codePiece = "CAPOT"
        Case 8: codePiece = "PIED"
        Case 9: codePiece = "CABLE"
    End Select
    
    Me.Controls("lblPrix" & index).Caption = "0.00€"
End Sub

Private Sub cmdMettreAJourResume_Click()
    MettreAJourResume
End Sub

Private Sub MettreAJourResume()
    Dim resumeTexte As String
    Dim nbPiecesRecuperees As Integer
    
    resumeTexte = "RÉCUPÉRATION - " & referenceFrigo & " | "
    resumeTexte = resumeTexte & "Frigoriste: " & nomFrigoriste & " | Date: " & Format(Now, "dd/mm/yyyy hh:nn:ss") & " | "
    
    Dim pieces(9) As String
    pieces(0) = "Compresseur"
    pieces(1) = "Éclairage LED"
    pieces(2) = "Vitre principale"
    pieces(3) = "Thermostat digital"
    pieces(4) = "Joints de porte"
    pieces(5) = "Grilles métalliques"
    pieces(6) = "Ventilateur"
    pieces(7) = "Capot arrière"
    pieces(8) = "Pieds réglables"
    pieces(9) = "Câblage électrique"
    
    Dim codes(9) As String
    codes(0) = "COMP"
    codes(1) = "LED"
    codes(2) = "VITRE"
    codes(3) = "THERMO"
    codes(4) = "JOINT"
    codes(5) = "GRILLE"
    codes(6) = "VENTILO"
    codes(7) = "CAPOT"
    codes(8) = "PIED"
    codes(9) = "CABLE"
    
    For i = 0 To 9
        If Me.Controls("chk" & i).Value = 1 And Val(Me.Controls("txtQte" & i).Text) > 0 Then
            Dim qte As Integer
            
            qte = Val(Me.Controls("txtQte" & i).Text)
            
            resumeTexte = resumeTexte & pieces(i) & " (" & codes(i) & "): " & qte & " | "
            
            nbPiecesRecuperees = nbPiecesRecuperees + qte
        End If
    Next i
    
    If nbPiecesRecuperees = 0 Then
        resumeTexte = resumeTexte & "AUCUNE PIÈCE SÉLECTIONNÉE"
    Else
        resumeTexte = resumeTexte & "TOTAL: " & nbPiecesRecuperees & " pièces récupérées"
    End If
    
    Me.Controls("txtResume").Text = resumeTexte
End Sub

Private Sub cmdValiderRecuperation_Click()
    MettreAJourResume
    
    If InStr(Me.Controls("txtResume").Text, "AUCUNE PIÈCE") > 0 Then
        MsgBox "Veuillez sélectionner au moins une pièce à récupérer !", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Confirmer la récupération des pièces sélectionnées ?", vbYesNo + vbQuestion) = vbYes Then
        SauvegarderRecuperation
        AjouterAuStockPieces
        
        MsgBox "Récupération validée avec succès !" & vbCrLf & "Les pièces ont été ajoutées au stock.", vbInformation
        Me.Hide
    End If
End Sub

Private Sub SauvegarderRecuperation()
    Dim fichier As String
    Dim numeroFichier As Integer
    
    ' Créer le répertoire s'il n'existe pas
    If Dir(App.Path & "\Recuperations", vbDirectory) = "" Then
        MkDir App.Path & "\Recuperations"
    End If
    
    fichier = App.Path & "\Recuperations\Recup_" & referenceFrigo & "_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    numeroFichier = FreeFile
    
    Open fichier For Output As #numeroFichier
    Print #numeroFichier, Me.Controls("txtResume").Text
    Close #numeroFichier
End Sub

Private Sub AjouterAuStockPieces()
    Dim fichier As String
    Dim numeroFichier As Integer
    
    fichier = App.Path & "\StockPieces.txt"
    numeroFichier = FreeFile
    
    ' En-tête si fichier n'existe pas
    If Dir(fichier) = "" Then
        Open fichier For Output As #numeroFichier
        Print #numeroFichier, "CODE|PIECE|QUANTITE|ETAT|ORIGINE|DATE|PRIX"
        Close #numeroFichier
    End If
    
    Open fichier For Append As #numeroFichier
    
    Dim codes(9) As String
    codes(0) = "COMP": codes(1) = "LED": codes(2) = "VITRE": codes(3) = "THERMO": codes(4) = "JOINT"
    codes(5) = "GRILLE": codes(6) = "VENTILO": codes(7) = "CAPOT": codes(8) = "PIED": codes(9) = "CABLE"
    
    Dim pieces(9) As String
    pieces(0) = "Compresseur": pieces(1) = "Eclairage LED": pieces(2) = "Vitre principale"
    pieces(3) = "Thermostat digital": pieces(4) = "Joints de porte": pieces(5) = "Grilles metalliques"
    pieces(6) = "Ventilateur": pieces(7) = "Capot arriere": pieces(8) = "Pieds reglables": pieces(9) = "Cablage electrique"
    
    For i = 0 To 9
        If Me.Controls("chk" & i).Value = 1 And Val(Me.Controls("txtQte" & i).Text) > 0 Then
            Dim ligne As String
            ligne = codes(i) & "|" & pieces(i) & "|" & Me.Controls("txtQte" & i).Text & "|"
            ligne = ligne & "Recupere" & "|" & referenceFrigo & "|"
            ligne = ligne & Format(Now, "dd/mm/yyyy hh:nn:ss") & "|0.00"
            Print #numeroFichier, ligne
        End If
    Next i
    
    Close #numeroFichier
End Sub

Private Sub cmdAnnuler_Click()
    If MsgBox("Annuler la récupération des pièces ?", vbYesNo + vbQuestion) = vbYes Then
        Me.Hide
    End If
End Sub
