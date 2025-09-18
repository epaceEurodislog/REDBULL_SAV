VERSION 5.00
Begin VB.Form frmAffectationPieces 
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
Attribute VB_Name = "frmAffectationPieces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' === FRMAFFECTATIONPIECES.FRM - AFFECTATION DES PIÈCES ===

Private referenceFrigoReparable As String
Private numeroSerieFrigo As String
Private nomFrigoriste As String
Private stockPieces() As String ' Pour stocker les données du stock
Private nombrePiecesAffichees As Integer

' DÉCLARATIONS WITHEVENTS POUR LES BOUTONS +/-
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
Private WithEvents cmdPlus13 As CommandButton
Attribute cmdPlus13.VB_VarHelpID = -1
Private WithEvents cmdMoins13 As CommandButton
Attribute cmdMoins13.VB_VarHelpID = -1
Private WithEvents cmdPlus14 As CommandButton
Attribute cmdPlus14.VB_VarHelpID = -1
Private WithEvents cmdMoins14 As CommandButton
Attribute cmdMoins14.VB_VarHelpID = -1
Private WithEvents cmdPlus15 As CommandButton
Attribute cmdPlus15.VB_VarHelpID = -1
Private WithEvents cmdMoins15 As CommandButton
Attribute cmdMoins15.VB_VarHelpID = -1
Private WithEvents cmdPlus16 As CommandButton
Attribute cmdPlus16.VB_VarHelpID = -1
Private WithEvents cmdMoins16 As CommandButton
Attribute cmdMoins16.VB_VarHelpID = -1
Private WithEvents cmdPlus17 As CommandButton
Attribute cmdPlus17.VB_VarHelpID = -1
Private WithEvents cmdMoins17 As CommandButton
Attribute cmdMoins17.VB_VarHelpID = -1
Private WithEvents cmdPlus18 As CommandButton
Attribute cmdPlus18.VB_VarHelpID = -1
Private WithEvents cmdMoins18 As CommandButton
Attribute cmdMoins18.VB_VarHelpID = -1
Private WithEvents cmdPlus19 As CommandButton
Attribute cmdPlus19.VB_VarHelpID = -1
Private WithEvents cmdMoins19 As CommandButton
Attribute cmdMoins19.VB_VarHelpID = -1

Private Sub Form_Load()
    Me.BackColor = RGB(240, 240, 240)
    Me.Caption = "Affectation Pièces - " & referenceFrigoReparable
    Me.Width = 15500
    Me.Height = 12000
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    CreerInterfaceAffectation
    ChargerStockPieces
End Sub

Public Sub InitialiserAvecFrigo(reference As String, numeroSerie As String, frigoriste As String)
    referenceFrigoReparable = reference
    numeroSerieFrigo = numeroSerie
    nomFrigoriste = frigoriste
    Me.Caption = "Affectation Pièces - " & referenceFrigoReparable & " (" & numeroSerie & ")"
End Sub
Private Sub CreerInterfaceAffectation()
    Dim ctrl As Object
    
    ' Titre principal
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitre")
    ctrl.Left = 240
    ctrl.Top = 120
    ctrl.Width = 12000
    ctrl.Height = 400
    ctrl.Caption = "AFFECTATION DES PIÈCES AU FRIGO RÉPARABLE"
    ctrl.BackColor = RGB(100, 150, 255)
    ctrl.ForeColor = RGB(255, 255, 255)
    ctrl.Font.Size = 16
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' Informations frigo cible
    Set ctrl = Me.Controls.Add("VB.Label", "lblInfoCible")
    ctrl.Left = 240
    ctrl.Top = 600
    ctrl.Width = 12000
    ctrl.Height = 300
    ctrl.Caption = "FRIGO: " & referenceFrigoReparable & " | N° SÉRIE: " & numeroSerieFrigo & " | FRIGORISTE: " & nomFrigoriste & " | " & Format(Now, "dd/mm/yyyy hh:nn:ss")
    ctrl.BackColor = RGB(200, 255, 200)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' Instructions
    Set ctrl = Me.Controls.Add("VB.Label", "lblInstructions")
    ctrl.Left = 240
    ctrl.Top = 960
    ctrl.Width = 12000
    ctrl.Height = 300
    ctrl.Caption = "Sélectionnez les pièces du stock à affecter au frigo réparable. Les quantités seront automatiquement déduites du stock disponible."
    ctrl.Font.Size = 10
    ctrl.Alignment = 2
    ctrl.BackColor = RGB(255, 255, 200)
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' === SECTION STOCK DISPONIBLE ===
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitreStock")
    ctrl.Left = 240
    ctrl.Top = 1320
    ctrl.Width = 12000
    ctrl.Height = 300
    ctrl.Caption = "STOCK DE PIÈCES DISPONIBLES"
    ctrl.BackColor = RGB(255, 150, 50)
    ctrl.ForeColor = RGB(255, 255, 255)
    ctrl.Font.Bold = True
    ctrl.Font.Size = 12
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' En-têtes colonnes stock
    CreerEnTetesStock
    
    ' === SECTION PIÈCES SÉLECTIONNÉES ===
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitreSelection")
    ctrl.Left = 240
    ctrl.Top = 6400
    ctrl.Width = 12000
    ctrl.Height = 300
    ctrl.Caption = "PIÈCES SÉLECTIONNÉES POUR AFFECTATION"
    ctrl.BackColor = RGB(50, 200, 50)
    ctrl.ForeColor = RGB(255, 255, 255)
    ctrl.Font.Bold = True
    ctrl.Font.Size = 12
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' Label pour affichage des pièces sélectionnées
    Set ctrl = Me.Controls.Add("VB.Label", "txtPiecesSelectionnees")
    ctrl.Left = 240
    ctrl.Top = 6760
    ctrl.Width = 12000
    ctrl.Height = 1000
    ctrl.BackColor = RGB(255, 255, 240)
    ctrl.Font.Size = 10
    ctrl.Caption = "Aucune pièce sélectionnée pour le moment..."
    ctrl.BorderStyle = 1
    ctrl.Alignment = 0
    ctrl.Visible = True
    
    ' Informations de résumé
    Set ctrl = Me.Controls.Add("VB.Label", "lblResume")
    ctrl.Left = 240
    ctrl.Top = 7820
    ctrl.Width = 6000
    ctrl.Height = 300
    ctrl.Caption = "PIÈCES: 0"
    ctrl.Font.Bold = True
    ctrl.Font.Size = 11
    ctrl.BackColor = RGB(255, 255, 200)
    ctrl.BorderStyle = 1
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Boutons d'action
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdActualiser")
    ctrl.Left = 6600
    ctrl.Top = 7820
    ctrl.Width = 1500
    ctrl.Height = 300
    ctrl.Caption = "Actualiser"
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdValiderAffectation")
    ctrl.Left = 8200
    ctrl.Top = 7820
    ctrl.Width = 2000
    ctrl.Height = 300
    ctrl.Caption = "VALIDER AFFECTATION"
    ctrl.Font.Bold = True
    ctrl.Font.Size = 10
    ctrl.BackColor = RGB(128, 255, 128)
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdAnnuler")
    ctrl.Left = 10400
    ctrl.Top = 7820
    ctrl.Width = 1500
    ctrl.Height = 300
    ctrl.Caption = "ANNULER"
    ctrl.Font.Bold = True
    ctrl.BackColor = RGB(255, 128, 128)
    ctrl.Visible = True
    
    ' === SECTION TEMPS DE RÉPARATION ===
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitreTempsReparation")
    ctrl.Left = 240
    ctrl.Top = 8200
    ctrl.Width = 12000
    ctrl.Height = 300
    ctrl.Caption = "TEMPS DE RÉPARATION"
    ctrl.BackColor = RGB(100, 150, 255)
    ctrl.ForeColor = RGB(255, 255, 255)
    ctrl.Font.Bold = True
    ctrl.Font.Size = 12
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkReparationEffectuee")
    ctrl.Left = 500
    ctrl.Top = 8600
    ctrl.Width = 2000
    ctrl.Caption = "RÉPARATION EFFECTUÉE"
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.Label", "lblTempsReparation")
    ctrl.Left = 2700
    ctrl.Top = 8600
    ctrl.Width = 1500
    ctrl.Caption = "TEMPS PASSÉ (min):"
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtTempsReparation")
    ctrl.Left = 4300
    ctrl.Top = 8600
    ctrl.Width = 1000
    ctrl.Height = 300
    ctrl.Alignment = 2
    ctrl.Font.Bold = True
    ctrl.Text = "0"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.Label", "lblCommentaireReparation")
    ctrl.Left = 5500
    ctrl.Top = 8600
    ctrl.Width = 1500
    ctrl.Caption = "COMMENTAIRE:"
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtCommentaireReparation")
    ctrl.Left = 7100
    ctrl.Top = 8600
    ctrl.Width = 5000
    ctrl.Height = 300
    ctrl.Visible = True
    
    ' Bouton validation finale repositionné
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdValiderReparationComplete")
    ctrl.Left = 4000
    ctrl.Top = 9100
    ctrl.Width = 3000
    ctrl.Height = 400
    ctrl.Caption = "VALIDER RÉPARATION COMPLÈTE"
    ctrl.Font.Bold = True
    ctrl.Font.Size = 12
    ctrl.BackColor = RGB(50, 255, 50)
    ctrl.Visible = True
End Sub

Private Sub CreerEnTetesStock()
    Dim ctrl As Object
    Dim topPos As Long
    topPos = 1680
    
    Dim positions(5) As Long
    Dim largeurs(5) As Long
    Dim titres(5) As String
    
    ' Nouvelles positions pour prendre toute la largeur (12000 pixels)
    positions(0) = 240   ' CODE
    positions(1) = 1400  ' DÉSIGNATION
    positions(2) = 6200  ' QUANTITÉ DISPO
    positions(3) = 7800  ' BOUTON -
    positions(4) = 9000  ' QTÉ AFFECT
    positions(5) = 10600 ' BOUTON +
    
    ' Nouvelles largeurs pour s'étendre sur toute la largeur
    largeurs(0) = 1160   ' CODE (élargi)
    largeurs(1) = 4800   ' DÉSIGNATION (beaucoup plus large)
    largeurs(2) = 1600   ' QUANTITÉ DISPO (élargi)
    largeurs(3) = 1200   ' BOUTON -
    largeurs(4) = 1600   ' QTÉ AFFECT (élargi)
    largeurs(5) = 1200   ' BOUTON +
    
    titres(0) = "CODE"
    titres(1) = "DÉSIGNATION PIÈCE"
    titres(2) = "QTÉ DISPO"
    titres(3) = "-"
    titres(4) = "QTÉ AFFECT"
    titres(5) = "+"
    
    For i = 0 To 5
        Set ctrl = Me.Controls.Add("VB.Label", "lblHeader" & i)
        ctrl.Left = positions(i)
        ctrl.Top = topPos
        ctrl.Width = largeurs(i)
        ctrl.Height = 280
        ctrl.Caption = titres(i)
        ctrl.BackColor = RGB(120, 120, 120)
        ctrl.ForeColor = RGB(255, 255, 255)
        ctrl.Font.Bold = True
        ctrl.Font.Size = 9
        ctrl.Alignment = 2
        ctrl.BorderStyle = 1
        ctrl.Visible = True
    Next i
End Sub

Private Sub ChargerStockPieces()
    Dim fichier As String
    Dim numeroFichier As Integer
    Dim ligne As String
    Dim elements() As String
    Dim topPosition As Long
    Dim compteur As Integer
    
    fichier = App.Path & "\StockPieces.txt"
    topPosition = 1960
    compteur = 0
    nombrePiecesAffichees = 0
    
    If Dir(fichier) = "" Then
        MsgBox "Fichier StockPieces.txt introuvable !" & vbCrLf & "Aucune pièce disponible pour l'affectation.", vbExclamation
        Exit Sub
    End If
    
    numeroFichier = FreeFile
    Open fichier For Input As #numeroFichier
    
    ' Ignorer l'en-tête
    If Not EOF(numeroFichier) Then Line Input #numeroFichier, ligne
    
    ' Redimensionner le tableau pour stocker les données
    ReDim stockPieces(20, 6) ' Maximum 20 pièces, 7 colonnes
    
    Do While Not EOF(numeroFichier) And compteur < 20
        Line Input #numeroFichier, ligne
        If Len(Trim(ligne)) > 0 Then
            elements = Split(ligne, "|")
            If UBound(elements) >= 6 And Val(elements(2)) > 0 Then ' Seulement si quantité > 0
                ' Stocker les données
                For j = 0 To 6
                    stockPieces(compteur, j) = elements(j)
                Next j
                
                ' MODIFICATION : On ne passe plus état, date, prix - seulement origine (elements(4))
                CreerLignePieceStock topPosition, elements(0), elements(1), elements(2), elements(4), compteur
                topPosition = topPosition + 300
                compteur = compteur + 1
            End If
        End If
    Loop
    
    Close #numeroFichier
    nombrePiecesAffichees = compteur
    
    If compteur = 0 Then
        MsgBox "Aucune pièce disponible en stock pour l'affectation.", vbInformation
    Else
        MsgBox compteur & " pièce(s) disponible(s) chargée(s) du stock.", vbInformation
    End If
End Sub

Private Sub CreerLignePieceStock(topPos As Long, code As String, piece As String, quantite As String, origine As String, index As Integer)
    Dim ctrl As Object
    
    Dim positions(5) As Long
    Dim largeurs(5) As Long
    
    ' Nouvelles positions pour étendre sur toute la largeur (12000 pixels)
    positions(0) = 240   ' CODE
    positions(1) = 1400  ' DÉSIGNATION
    positions(2) = 6200  ' QUANTITÉ DISPO
    positions(3) = 7800  ' BOUTON -
    positions(4) = 9000  ' QTÉ AFFECT
    positions(5) = 10600 ' BOUTON +
    
    ' Nouvelles largeurs pour s'étendre sur toute la largeur
    largeurs(0) = 1160   ' CODE (élargi)
    largeurs(1) = 4800   ' DÉSIGNATION (beaucoup plus large)
    largeurs(2) = 1600   ' QUANTITÉ DISPO (élargi)
    largeurs(3) = 1200   ' BOUTON -
    largeurs(4) = 1600   ' QTÉ AFFECT (élargi)
    largeurs(5) = 1200   ' BOUTON +
    
    ' Code de la pièce
    Set ctrl = Me.Controls.Add("VB.Label", "lblCode" & index)
    ctrl.Left = positions(0)
    ctrl.Top = topPos
    ctrl.Width = largeurs(0)
    ctrl.Height = 280
    ctrl.Caption = code
    ctrl.BackColor = RGB(255, 255, 255)
    ctrl.BorderStyle = 1
    ctrl.Font.Bold = True
    ctrl.Font.Size = 9
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Désignation de la pièce
    Set ctrl = Me.Controls.Add("VB.Label", "lblPiece" & index)
    ctrl.Left = positions(1)
    ctrl.Top = topPos
    ctrl.Width = largeurs(1)
    ctrl.Height = 280
    ctrl.Caption = piece
    ctrl.BackColor = RGB(255, 255, 255)
    ctrl.BorderStyle = 1
    ctrl.Font.Size = 9
    ctrl.Visible = True
    
    ' Quantité disponible
    Set ctrl = Me.Controls.Add("VB.Label", "lblQteDispo" & index)
    ctrl.Left = positions(2)
    ctrl.Top = topPos
    ctrl.Width = largeurs(2)
    ctrl.Height = 280
    ctrl.Caption = quantite
    ctrl.BackColor = RGB(255, 255, 255)
    ctrl.BorderStyle = 1
    ctrl.Font.Bold = True
    ctrl.Font.Size = 10
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' BOUTONS AVEC WITHEVENTS
    Select Case index
        Case 0
            Set cmdMoins0 = Me.Controls.Add("VB.CommandButton", "cmdMoins" & index)
            Set cmdPlus0 = Me.Controls.Add("VB.CommandButton", "cmdPlus" & index)
        Case 1
            Set cmdMoins1 = Me.Controls.Add("VB.CommandButton", "cmdMoins" & index)
            Set cmdPlus1 = Me.Controls.Add("VB.CommandButton", "cmdPlus" & index)
        Case 2
            Set cmdMoins2 = Me.Controls.Add("VB.CommandButton", "cmdMoins" & index)
            Set cmdPlus2 = Me.Controls.Add("VB.CommandButton", "cmdPlus" & index)
        Case 3
            Set cmdMoins3 = Me.Controls.Add("VB.CommandButton", "cmdMoins" & index)
            Set cmdPlus3 = Me.Controls.Add("VB.CommandButton", "cmdPlus" & index)
        Case 4
            Set cmdMoins4 = Me.Controls.Add("VB.CommandButton", "cmdMoins" & index)
            Set cmdPlus4 = Me.Controls.Add("VB.CommandButton", "cmdPlus" & index)
        Case 5
            Set cmdMoins5 = Me.Controls.Add("VB.CommandButton", "cmdMoins" & index)
            Set cmdPlus5 = Me.Controls.Add("VB.CommandButton", "cmdPlus" & index)
        Case 6
            Set cmdMoins6 = Me.Controls.Add("VB.CommandButton", "cmdMoins" & index)
            Set cmdPlus6 = Me.Controls.Add("VB.CommandButton", "cmdPlus" & index)
        Case 7
            Set cmdMoins7 = Me.Controls.Add("VB.CommandButton", "cmdMoins" & index)
            Set cmdPlus7 = Me.Controls.Add("VB.CommandButton", "cmdPlus" & index)
        Case 8
            Set cmdMoins8 = Me.Controls.Add("VB.CommandButton", "cmdMoins" & index)
            Set cmdPlus8 = Me.Controls.Add("VB.CommandButton", "cmdPlus" & index)
        Case 9
            Set cmdMoins9 = Me.Controls.Add("VB.CommandButton", "cmdMoins" & index)
            Set cmdPlus9 = Me.Controls.Add("VB.CommandButton", "cmdPlus" & index)
        Case 10
            Set cmdMoins10 = Me.Controls.Add("VB.CommandButton", "cmdMoins" & index)
            Set cmdPlus10 = Me.Controls.Add("VB.CommandButton", "cmdPlus" & index)
        Case 11
            Set cmdMoins11 = Me.Controls.Add("VB.CommandButton", "cmdMoins" & index)
            Set cmdPlus11 = Me.Controls.Add("VB.CommandButton", "cmdPlus" & index)
        Case 12
            Set cmdMoins12 = Me.Controls.Add("VB.CommandButton", "cmdMoins" & index)
            Set cmdPlus12 = Me.Controls.Add("VB.CommandButton", "cmdPlus" & index)
        Case 13
            Set cmdMoins13 = Me.Controls.Add("VB.CommandButton", "cmdMoins" & index)
            Set cmdPlus13 = Me.Controls.Add("VB.CommandButton", "cmdPlus" & index)
        Case 14
            Set cmdMoins14 = Me.Controls.Add("VB.CommandButton", "cmdMoins" & index)
            Set cmdPlus14 = Me.Controls.Add("VB.CommandButton", "cmdPlus" & index)
        Case 15
            Set cmdMoins15 = Me.Controls.Add("VB.CommandButton", "cmdMoins" & index)
            Set cmdPlus15 = Me.Controls.Add("VB.CommandButton", "cmdPlus" & index)
        Case 16
            Set cmdMoins16 = Me.Controls.Add("VB.CommandButton", "cmdMoins" & index)
            Set cmdPlus16 = Me.Controls.Add("VB.CommandButton", "cmdPlus" & index)
        Case 17
            Set cmdMoins17 = Me.Controls.Add("VB.CommandButton", "cmdMoins" & index)
            Set cmdPlus17 = Me.Controls.Add("VB.CommandButton", "cmdPlus" & index)
        Case 18
            Set cmdMoins18 = Me.Controls.Add("VB.CommandButton", "cmdMoins" & index)
            Set cmdPlus18 = Me.Controls.Add("VB.CommandButton", "cmdPlus" & index)
        Case 19
            Set cmdMoins19 = Me.Controls.Add("VB.CommandButton", "cmdMoins" & index)
            Set cmdPlus19 = Me.Controls.Add("VB.CommandButton", "cmdPlus" & index)
    End Select
    
    ' Configuration bouton -
    Dim btnMoins As CommandButton
    Set btnMoins = Me.Controls("cmdMoins" & index)
    btnMoins.Left = positions(3)
    btnMoins.Top = topPos + 20
    btnMoins.Width = largeurs(3)
    btnMoins.Height = 240
    btnMoins.Caption = "-"
    btnMoins.Font.Size = 14
    btnMoins.Font.Bold = True
    btnMoins.Visible = True
    
    ' Zone quantité à affecter
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtQteAffecter" & index)
    ctrl.Left = positions(4) + 10
    ctrl.Top = topPos + 40
    ctrl.Width = largeurs(4) - 20
    ctrl.Height = 200
    ctrl.Text = "0"
    ctrl.Alignment = 2
    ctrl.Font.Bold = True
    ctrl.Font.Size = 10
    ctrl.Tag = quantite
    ctrl.Visible = True
    
    ' Configuration bouton +
    Dim btnPlus As CommandButton
    Set btnPlus = Me.Controls("cmdPlus" & index)
    btnPlus.Left = positions(5)
    btnPlus.Top = topPos + 20
    btnPlus.Width = largeurs(5)
    btnPlus.Height = 240
    btnPlus.Caption = "+"
    btnPlus.Font.Size = 14
    btnPlus.Font.Bold = True
    btnPlus.Visible = True
End Sub

Private Sub cmdActualiser_Click()
    MettreAJourSelection
End Sub

Private Sub MettreAJourSelection()
    Dim selection As String
    Dim nbPiecesTotal As Integer
    Dim nbTypePieces As Integer
    
    selection = "AFFECTATION POUR LE FRIGO: " & referenceFrigoReparable & " (Série: " & numeroSerieFrigo & ")" & vbCrLf
    selection = selection & "Frigoriste responsable: " & nomFrigoriste & vbCrLf
    selection = selection & "Date de sélection: " & Format(Now, "dd/mm/yyyy hh:nn:ss") & vbCrLf
    selection = selection & String(80, "=") & vbCrLf & vbCrLf
    
    For i = 0 To nombrePiecesAffichees - 1
        On Error Resume Next
        ' NOUVELLE LOGIQUE : Vérifier directement la quantité
        If Val(Me.Controls("txtQteAffecter" & i).Text) > 0 Then
            Dim code As String
            Dim piece As String
            Dim qteAffecter As Integer
            Dim qteMax As Integer
            
            code = Me.Controls("lblCode" & i).Caption
            piece = Me.Controls("lblPiece" & i).Caption
            qteAffecter = Val(Me.Controls("txtQteAffecter" & i).Text)
            qteMax = Val(Me.Controls("txtQteAffecter" & i).Tag)
            
            ' Validation de la quantité
            If qteAffecter > qteMax Then
                qteAffecter = qteMax
                Me.Controls("txtQteAffecter" & i).Text = qteMax
            End If
            
            selection = selection & nbTypePieces + 1 & ". " & code & " - " & piece & vbCrLf
            selection = selection & "   Quantité: " & qteAffecter & " sur " & qteMax & " disponible(s)" & vbCrLf & vbCrLf
            
            nbPiecesTotal = nbPiecesTotal + qteAffecter
            nbTypePieces = nbTypePieces + 1
        End If
    Next i
    
    If nbTypePieces = 0 Then
        selection = selection & "AUCUNE PIÈCE SÉLECTIONNÉE" & vbCrLf & vbCrLf
        selection = selection & "Utilisez les boutons + et - pour sélectionner les quantités de pièces à affecter."
        Me.Controls("lblResume").Caption = "PIÈCES SÉLECTIONNÉES: 0"
    Else
        selection = selection & String(80, "-") & vbCrLf
        selection = selection & "RÉSUMÉ DE L'AFFECTATION:" & vbCrLf
        selection = selection & "Types de pièces sélectionnées: " & nbTypePieces & vbCrLf
        selection = selection & "Nombre total de pièces: " & nbPiecesTotal & vbCrLf & vbCrLf
        selection = selection & "ATTENTION: Ces pièces seront automatiquement déduites du stock lors de la validation !"
        Me.Controls("lblResume").Caption = "PIÈCES SÉLECTIONNÉES: " & nbPiecesTotal & " (" & nbTypePieces & " types)"
    End If
    
    Me.Controls("txtPiecesSelectionnees").Caption = selection
End Sub

Private Sub cmdValiderAffectation_Click()
    ' Mettre à jour la sélection avant validation
    MettreAJourSelection
    
    ' MODIFICATION : Utilisation de Caption au lieu de Text
    If InStr(Me.Controls("txtPiecesSelectionnees").Caption, "AUCUNE PIÈCE") > 0 Then
        MsgBox "Aucune pièce sélectionnée !" & vbCrLf & vbCrLf & "Veuillez sélectionner au moins une pièce pour procéder à l'affectation.", vbExclamation, "Sélection requise"
        Exit Sub
    End If
    
    ' Validation des quantités
    If Not ValiderQuantites() Then Exit Sub
    
    ' Confirmation finale
    Dim message As String
    message = "CONFIRMER L'AFFECTATION DES PIÈCES ?" & vbCrLf & vbCrLf
    message = message & "Frigo cible: " & referenceFrigoReparable & " (" & numeroSerieFrigo & ")" & vbCrLf
    message = message & "Frigoriste: " & nomFrigoriste & vbCrLf & vbCrLf
    message = message & "Cette opération va:" & vbCrLf
    message = message & "• Déduire les pièces sélectionnées du stock" & vbCrLf
    message = message & "• Les affecter au frigo pour réparation" & vbCrLf
    message = message & "• Marquer le frigo comme 'EN_COURS_REPARATION'" & vbCrLf
    message = message & "• Créer un historique de l'affectation" & vbCrLf & vbCrLf
    message = message & "Cette action est définitive !"
    
    If MsgBox(message, vbYesNo + vbQuestion + vbDefaultButton2, "Confirmation d'affectation") = vbYes Then
        ValiderAffectationComplete
    End If
End Sub

Private Function ValiderQuantites() As Boolean
    Dim erreurs As String
    
    For i = 0 To nombrePiecesAffichees - 1
        On Error Resume Next
        Dim qteAffecter As Integer
        Dim qteMax As Integer
        Dim code As String
        
        qteAffecter = Val(Me.Controls("txtQteAffecter" & i).Text)
        qteMax = Val(Me.Controls("txtQteAffecter" & i).Tag)
        code = Me.Controls("lblCode" & i).Caption
        
        If qteAffecter > 0 Then  ' Seulement vérifier si quantité > 0
            If qteAffecter > qteMax Then
                erreurs = erreurs & "• " & code & ": Quantité demandée (" & qteAffecter & ") > disponible (" & qteMax & ")" & vbCrLf
            End If
        End If
    Next i
    
    If Len(erreurs) > 0 Then
        MsgBox "Erreurs de quantités détectées:" & vbCrLf & vbCrLf & erreurs & vbCrLf & "Veuillez corriger ces erreurs avant de valider.", vbExclamation, "Erreurs de validation"
        ValiderQuantites = False
    Else
        ValiderQuantites = True
    End If
End Function

Private Sub ValiderAffectationComplete()
    On Error GoTo GestionErreur
    
    ' 1. Sauvegarder l'affectation
    SauvegarderAffectation
    
    ' 2. Déduire les pièces du stock
    DeduireStockPieces
    
    ' 3. Mettre à jour le statut du frigo
    MettreAJourStatutFrigo "EN_COURS_REPARATION"
    
    ' 4. Créer l'historique de l'opération
    CreerHistoriqueAffectation
    
    MsgBox "Affectation réalisée avec succès !" & vbCrLf & vbCrLf & _
           "• Pièces déduites du stock" & vbCrLf & _
           "• Frigo " & referenceFrigoReparable & " en cours de réparation" & vbCrLf & _
           "• Historique mis à jour", vbInformation, "Succès"
    
    Me.Hide
    Exit Sub
    
GestionErreur:
    MsgBox "Erreur lors de l'affectation:" & vbCrLf & Err.description, vbCritical, "Erreur"
End Sub

Private Sub SauvegarderAffectation()
    Dim fichier As String
    Dim numeroFichier As Integer
    
    ' Créer le répertoire s'il n'existe pas
    If Dir(App.Path & "\Affectations", vbDirectory) = "" Then
        MkDir App.Path & "\Affectations"
    End If
    
    fichier = App.Path & "\Affectations\Affectation_" & referenceFrigoReparable & "_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    numeroFichier = FreeFile
    
    Open fichier For Output As #numeroFichier
    ' MODIFICATION : Utilisation de Caption au lieu de Text
    Print #numeroFichier, Me.Controls("txtPiecesSelectionnees").Caption
    Close #numeroFichier
End Sub

Private Sub DeduireStockPieces()
    On Error GoTo GestionErreurStock
    
    Dim fichier As String
    Dim fichierTemp As String
    Dim numeroFichier As Integer
    Dim numeroTemp As Integer
    Dim ligne As String
    Dim elements() As String
    
    fichier = App.Path & "\StockPieces.txt"
    fichierTemp = App.Path & "\StockPieces_temp.txt"
    
    If Dir(fichier) = "" Then
        MsgBox "Fichier stock introuvable pour mise à jour !", vbExclamation
        Exit Sub
    End If
    
    numeroFichier = FreeFile
    numeroTemp = FreeFile + 1
    
    Open fichier For Input As #numeroFichier
    Open fichierTemp For Output As #numeroTemp
    
    ' Copier l'en-tête
    If Not EOF(numeroFichier) Then
        Line Input #numeroFichier, ligne
        Print #numeroTemp, ligne
    End If
    
    ' Traiter chaque ligne du stock
    Do While Not EOF(numeroFichier)
        Line Input #numeroFichier, ligne
        If Len(Trim(ligne)) > 0 Then
            elements = Split(ligne, "|")
            If UBound(elements) >= 2 Then
                ' Vérifier si cette pièce est dans notre sélection
                For i = 0 To nombrePiecesAffichees - 1
                    On Error Resume Next
                    ' NOUVELLE LOGIQUE : Vérifier directement la quantité
                    If Val(Me.Controls("txtQteAffecter" & i).Text) > 0 And elements(0) = Me.Controls("lblCode" & i).Caption Then
                        ' Déduire la quantité affectée
                        Dim nouvelleQuantite As Integer
                        nouvelleQuantite = Val(elements(2)) - Val(Me.Controls("txtQteAffecter" & i).Text)
                        If nouvelleQuantite < 0 Then nouvelleQuantite = 0
                        elements(2) = CStr(nouvelleQuantite)
                        Exit For
                    End If
                Next i
                
                ' Reconstruire la ligne
                ligne = Join(elements, "|")
            End If
        End If
        Print #numeroTemp, ligne
    Loop
    
    Close #numeroFichier
    Close #numeroTemp
    
    ' Remplacer le fichier original
    Kill fichier
    Name fichierTemp As fichier
    
    Exit Sub
    
GestionErreurStock:
    MsgBox "Erreur lors de la mise à jour du stock: " & Err.description, vbCritical
    On Error Resume Next
    Close #numeroFichier
    Close #numeroTemp
End Sub

Private Sub MettreAJourStatutFrigo(nouveauStatut As String)
    On Error GoTo GestionErreurStatut
    
    Dim fichier As String
    Dim fichierTemp As String
    Dim numeroFichier As Integer
    Dim numeroTemp As Integer
    Dim ligne As String
    Dim elements() As String
    Dim trouve As Boolean
    
    fichier = App.Path & "\StockReparable.txt"
    fichierTemp = App.Path & "\StockReparable_temp.txt"
    trouve = False
    
    If Dir(fichier) = "" Then
        ' Créer le fichier s'il n'existe pas
        numeroFichier = FreeFile
        Open fichier For Output As #numeroFichier
        Print #numeroFichier, "REFERENCE|NUMERO_SERIE|DATE_ENTREE|STATUT|FRIGORISTE|COMMENTAIRE"
        Print #numeroFichier, referenceFrigoReparable & "|" & numeroSerieFrigo & "|" & Format(Now, "dd/mm/yyyy hh:nn:ss") & "|" & nouveauStatut & "|" & nomFrigoriste & "|Affectation_pieces"
        Close #numeroFichier
        Exit Sub
    End If
    
    numeroFichier = FreeFile
    numeroTemp = FreeFile + 1
    
    Open fichier For Input As #numeroFichier
    Open fichierTemp For Output As #numeroTemp
    
    ' Copier l'en-tête s'il existe
    If Not EOF(numeroFichier) Then
        Line Input #numeroFichier, ligne
        Print #numeroTemp, ligne
    End If
    
    ' Traiter chaque ligne
    Do While Not EOF(numeroFichier)
        Line Input #numeroFichier, ligne
        If Len(Trim(ligne)) > 0 Then
            elements = Split(ligne, "|")
            If UBound(elements) >= 3 And (elements(0) = referenceFrigoReparable Or elements(1) = numeroSerieFrigo) Then
                ' Mettre à jour le statut de ce frigo
                If UBound(elements) >= 3 Then elements(3) = nouveauStatut
                If UBound(elements) >= 4 Then elements(4) = nomFrigoriste
                If UBound(elements) >= 5 Then elements(5) = "Affectation_pieces_" & Format(Now, "dd/mm/yyyy")
                ligne = Join(elements, "|")
                trouve = True
            End If
        End If
        Print #numeroTemp, ligne
    Loop
    
    ' Si pas trouvé, ajouter une nouvelle ligne
    If Not trouve Then
        Print #numeroTemp, referenceFrigoReparable & "|" & numeroSerieFrigo & "|" & Format(Now, "dd/mm/yyyy hh:nn:ss") & "|" & nouveauStatut & "|" & nomFrigoriste & "|Affectation_pieces"
    End If
    
    Close #numeroFichier
    Close #numeroTemp
    
    ' Remplacer le fichier original
    Kill fichier
    Name fichierTemp As fichier
    
    Exit Sub
    
GestionErreurStatut:
    MsgBox "Erreur lors de la mise à jour du statut frigo: " & Err.description, vbCritical
    On Error Resume Next
    Close #numeroFichier
    Close #numeroTemp
End Sub

Private Sub CreerHistoriqueAffectation()
    On Error GoTo GestionErreurHistorique
    
    Dim fichier As String
    Dim numeroFichier As Integer
    
    fichier = App.Path & "\HistoriqueAffectations.txt"
    numeroFichier = FreeFile
    
    ' Créer l'en-tête si nouveau fichier
    If Dir(fichier) = "" Then
        Open fichier For Output As #numeroFichier
        Print #numeroFichier, "DATE|FRIGO_CIBLE|NUMERO_SERIE|FRIGORISTE|NB_PIECES|NB_TYPES|COUT_TOTAL|STATUT|DETAILS"
        Close #numeroFichier
    End If
    
    ' Calculer les totaux
    Dim nbPiecesTotal As Integer
    Dim nbTypes As Integer
    Dim coutTotal As Double
    Dim details As String
    
    For i = 0 To nombrePiecesAffichees - 1
        On Error Resume Next
        If Me.Controls("chkSelectionne" & i).Value = 1 Then
            Dim code As String
            Dim qte As Integer
            Dim prix As Double
            
            code = Me.Controls("lblCode" & i).Caption
            qte = Val(Me.Controls("txtQteAffecter" & i).Text)
            prix = Val(Replace(Me.Controls("lblPrixPiece" & i).Caption, "€", ""))
            
            nbPiecesTotal = nbPiecesTotal + qte
            nbTypes = nbTypes + 1
            coutTotal = coutTotal + (prix * qte)
            
            If Len(details) > 0 Then details = details & ";"
            details = details & code & ":" & qte
        End If
    Next i
    
    ' Ajouter l'entrée d'historique
    Open fichier For Append As #numeroFichier
    Print #numeroFichier, Format(Now, "dd/mm/yyyy hh:nn:ss") & "|" & referenceFrigoReparable & "|" & numeroSerieFrigo & "|" & nomFrigoriste & "|" & nbPiecesTotal & "|" & nbTypes & "|" & Format(coutTotal, "0.00") & "|EN_COURS_REPARATION|" & details
    Close #numeroFichier
    
    Exit Sub
    
GestionErreurHistorique:
    MsgBox "Erreur lors de la création de l'historique: " & Err.description, vbCritical
    On Error Resume Next
    Close #numeroFichier
End Sub

Private Sub cmdAnnuler_Click()
    If MsgBox("Annuler l'affectation des pièces ?" & vbCrLf & vbCrLf & "Toutes les sélections seront perdues.", vbYesNo + vbQuestion + vbDefaultButton2, "Confirmation d'annulation") = vbYes Then
        Me.Hide
    End If
End Sub

' === ÉVÉNEMENTS DES BOUTONS +/- ===
Private Sub cmdPlus0_Click(): AjusterQuantiteAffectation 0, 1: End Sub
Private Sub cmdMoins0_Click(): AjusterQuantiteAffectation 0, -1: End Sub
Private Sub cmdPlus1_Click(): AjusterQuantiteAffectation 1, 1: End Sub
Private Sub cmdMoins1_Click(): AjusterQuantiteAffectation 1, -1: End Sub
Private Sub cmdPlus2_Click(): AjusterQuantiteAffectation 2, 1: End Sub
Private Sub cmdMoins2_Click(): AjusterQuantiteAffectation 2, -1: End Sub
Private Sub cmdPlus3_Click(): AjusterQuantiteAffectation 3, 1: End Sub
Private Sub cmdMoins3_Click(): AjusterQuantiteAffectation 3, -1: End Sub
Private Sub cmdPlus4_Click(): AjusterQuantiteAffectation 4, 1: End Sub
Private Sub cmdMoins4_Click(): AjusterQuantiteAffectation 4, -1: End Sub
Private Sub cmdPlus5_Click(): AjusterQuantiteAffectation 5, 1: End Sub
Private Sub cmdMoins5_Click(): AjusterQuantiteAffectation 5, -1: End Sub
Private Sub cmdPlus6_Click(): AjusterQuantiteAffectation 6, 1: End Sub
Private Sub cmdMoins6_Click(): AjusterQuantiteAffectation 6, -1: End Sub
Private Sub cmdPlus7_Click(): AjusterQuantiteAffectation 7, 1: End Sub
Private Sub cmdMoins7_Click(): AjusterQuantiteAffectation 7, -1: End Sub
Private Sub cmdPlus8_Click(): AjusterQuantiteAffectation 8, 1: End Sub
Private Sub cmdMoins8_Click(): AjusterQuantiteAffectation 8, -1: End Sub
Private Sub cmdPlus9_Click(): AjusterQuantiteAffectation 9, 1: End Sub
Private Sub cmdMoins9_Click(): AjusterQuantiteAffectation 9, -1: End Sub
Private Sub cmdPlus10_Click(): AjusterQuantiteAffectation 10, 1: End Sub
Private Sub cmdMoins10_Click(): AjusterQuantiteAffectation 10, -1: End Sub
Private Sub cmdPlus11_Click(): AjusterQuantiteAffectation 11, 1: End Sub
Private Sub cmdMoins11_Click(): AjusterQuantiteAffectation 11, -1: End Sub
Private Sub cmdPlus12_Click(): AjusterQuantiteAffectation 12, 1: End Sub
Private Sub cmdMoins12_Click(): AjusterQuantiteAffectation 12, -1: End Sub
Private Sub cmdPlus13_Click(): AjusterQuantiteAffectation 13, 1: End Sub
Private Sub cmdMoins13_Click(): AjusterQuantiteAffectation 13, -1: End Sub
Private Sub cmdPlus14_Click(): AjusterQuantiteAffectation 14, 1: End Sub
Private Sub cmdMoins14_Click(): AjusterQuantiteAffectation 14, -1: End Sub
Private Sub cmdPlus15_Click(): AjusterQuantiteAffectation 15, 1: End Sub
Private Sub cmdMoins15_Click(): AjusterQuantiteAffectation 15, -1: End Sub
Private Sub cmdPlus16_Click(): AjusterQuantiteAffectation 16, 1: End Sub
Private Sub cmdMoins16_Click(): AjusterQuantiteAffectation 16, -1: End Sub
Private Sub cmdPlus17_Click(): AjusterQuantiteAffectation 17, 1: End Sub
Private Sub cmdMoins17_Click(): AjusterQuantiteAffectation 17, -1: End Sub
Private Sub cmdPlus18_Click(): AjusterQuantiteAffectation 18, 1: End Sub
Private Sub cmdMoins18_Click(): AjusterQuantiteAffectation 18, -1: End Sub
Private Sub cmdPlus19_Click(): AjusterQuantiteAffectation 19, 1: End Sub
Private Sub cmdMoins19_Click(): AjusterQuantiteAffectation 19, -1: End Sub

Private Sub AjusterQuantiteAffectation(index As Integer, ajustement As Integer)
    On Error GoTo ErrorHandler
    
    Dim quantiteActuelle As Integer
    Dim nouvelleQuantite As Integer
    Dim quantiteMax As Integer
    
    quantiteActuelle = Val(Me.Controls("txtQteAffecter" & index).Text)
    nouvelleQuantite = quantiteActuelle + ajustement
    quantiteMax = Val(Me.Controls("txtQteAffecter" & index).Tag)
    
    If nouvelleQuantite < 0 Then nouvelleQuantite = 0
    If nouvelleQuantite > quantiteMax Then nouvelleQuantite = quantiteMax
    
    Me.Controls("txtQteAffecter" & index).Text = CStr(nouvelleQuantite)
    
    ' Mise à jour automatique du résumé
    MettreAJourSelection
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur lors de l'ajustement de la quantité: " & Err.description, vbCritical
End Sub



