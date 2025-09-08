VERSION 5.00
Begin VB.Form frmAffectatonPieces 
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
Attribute VB_Name = "frmAffectatonPieces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' === FRMAFFECTATIONPIECES.FRM - AFFECTATION DES PIÈCES ===

Private referenceFrigoReparable As String
Private nomFrigoriste As String

Private Sub Form_Load()
    Me.BackColor = RGB(240, 240, 240)
    Me.Caption = "Affectation Pièces - " & referenceFrigoReparable
    Me.Width = 15000
    Me.Height = 11000
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    CreerInterfaceAffectation
    ChargerStockPieces
End Sub

Public Sub InitialiserAvecFrigo(reference As String, frigoriste As String)
    referenceFrigoReparable = reference
    nomFrigoriste = frigoriste
    Me.Caption = "Affectation Pièces - " & referenceFrigoReparable
End Sub

Private Sub CreerInterfaceAffectation()
    Dim ctrl As Object
    
    ' Titre
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitre")
    ctrl.Left = 240
    ctrl.Top = 120
    ctrl.Width = 11000
    ctrl.Height = 400
    ctrl.Caption = "?? AFFECTATION DES PIÈCES AU FRIGO RÉPARABLE ??"
    ctrl.BackColor = RGB(100, 150, 255)
    ctrl.ForeColor = RGB(255, 255, 255)
    ctrl.Font.Size = 16
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Info frigo cible
    Set ctrl = Me.Controls.Add("VB.Label", "lblInfoCible")
    ctrl.Left = 240
    ctrl.Top = 600
    ctrl.Width = 11000
    ctrl.Height = 300
    ctrl.Caption = "FRIGO CIBLE: " & referenceFrigoReparable & " | FRIGORISTE: " & nomFrigoriste & " | " & Format(Now, "dd/mm/yyyy hh:nn:ss")
    ctrl.BackColor = RGB(200, 255, 200)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Instructions
    Set ctrl = Me.Controls.Add("VB.Label", "lblInstructions")
    ctrl.Left = 240
    ctrl.Top = 960
    ctrl.Width = 11000
    ctrl.Height = 300
    ctrl.Caption = "Sélectionnez les pièces du stock à affecter au frigo réparable. Les quantités seront déduites du stock."
    ctrl.Font.Size = 10
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' === STOCK DISPONIBLE ===
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitreStock")
    ctrl.Left = 240
    ctrl.Top = 1320
    ctrl.Width = 11000
    ctrl.Height = 300
    ctrl.Caption = "=== STOCK DE PIÈCES DISPONIBLES ==="
    ctrl.BackColor = RGB(255, 200, 100)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' En-têtes colonnes stock
    CreerEnTetesStock
    
    ' === PIÈCES SÉLECTIONNÉES ===
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitreSelection")
    ctrl.Left = 240
    ctrl.Top = 6000
    ctrl.Width = 11000
    ctrl.Height = 300
    ctrl.Caption = "=== PIÈCES SÉLECTIONNÉES POUR AFFECTATION ==="
    ctrl.BackColor = RGB(100, 255, 100)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtPiecesSelectionnees")
    ctrl.Left = 240
    ctrl.Top = 6360
    ctrl.Width = 11000
    ctrl.Height = 1200
    ctrl.MultiLine = True
    ctrl.ScrollBars = 2
    ctrl.BackColor = RGB(255, 255, 240)
    ctrl.Text = "Aucune pièce sélectionnée"
    ctrl.Visible = True
    
    ' Boutons d'action
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdMettreAJourSelection")
    ctrl.Left = 240
    ctrl.Top = 7640
    ctrl.Width = 2000
    ctrl.Height = 400
    ctrl.Caption = "?? Actualiser Sélection"
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdValiderAffectation")
    ctrl.Left = 2400
    ctrl.Top = 7640
    ctrl.Width = 2500
    ctrl.Height = 400
    ctrl.Caption = "? VALIDER AFFECTATION"
    ctrl.Font.Bold = True
    ctrl.Font.Size = 11
    ctrl.BackColor = RGB(128, 255, 128)
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdSimulerReparation")
    ctrl.Left = 5000
    ctrl.Top = 7640
    ctrl.Width = 2500
    ctrl.Height = 400
    ctrl.Caption = "?? SIMULER RÉPARATION"
    ctrl.Font.Bold = True
    ctrl.BackColor = RGB(255, 200, 100)
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdAnnuler")
    ctrl.Left = 7600
    ctrl.Top = 7640
    ctrl.Width = 2000
    ctrl.Height = 400
    ctrl.Caption = "? ANNULER"
    ctrl.Font.Bold = True
    ctrl.BackColor = RGB(255, 128, 128)
    ctrl.Visible = True
    
    ' Coût total
    Set ctrl = Me.Controls.Add("VB.Label", "lblCoutTotal")
    ctrl.Left = 10000
    ctrl.Top = 7640
    ctrl.Width = 1500
    ctrl.Height = 400
    ctrl.Caption = "Coût: 0.00€"
    ctrl.Font.Bold = True
    ctrl.Font.Size = 12
    ctrl.BackColor = RGB(255, 255, 200)
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
End Sub

Private Sub CreerEnTetesStock()
    Dim ctrl As Object
    Dim topPos As Long
    topPos = 1680
    
    ' Code
    Set ctrl = Me.Controls.Add("VB.Label", "lblHCode")
    ctrl.Left = 240
    ctrl.Top = topPos
    ctrl.Width = 800
    ctrl.Height = 250
    ctrl.Caption = "CODE"
    ctrl.BackColor = RGB(180, 180, 180)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' Pièce
    Set ctrl = Me.Controls.Add("VB.Label", "lblHPiece")
    ctrl.Left = 1040
    ctrl.Top = topPos
    ctrl.Width = 2500
    ctrl.Height = 250
    ctrl.Caption = "PIÈCE"
    ctrl.BackColor = RGB(180, 180, 180)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' Quantité
    Set ctrl = Me.Controls.Add("VB.Label", "lblHQte")
    ctrl.Left = 3540
    ctrl.Top = topPos
    ctrl.Width = 800
    ctrl.Height = 250
    ctrl.Caption = "QTÉ"
    ctrl.BackColor = RGB(180, 180, 180)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' État
    Set ctrl = Me.Controls.Add("VB.Label", "lblHEtat")
    ctrl.Left = 4340
    ctrl.Top = topPos
    ctrl.Width = 1200
    ctrl.Height = 250
    ctrl.Caption = "ÉTAT"
    ctrl.BackColor = RGB(180, 180, 180)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' Origine
    Set ctrl = Me.Controls.Add("VB.Label", "lblHOrigine")
    ctrl.Left = 5540
    ctrl.Top = topPos
    ctrl.Width = 1500
    ctrl.Height = 250
    ctrl.Caption = "ORIGINE"
    ctrl.BackColor = RGB(180, 180, 180)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' Date
    Set ctrl = Me.Controls.Add("VB.Label", "lblHDate")
    ctrl.Left = 7040
    ctrl.Top = topPos
    ctrl.Width = 1200
    ctrl.Height = 250
    ctrl.Caption = "DATE"
    ctrl.BackColor = RGB(180, 180, 180)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' Prix
    Set ctrl = Me.Controls.Add("VB.Label", "lblHPrix")
    ctrl.Left = 8240
    ctrl.Top = topPos
    ctrl.Width = 800
    ctrl.Height = 250
    ctrl.Caption = "PRIX"
    ctrl.BackColor = RGB(180, 180, 180)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' Action
    Set ctrl = Me.Controls.Add("VB.Label", "lblHAction")
    ctrl.Left = 9040
    ctrl.Top = topPos
    ctrl.Width = 1000
    ctrl.Height = 250
    ctrl.Caption = "ACTION"
    ctrl.BackColor = RGB(180, 180, 180)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' Sélection
    Set ctrl = Me.Controls.Add("VB.Label", "lblHSelect")
    ctrl.Left = 10040
    ctrl.Top = topPos
    ctrl.Width = 1200
    ctrl.Height = 250
    ctrl.Caption = "SÉLECTION"
    ctrl.BackColor = RGB(180, 180, 180)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
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
    
    If Dir(fichier) = "" Then
        MsgBox "Aucun stock de pièces disponible", vbInformation
        Exit Sub
    End If
    
    numeroFichier = FreeFile
    Open fichier For Input As #numeroFichier
    
    ' Ignorer l'en-tête
    If Not EOF(numeroFichier) Then Line Input #numeroFichier, ligne
    
    Do While Not EOF(numeroFichier) And compteur < 15 ' Limiter à 15 pièces pour l'affichage
        Line Input #numeroFichier, ligne
        If Len(Trim(ligne)) > 0 Then
            elements = Split(ligne, "|")
            If UBound(elements) >= 6 And Val(elements(2)) > 0 Then ' Seulement si quantité > 0
                CreerLignePieceStock topPosition, elements(0), elements(1), elements(2), elements(3), elements(4), elements(5), elements(6), compteur
                topPosition = topPosition + 280
                compteur = compteur + 1
            End If
        End If
    Loop
    
    Close #numeroFichier
    
    If compteur = 0 Then
        MsgBox "Aucune pièce disponible en stock", vbInformation
    End If
End Sub

Private Sub CreerLignePieceStock(topPos As Long, code As String, piece As String, quantite As String, etat As String, origine As String, dateAjout As String, prix As String, index As Integer)
    Dim ctrl As Object
    
    ' Code
    Set ctrl = Me.Controls.Add("VB.Label", "lblCode" & index)
    ctrl.Left = 240
    ctrl.Top = topPos
    ctrl.Width = 800
    ctrl.Height = 250
    ctrl.Caption = code
    ctrl.BackColor = RGB(255, 255, 255)
    ctrl.BorderStyle = 1
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Pièce
    Set ctrl = Me.Controls.Add("VB.Label", "lblPiece" & index)
    ctrl.Left = 1040
    ctrl.Top = topPos
    ctrl.Width = 2500
    ctrl.Height = 250
    ctrl.Caption = piece
    ctrl.BackColor = RGB(255, 255, 255)
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' Quantité disponible
    Set ctrl = Me.Controls.Add("VB.Label", "lblQteDispo" & index)
    ctrl.Left = 3540
    ctrl.Top = topPos
    ctrl.Width = 800
    ctrl.Height = 250
    ctrl.Caption = quantite
    ctrl.BackColor = RGB(255, 255, 255)
    ctrl.BorderStyle = 1
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' État avec couleur
    Set ctrl = Me.Controls.Add("VB.Label", "lblEtat" & index)
    ctrl.Left = 4340
    ctrl.Top = topPos
    ctrl.Width = 1200
    ctrl.Height = 250
    ctrl.Caption = etat
    ctrl.BorderStyle = 1
    ctrl.Alignment = 2
    Select Case UCase(etat)
        Case "EXCELLENT"
            ctrl.BackColor = RGB(100, 255, 100)
        Case "BON"
            ctrl.BackColor = RGB(150, 255, 150)
        Case "MOYEN"
            ctrl.BackColor = RGB(255, 255, 150)
        Case "DEFECTUEUX"
            ctrl.BackColor = RGB(255, 150, 150)
        Case Else
            ctrl.BackColor = RGB(255, 255, 255)
    End Select
    ctrl.Visible = True
    
    ' Origine
    Set ctrl = Me.Controls.Add("VB.Label", "lblOrigine" & index)
    ctrl.Left = 5540
    ctrl.Top = topPos
    ctrl.Width = 1500
    ctrl.Height = 250
    ctrl.Caption = origine
    ctrl.BackColor = RGB(255, 255, 255)
    ctrl.BorderStyle = 1
    ctrl.Font.Size = 8
    ctrl.Visible = True
    
    ' Date
    Set ctrl = Me.Controls.Add("VB.Label", "lblDatePiece" & index)
    ctrl.Left = 7040
    ctrl.Top = topPos
    ctrl.Width = 1200
    ctrl.Height = 250
    ctrl.Caption = Left(dateAjout, 10)
    ctrl.BackColor = RGB(255, 255, 255)
    ctrl.BorderStyle = 1
    ctrl.Font.Size = 8
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Prix
    Set ctrl = Me.Controls.Add("VB.Label", "lblPrixPiece" & index)
    ctrl.Left = 8240
    ctrl.Top = topPos
    ctrl.Width = 800
    ctrl.Height = 250
    ctrl.Caption = prix & "€"
    ctrl.BackColor = RGB(255, 255, 255)
    ctrl.BorderStyle = 1
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Bouton Sélectionner
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdSelectionner" & index)
    ctrl.Left = 9040
    ctrl.Top = topPos
    ctrl.Width = 1000
    ctrl.Height = 250
    ctrl.Caption = "? Ajouter"
    ctrl.Font.Size = 8
    ctrl.Tag = index
    ctrl.Visible = True
    
    ' CheckBox sélection
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkSelectionne" & index)
    ctrl.Left = 10140
    ctrl.Top = topPos + 50
    ctrl.Width = 300
    ctrl.Height = 150
    ctrl.Caption = ""
    ctrl.Visible = True
    
    ' Quantité à affecter (initialement cachée)
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtQteAffecter" & index)
    ctrl.Left = 10440
    ctrl.Top = topPos + 25
    ctrl.Width = 600
    ctrl.Height = 200
    ctrl.Text = "1"
    ctrl.Alignment = 2
    ctrl.Font.Bold = True
    ctrl.Visible = False
    ctrl.Tag = quantite ' Stocker la quantité max disponible
End Sub

' Gestion des boutons de sélection (génération dynamique)
Private Sub cmdSelectionner0_Click()
    GererSelectionPiece 0
End Sub
Private Sub cmdSelectionner1_Click()
    GererSelectionPiece 1
End Sub
Private Sub cmdSelectionner2_Click()
    GererSelectionPiece 2
End Sub
Private Sub cmdSelectionner3_Click()
    GererSelectionPiece 3
End Sub
Private Sub cmdSelectionner4_Click()
    GererSelectionPiece 4
End Sub
Private Sub cmdSelectionner5_Click()
    GererSelectionPiece 5
End Sub
Private Sub cmdSelectionner6_Click()
    GererSelectionPiece 6
End Sub
Private Sub cmdSelectionner7_Click()
    GererSelectionPiece 7
End Sub
Private Sub cmdSelectionner8_Click()
    GererSelectionPiece 8
End Sub
Private Sub cmdSelectionner9_Click()
    GererSelectionPiece 9
End Sub
Private Sub cmdSelectionner10_Click()
    GererSelectionPiece 10
End Sub
Private Sub cmdSelectionner11_Click()
    GererSelectionPiece 11
End Sub
Private Sub cmdSelectionner12_Click()
    GererSelectionPiece 12
End Sub
Private Sub cmdSelectionner13_Click()
    GererSelectionPiece 13
End Sub
Private Sub cmdSelectionner14_Click()
    GererSelectionPiece 14
End Sub

Private Sub GererSelectionPiece(index As Integer)
    On Error Resume Next
    
    If Me.Controls("chkSelectionne" & index).Value = 0 Then
        ' Sélectionner la pièce
        Me.Controls("chkSelectionne" & index).Value = 1
        Me.Controls("txtQteAffecter" & index).Visible = True
        Me.Controls("cmdSelectionner" & index).Caption = "? Retirer"
        Me.Controls("cmdSelectionner" & index).BackColor = RGB(255, 200, 200)
        
        ' Vérifier la quantité max
        Dim qteMax As Integer
        qteMax = Val(Me.Controls("txtQteAffecter" & index).Tag)
        If Val(Me.Controls("txtQteAffecter" & index).Text) > qteMax Then
            Me.Controls("txtQteAffecter" & index).Text = qteMax
        End If
    Else
        ' Désélectionner la pièce
        Me.Controls("chkSelectionne" & index).Value = 0
        Me.Controls("txtQteAffecter" & index).Visible = False
        Me.Controls("cmdSelectionner" & index).Caption = "? Ajouter"
        Me.Controls("cmdSelectionner" & index).BackColor = vbButtonFace
    End If
    
    ' Mettre à jour automatiquement la sélection
    MettreAJourSelection
End Sub

Private Sub cmdMettreAJourSelection_Click()
    MettreAJourSelection
End Sub

Private Sub MettreAJourSelection()
    Dim selection As String
    Dim coutTotal As Double
    Dim nbPiecesSelectionnees As Integer
    
    selection = "PIÈCES SÉLECTIONNÉES POUR LE FRIGO: " & referenceFrigoReparable & vbCrLf
    selection = selection & "Date de sélection: " & Format(Now, "dd/mm/yyyy hh:nn:ss") & vbCrLf
    selection = selection & String(70, "=") & vbCrLf & vbCrLf
    
    For i = 0 To 14
        On Error Resume Next
        If Me.Controls("chkSelectionne" & i).Value = 1 Then
            Dim code As String
            Dim piece As String
            Dim qteAffecter As String
            Dim etat As String
            Dim prix As Double
            
            code = Me.Controls("lblCode" & i).Caption
            piece = Me.Controls("lblPiece" & i).Caption
            qteAffecter = Me.Controls("txtQteAffecter" & i).Text
            etat = Me.Controls("lblEtat" & i).Caption
            prix = Val(Replace(Me.Controls("lblPrixPiece" & i).Caption, "€", ""))
            
            selection = selection & "• " & code & " - " & piece & vbCrLf
            selection = selection & "  Quantité affectée: " & qteAffecter & " | État: " & etat & " | Prix unitaire: " & Format(prix, "0.00") & "€" & vbCrLf
            selection = selection & "  Coût total pièce: " & Format(prix * Val(qteAffecter), "0.00") & "€" & vbCrLf & vbCrLf
            
            coutTotal = coutTotal + (prix * Val(qteAffecter))
            nbPiecesSelectionnees = nbPiecesSelectionnees + Val(qteAffecter)
        End If
    Next i
    
    If nbPiecesSelectionnees = 0 Then
        selection = selection & "? AUCUNE PIÈCE SÉLECTIONNÉE" & vbCrLf
        selection = selection & "Veuillez sélectionner au moins une pièce pour l'affectation."
    Else
        selection = selection & String(70, "-") & vbCrLf
        selection = selection & "RÉSUMÉ FINAL:" & vbCrLf
        selection = selection & "Nombre de pièces: " & nbPiecesSelectionnees & vbCrLf
        selection = selection & "Coût total de la réparation: " & Format(coutTotal, "0.00") & "€" & vbCrLf
        selection = selection & vbCrLf & "?? Ces pièces seront déduites du stock lors de la validation."
    End If
    
    Me.Controls("txtPiecesSelectionnees").Text = selection
    Me.Controls("lblCoutTotal").Caption = "Coût: " & Format(coutTotal, "0.00") & "€"
End Sub

Private Sub cmdValiderAffectation_Click()
    MettreAJourSelection
    
    If InStr(Me.Controls("txtPiecesSelectionnees").Text, "AUCUNE PIÈCE") > 0 Then
        MsgBox "Veuillez sélectionner au moins une pièce à affecter !", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("CONFIRMER L'AFFECTATION DES PIÈCES ?" & vbCrLf & vbCrLf & "?? Cette action va:" & vbCrLf & "- Déduire les pièces du stock" & vbCrLf & "- Les affecter au frigo " & referenceFrigoReparable & vbCrLf & "- Marquer le frigo comme EN_COURS", vbYesNo + vbQuestion) = vbYes Then
        
        ValiderAffectation
        MsgBox "Affectation réussie !" & vbCrLf & "Le frigo est maintenant en cours de réparation.", vbInformation
        Me.Hide
    End If
End Sub

Private Sub ValiderAffectation()
    ' 1. Sauvegarder l'affectation
    SauvegarderAffectation
    
    ' 2. Déduire les pièces du stock
    DeduireStockPieces
    
    ' 3. Mettre à jour le statut du frigo
    MettreAJourStatutFrigo "EN_COURS"
    
    ' 4. Créer l'historique de l'opération
    CreerHistoriqueAffectation
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
    Print #numeroFichier, Me.Controls("txtPiecesSelectionnees").Text
    Close #numeroFichier
End Sub

Private Sub DeduireStockPieces()
    ' Lire le stock actuel et mettre à jour
    Dim fichier As String
    Dim fichierTemp As String
    Dim numeroFichier As Integer
    Dim numeroTemp As Integer
    Dim ligne As String
    Dim elements() As String
    
    fichier = App.Path & "\StockPieces.txt"
    fichierTemp = App.Path & "\StockPieces_temp.txt"
    
    numeroFichier = FreeFile
    numeroTemp = FreeFile + 1
    
    Open fichier For Input As #numeroFichier
    Open fichierTemp For Output As #numeroTemp
    
    ' Copier l'en-tête
    If Not EOF(numeroFichier) Then
        Line Input #numeroFichier, ligne
        Print #numeroTemp, ligne
    End If
    
    Do While Not EOF(numeroFichier)
        Line Input #numeroFichier, ligne
        If Len(Trim(ligne)) > 0 Then
            elements = Split(ligne, "|")
            If UBound(elements) >= 2 Then
                ' Vérifier si cette pièce est dans notre sélection
                For i = 0 To 14
                    On Error Resume Next
                    If Me.Controls("chkSelectionne" & i).Value = 1 And elements(0) = Me.Controls("lblCode" & i).Caption Then
                        ' Déduire la quantité
                        Dim nouvelleQuantite As Integer
                        nouvelleQuantite = Val(elements(2)) - Val(Me.Controls("txtQteAffecter" & i).Text)
                        If nouvelleQuantite < 0 Then nouvelleQuantite = 0
                        elements(2) = nouvelleQuantite
                        Exit For
                    End If
                Next i
                
                ' Réécrire la ligne (même si pas modifiée)
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
End Sub

Private Sub MettreAJourStatutFrigo(nouveauStatut As String)
    Dim fichier As String
    Dim fichierTemp As String
    Dim numeroFichier As Integer
    Dim numeroTemp As Integer
    Dim ligne As String
    Dim elements() As String
    
    fichier = App.Path & "\StockReparable.txt"
    fichierTemp = App.Path & "\StockReparable_temp.txt"
    
    numeroFichier = FreeFile
    numeroTemp = FreeFile + 1
    
    Open fichier For Input As #numeroFichier
    Open fichierTemp For Output As #numeroTemp
    
    Do While Not EOF(numeroFichier)
        Line Input #numeroFichier, ligne
        If Len(Trim(ligne)) > 0 Then
            elements = Split(ligne, "|")
            If UBound(elements) >= 3 And elements(0) = referenceFrigoReparable Then
                elements(3) = nouveauStatut
                ligne = Join(elements, "|")
            End If
        End If
        Print #numeroTemp, ligne
    Loop
    
    Close #numeroFichier
    Close #numeroTemp
    
    Kill fichier
    Name fichierTemp As fichier
End Sub

Private Sub CreerHistoriqueAffectation()
    Dim fichier As String
    Dim numeroFichier As Integer
    
    fichier = App.Path & "\HistoriqueAffectations.txt"
    numeroFichier = FreeFile
    
    ' En-tête si nouveau fichier
    If Dir(fichier) = "" Then
        Open fichier For Output As #numeroFichier
        Print #numeroFichier, "DATE|FRIGO_CIBLE|FRIGORISTE|PIECES_AFFECTEES|COUT_TOTAL|STATUT"
        Close #numeroFichier
    End If
    
    Open fichier For Append As #numeroFichier
    
    ' Compter les pièces et calculer le coût
    Dim nbPieces As Integer
    Dim cout As Double
    For i = 0 To 14
        On Error Resume Next
        If Me.Controls("chkSelectionne" & i).Value = 1 Then
            nbPieces = nbPieces + Val(Me.Controls("txtQteAffecter" & i).Text)
            cout = cout + (Val(Replace(Me.Controls("lblPrixPiece" & i).Caption, "€", "")) * Val(Me.Controls("txtQteAffecter" & i).Text))
        End If
    Next i
    
    Print #numeroFichier, Format(Now, "dd/mm/yyyy hh:nn:ss") & "|" & referenceFrigoReparable & "|" & nomFrigoriste & "|" & nbPieces & "|" & Format(cout, "0.00") & "|EN_COURS"
    Close #numeroFichier
End Sub

Private Sub cmdSimulerReparation_Click()
    MsgBox "?? SIMULATION DE RÉPARATION ??" & vbCrLf & vbCrLf & _
           "1. Pièces sélectionnées ? Installées sur le frigo" & vbCrLf & _
           "2. Tests de fonctionnement ? OK" & vbCrLf & _
           "3. Contrôle qualité ? Validé" & vbCrLf & _
           "4. Frigo prêt pour remise en service" & vbCrLf & vbCrLf & _
           "Le frigo " & referenceFrigoReparable & " sera comme neuf !", vbInformation, "Simulation"
End Sub

Private Sub cmdAnnuler_Click()
    If MsgBox("Annuler l'affectation des pièces ?", vbYesNo + vbQuestion) = vbYes Then
        Me.Hide
    End If
End Sub

