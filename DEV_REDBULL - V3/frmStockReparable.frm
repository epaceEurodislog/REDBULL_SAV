VERSION 5.00
Begin VB.Form frmStockReparable 
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
Attribute VB_Name = "frmStockReparable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' === FRMSTOCKREPARABLE.FRM - STOCK FRIGOS RÉPARABLES ===

Private Sub Form_Load()
    Me.BackColor = RGB(240, 240, 240)
    Me.Caption = "Stock des Frigos Réparables"
    Me.Width = 14000
    Me.Height = 10000
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    CreerInterfaceStock
    ChargerStockReparable
End Sub

Private Sub CreerInterfaceStock()
    Dim ctrl As Object
    
    ' Titre
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitre")
    ctrl.Left = 240
    ctrl.Top = 120
    ctrl.Width = 10000
    ctrl.Height = 400
    ctrl.Caption = "?? STOCK DES FRIGOS RÉPARABLES ??"
    ctrl.BackColor = RGB(100, 200, 100)
    ctrl.ForeColor = RGB(255, 255, 255)
    ctrl.Font.Size = 16
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Statistiques
    Set ctrl = Me.Controls.Add("VB.Label", "lblStats")
    ctrl.Left = 240
    ctrl.Top = 600
    ctrl.Width = 10000
    ctrl.Height = 300
    ctrl.Caption = "Chargement des statistiques..."
    ctrl.BackColor = RGB(200, 255, 200)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Boutons d'action
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdActualiser")
    ctrl.Left = 240
    ctrl.Top = 960
    ctrl.Width = 1500
    ctrl.Height = 400
    ctrl.Caption = "?? Actualiser"
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdVoirStock")
    ctrl.Left = 1800
    ctrl.Top = 960
    ctrl.Width = 1800
    ctrl.Height = 400
    ctrl.Caption = "?? Voir Stock Pièces"
    ctrl.Font.Bold = True
    ctrl.BackColor = RGB(200, 200, 255)
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdExporter")
    ctrl.Left = 3660
    ctrl.Top = 960
    ctrl.Width = 1500
    ctrl.Height = 400
    ctrl.Caption = "?? Exporter"
    ctrl.Font.Bold = True
    ctrl.BackColor = RGB(255, 200, 100)
    ctrl.Visible = True
    
    ' Liste des frigos réparables
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitreListe")
    ctrl.Left = 240
    ctrl.Top = 1440
    ctrl.Width = 10000
    ctrl.Height = 300
    ctrl.Caption = "=== FRIGOS EN ATTENTE DE RÉPARATION ==="
    ctrl.BackColor = RGB(200, 200, 200)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' En-têtes de colonnes
    Set ctrl = Me.Controls.Add("VB.Label", "lblColRef")
    ctrl.Left = 240
    ctrl.Top = 1800
    ctrl.Width = 2000
    ctrl.Height = 250
    ctrl.Caption = "RÉFÉRENCE"
    ctrl.BackColor = RGB(180, 180, 180)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.Label", "lblColFrigo")
    ctrl.Left = 2240
    ctrl.Top = 1800
    ctrl.Width = 1800
    ctrl.Height = 250
    ctrl.Caption = "FRIGORISTE"
    ctrl.BackColor = RGB(180, 180, 180)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.Label", "lblColDate")
    ctrl.Left = 4040
    ctrl.Top = 1800
    ctrl.Width = 1500
    ctrl.Height = 250
    ctrl.Caption = "DATE"
    ctrl.BackColor = RGB(180, 180, 180)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.Label", "lblColStatut")
    ctrl.Left = 5540
    ctrl.Top = 1800
    ctrl.Width = 1200
    ctrl.Height = 250
    ctrl.Caption = "STATUT"
    ctrl.BackColor = RGB(180, 180, 180)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.Label", "lblColAction")
    ctrl.Left = 6740
    ctrl.Top = 1800
    ctrl.Width = 2000
    ctrl.Height = 250
    ctrl.Caption = "ACTION"
    ctrl.BackColor = RGB(180, 180, 180)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.Label", "lblColComm")
    ctrl.Left = 8740
    ctrl.Top = 1800
    ctrl.Width = 1500
    ctrl.Height = 250
    ctrl.Caption = "COMMENTAIRE"
    ctrl.BackColor = RGB(180, 180, 180)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' Zone de détails
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitreDetails")
    ctrl.Left = 240
    ctrl.Top = 6000
    ctrl.Width = 10000
    ctrl.Height = 300
    ctrl.Caption = "=== DÉTAILS DU FRIGO SÉLECTIONNÉ ==="
    ctrl.BackColor = RGB(255, 200, 100)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtDetails")
    ctrl.Left = 240
    ctrl.Top = 6360
    ctrl.Width = 10000
    ctrl.Height = 1200
    ctrl.MultiLine = True
    ctrl.ScrollBars = 2
    ctrl.BackColor = RGB(255, 255, 240)
    ctrl.Text = "Sélectionnez un frigo pour voir les détails..."
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdAffecterPieces")
    ctrl.Left = 2000
    ctrl.Top = 7640
    ctrl.Width = 2500
    ctrl.Height = 400
    ctrl.Caption = "?? AFFECTER PIÈCES"
    ctrl.Font.Bold = True
    ctrl.Font.Size = 11
    ctrl.BackColor = RGB(100, 255, 100)
    ctrl.Enabled = False
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdMarquerRepare")
    ctrl.Left = 4600
    ctrl.Top = 7640
    ctrl.Width = 2500
    ctrl.Height = 400
    ctrl.Caption = "? MARQUER RÉPARÉ"
    ctrl.Font.Bold = True
    ctrl.Font.Size = 11
    ctrl.BackColor = RGB(128, 255, 128)
    ctrl.Enabled = False
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdFermer")
    ctrl.Left = 7200
    ctrl.Top = 7640
    ctrl.Width = 1500
    ctrl.Height = 400
    ctrl.Caption = "? FERMER"
    ctrl.Font.Bold = True
    ctrl.BackColor = RGB(255, 128, 128)
    ctrl.Visible = True
End Sub

Private Sub ChargerStockReparable()
    Dim fichier As String
    Dim numeroFichier As Integer
    Dim ligne As String
    Dim elements() As String
    Dim topPosition As Long
    Dim compteur As Integer
    
    fichier = App.Path & "\StockReparable.txt"
    topPosition = 2080
    compteur = 0
    
    If Dir(fichier) = "" Then
        Me.Controls("lblStats").Caption = "Aucun frigo réparable en stock"
        Exit Sub
    End If
    
    numeroFichier = FreeFile
    Open fichier For Input As #numeroFichier
    
    Do While Not EOF(numeroFichier)
        Line Input #numeroFichier, ligne
        If Len(Trim(ligne)) > 0 Then
            elements = Split(ligne, "|")
            If UBound(elements) >= 4 Then
                CreerLigneFrigo topPosition, elements(0), elements(1), elements(2), elements(3), elements(4), compteur
                topPosition = topPosition + 400
                compteur = compteur + 1
            End If
        End If
    Loop
    
    Close #numeroFichier
    
    ' Mettre à jour les statistiques
    Me.Controls("lblStats").Caption = "Total: " & compteur & " frigos réparables en stock | Dernière mise à jour: " & Format(Now, "dd/mm/yyyy hh:nn:ss")
End Sub

Private Sub CreerLigneFrigo(topPos As Long, reference As String, frigoriste As String, dateAjout As String, statut As String, commentaire As String, index As Integer)
    Dim ctrl As Object
    
    ' Référence
    Set ctrl = Me.Controls.Add("VB.Label", "lblRef" & index)
    ctrl.Left = 240
    ctrl.Top = topPos
    ctrl.Width = 2000
    ctrl.Height = 350
    ctrl.Caption = reference
    ctrl.BackColor = RGB(255, 255, 255)
    ctrl.BorderStyle = 1
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Frigoriste
    Set ctrl = Me.Controls.Add("VB.Label", "lblFrig" & index)
    ctrl.Left = 2240
    ctrl.Top = topPos
    ctrl.Width = 1800
    ctrl.Height = 350
    ctrl.Caption = frigoriste
    ctrl.BackColor = RGB(255, 255, 255)
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' Date
    Set ctrl = Me.Controls.Add("VB.Label", "lblDate" & index)
    ctrl.Left = 4040
    ctrl.Top = topPos
    ctrl.Width = 1500
    ctrl.Height = 350
    ctrl.Caption = Left(dateAjout, 10)
    ctrl.BackColor = RGB(255, 255, 255)
    ctrl.BorderStyle = 1
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Statut avec couleur
    Set ctrl = Me.Controls.Add("VB.Label", "lblStat" & index)
    ctrl.Left = 5540
    ctrl.Top = topPos
    ctrl.Width = 1200
    ctrl.Height = 350
    ctrl.Caption = statut
    ctrl.BorderStyle = 1
    ctrl.Alignment = 2
    ctrl.Font.Bold = True
    Select Case UCase(statut)
        Case "DISPONIBLE"
            ctrl.BackColor = RGB(128, 255, 128)
        Case "EN_COURS"
            ctrl.BackColor = RGB(255, 255, 128)
        Case "REPARE"
            ctrl.BackColor = RGB(100, 255, 100)
        Case Else
            ctrl.BackColor = RGB(255, 255, 255)
    End Select
    ctrl.Visible = True
    
    ' Bouton Sélectionner
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdSelect" & index)
    ctrl.Left = 6740
    ctrl.Top = topPos
    ctrl.Width = 1000
    ctrl.Height = 350
    ctrl.Caption = "?? Sélect."
    ctrl.Font.Size = 9
    ctrl.Tag = index
    ctrl.Visible = True
    
    ' Bouton Détails
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdDetails" & index)
    ctrl.Left = 7740
    ctrl.Top = topPos
    ctrl.Width = 1000
    ctrl.Height = 350
    ctrl.Caption = "?? Détails"
    ctrl.Font.Size = 9
    ctrl.Tag = index
    ctrl.Visible = True
    
    ' Commentaire (tronqué)
    Set ctrl = Me.Controls.Add("VB.Label", "lblComm" & index)
    ctrl.Left = 8740
    ctrl.Top = topPos
    ctrl.Width = 1500
    ctrl.Height = 350
    ctrl.Caption = Left(commentaire, 25) & IIf(Len(commentaire) > 25, "...", "")
    ctrl.BackColor = RGB(255, 255, 255)
    ctrl.BorderStyle = 1
    ctrl.Visible = True
End Sub

' Gestion dynamique des clics sur les boutons (à adapter selon le nombre de lignes)
Private Sub cmdSelect0_Click()
    SelectionnerFrigo 0
End Sub
Private Sub cmdSelect1_Click()
    SelectionnerFrigo 1
End Sub
Private Sub cmdSelect2_Click()
    SelectionnerFrigo 2
End Sub
Private Sub cmdSelect3_Click()
    SelectionnerFrigo 3
End Sub
Private Sub cmdSelect4_Click()
    SelectionnerFrigo 4
End Sub
Private Sub cmdSelect5_Click()
    SelectionnerFrigo 5
End Sub
Private Sub cmdSelect6_Click()
    SelectionnerFrigo 6
End Sub
Private Sub cmdSelect7_Click()
    SelectionnerFrigo 7
End Sub
Private Sub cmdSelect8_Click()
    SelectionnerFrigo 8
End Sub
Private Sub cmdSelect9_Click()
    SelectionnerFrigo 9
End Sub

Private Sub cmdDetails0_Click()
    AfficherDetails 0
End Sub
Private Sub cmdDetails1_Click()
    AfficherDetails 1
End Sub
Private Sub cmdDetails2_Click()
    AfficherDetails 2
End Sub
Private Sub cmdDetails3_Click()
    AfficherDetails 3
End Sub
Private Sub cmdDetails4_Click()
    AfficherDetails 4
End Sub
Private Sub cmdDetails5_Click()
    AfficherDetails 5
End Sub
Private Sub cmdDetails6_Click()
    AfficherDetails 6
End Sub
Private Sub cmdDetails7_Click()
    AfficherDetails 7
End Sub
Private Sub cmdDetails8_Click()
    AfficherDetails 8
End Sub
Private Sub cmdDetails9_Click()
    AfficherDetails 9
End Sub

Private frigoSelectionne As Integer
Private referenceSelectionnee As String

Private Sub SelectionnerFrigo(index As Integer)
    frigoSelectionne = index
    referenceSelectionnee = Me.Controls("lblRef" & index).Caption
    
    ' Désélectionner tous les autres
    For i = 0 To 9
        On Error Resume Next
        Me.Controls("lblRef" & i).BackColor = RGB(255, 255, 255)
        Me.Controls("cmdSelect" & i).Caption = "?? Sélect."
        Me.Controls("cmdSelect" & i).BackColor = vbButtonFace
    Next i
    
    ' Sélectionner le frigo actuel
    Me.Controls("lblRef" & index).BackColor = RGB(255, 200, 100)
    Me.Controls("cmdSelect" & index).Caption = "? Sélect."
    Me.Controls("cmdSelect" & index).BackColor = RGB(100, 255, 100)
    
    ' Activer les boutons d'action
    Me.Controls("cmdAffecterPieces").Enabled = True
    Me.Controls("cmdMarquerRepare").Enabled = True
    
    ' Afficher les détails
    AfficherDetails index
    
    MsgBox "Frigo sélectionné: " & referenceSelectionnee, vbInformation
End Sub

Private Sub AfficherDetails(index As Integer)
    Dim details As String
    Dim reference As String
    
    reference = Me.Controls("lblRef" & index).Caption
    
    details = "=== DÉTAILS DU FRIGO ===" & vbCrLf
    details = details & "Référence: " & Me.Controls("lblRef" & index).Caption & vbCrLf
    details = details & "Frigoriste: " & Me.Controls("lblFrig" & index).Caption & vbCrLf
    details = details & "Date d'ajout: " & Me.Controls("lblDate" & index).Caption & vbCrLf
    details = details & "Statut: " & Me.Controls("lblStat" & index).Caption & vbCrLf
    details = details & "Commentaire: " & Me.Controls("lblComm" & index).Caption & vbCrLf
    details = details & String(50, "-") & vbCrLf & vbCrLf
    
    ' Récupérer plus d'infos depuis la fiche originale
    details = details & "INFORMATIONS SUPPLÉMENTAIRES:" & vbCrLf
    details = details & ObtenirDetailsFiche(reference)
    
    ' Afficher les pièces disponibles pour affectation
    details = details & vbCrLf & "PIÈCES DISPONIBLES POUR AFFECTATION:" & vbCrLf
    details = details & ObtenirStockPiecesDisponibles()
    
    Me.Controls("txtDetails").Text = details
End Sub

Private Function ObtenirDetailsFiche(reference As String) As String
    Dim fichier As String
    Dim details As String
    
    ' Chercher dans le répertoire des fiches
    fichier = Dir(App.Path & "\Fiches\Fiche_" & reference & "_*.txt")
    
    If fichier <> "" Then
        details = "Fiche trouvée: " & fichier & vbCrLf
        details = details & "Type de problème diagnostiqué disponible dans la fiche" & vbCrLf
    Else
        details = "Fiche détaillée non trouvée" & vbCrLf
        details = details & "Informations limitées disponibles" & vbCrLf
    End If
    
    ObtenirDetailsFiche = details
End Function

Private Function ObtenirStockPiecesDisponibles() As String
    Dim fichier As String
    Dim numeroFichier As Integer
    Dim ligne As String
    Dim elements() As String
    Dim stock As String
    Dim compteurPieces As String
    
    fichier = App.Path & "\StockPieces.txt"
    stock = ""
    
    If Dir(fichier) = "" Then
        stock = "Aucun stock de pièces disponible"
    Else
        numeroFichier = FreeFile
        Open fichier For Input As #numeroFichier
        
        ' Ignorer l'en-tête
        If Not EOF(numeroFichier) Then Line Input #numeroFichier, ligne
        
        Do While Not EOF(numeroFichier)
            Line Input #numeroFichier, ligne
            If Len(Trim(ligne)) > 0 Then
                elements = Split(ligne, "|")
                If UBound(elements) >= 6 Then
                    stock = stock & "• " & elements(0) & " - " & elements(1) & " (Qté: " & elements(2) & ", État: " & elements(3) & ")" & vbCrLf
                End If
            End If
        Loop
        
        Close #numeroFichier
    End If
    
    If stock = "" Then stock = "Aucune pièce en stock actuellement"
    
    ObtenirStockPiecesDisponibles = stock
End Function

Private Sub cmdAffecterPieces_Click()
    If frigoSelectionne < 0 Then
        MsgBox "Veuillez d'abord sélectionner un frigo !", vbExclamation
        Exit Sub
    End If
    
    Load frmAffectationPieces
    frmAffectationPieces.InitialiserAvecFrigo referenceSelectionnee, Me.Controls("lblFrig" & frigoSelectionne).Caption
    frmAffectationPieces.Show
End Sub

Private Sub cmdMarquerRepare_Click()
    If frigoSelectionne < 0 Then
        MsgBox "Veuillez d'abord sélectionner un frigo !", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Marquer le frigo " & referenceSelectionnee & " comme RÉPARÉ ?", vbYesNo + vbQuestion) = vbYes Then
        MarquerFrigoRepare
        cmdActualiser_Click
        MsgBox "Frigo marqué comme RÉPARÉ avec succès !", vbInformation
    End If
End Sub

Private Sub MarquerFrigoRepare()
    ' Mettre à jour le statut dans le fichier
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
            If UBound(elements) >= 4 And elements(0) = referenceSelectionnee Then
                ' Modifier le statut
                elements(3) = "REPARE"
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
    
    ' Ajouter au stock des frigos réparés
    AjouterAuStockRepare
End Sub

Private Sub AjouterAuStockRepare()
    Dim fichier As String
    Dim numeroFichier As Integer
    
    fichier = App.Path & "\StockRepare.txt"
    numeroFichier = FreeFile
    
    ' En-tête si nouveau fichier
    If Dir(fichier) = "" Then
        Open fichier For Output As #numeroFichier
        Print #numeroFichier, "REFERENCE|FRIGORISTE|DATE_REPARATION|PIECES_UTILISEES|COUT_REPARATION"
        Close #numeroFichier
    End If
    
    Open fichier For Append As #numeroFichier
    Print #numeroFichier, referenceSelectionnee & "|" & Me.Controls("lblFrig" & frigoSelectionne).Caption & "|" & Format(Now, "dd/mm/yyyy hh:nn:ss") & "|A_DETERMINER|0.00"
    Close #numeroFichier
End Sub

Private Sub cmdActualiser_Click()
    ' Effacer les contrôles existants des lignes
    For i = 0 To 20
        On Error Resume Next
        Me.Controls.Remove "lblRef" & i
        Me.Controls.Remove "lblFrig" & i
        Me.Controls.Remove "lblDate" & i
        Me.Controls.Remove "lblStat" & i
        Me.Controls.Remove "cmdSelect" & i
        Me.Controls.Remove "cmdDetails" & i
        Me.Controls.Remove "lblComm" & i
    Next i
    
    ' Recharger
    ChargerStockReparable
    Me.Controls("txtDetails").Text = "Données actualisées - Sélectionnez un frigo pour voir les détails..."
    
    ' Réinitialiser la sélection
    frigoSelectionne = -1
    referenceSelectionnee = ""
    Me.Controls("cmdAffecterPieces").Enabled = False
    Me.Controls("cmdMarquerRepare").Enabled = False
End Sub

Private Sub cmdVoirStock_Click()
    Load frmStockPieces
    frmStockPieces.Show
End Sub

Private Sub cmdExporter_Click()
    ExporterStockCSV
End Sub

Private Sub ExporterStockCSV()
    Dim fichier As String
    Dim fichierCSV As String
    Dim numeroFichier As Integer
    Dim numeroCSV As Integer
    Dim ligne As String
    
    fichier = App.Path & "\StockReparable.txt"
    fichierCSV = App.Path & "\Export_StockReparable_" & Format(Now, "yyyymmdd_hhnnss") & ".csv"
    
    If Dir(fichier) = "" Then
        MsgBox "Aucune donnée à exporter", vbInformation
        Exit Sub
    End If
    
    numeroFichier = FreeFile
    numeroCSV = FreeFile + 1
    
    Open fichier For Input As #numeroFichier
    Open fichierCSV For Output As #numeroCSV
    
    ' En-tête CSV
    Print #numeroCSV, "REFERENCE;FRIGORISTE;DATE_AJOUT;STATUT;COMMENTAIRE"
    
    Do While Not EOF(numeroFichier)
        Line Input #numeroFichier, ligne
        If Len(Trim(ligne)) > 0 Then
            ' Remplacer | par ;
            ligne = Replace(ligne, "|", ";")
            Print #numeroCSV, ligne
        End If
    Loop
    
    Close #numeroFichier
    Close #numeroCSV
    
    MsgBox "Export terminé !" & vbCrLf & "Fichier: " & fichierCSV, vbInformation
End Sub

Private Sub cmdFermer_Click()
    Me.Hide
End Sub
