VERSION 5.00
Begin VB.Form frmStockPieces 
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
Attribute VB_Name = "frmStockPieces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' === FRMSTOCKPIECES.FRM - VISUALISATION STOCK PIÈCES ===

Private Sub Form_Load()
    Me.BackColor = RGB(240, 240, 240)
    Me.Caption = "Stock des Pièces Récupérées"
    Me.Width = 14000
    Me.Height = 10000
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    CreerInterfaceStockPieces
    ChargerStockPieces
End Sub

Private Sub CreerInterfaceStockPieces()
    Dim ctrl As Object
    
    ' Titre
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitre")
    ctrl.Left = 240
    ctrl.Top = 120
    ctrl.Width = 10000
    ctrl.Height = 400
    ctrl.Caption = "?? STOCK DES PIÈCES RÉCUPÉRÉES ??"
    ctrl.BackColor = RGB(100, 100, 255)
    ctrl.ForeColor = RGB(255, 255, 255)
    ctrl.Font.Size = 16
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Statistiques
    Set ctrl = Me.Controls.Add("VB.Label", "lblStatistiques")
    ctrl.Left = 240
    ctrl.Top = 600
    ctrl.Width = 10000
    ctrl.Height = 300
    ctrl.Caption = "Chargement des statistiques..."
    ctrl.BackColor = RGB(200, 200, 255)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Boutons de gestion
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdActualiser")
    ctrl.Left = 240
    ctrl.Top = 960
    ctrl.Width = 1500
    ctrl.Height = 400
    ctrl.Caption = "?? Actualiser"
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdFiltrer")
    ctrl.Left = 1800
    ctrl.Top = 960
    ctrl.Width = 1500
    ctrl.Height = 400
    ctrl.Caption = "?? Filtrer"
    ctrl.Font.Bold = True
    ctrl.BackColor = RGB(255, 200, 100)
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdExporter")
    ctrl.Left = 3360
    ctrl.Top = 960
    ctrl.Width = 1500
    ctrl.Height = 400
    ctrl.Caption = "?? Exporter"
    ctrl.Font.Bold = True
    ctrl.BackColor = RGB(200, 255, 200)
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdInventaire")
    ctrl.Left = 4920
    ctrl.Top = 960
    ctrl.Width = 1500
    ctrl.Height = 400
    ctrl.Caption = "?? Inventaire"
    ctrl.Font.Bold = True
    ctrl.BackColor = RGB(255, 255, 200)
    ctrl.Visible = True
    
    ' Filtre par état
    Set ctrl = Me.Controls.Add("VB.Label", "lblFiltreEtat")
    ctrl.Left = 6600
    ctrl.Top = 960
    ctrl.Width = 1000
    ctrl.Height = 200
    ctrl.Caption = "Filtre état:"
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.ComboBox", "cmbFiltreEtat")
    ctrl.Left = 6600
    ctrl.Top = 1160
    ctrl.Width = 1500
    ctrl.Height = 200
    ctrl.AddItem "Tous"
    ctrl.AddItem "Excellent"
    ctrl.AddItem "Bon"
    ctrl.AddItem "Moyen"
    ctrl.AddItem "Défectueux"
    ctrl.ListIndex = 0
    ctrl.Visible = True
    
    ' En-têtes du tableau
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitreTableau")
    ctrl.Left = 240
    ctrl.Top = 1440
    ctrl.Width = 10000
    ctrl.Height = 300
    ctrl.Caption = "=== DÉTAIL DU STOCK ==="
    ctrl.BackColor = RGB(200, 200, 200)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    CreerEnTetesTableau
    
    ' Zone de résumé
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitreResume")
    ctrl.Left = 240
    ctrl.Top = 7200
    ctrl.Width = 10000
    ctrl.Height = 300
    ctrl.Caption = "=== RÉSUMÉ PAR CATÉGORIE ==="
    ctrl.BackColor = RGB(255, 200, 100)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtResume")
    ctrl.Left = 240
    ctrl.Top = 7560
    ctrl.Width = 10000
    ctrl.Height = 1200
    ctrl.MultiLine = True
    ctrl.ScrollBars = 2
    ctrl.BackColor = RGB(255, 255, 240)
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdFermer")
    ctrl.Left = 8500
    ctrl.Top = 8840
    ctrl.Width = 1500
    ctrl.Height = 400
    ctrl.Caption = "? FERMER"
    ctrl.Font.Bold = True
    ctrl.BackColor = RGB(255, 128, 128)
    ctrl.Visible = True
End Sub

Private Sub CreerEnTetesTableau()
    Dim ctrl As Object
    Dim topPos As Long
    topPos = 1800
    
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
    ctrl.Width = 2000
    ctrl.Height = 250
    ctrl.Caption = "PIÈCE"
    ctrl.BackColor = RGB(180, 180, 180)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' Quantité
    Set ctrl = Me.Controls.Add("VB.Label", "lblHQuantite")
    ctrl.Left = 3040
    ctrl.Top = topPos
    ctrl.Width = 600
    ctrl.Height = 250
    ctrl.Caption = "QTÉ"
    ctrl.BackColor = RGB(180, 180, 180)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' État
    Set ctrl = Me.Controls.Add("VB.Label", "lblHEtat")
    ctrl.Left = 3640
    ctrl.Top = topPos
    ctrl.Width = 1000
    ctrl.Height = 250
    ctrl.Caption = "ÉTAT"
    ctrl.BackColor = RGB(180, 180, 180)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' Origine
    Set ctrl = Me.Controls.Add("VB.Label", "lblHOrigine")
    ctrl.Left = 4640
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
    ctrl.Left = 6140
    ctrl.Top = topPos
    ctrl.Width = 1200
    ctrl.Height = 250
    ctrl.Caption = "DATE"
    ctrl.BackColor = RGB(180, 180, 180)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' Prix unitaire
    Set ctrl = Me.Controls.Add("VB.Label", "lblHPrix")
    ctrl.Left = 7340
    ctrl.Top = topPos
    ctrl.Width = 800
    ctrl.Height = 250
    ctrl.Caption = "PRIX U."
    ctrl.BackColor = RGB(180, 180, 180)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' Valeur totale
    Set ctrl = Me.Controls.Add("VB.Label", "lblHValeur")
    ctrl.Left = 8140
    ctrl.Top = topPos
    ctrl.Width = 800
    ctrl.Height = 250
    ctrl.Caption = "VAL. TOT."
    ctrl.BackColor = RGB(180, 180, 180)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' Statut
    Set ctrl = Me.Controls.Add("VB.Label", "lblHStatut")
    ctrl.Left = 8940
    ctrl.Top = topPos
    ctrl.Width = 1000
    ctrl.Height = 250
    ctrl.Caption = "STATUT"
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
    Dim filtreEtat As String
    
    fichier = App.Path & "\StockPieces.txt"
    topPosition = 2080
    compteur = 0
    filtreEtat = Me.Controls("cmbFiltreEtat").Text
    
    ' Nettoyer les anciennes lignes
    For i = 0 To 50
        On Error Resume Next
        Me.Controls.Remove "lblCode" & i
        Me.Controls.Remove "lblPiece" & i
        Me.Controls.Remove "lblQte" & i
        Me.Controls.Remove "lblEtat" & i
        Me.Controls.Remove "lblOrigine" & i
        Me.Controls.Remove "lblDate" & i
        Me.Controls.Remove "lblPrix" & i
        Me.Controls.Remove "lblValeur" & i
        Me.Controls.Remove "lblStatut" & i
    Next i
    
    If Dir(fichier) = "" Then
        Me.Controls("lblStatistiques").Caption = "Aucun stock de pièces disponible"
        Me.Controls("txtResume").Text = "Aucune donnée à afficher"
        Exit Sub
    End If
    
    numeroFichier = FreeFile
    Open fichier For Input As #numeroFichier
    
    ' Ignorer l'en-tête
    If Not EOF(numeroFichier) Then Line Input #numeroFichier, ligne
    
    Do While Not EOF(numeroFichier) And compteur < 20 ' Limiter l'affichage
        Line Input #numeroFichier, ligne
        If Len(Trim(ligne)) > 0 Then
            elements = Split(ligne, "|")
            If UBound(elements) >= 6 Then
                ' Appliquer le filtre
                If filtreEtat = "Tous" Or elements(3) = filtreEtat Then
                    CreerLignePiece topPosition, elements(0), elements(1), elements(2), elements(3), elements(4), elements(5), elements(6), compteur
                    topPosition = topPosition + 280
                    compteur = compteur + 1
                End If
            End If
        End If
    Loop
    
    Close #numeroFichier
    
    ' Calculer et afficher les statistiques
    CalculerStatistiques
    GenererResume
End Sub

Private Sub CreerLignePiece(topPos As Long, code As String, piece As String, quantite As String, etat As String, origine As String, dateAjout As String, prix As String, index As Integer)
    Dim ctrl As Object
    Dim valeurTotale As Double
    
    valeurTotale = Val(prix) * Val(quantite)
    
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
    ctrl.Font.Size = 8
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Pièce
    Set ctrl = Me.Controls.Add("VB.Label", "lblPiece" & index)
    ctrl.Left = 1040
    ctrl.Top = topPos
    ctrl.Width = 2000
    ctrl.Height = 250
    ctrl.Caption = piece
    ctrl.BackColor = RGB(255, 255, 255)
    ctrl.BorderStyle = 1
    ctrl.Font.Size = 8
    ctrl.Visible = True
    
    ' Quantité
    Set ctrl = Me.Controls.Add("VB.Label", "lblQte" & index)
    ctrl.Left = 3040
    ctrl.Top = topPos
    ctrl.Width = 600
    ctrl.Height = 250
    ctrl.Caption = quantite
    ctrl.BackColor = IIf(Val(quantite) = 0, RGB(255, 200, 200), RGB(255, 255, 255))
    ctrl.BorderStyle = 1
    ctrl.Font.Bold = True
    ctrl.Font.Size = 8
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' État avec couleur
    Set ctrl = Me.Controls.Add("VB.Label", "lblEtat" & index)
    ctrl.Left = 3640
    ctrl.Top = topPos
    ctrl.Width = 1000
    ctrl.Height = 250
    ctrl.Caption = etat
    ctrl.BorderStyle = 1
    ctrl.Font.Size = 8
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
    ctrl.Left = 4640
    ctrl.Top = topPos
    ctrl.Width = 1500
    ctrl.Height = 250
    ctrl.Caption = origine
    ctrl.BackColor = RGB(255, 255, 255)
    ctrl.BorderStyle = 1
    ctrl.Font.Size = 8
    ctrl.Visible = True
    
    ' Date
    Set ctrl = Me.Controls.Add("VB.Label", "lblDate" & index)
    ctrl.Left = 6140
    ctrl.Top = topPos
    ctrl.Width = 1200
    ctrl.Height = 250
    ctrl.Caption = Left(dateAjout, 10)
    ctrl.BackColor = RGB(255, 255, 255)
    ctrl.BorderStyle = 1
    ctrl.Font.Size = 8
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Prix unitaire
    Set ctrl = Me.Controls.Add("VB.Label", "lblPrix" & index)
    ctrl.Left = 7340
    ctrl.Top = topPos
    ctrl.Width = 800
    ctrl.Height = 250
    ctrl.Caption = Format(Val(prix), "0.00") & "€"
    ctrl.BackColor = RGB(255, 255, 255)
    ctrl.BorderStyle = 1
    ctrl.Font.Bold = True
    ctrl.Font.Size = 8
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Valeur totale
    Set ctrl = Me.Controls.Add("VB.Label", "lblValeur" & index)
    ctrl.Left = 8140
    ctrl.Top = topPos
    ctrl.Width = 800
    ctrl.Height = 250
    ctrl.Caption = Format(valeurTotale, "0.00") & "€"
    ctrl.BackColor = RGB(255, 255, 255)
    ctrl.BorderStyle = 1
    ctrl.Font.Bold = True
    ctrl.Font.Size = 8
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Statut
    Set ctrl = Me.Controls.Add("VB.Label", "lblStatut" & index)
    ctrl.Left = 8940
    ctrl.Top = topPos
    ctrl.Width = 1000
    ctrl.Height = 250
    If Val(quantite) > 0 Then
        ctrl.Caption = "DISPONIBLE"
        ctrl.BackColor = RGB(128, 255, 128)
    Else
        ctrl.Caption = "ÉPUISÉ"
        ctrl.BackColor = RGB(255, 128, 128)
    End If
    ctrl.BorderStyle = 1
    ctrl.Font.Bold = True
    ctrl.Font.Size = 8
    ctrl.Alignment = 2
    ctrl.Visible = True
End Sub

Private Sub CalculerStatistiques()
    Dim fichier As String
    Dim numeroFichier As Integer
    Dim ligne As String
    Dim elements() As String
    Dim totalPieces As Integer
    Dim totalValeur As Double
    Dim piecesDisponibles As Integer
    Dim piecesEpuisees As Integer
    
    fichier = App.Path & "\StockPieces.txt"
    
    If Dir(fichier) = "" Then Exit Sub
    
    numeroFichier = FreeFile
    Open fichier For Input As #numeroFichier
    
    ' Ignorer l'en-tête
    If Not EOF(numeroFichier) Then Line Input #numeroFichier, ligne
    
    Do While Not EOF(numeroFichier)
        Line Input #numeroFichier, ligne
        If Len(Trim(ligne)) > 0 Then
            elements = Split(ligne, "|")
            If UBound(elements) >= 6 Then
                Dim qte As Integer
                Dim prix As Double
                
                qte = Val(elements(2))
                prix = Val(elements(6))
                
                totalPieces = totalPieces + qte
                totalValeur = totalValeur + (qte * prix)
                
                If qte > 0 Then
                    piecesDisponibles = piecesDisponibles + 1
                Else
                    piecesEpuisees = piecesEpuisees + 1
                End If
            End If
        End If
    Loop
    
    Close #numeroFichier
    
    Me.Controls("lblStatistiques").Caption = "Total pièces: " & totalPieces & " | Valeur totale: " & Format(totalValeur, "0.00") & "€ | Références disponibles: " & piecesDisponibles & " | Épuisées: " & piecesEpuisees
End Sub

Private Sub GenererResume()
    Dim fichier As String
    Dim numeroFichier As Integer
    Dim ligne As String
    Dim elements() As String
    Dim resume As String
    
    ' Compteurs par type de pièce
    Dim compresseurs As Integer, leds As Integer, vitres As Integer, thermostats As Integer
    Dim joints As Integer, grilles As Integer, ventilos As Integer, capots As Integer
    Dim pieds As Integer, cables As Integer
    
    ' Valeurs par type
    Dim valCompresseurs As Double, valLeds As Double, valVitres As Double, valThermostats As Double
    Dim valJoints As Double, valGrilles As Double, valVentilos As Double, valCapots As Double
    Dim valPieds As Double, valCables As Double
    
    fichier = App.Path & "\StockPieces.txt"
    
    resume = "RÉSUMÉ DU STOCK PAR CATÉGORIE" & vbCrLf
    resume = resume & "Dernière mise à jour: " & Format(Now, "dd/mm/yyyy hh:nn:ss") & vbCrLf
    resume = resume & String(60, "=") & vbCrLf & vbCrLf
    
    If Dir(fichier) = "" Then
        resume = resume & "Aucune donnée disponible"
        Me.Controls("txtResume").Text = resume
        Exit Sub
    End If
    
    numeroFichier = FreeFile
    Open fichier For Input As #numeroFichier
    
    ' Ignorer l'en-tête
    If Not EOF(numeroFichier) Then Line Input #numeroFichier, ligne
    
    Do While Not EOF(numeroFichier)
        Line Input #numeroFichier, ligne
        If Len(Trim(ligne)) > 0 Then
            elements = Split(ligne, "|")
            If UBound(elements) >= 6 Then
                Dim code As String
                Dim qte As Integer
                Dim prix As Double
                
                code = elements(0)
                qte = Val(elements(2))
                prix = Val(elements(6))
                
                Select Case code
                    Case "COMP"
                        compresseurs = compresseurs + qte
                        valCompresseurs = valCompresseurs + (qte * prix)
                    Case "LED"
                        leds = leds + qte
                        valLeds = valLeds + (qte * prix)
                    Case "VITRE"
                        vitres = vitres + qte
                        valVitres = valVitres + (qte * prix)
                    Case "THERMO"
                        thermostats = thermostats + qte
                        valThermostats = valThermostats + (qte * prix)
                    Case "JOINT"
                        joints = joints + qte
                        valJoints = valJoints + (qte * prix)
                    Case "GRILLE"
                        grilles = grilles + qte
                        valGrilles = valGrilles + (qte * prix)
                    Case "VENTILO"
                        ventilos = ventilos + qte
                        valVentilos = valVentilos + (qte * prix)
                    Case "CAPOT"
                        capots = capots + qte
                        valCapots = valCapots + (qte * prix)
                    Case "PIED"
                        pieds = pieds + qte
                        valPieds = valPieds + (qte * prix)
                    Case "CABLE"
                        cables = cables + qte
                        valCables = valCables + (qte * prix)
                End Select
            End If
        End If
    Loop
    
    Close #numeroFichier
    
    ' Construire le résumé
    resume = resume & "?? COMPRESSEURS: " & compresseurs & " unités (" & Format(valCompresseurs, "0.00") & "€)" & vbCrLf
    resume = resume & "?? ÉCLAIRAGES LED: " & leds & " unités (" & Format(valLeds, "0.00") & "€)" & vbCrLf
    resume = resume & "?? VITRES: " & vitres & " unités (" & Format(valVitres, "0.00") & "€)" & vbCrLf
    resume = resume & "??? THERMOSTATS: " & thermostats & " unités (" & Format(valThermostats, "0.00") & "€)" & vbCrLf
    resume = resume & "?? JOINTS: " & joints & " unités (" & Format(valJoints, "0.00") & "€)" & vbCrLf
    resume = resume & "?? GRILLES: " & grilles & " unités (" & Format(valGrilles, "0.00") & "€)" & vbCrLf
    resume = resume & "?? VENTILATEURS: " & ventilos & " unités (" & Format(valVentilos, "0.00") & "€)" & vbCrLf
    resume = resume & "??? CAPOTS: " & capots & " unités (" & Format(valCapots, "0.00") & "€)" & vbCrLf
    resume = resume & "?? PIEDS: " & pieds & " unités (" & Format(valPieds, "0.00") & "€)" & vbCrLf
    resume = resume & "? CÂBLAGES: " & cables & " unités (" & Format(valCables, "0.00") & "€)" & vbCrLf
    resume = resume & vbCrLf & String(60, "-") & vbCrLf
    
    Dim totalGeneral As Double
    totalGeneral = valCompresseurs + valLeds + valVitres + valThermostats + valJoints + valGrilles + valVentilos + valCapots + valPieds + valCables
    
    resume = resume & "?? VALEUR TOTALE DU STOCK: " & Format(totalGeneral, "0.00") & "€" & vbCrLf
    
    ' Alertes stock faible
    resume = resume & vbCrLf & "?? ALERTES STOCK FAIBLE:" & vbCrLf
    If compresseurs = 0 Then resume = resume & "- COMPRESSEURS: STOCK ÉPUISÉ" & vbCrLf
    If vitres = 0 Then resume = resume & "- VITRES: STOCK ÉPUISÉ" & vbCrLf
    If thermostats = 0 Then resume = resume & "- THERMOSTATS: STOCK ÉPUISÉ" & vbCrLf
    If ventilos = 0 Then resume = resume & "- VENTILATEURS: STOCK ÉPUISÉ" & vbCrLf
    
    If compresseurs <= 1 And compresseurs > 0 Then resume = resume & "- COMPRESSEURS: STOCK FAIBLE (" & compresseurs & ")" & vbCrLf
    If vitres <= 1 And vitres > 0 Then resume = resume & "- VITRES: STOCK FAIBLE (" & vitres & ")" & vbCrLf
    
    Me.Controls("txtResume").Text = resume
End Sub

Private Sub cmdActualiser_Click()
    ChargerStockPieces
End Sub

Private Sub cmdFiltrer_Click()
    ChargerStockPieces
End Sub

Private Sub cmbFiltreEtat_Click()
    ChargerStockPieces
End Sub

Private Sub cmdExporter_Click()
    ExporterStockCSV
End Sub

Private Sub ExporterStockCSV()
    Dim fichierSource As String
    Dim fichierCSV As String
    Dim numeroSource As Integer
    Dim numeroCSV As Integer
    Dim ligne As String
    
    fichierSource = App.Path & "\StockPieces.txt"
    fichierCSV = App.Path & "\Export_StockPieces_" & Format(Now, "yyyymmdd_hhnnss") & ".csv"
    
    If Dir(fichierSource) = "" Then
        MsgBox "Aucun stock à exporter", vbInformation
        Exit Sub
    End If
    
    numeroSource = FreeFile
    numeroCSV = FreeFile + 1
    
    Open fichierSource For Input As #numeroSource
    Open fichierCSV For Output As #numeroCSV
    
    Do While Not EOF(numeroSource)
        Line Input #numeroSource, ligne
        If Len(Trim(ligne)) > 0 Then
            ' Remplacer | par ;
            ligne = Replace(ligne, "|", ";")
            Print #numeroCSV, ligne
        End If
    Loop
    
    Close #numeroSource
    Close #numeroCSV
    
    MsgBox "Export terminé !" & vbCrLf & "Fichier: " & fichierCSV, vbInformation
End Sub

Private Sub cmdInventaire_Click()
    Dim message As String
    message = "?? RAPPORT D'INVENTAIRE RAPIDE ??" & vbCrLf & vbCrLf
    message = message & "Date: " & Format(Now, "dd/mm/yyyy hh:nn:ss") & vbCrLf
    message = message & Me.Controls("lblStatistiques").Caption & vbCrLf & vbCrLf
    message = message & "Voulez-vous générer un rapport détaillé ?"
    
    If MsgBox(message, vbYesNo + vbQuestion) = vbYes Then
        GenererRapportInventaire
    End If
End Sub

Private Sub GenererRapportInventaire()
    Dim fichier As String
    Dim numeroFichier As Integer
    
    fichier = App.Path & "\Inventaire_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    numeroFichier = FreeFile
    
    Open fichier For Output As #numeroFichier
    Print #numeroFichier, "=== RAPPORT D'INVENTAIRE STOCK PIÈCES ==="
    Print #numeroFichier, "Date: " & Format(Now, "dd/mm/yyyy hh:nn:ss")
    Print #numeroFichier, ""
    Print #numeroFichier, Me.Controls("lblStatistiques").Caption
    Print #numeroFichier, ""
    Print #numeroFichier, Me.Controls("txtResume").Text
    Close #numeroFichier
    
    MsgBox "Rapport d'inventaire généré !" & vbCrLf & "Fichier: " & fichier, vbInformation
End Sub

Private Sub cmdFermer_Click()
    Me.Hide
End Sub
