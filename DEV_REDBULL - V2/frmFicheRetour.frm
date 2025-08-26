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
' === FRMFICHERETOUR.FRM - FORMULAIRE FICHE RETOUR ===

Private referenceFrigo As String

Private Sub Form_Load()
    Me.BackColor = RGB(240, 240, 240)
    Me.Caption = "Fiche Retour SAV - " & referenceFrigo
    Me.Width = 12000
    Me.Height = 10000
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    CreerInterfaceFiche
End Sub

Public Sub InitialiserAvecReference(reference As String)
    referenceFrigo = reference
    Me.Caption = "Fiche Retour SAV - " & referenceFrigo
End Sub

Private Sub CreerInterfaceFiche()
    Dim ctrl As Object
    
    ' Titre
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitre")
    ctrl.Left = 240
    ctrl.Top = 120
    ctrl.Width = 8295
    ctrl.Height = 375
    ctrl.Caption = "?? FICHE RETOUR SAV RED BULL ??"
    ctrl.BackColor = RGB(51, 102, 255)
    ctrl.ForeColor = RGB(255, 255, 255)
    ctrl.Font.Size = 14
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' R�f�rence (pr�-remplie)
    Set ctrl = Me.Controls.Add("VB.Label", "lblRef")
    ctrl.Left = 480
    ctrl.Top = 600
    ctrl.Width = 1500
    ctrl.Height = 255
    ctrl.Caption = "R�f�rence produit:"
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtReference")
    ctrl.Left = 480
    ctrl.Top = 840
    ctrl.Width = 4000
    ctrl.Height = 285
    ctrl.Text = referenceFrigo
    ctrl.Enabled = False
    ctrl.BackColor = RGB(240, 240, 240)
    ctrl.Visible = True
    
    ' Informations frigoriste
    Set ctrl = Me.Controls.Add("VB.Label", "lblFrigoriste")
    ctrl.Left = 480
    ctrl.Top = 1200
    ctrl.Width = 1500
    ctrl.Height = 255
    ctrl.Caption = "Nom frigoriste:"
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtFrigoriste")
    ctrl.Left = 480
    ctrl.Top = 1440
    ctrl.Width = 4000
    ctrl.Height = 285
    ctrl.Visible = True
    
    ' Date
    Set ctrl = Me.Controls.Add("VB.Label", "lblDate")
    ctrl.Left = 480
    ctrl.Top = 1800
    ctrl.Width = 1500
    ctrl.Height = 255
    ctrl.Caption = "Date d'intervention:"
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtDate")
    ctrl.Left = 480
    ctrl.Top = 2040
    ctrl.Width = 2000
    ctrl.Height = 285
    ctrl.Text = Format(Date, "dd/mm/yyyy")
    ctrl.Visible = True
    
    ' Motif du retour
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitreMotif")
    ctrl.Left = 240
    ctrl.Top = 2400
    ctrl.Width = 8295
    ctrl.Height = 300
    ctrl.Caption = "=== MOTIF DU RETOUR ==="
    ctrl.BackColor = RGB(200, 200, 200)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.OptionButton", "optMecanique")
    ctrl.Left = 480
    ctrl.Top = 2760
    ctrl.Width = 1575
    ctrl.Height = 255
    ctrl.Caption = "?? M�CANIQUE"
    ctrl.Value = True
    ctrl.BackColor = RGB(255, 128, 128)
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.OptionButton", "optEsthetique")
    ctrl.Left = 2400
    ctrl.Top = 2760
    ctrl.Width = 1575
    ctrl.Height = 255
    ctrl.Caption = "?? ESTH�TIQUE"
    ctrl.Visible = True
    
    ' Diagnostic technique
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitreDiag")
    ctrl.Left = 240
    ctrl.Top = 3120
    ctrl.Width = 8295
    ctrl.Height = 300
    ctrl.Caption = "=== DIAGNOSTIC TECHNIQUE ==="
    ctrl.BackColor = RGB(200, 200, 200)
    ctrl.Font.Bold = True
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkCompresseur")
    ctrl.Left = 480
    ctrl.Top = 3480
    ctrl.Width = 3000
    ctrl.Height = 255
    ctrl.Caption = "?? Probl�me compresseur"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkEclairage")
    ctrl.Left = 480
    ctrl.Top = 3760
    ctrl.Width = 3000
    ctrl.Height = 255
    ctrl.Caption = "?? Probl�me �clairage"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkVitre")
    ctrl.Left = 480
    ctrl.Top = 4040
    ctrl.Width = 3000
    ctrl.Height = 255
    ctrl.Caption = "?? Vitre cass�e/ray�e"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkThermostat")
    ctrl.Left = 4000
    ctrl.Top = 3480
    ctrl.Width = 3000
    ctrl.Height = 255
    ctrl.Caption = "??? Thermostat d�faillant"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkJoints")
    ctrl.Left = 4000
    ctrl.Top = 3760
    ctrl.Width = 3000
    ctrl.Height = 255
    ctrl.Caption = "?? Joints ab�m�s"
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CheckBox", "chkAutre")
    ctrl.Left = 4000
    ctrl.Top = 4040
    ctrl.Width = 3000
    ctrl.Height = 255
    ctrl.Caption = "? Autre probl�me"
    ctrl.Visible = True
    
    ' Zone commentaires
    Set ctrl = Me.Controls.Add("VB.Label", "lblCommentaires")
    ctrl.Left = 480
    ctrl.Top = 4400
    ctrl.Width = 2000
    ctrl.Height = 255
    ctrl.Caption = "Commentaires d�taill�s:"
    ctrl.Font.Bold = True
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtCommentaires")
    ctrl.Left = 480
    ctrl.Top = 4640
    ctrl.Width = 6000
    ctrl.Height = 800
    ctrl.MultiLine = True
    ctrl.ScrollBars = 2 ' Vertical
    ctrl.Visible = True
    
    ' D�CISION FINALE - HS ou R�PARABLE
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitreDecision")
    ctrl.Left = 240
    ctrl.Top = 5520
    ctrl.Width = 8295
    ctrl.Height = 400
    ctrl.Caption = "=== D�CISION FINALE ==="
    ctrl.BackColor = RGB(255, 200, 100)
    ctrl.Font.Bold = True
    ctrl.Font.Size = 12
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.OptionButton", "optHS")
    ctrl.Left = 1000
    ctrl.Top = 6000
    ctrl.Width = 2500
    ctrl.Height = 400
    ctrl.Caption = "? HORS SERVICE (HS)" & vbCrLf & "R�cup�ration pi�ces"
    ctrl.Font.Size = 11
    ctrl.Font.Bold = True
    ctrl.BackColor = RGB(255, 128, 128)
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.OptionButton", "optReparable")
    ctrl.Left = 4000
    ctrl.Top = 6000
    ctrl.Width = 2500
    ctrl.Height = 400
    ctrl.Caption = "? R�PARABLE" & vbCrLf & "Mise en stock r�parable"
    ctrl.Font.Size = 11
    ctrl.Font.Bold = True
    ctrl.BackColor = RGB(128, 255, 128)
    ctrl.Visible = True
    
    ' Boutons d'action
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdValider")
    ctrl.Left = 2000
    ctrl.Top = 6600
    ctrl.Width = 1800
    ctrl.Height = 500
    ctrl.Caption = "? VALIDER FICHE"
    ctrl.Font.Bold = True
    ctrl.Font.Size = 12
    ctrl.BackColor = RGB(128, 255, 128)
    ctrl.Visible = True
    
    Set ctrl = Me.Controls.Add("VB.CommandButton", "cmdAnnuler")
    ctrl.Left = 4000
    ctrl.Top = 6600
    ctrl.Width = 1800
    ctrl.Height = 500
    ctrl.Caption = "? ANNULER"
    ctrl.Font.Bold = True
    ctrl.Font.Size = 12
    ctrl.BackColor = RGB(255, 128, 128)
    ctrl.Visible = True
    
    ' Pr�-remplir la r�f�rence si disponible
    If Len(referenceFrigo) > 0 Then
        Me.Controls("txtReference").Text = referenceFrigo
    End If
End Sub

Private Sub cmdValider_Click()
    If Not ValiderFormulaire() Then Exit Sub
    
    If Me.Controls("optHS").Value = True Then
        ' Frigo HS - Ouvrir formulaire de r�cup�ration pi�ces
        TraiterFrigoHS
    ElseIf Me.Controls("optReparable").Value = True Then
        ' Frigo r�parable - Ajouter au stock r�parable
        TraiterFrigoReparable
    Else
        MsgBox "Veuillez choisir si le frigo est HS ou R�parable !", vbExclamation
        Exit Sub
    End If
End Sub

Private Function ValiderFormulaire() As Boolean
    If Len(Trim(Me.Controls("txtFrigoriste").Text)) = 0 Then
        MsgBox "Veuillez saisir le nom du frigoriste !", vbExclamation
        Me.Controls("txtFrigoriste").SetFocus
        ValiderFormulaire = False
        Exit Function
    End If
    
    If Len(Trim(Me.Controls("txtCommentaires").Text)) < 10 Then
        MsgBox "Veuillez ajouter des commentaires d�taill�s (minimum 10 caract�res) !", vbExclamation
        Me.Controls("txtCommentaires").SetFocus
        ValiderFormulaire = False
        Exit Function
    End If
    
    ValiderFormulaire = True
End Function

Private Sub TraiterFrigoHS()
    ' Sauvegarder la fiche
    SauvegarderFiche "HS"
    
    ' Ouvrir formulaire de r�cup�ration des pi�ces
    Load frmRecuperationPieces
    frmRecuperationPieces.InitialiserAvecFrigo referenceFrigo, Me.Controls("txtFrigoriste").Text
    frmRecuperationPieces.Show
    
    MsgBox "Frigo marqu� comme HORS SERVICE." & vbCrLf & "Ouverture du formulaire de r�cup�ration des pi�ces...", vbInformation
    Me.Hide
End Sub

Private Sub TraiterFrigoReparable()
    ' Sauvegarder la fiche
    SauvegarderFiche "REPARABLE"
    
    ' Ajouter au fichier stock r�parable (simulation base de donn�es)
    AjouterAuStockReparable
    
    MsgBox "Frigo ajout� au stock R�PARABLE avec succ�s !" & vbCrLf & "Il est maintenant disponible pour recevoir des pi�ces.", vbInformation
    Me.Hide
End Sub

Private Sub SauvegarderFiche(statut As String)
    Dim fichier As String
    Dim numeroFichier As Integer
    
    fichier = App.Path & "\Fiches\Fiche_" & referenceFrigo & "_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    numeroFichier = FreeFile
    
    ' Cr�er le r�pertoire s'il n'existe pas
    If Dir(App.Path & "\Fiches", vbDirectory) = "" Then
        MkDir App.Path & "\Fiches"
    End If
    
    Open fichier For Output As #numeroFichier
    Print #numeroFichier, "=== FICHE RETOUR SAV RED BULL ==="
    Print #numeroFichier, "R�f�rence: " & referenceFrigo
    Print #numeroFichier, "Frigoriste: " & Me.Controls("txtFrigoriste").Text
    Print #numeroFichier, "Date: " & Me.Controls("txtDate").Text
    Print #numeroFichier, "Statut final: " & statut
    Print #numeroFichier, ""
    Print #numeroFichier, "MOTIF:"
    If Me.Controls("optMecanique").Value Then Print #numeroFichier, "- M�CANIQUE"
    If Me.Controls("optEsthetique").Value Then Print #numeroFichier, "- ESTH�TIQUE"
    Print #numeroFichier, ""
    Print #numeroFichier, "DIAGNOSTIC:"
    If Me.Controls("chkCompresseur").Value = 1 Then Print #numeroFichier, "- Probl�me compresseur"
    If Me.Controls("chkEclairage").Value = 1 Then Print #numeroFichier, "- Probl�me �clairage"
    If Me.Controls("chkVitre").Value = 1 Then Print #numeroFichier, "- Vitre cass�e/ray�e"
    If Me.Controls("chkThermostat").Value = 1 Then Print #numeroFichier, "- Thermostat d�faillant"
    If Me.Controls("chkJoints").Value = 1 Then Print #numeroFichier, "- Joints ab�m�s"
    If Me.Controls("chkAutre").Value = 1 Then Print #numeroFichier, "- Autre probl�me"
    Print #numeroFichier, ""
    Print #numeroFichier, "COMMENTAIRES:"
    Print #numeroFichier, Me.Controls("txtCommentaires").Text
    Print #numeroFichier, ""
    Print #numeroFichier, "Date cr�ation: " & Now
    Close #numeroFichier
End Sub

Private Sub AjouterAuStockReparable()
    Dim fichier As String
    Dim numeroFichier As Integer
    
    fichier = App.Path & "\StockReparable.txt"
    numeroFichier = FreeFile
    
    Open fichier For Append As #numeroFichier
    Print #numeroFichier, referenceFrigo & "|" & Me.Controls("txtFrigoriste").Text & "|" & Format(Now, "dd/mm/yyyy hh:nn:ss") & "|DISPONIBLE|" & Me.Controls("txtCommentaires").Text
    Close #numeroFichier
End Sub

Private Sub cmdAnnuler_Click()
    If MsgBox("�tes-vous s�r de vouloir annuler cette fiche ?", vbYesNo + vbQuestion) = vbYes Then
        Me.Hide
    End If
End Sub
