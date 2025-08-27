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
Private referenceValidee As String
Private WithEvents cmdValider As CommandButton
Attribute cmdValider.VB_VarHelpID = -1
Private WithEvents cmdOuvrirFiche As CommandButton
Attribute cmdOuvrirFiche.VB_VarHelpID = -1

Private Sub Form_Load()
    Me.Caption = "SAV Red Bull - Scanner"
    Me.Width = 8000
    Me.Height = 6000
    referenceValidee = ""
    
    ' Cr�er les contr�les dynamiquement
    CreerControles
End Sub

Private Sub CreerControles()
    Dim ctrl As Object
    
    ' Label titre
    Set ctrl = Me.Controls.Add("VB.Label", "lblTitre")
    ctrl.Left = 600
    ctrl.Top = 400
    ctrl.Width = 6800
    ctrl.Height = 400
    ctrl.Caption = "SAV RED BULL - SCANNER"
    ctrl.BackColor = RGB(0, 100, 200)
    ctrl.Alignment = 2
    ctrl.Visible = True
    
    ' Label "R�f�rence frigo"
    Set ctrl = Me.Controls.Add("VB.Label", "lblRef")
    ctrl.Left = 600
    ctrl.Top = 1000
    ctrl.Width = 1500
    ctrl.Caption = "R�f�rence frigo:"
    ctrl.Visible = True
    
    ' TextBox pour saisie
    Set ctrl = Me.Controls.Add("VB.TextBox", "txtReference")
    ctrl.Left = 2200
    ctrl.Top = 1000
    ctrl.Width = 3000
    ctrl.Height = 300
    ctrl.Visible = True
    
    ' Bouton Valider
    Set cmdValider = Me.Controls.Add("VB.CommandButton", "cmdValider")
    cmdValider.Left = 5400
    cmdValider.Top = 1000
    cmdValider.Width = 1000
    cmdValider.Height = 300
    cmdValider.Caption = "VALIDER"
    cmdValider.Visible = True
    
    ' Zone d'information
    Set ctrl = Me.Controls.Add("VB.Label", "lblInfo")
    ctrl.Left = 600
    ctrl.Top = 1500
    ctrl.Width = 6800
    ctrl.Height = 1500
    ctrl.Caption = "Saisissez la r�f�rence d'un frigo et cliquez VALIDER"
    ctrl.BackColor = RGB(240, 240, 240)
    ctrl.BorderStyle = 1
    ctrl.Visible = True
    
    ' Bouton Ouvrir Fiche (d�sactiv� au d�but)
    Set cmdOuvrirFiche = Me.Controls.Add("VB.CommandButton", "cmdOuvrirFiche")
    cmdOuvrirFiche.Left = 2400
    cmdOuvrirFiche.Top = 3200
    cmdOuvrirFiche.Width = 2000
    cmdOuvrirFiche.Height = 400
    cmdOuvrirFiche.Caption = "OUVRIR FICHE RETOUR"
    cmdOuvrirFiche.Enabled = False
    cmdOuvrirFiche.BackColor = RGB(150, 150, 150)
    cmdOuvrirFiche.Visible = True
End Sub

Private Sub cmdValider_Click()
    Dim ref As String
    ref = Trim(Me.Controls("txtReference").Text)
    
    If Len(ref) = 0 Then
        MsgBox "Veuillez saisir une r�f�rence !", vbExclamation
        Exit Sub
    End If
    
    ' Valider et afficher info
    referenceValidee = ref
    
    Dim info As String
    Select Case Left(ref, 6)
        Case "VC2286"
            info = "FRIGO VITRINE VC2286" & vbCrLf & _
                   "Capacit�: 250L" & vbCrLf & _
                   "Temp�rature: +2�C � +8�C" & vbCrLf & _
                   "Prix neuf: 1,250�" & vbCrLf & vbCrLf & _
                   "R�f�rence valid�e - Vous pouvez ouvrir la fiche retour"
        Case "RB4458"
            info = "FRIGO RED BULL RB4458" & vbCrLf & _
                   "Capacit�: 180L" & vbCrLf & _
                   "Temp�rature: +1�C � +6�C" & vbCrLf & _
                   "Prix neuf: 1,580�" & vbCrLf & vbCrLf & _
                   "R�f�rence valid�e - Vous pouvez ouvrir la fiche retour"
        Case Else
            info = "FRIGO G�N�RIQUE" & vbCrLf & _
                   "R�f�rence: " & ref & vbCrLf & _
                   "Mod�le non identifi�" & vbCrLf & vbCrLf & _
                   "R�f�rence accept�e - Vous pouvez ouvrir la fiche retour"
    End Select
    
    Me.Controls("lblInfo").Caption = info
    Me.Controls("lblInfo").BackColor = RGB(200, 255, 200)
    
    ' Activer le bouton fiche
    Me.Controls("cmdOuvrirFiche").Enabled = True
    Me.Controls("cmdOuvrirFiche").BackColor = RGB(0, 150, 0)
End Sub

Private Sub cmdOuvrirFiche_Click()
    If Len(referenceValidee) = 0 Then
        MsgBox "Aucune r�f�rence valid�e !", vbExclamation
        Exit Sub
    End If
    
    ' Ouvrir la fiche retour
    Load frmFicheRetour
    frmFicheRetour.InitialiserAvecReference referenceValidee
    frmFicheRetour.Show vbModal
    
    ' Reset apr�s fermeture
    Me.Controls("txtReference").Text = ""
    Me.Controls("lblInfo").Caption = "Fiche trait�e. Saisissez une nouvelle r�f�rence."
    Me.Controls("cmdOuvrirFiche").Enabled = False
    Me.Controls("cmdOuvrirFiche").BackColor = RGB(150, 150, 150)
    referenceValidee = ""
End Sub

