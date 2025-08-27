Attribute VB_Name = "Module1"
' === MODULE1.BAS - FONCTIONS GLOBALES SAV RED BULL ===

' D�clarations globales
Public Const VERSION_APP = "v2.1"
Public Const NOM_APP = "SAV Red Bull Scanner Pro"

' Chemins des fichiers de donn�es
Public Const FICHIER_HISTORIQUE = "\HistoriqueScans.txt"
Public Const FICHIER_STOCK_PIECES = "\StockPieces.txt"
Public Const FICHIER_STOCK_REPARABLE = "\StockReparable.txt"

' Structure pour les donn�es SAV
Public Type TypeSAV
    numeroEnlevement As String
    NumeroReception As String
    DateRetour As String
    ReferenceProduit As String
    MotifRetour As String
    CoherenceBoutique As Boolean
    DiagnosticPiece As Boolean
    DiagnosticTechnique As Boolean
    DiagnosticRayures As Boolean
    DateCreation As Date
    Statut As String
End Type

' Structure pour les pi�ces
Public Type TypePiece
    Code As String
    Nom As String
    Quantite As Integer
    Etat As String
    Origine As String
    DateAjout As Date
    Prix As Double
End Type

' === FONCTIONS DE GESTION DES FICHIERS ===

' Fonction pour �crire dans l'historique des scans
Public Sub EcrireHistoriqueScan(reference As String, modele As String)
    Dim fichier As String
    Dim numeroFichier As Integer
    
    fichier = App.Path & FICHIER_HISTORIQUE
    numeroFichier = FreeFile
    
    Open fichier For Append As #numeroFichier
    Print #numeroFichier, Format(Now, "dd/mm/yy hh:nn:ss") & " - " & reference & " - " & modele
    Close #numeroFichier
End Sub

' Fonction pour lire l'historique des scans
Public Function LireHistoriqueScan() As String
    Dim fichier As String
    Dim numeroFichier As Integer
    Dim ligne As String
    Dim historique As String
    
    fichier = App.Path & FICHIER_HISTORIQUE
    
    If Dir(fichier) = "" Then
        LireHistoriqueScan = "Aucun historique disponible"
        Exit Function
    End If
    
    numeroFichier = FreeFile
    Open fichier For Input As #numeroFichier
    
    Do While Not EOF(numeroFichier)
        Line Input #numeroFichier, ligne
        historique = historique & ligne & vbCrLf
    Loop
    
    Close #numeroFichier
    LireHistoriqueScan = historique
End Function

' Fonction pour effacer l'historique
Public Sub EffacerHistoriqueScan()
    Dim fichier As String
    fichier = App.Path & FICHIER_HISTORIQUE
    
    If Dir(fichier) <> "" Then
        Kill fichier
    End If
End Sub

' === FONCTIONS DE VALIDATION ===

' Fonction pour valider un code-barres
Public Function ValiderCodeBarre(codeBarre As String) As Boolean
    ' Supprime les espaces
    codeBarre = Trim(codeBarre)
    
    ' V�rifications de base
    If Len(codeBarre) < 6 Then
        ValiderCodeBarre = False
        Exit Function
    End If
    
    ' V�rifie le format basique (lettres + chiffres + tirets)
    Dim i As Integer
    For i = 1 To Len(codeBarre)
        Dim char As String
        char = UCase(Mid(codeBarre, i, 1))
        If Not ((char >= "A" And char <= "Z") Or (char >= "0" And char <= "9") Or char = "-") Then
            ValiderCodeBarre = False
            Exit Function
        End If
    Next i
    
    ValiderCodeBarre = True
End Function

' Fonction pour extraire le mod�le du code-barres
Public Function ExtraireModele(codeBarre As String) As String
    Dim prefixe As String
    prefixe = Left(codeBarre, 6)
    
    Select Case prefixe
        Case "VC2286"
            ExtraireModele = "Vitrine VC2286"
        Case "RB4458"
            ExtraireModele = "Red Bull RB4458"
        Case "CF3401"
            ExtraireModele = "Cong�lateur CF3401"
        Case "RB2024"
            ExtraireModele = "Red Bull Premium 2024"
        Case Else
            ExtraireModele = "Mod�le non identifi�"
    End Select
End Function

' === FONCTIONS DE GESTION DU STOCK ===

' Fonction pour initialiser le fichier stock pi�ces s'il n'existe pas
Public Sub InitialiserStockPieces()
    Dim fichier As String
    Dim numeroFichier As Integer
    
    fichier = App.Path & FICHIER_STOCK_PIECES
    
    If Dir(fichier) = "" Then
        numeroFichier = FreeFile
        Open fichier For Output As #numeroFichier
        Print #numeroFichier, "CODE|PIECE|QUANTITE|ETAT|ORIGINE|DATE|PRIX"
        
        ' Ajouter quelques pi�ces d'exemple
        Print #numeroFichier, "COMP|Compresseur Standard|2|Bon|DEMO001|" & Format(Now, "dd/mm/yyyy hh:nn:ss") & "|450.00"
        Print #numeroFichier, "LED|Eclairage LED|5|Excellent|DEMO002|" & Format(Now, "dd/mm/yyyy hh:nn:ss") & "|35.00"
        Print #numeroFichier, "VITRE|Vitre principale|1|Excellent|DEMO003|" & Format(Now, "dd/mm/yyyy hh:nn:ss") & "|120.00"
        Print #numeroFichier, "THERMO|Thermostat digital|3|Bon|DEMO004|" & Format(Now, "dd/mm/yyyy hh:nn:ss") & "|85.00"
        Print #numeroFichier, "JOINT|Joints de porte|8|Moyen|DEMO005|" & Format(Now, "dd/mm/yyyy hh:nn:ss") & "|25.00"
        
        Close #numeroFichier
    End If
End Sub

' Fonction pour initialiser le fichier stock r�parable
Public Sub InitialiserStockReparable()
    Dim fichier As String
    
    fichier = App.Path & FICHIER_STOCK_REPARABLE
    
    If Dir(fichier) = "" Then
        ' Le fichier sera cr�� lors de la premi�re fiche retour
    End If
End Sub

' === FONCTIONS UTILITAIRES ===

' Fonction pour cr�er les r�pertoires n�cessaires
Public Sub CreerRepertoires()
    Dim repertoires() As String
    Dim i As Integer
    
    ' Liste des r�pertoires � cr�er
    ReDim repertoires(4)
    repertoires(0) = App.Path & "\Fiches"
    repertoires(1) = App.Path & "\Recuperations"
    repertoires(2) = App.Path & "\Affectations"
    repertoires(3) = App.Path & "\Sauvegardes"
    repertoires(4) = App.Path & "\Exports"
    
    ' Cr�er chaque r�pertoire s'il n'existe pas
    For i = 0 To UBound(repertoires)
        If Dir(repertoires(i), vbDirectory) = "" Then
            MkDir repertoires(i)
        End If
    Next i
End Sub

' Fonction pour obtenir la date/heure format�e
Public Function ObtenirDateTimeFormatee() As String
    ObtenirDateTimeFormatee = Format(Now, "dd/mm/yyyy hh:nn:ss")
End Function

' Fonction pour g�n�rer un nom de fichier unique
Public Function GenererNomFichierUnique(prefixe As String, extension As String) As String
    Dim timestamp As String
    timestamp = Format(Now, "yyyymmdd_hhnnss")
    GenererNomFichierUnique = prefixe & "_" & timestamp & "." & extension
End Function

' === FONCTIONS DE D�MARRAGE ===

' Fonction d'initialisation appel�e au d�marrage
Public Sub InitialiserApplication()
    ' Cr�er les r�pertoires n�cessaires
    CreerRepertoires
    
    ' Initialiser les fichiers de stock
    InitialiserStockPieces
    InitialiserStockReparable
    
    ' Nettoyer les fichiers temporaires anciens
    NettoyerFichiersTemporaires
End Sub

' Fonction pour nettoyer les fichiers temporaires
Public Sub NettoyerFichiersTemporaires()
    On Error Resume Next
    
    Dim fichier As String
    Dim chemin As String
    
    ' Nettoyer les fichiers temporaires de plus de 7 jours
    chemin = App.Path & "\Temp\"
    
    If Dir(chemin, vbDirectory) <> "" Then
        fichier = Dir(chemin & "*.*")
        Do While fichier <> ""
            Dim cheminComplet As String
            cheminComplet = chemin & fichier
            
            ' Supprimer si plus vieux que 7 jours
            If DateDiff("d", FileDateTime(cheminComplet), Now) > 7 Then
                Kill cheminComplet
            End If
            
            fichier = Dir
        Loop
    End If
    
    On Error GoTo 0
End Sub

' === FONCTIONS DE DIAGNOSTIC ===

' Fonction pour obtenir des informations syst�me
Public Function ObtenirInfosSysteme() As String
    Dim infos As String
    
    infos = "=== INFORMATIONS SYST�ME ===" & vbCrLf
    infos = infos & "Application: " & NOM_APP & " " & VERSION_APP & vbCrLf
    infos = infos & "Chemin: " & App.Path & vbCrLf
    infos = infos & "Date syst�me: " & Format(Now, "dd/mm/yyyy hh:nn:ss") & vbCrLf
    infos = infos & "Utilisateur: " & Environ("USERNAME") & vbCrLf
    infos = infos & "Ordinateur: " & Environ("COMPUTERNAME") & vbCrLf
    
    ObtenirInfosSysteme = infos
End Function

' Fonction pour v�rifier l'int�grit� des fichiers
Public Function VerifierIntegriteFichiers() As Boolean
    Dim fichiers() As String
    Dim i As Integer
    Dim tousExistent As Boolean
    
    tousExistent = True
    
    ' Liste des fichiers critiques
    ReDim fichiers(1)
    fichiers(0) = App.Path & FICHIER_STOCK_PIECES
    fichiers(1) = App.Path & FICHIER_STOCK_REPARABLE
    
    For i = 0 To UBound(fichiers)
        If Dir(fichiers(i)) = "" Then
            tousExistent = False
        End If
    Next i
    
    VerifierIntegriteFichiers = tousExistent
End Function

' === FONCTIONS DE SAUVEGARDE ===

' Fonction pour cr�er une sauvegarde compl�te
Public Sub CreerSauvegardeComplete()
    Dim dateStr As String
    Dim repertoireSauvegarde As String
    
    dateStr = Format(Now, "yyyymmdd_hhnnss")
    repertoireSauvegarde = App.Path & "\Sauvegardes\Sauvegarde_" & dateStr & "\"
    
    ' Cr�er le r�pertoire de sauvegarde
    If Dir(repertoireSauvegarde, vbDirectory) = "" Then
        MkDir repertoireSauvegarde
    End If
    
    ' Copier les fichiers importants
    On Error Resume Next
    FileCopy App.Path & FICHIER_HISTORIQUE, repertoireSauvegarde & "HistoriqueScans.txt"
    FileCopy App.Path & FICHIER_STOCK_PIECES, repertoireSauvegarde & "StockPieces.txt"
    FileCopy App.Path & FICHIER_STOCK_REPARABLE, repertoireSauvegarde & "StockReparable.txt"
    On Error GoTo 0
End Sub

' === ANCIENNES FONCTIONS (COMPATIBILIT�) ===

Public Function ValiderDonnees(donnees As TypeSAV) As Boolean
    ' Conserv� pour compatibilit�
    ValiderDonnees = True
End Function

Public Function GenererNumeroSerie() As String
    GenererNumeroSerie = "SAV" & Format(Now, "yyyymmddhhnnss")
End Function

Public Function FormaterDateFrancaise(laDate As Date) As String
    FormaterDateFrancaise = Format(laDate, "dd/mm/yyyy")
End Function

Public Function CreerNomFichier(numeroEnlevement As String) As String
    CreerNomFichier = App.Path & "\Sauvegardes\SAV_" & numeroEnlevement & "_" & Format(Now, "yyyymmdd") & ".txt"
End Function

Public Sub CreerRepertoireSauvegarde()
    If Dir(App.Path & "\Sauvegardes", vbDirectory) = "" Then
        MkDir App.Path & "\Sauvegardes"
    End If
End Sub

Public Sub SauvegardeAutomatique()
    CreerSauvegardeComplete
End Sub
