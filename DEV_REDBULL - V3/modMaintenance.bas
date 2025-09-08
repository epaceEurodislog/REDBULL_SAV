Attribute VB_Name = "modMaintenance"
' === modMaintenance.bas - MODULE COMPLET ===
Option Explicit

' === FONCTIONS DE MAINTENANCE ===

' Fonction de maintenance rapide
Public Sub MaintenanceRapide()
    Dim message As String
    message = "Maintenance en cours..." & vbCrLf
    
    ' V�rifier l'int�grit� des fichiers
    If VerifierIntegriteFichiers() Then
        message = message & "? Fichiers : OK" & vbCrLf
    Else
        message = message & "? Fichiers : Probl�me d�tect�" & vbCrLf
        InitialiserStockPieces
        InitialiserStockReparable
        message = message & "? Fichiers : R�par�s" & vbCrLf
    End If
    
    ' Nettoyer les fichiers temporaires
    NettoyerFichiersTemporaires
    message = message & "? Nettoyage : OK" & vbCrLf
    
    ' Test connexion BDD
    If ConnecterBDD() Then
        message = message & "? BDD : Connexion OK" & vbCrLf
    Else
        message = message & "? BDD : Connexion �chou�e" & vbCrLf
    End If
    
    ' Cr�er sauvegarde
    CreerSauvegardeComplete
    message = message & "? Sauvegarde : OK" & vbCrLf
    
    message = message & vbCrLf & "Maintenance termin�e"
    MsgBox message, vbInformation, "Maintenance"
End Sub

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

' Fonction pour tester compl�tement le syst�me
Public Sub TesterSystemeComplet()
    Dim message As String
    
    message = "=== TEST SYST�ME SAV RED BULL ===" & vbCrLf & vbCrLf
    
    ' Test connexion BDD
    If VerifierConnexionBDD() Then
        message = message & "? Connexion BDD : OK" & vbCrLf
        message = message & "  Serveur: " & SERVER_NAME & vbCrLf
        message = message & "  Base: " & DATABASE_NAME & vbCrLf & vbCrLf
        
        ' Test de requ�te sp�cifique
        Dim rsTest As ADODB.Recordset
        Set rsTest = ObtenirArticlesRB()
        
        If Not rsTest Is Nothing Then
            Dim compteurTotal As Integer
            Dim compteurAvecSerie As Integer
            compteurTotal = 0
            compteurAvecSerie = 0
            
            Do While Not rsTest.EOF
                compteurTotal = compteurTotal + 1
                If Not IsNull(rsTest!nse_nums) And Len(rsTest!nse_nums) > 0 Then
                    compteurAvecSerie = compteurAvecSerie + 1
                End If
                rsTest.MoveNext
            Loop
            
            message = message & "? Requ�te articles RB : OK" & vbCrLf
            message = message & "  Articles totaux Red Bull: " & compteurTotal & vbCrLf
            message = message & "  Avec num�ro de s�rie: " & compteurAvecSerie & vbCrLf
            message = message & "  Sans num�ro de s�rie: " & (compteurTotal - compteurAvecSerie) & vbCrLf & vbCrLf
            
            rsTest.Close
            Set rsTest = Nothing
        Else
            message = message & "? Requ�te articles RB : �CHEC" & vbCrLf
        End If
    Else
        message = message & "? Connexion BDD : �CHEC" & vbCrLf
    End If
    
    ' Test fichiers
    If VerifierIntegriteFichiers() Then
        message = message & "? Fichiers locaux : OK" & vbCrLf
    Else
        message = message & "? Fichiers locaux : MANQUANTS" & vbCrLf
    End If
    
    message = message & vbCrLf & "Test termin� : " & ObtenirDateTimeFormatee()
    
    MsgBox message, vbInformation, "Test Syst�me SAV Red Bull"
End Sub

' Fonction pour tester la requ�te filtr�e 92 codes
Public Sub TesterRequeteFiltree92Codes()
    On Error GoTo ErrorHandler
    
    If Not VerifierConnexionBDD() Then
        MsgBox "Pas de connexion BDD pour tester la requ�te filtr�e", vbExclamation
        Exit Sub
    End If
    
    Dim message As String
    message = "=== TEST REQU�TE FILTR�E - 92 CODES ARTICLES ===" & vbCrLf & vbCrLf
    
    ' Tester la requ�te filtr�e
    Dim rsTest As ADODB.Recordset
    Set rsTest = ObtenirArticlesRB()
    
    If Not rsTest Is Nothing Then
        message = message & "? Requ�te filtr�e ex�cut�e avec succ�s !" & vbCrLf & vbCrLf
        
        ' Compter les r�sultats
        Dim compteur As Integer
        compteur = 0
        
        Do While Not rsTest.EOF
            compteur = compteur + 1
            rsTest.MoveNext
        Loop
        
        message = message & "=== STATISTIQUES FILTRAGE ===" & vbCrLf
        message = message & "� Codes articles dans la liste: 92" & vbCrLf
        message = message & "� Num�ros de s�rie trouv�s: " & compteur & vbCrLf
        message = message & "� Filtrage: INNER JOIN + DISTINCT" & vbCrLf
        message = message & "� Validation: stricte (non NULL, non vide)" & vbCrLf & vbCrLf
        
        If compteur > 0 Then
            message = message & "? " & compteur & " �quipements Red Bull autoris�s trouv�s"
        Else
            message = message & "? Aucun �quipement trouv� - v�rifier les donn�es"
        End If
        
        rsTest.Close
        Set rsTest = Nothing
    Else
        message = message & "? Erreur lors de l'ex�cution de la requ�te filtr�e"
    End If
    
    MsgBox message, vbInformation, "Test Requ�te Filtr�e - 92 Codes"
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur lors du test de la requ�te filtr�e : " & Err.description, vbCritical
End Sub

' === FONCTIONS DE DIAGNOSTIC AVANC� ===

' Fonction pour diagnostiquer la performance du syst�me
Public Sub DiagnosticPerformance()
    Dim debut As Date
    Dim fin As Date
    Dim duree As Double
    Dim rapport As String
    
    rapport = "=== DIAGNOSTIC PERFORMANCE ===" & vbCrLf & vbCrLf
    
    ' Test vitesse connexion BDD
    debut = Now
    Dim connexionOK As Boolean
    connexionOK = ConnecterBDD()
    fin = Now
    duree = (fin - debut) * 24 * 60 * 60 ' Convertir en secondes
    
    rapport = rapport & "Temps connexion BDD: " & Format(duree, "0.00") & " secondes" & vbCrLf
    
    If connexionOK Then
        ' Test vitesse requ�te
        debut = Now
        Dim rs As ADODB.Recordset
        Set rs = ObtenirArticlesRB()
        fin = Now
        duree = (fin - debut) * 24 * 60 * 60
        
        rapport = rapport & "Temps requ�te filtrage: " & Format(duree, "0.00") & " secondes" & vbCrLf
        
        If Not rs Is Nothing Then
            rs.Close
            Set rs = Nothing
        End If
    End If
    
    ' Test vitesse lecture fichiers
    debut = Now
    Dim historique As String
    historique = LireHistoriqueScan()
    fin = Now
    duree = (fin - debut) * 24 * 60 * 60
    
    rapport = rapport & "Temps lecture historique: " & Format(duree, "0.00") & " secondes" & vbCrLf
    rapport = rapport & vbCrLf & "=== RECOMMANDATIONS ===" & vbCrLf
    
    If duree > 1 Then
        rapport = rapport & "� Fichier historique volumineux - archivage recommand�" & vbCrLf
    End If
    
    MsgBox rapport, vbInformation, "Diagnostic Performance"
End Sub

' Fonction pour nettoyer les fichiers anciens
Public Sub NettoyageAvance()
    Dim reponse As Integer
    reponse = MsgBox("Nettoyer les fichiers de plus de 30 jours ?", vbYesNo + vbQuestion, "Nettoyage avanc�")
    
    If reponse = vbYes Then
        ' Nettoyer les sauvegardes anciennes
        NettoyerSauvegardesAnciennes 30
        
        ' Nettoyer les logs anciens
        NettoyerLogsAnciens 30
        
        ' Archiver l'historique
        ArchiverHistorique
        
        MsgBox "Nettoyage termin�", vbInformation
    End If
End Sub

' Fonction pour nettoyer les sauvegardes anciennes
Private Sub NettoyerSauvegardesAnciennes(joursMax As Integer)
    On Error Resume Next
    
    Dim chemin As String
    Dim fichier As String
    Dim cheminComplet As String
    
    chemin = App.Path & "\Sauvegardes\"
    
    If Dir(chemin, vbDirectory) <> "" Then
        fichier = Dir(chemin & "*.*")
        
        Do While fichier <> ""
            cheminComplet = chemin & fichier
            
            ' V�rifier si le fichier est plus ancien que joursMax
            If DateDiff("d", FileDateTime(cheminComplet), Now) > joursMax Then
                Kill cheminComplet
            End If
            
            fichier = Dir
        Loop
    End If
End Sub

' Fonction pour archiver l'historique
Private Sub ArchiverHistorique()
    On Error Resume Next
    
    Dim fichierHistorique As String
    Dim fichierArchive As String
    
    fichierHistorique = App.Path & FICHIER_HISTORIQUE
    fichierArchive = App.Path & "\Archives\Historique_" & Format(Now, "yyyymmdd") & ".txt"
    
    ' Cr�er le r�pertoire Archives s'il n'existe pas
    If Dir(App.Path & "\Archives", vbDirectory) = "" Then
        MkDir App.Path & "\Archives"
    End If
    
    ' Copier l'historique vers les archives
    If Dir(fichierHistorique) <> "" Then
        FileCopy fichierHistorique, fichierArchive
        
        ' Vider l'historique actuel (garder seulement l'en-t�te)
        Dim numeroFichier As Integer
        numeroFichier = FreeFile
        Open fichierHistorique For Output As #numeroFichier
        Print #numeroFichier, Format(Now, "dd/mm/yyyy hh:nn:ss") & " - SYSTEM - Historique archiv�"
        Close #numeroFichier
    End If
End Sub

' === FONCTIONS DE R�PARATION SYST�ME ===

' Fonction pour r�parer automatiquement les probl�mes courants
Public Sub ReparationAutomatique()
    Dim problemes() As String
    Dim solutions() As String
    Dim i As Integer
    Dim nbProblemes As Integer
    
    nbProblemes = 0
    ReDim problemes(10)
    ReDim solutions(10)
    
    ' V�rifier les r�pertoires manquants
    If Dir(App.Path & "\Fiches", vbDirectory) = "" Then
        MkDir App.Path & "\Fiches"
        problemes(nbProblemes) = "R�pertoire \Fiches manquant"
        solutions(nbProblemes) = "R�pertoire cr��"
        nbProblemes = nbProblemes + 1
    End If
    
    If Dir(App.Path & "\Sauvegardes", vbDirectory) = "" Then
        MkDir App.Path & "\Sauvegardes"
        problemes(nbProblemes) = "R�pertoire \Sauvegardes manquant"
        solutions(nbProblemes) = "R�pertoire cr��"
        nbProblemes = nbProblemes + 1
    End If
    
    ' V�rifier les fichiers critiques
    If Dir(App.Path & FICHIER_STOCK_PIECES) = "" Then
        InitialiserStockPieces
        problemes(nbProblemes) = "Fichier stock pi�ces manquant"
        solutions(nbProblemes) = "Fichier recr�� avec donn�es d'exemple"
        nbProblemes = nbProblemes + 1
    End If
    
    ' Afficher le rapport de r�paration
    Dim rapport As String
    rapport = "=== R�PARATION AUTOMATIQUE ===" & vbCrLf & vbCrLf
    
    If nbProblemes = 0 Then
        rapport = rapport & "? Aucun probl�me d�tect�" & vbCrLf
    Else
        rapport = rapport & "Probl�mes r�solus :" & vbCrLf
        For i = 0 To nbProblemes - 1
            rapport = rapport & "� " & problemes(i) & " ? " & solutions(i) & vbCrLf
        Next i
    End If
    
    MsgBox rapport, vbInformation, "R�paration automatique termin�e"
End Sub

 'Fonction pour optimiser la base de donn�es
Public Sub OptimiserBDD()
    On Error GoTo ErrorHandler
    
    If Not VerifierConnexionBDD() Then
        MsgBox "Pas de connexion BDD pour l'optimisation", vbExclamation
        Exit Sub
    End If
    
    Dim reponse As Integer
    reponse = MsgBox("Lancer l'optimisation de la base de donn�es ?", vbYesNo + vbQuestion, "Optimisation BDD")
    
    If reponse = vbYes Then
        MsgBox "Optimisation BDD simul�e", vbInformation, "Optimisation"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur lors de l'optimisation BDD : " & Err.description, vbCritical
End Sub

' Fonction pour nettoyer les logs anciens
Private Sub NettoyerLogsAnciens(joursMax As Integer)
    On Error Resume Next
    
    Dim chemin As String
    Dim fichier As String
    Dim cheminComplet As String
    
    chemin = App.Path & "\Logs\"
    
    If Dir(chemin, vbDirectory) <> "" Then
        fichier = Dir(chemin & "*.txt")
        
        Do While fichier <> ""
            cheminComplet = chemin & fichier
            
            If DateDiff("d", FileDateTime(cheminComplet), Now) > joursMax Then
                Kill cheminComplet
            End If
            
            fichier = Dir
        Loop
    End If
End Sub



