Attribute VB_Name = "modGestionFichiers"
' === modGestionFichiers.bas - MODULE COMPLET ===
Option Explicit

' === FONCTIONS DE GESTION DES FICHIERS ===

' Fonction pour écrire dans l'historique des scans
Public Sub EcrireHistoriqueScan(reference As String, modele As String)
    Dim fichier As String
    Dim numeroFichier As Integer
    Dim valideBDD As String
    
    ' Vérifier avec la logique de requête
    If VerifierNumeroSerieBDD(reference) Then
        valideBDD = " [BDD: VALIDÉ]"
    Else
        valideBDD = " [BDD: NON TROUVÉ]"
    End If
    
    fichier = App.Path & FICHIER_HISTORIQUE
    numeroFichier = FreeFile
    
    Open fichier For Append As #numeroFichier
    Print #numeroFichier, Format(Now, "dd/mm/yy hh:nn:ss") & " - " & reference & " - " & modele & valideBDD
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

' Fonction pour créer une sauvegarde complète
Public Sub SauvegardeAutomatique()
    CreerSauvegardeComplete
End Sub

' Fonction pour créer une sauvegarde complète
Public Sub CreerSauvegardeComplete()
    Dim dateStr As String
    Dim repertoireSauvegarde As String
    
    dateStr = Format(Now, "yyyymmdd_hhnnss")
    repertoireSauvegarde = App.Path & "\Sauvegardes\Sauvegarde_" & dateStr & "\"
    
    ' Créer le répertoire de sauvegarde
    If Dir(repertoireSauvegarde, vbDirectory) = "" Then
        MkDir repertoireSauvegarde
    End If
    
    ' Copier les fichiers importants
    On Error Resume Next
    FileCopy App.Path & FICHIER_HISTORIQUE, repertoireSauvegarde & "HistoriqueScans.txt"
    FileCopy App.Path & FICHIER_STOCK_PIECES, repertoireSauvegarde & "StockPieces.txt"
    FileCopy App.Path & FICHIER_STOCK_REPARABLE, repertoireSauvegarde & "StockReparable.txt"
    
    ' Sauvegarder les infos système
    Dim numeroFichier As Integer
    numeroFichier = FreeFile
    Open repertoireSauvegarde & "InfosSysteme.txt" For Output As #numeroFichier
    Print #numeroFichier, ObtenirInfosSysteme()
    Close #numeroFichier
    
    ' Sauvegarder la requête SQL utilisée
    numeroFichier = FreeFile
    Open repertoireSauvegarde & "RequeteSQL.txt" For Output As #numeroFichier
    Print #numeroFichier, "=== REQUÊTE SQL UTILISÉE DANS L'APPLICATION ===" & vbCrLf
    Print #numeroFichier, "SELECT DISTINCT art.art_code, art.art_desl, nse.nse_nums"
    Print #numeroFichier, "FROM ART_PAR as art"
    Print #numeroFichier, "INNER JOIN nse_dat as nse ON"
    Print #numeroFichier, "nse.act_code = art.act_code AND nse.art_code = art.art_code"
    Print #numeroFichier, "AND nse.act_code = 'RB'" & vbCrLf
    Print #numeroFichier, "=== DESCRIPTION ===" & vbCrLf
    Print #numeroFichier, "Cette requête récupère :"
    Print #numeroFichier, "- art_code : Code de l'article"
    Print #numeroFichier, "- art_desl : Désignation de l'article"
    Print #numeroFichier, "- nse_nums : Numéro de série de l'équipement"
    Print #numeroFichier, "Filtré sur act_code = 'RB' pour Red Bull uniquement"
    Print #numeroFichier, "Et sur 92 codes articles autorisés"
    Close #numeroFichier
    
    On Error GoTo 0
End Sub

' === FONCTIONS DE GESTION DES LOGS ===

' Fonction pour logger les erreurs système
Public Sub LoggerErreur(source As String, description As String)
    On Error Resume Next
    
    Dim fichierLog As String
    Dim numeroFichier As Integer
    
    fichierLog = App.Path & "\Logs\Erreurs_" & Format(Now, "yyyymmdd") & ".txt"
    
    ' Créer le répertoire Logs s'il n'existe pas
    If Dir(App.Path & "\Logs", vbDirectory) = "" Then
        MkDir App.Path & "\Logs"
    End If
    
    numeroFichier = FreeFile
    Open fichierLog For Append As #numeroFichier
    Print #numeroFichier, Format(Now, "dd/mm/yyyy hh:nn:ss") & " - [" & source & "] " & description
    Close #numeroFichier
End Sub

' === FONCTIONS DE GESTION DES SESSIONS ===

' Fonction pour sauvegarder l'état de la session
Public Sub SauvegarderEtatSession(referenceValidee As String, numeroSerieValide As String)
    On Error Resume Next
    
    Dim fichierEtat As String
    Dim numeroFichier As Integer
    
    fichierEtat = App.Path & "\Session_" & Format(Now, "yyyymmdd") & ".tmp"
    numeroFichier = FreeFile
    
    Open fichierEtat For Output As #numeroFichier
    Print #numeroFichier, "DERNIERE_REFERENCE=" & referenceValidee
    Print #numeroFichier, "DERNIER_NUMERO_SERIE=" & numeroSerieValide
    Print #numeroFichier, "TIMESTAMP=" & Format(Now, "dd/mm/yyyy hh:nn:ss")
    Print #numeroFichier, "STATUT_BDD=" & IIf(VerifierConnexionBDD(), "CONNECTE", "DECONNECTE")
    Close #numeroFichier
End Sub

' Fonction pour récupérer l'état de la dernière session
Public Function RecupererEtatSession() As String
    On Error GoTo ErrorHandler
    
    Dim fichierEtat As String
    Dim numeroFichier As Integer
    Dim ligne As String
    Dim etatSession As String
    
    fichierEtat = App.Path & "\Session_" & Format(Now, "yyyymmdd") & ".tmp"
    
    If Dir(fichierEtat) <> "" Then
        numeroFichier = FreeFile
        Open fichierEtat For Input As #numeroFichier
        
        Do While Not EOF(numeroFichier)
            Line Input #numeroFichier, ligne
            etatSession = etatSession & ligne & vbCrLf
        Loop
        
        Close #numeroFichier
    End If
    
    RecupererEtatSession = etatSession
    Exit Function
    
ErrorHandler:
    RecupererEtatSession = "Erreur lors de la récupération de l'état de session"
End Function

' === FONCTIONS D'EXPORT ===

' Fonction pour exporter les données vers CSV
Public Sub ExporterVersCSV(donnees As String, nomFichier As String)
    On Error GoTo ErrorHandler
    
    Dim cheminExport As String
    Dim numeroFichier As Integer
    
    ' Créer le répertoire d'export s'il n'existe pas
    If Dir(App.Path & "\Exports", vbDirectory) = "" Then
        MkDir App.Path & "\Exports"
    End If
    
    cheminExport = App.Path & "\Exports\" & nomFichier & "_" & Format(Now, "yyyymmdd_hhnnss") & ".csv"
    numeroFichier = FreeFile
    
    Open cheminExport For Output As #numeroFichier
    Print #numeroFichier, donnees
    Close #numeroFichier
    
    MsgBox "Données exportées vers : " & cheminExport, vbInformation, "Export réussi"
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur lors de l'export : " & Err.description, vbCritical, "Erreur d'export"
End Sub
