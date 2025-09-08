Attribute VB_Name = "modGestionFichiers"
' === modGestionFichiers.bas - MODULE COMPLET ===
Option Explicit

' === FONCTIONS DE GESTION DES FICHIERS ===

' Fonction pour �crire dans l'historique des scans
Public Sub EcrireHistoriqueScan(reference As String, modele As String)
    Dim fichier As String
    Dim numeroFichier As Integer
    Dim valideBDD As String
    
    ' V�rifier avec la logique de requ�te
    If VerifierNumeroSerieBDD(reference) Then
        valideBDD = " [BDD: VALID�]"
    Else
        valideBDD = " [BDD: NON TROUV�]"
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

' Fonction pour cr�er une sauvegarde compl�te
Public Sub SauvegardeAutomatique()
    CreerSauvegardeComplete
End Sub

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
    
    ' Sauvegarder les infos syst�me
    Dim numeroFichier As Integer
    numeroFichier = FreeFile
    Open repertoireSauvegarde & "InfosSysteme.txt" For Output As #numeroFichier
    Print #numeroFichier, ObtenirInfosSysteme()
    Close #numeroFichier
    
    ' Sauvegarder la requ�te SQL utilis�e
    numeroFichier = FreeFile
    Open repertoireSauvegarde & "RequeteSQL.txt" For Output As #numeroFichier
    Print #numeroFichier, "=== REQU�TE SQL UTILIS�E DANS L'APPLICATION ===" & vbCrLf
    Print #numeroFichier, "SELECT DISTINCT art.art_code, art.art_desl, nse.nse_nums"
    Print #numeroFichier, "FROM ART_PAR as art"
    Print #numeroFichier, "INNER JOIN nse_dat as nse ON"
    Print #numeroFichier, "nse.act_code = art.act_code AND nse.art_code = art.art_code"
    Print #numeroFichier, "AND nse.act_code = 'RB'" & vbCrLf
    Print #numeroFichier, "=== DESCRIPTION ===" & vbCrLf
    Print #numeroFichier, "Cette requ�te r�cup�re :"
    Print #numeroFichier, "- art_code : Code de l'article"
    Print #numeroFichier, "- art_desl : D�signation de l'article"
    Print #numeroFichier, "- nse_nums : Num�ro de s�rie de l'�quipement"
    Print #numeroFichier, "Filtr� sur act_code = 'RB' pour Red Bull uniquement"
    Print #numeroFichier, "Et sur 92 codes articles autoris�s"
    Close #numeroFichier
    
    On Error GoTo 0
End Sub

' === FONCTIONS DE GESTION DES LOGS ===

' Fonction pour logger les erreurs syst�me
Public Sub LoggerErreur(source As String, description As String)
    On Error Resume Next
    
    Dim fichierLog As String
    Dim numeroFichier As Integer
    
    fichierLog = App.Path & "\Logs\Erreurs_" & Format(Now, "yyyymmdd") & ".txt"
    
    ' Cr�er le r�pertoire Logs s'il n'existe pas
    If Dir(App.Path & "\Logs", vbDirectory) = "" Then
        MkDir App.Path & "\Logs"
    End If
    
    numeroFichier = FreeFile
    Open fichierLog For Append As #numeroFichier
    Print #numeroFichier, Format(Now, "dd/mm/yyyy hh:nn:ss") & " - [" & source & "] " & description
    Close #numeroFichier
End Sub

' === FONCTIONS DE GESTION DES SESSIONS ===

' Fonction pour sauvegarder l'�tat de la session
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

' Fonction pour r�cup�rer l'�tat de la derni�re session
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
    RecupererEtatSession = "Erreur lors de la r�cup�ration de l'�tat de session"
End Function

' === FONCTIONS D'EXPORT ===

' Fonction pour exporter les donn�es vers CSV
Public Sub ExporterVersCSV(donnees As String, nomFichier As String)
    On Error GoTo ErrorHandler
    
    Dim cheminExport As String
    Dim numeroFichier As Integer
    
    ' Cr�er le r�pertoire d'export s'il n'existe pas
    If Dir(App.Path & "\Exports", vbDirectory) = "" Then
        MkDir App.Path & "\Exports"
    End If
    
    cheminExport = App.Path & "\Exports\" & nomFichier & "_" & Format(Now, "yyyymmdd_hhnnss") & ".csv"
    numeroFichier = FreeFile
    
    Open cheminExport For Output As #numeroFichier
    Print #numeroFichier, donnees
    Close #numeroFichier
    
    MsgBox "Donn�es export�es vers : " & cheminExport, vbInformation, "Export r�ussi"
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur lors de l'export : " & Err.description, vbCritical, "Erreur d'export"
End Sub
