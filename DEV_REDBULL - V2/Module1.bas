Attribute VB_Name = "Module1"


' Module de fonctions communes pour le système SAV Red Bull

' Déclarations globales
Public Const VERSION_APP = "v2.1"
Public Const NOM_APP = "SAV Red Bull Scanner Pro"

' Structure pour les données SAV
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

' Fonction pour valider les données d'entrée
Public Function ValiderDonnees(donnees As TypeSAV) As Boolean
    Dim erreurs As String
    
    ' Vérification des champs obligatoires
    If Len(Trim(donnees.numeroEnlevement)) = 0 Then
        erreurs = erreurs & "- Numéro d'enlèvement requis" & vbCrLf
    End If
    
    If Len(Trim(donnees.NumeroReception)) = 0 Then
        erreurs = erreurs & "- Numéro de réception requis" & vbCrLf
    End If
    
    If Not IsDate(donnees.DateRetour) Then
        erreurs = erreurs & "- Date invalide" & vbCrLf
    End If
    
    If Len(Trim(donnees.ReferenceProduit)) = 0 Then
        erreurs = erreurs & "- Référence produit requise" & vbCrLf
    End If
    
    ' Affichage des erreurs s'il y en a
    If Len(erreurs) > 0 Then
        MsgBox "Erreurs de validation :" & vbCrLf & vbCrLf & erreurs, vbExclamation, "Validation"
        ValiderDonnees = False
    Else
        ValiderDonnees = True
    End If
End Function

' Fonction pour générer un numéro de série unique
Public Function GenererNumeroSerie() As String
    Dim numero As String
    numero = Format(Now, "yyyymmddhhnnss")
    GenererNumeroSerie = "SAV" & numero
End Function

' Fonction pour formater la date au format français
Public Function FormaterDateFrancaise(laDate As Date) As String
    FormaterDateFrancaise = Format(laDate, "dd/mm/yyyy")
End Function

' Fonction pour créer le nom de fichier de sauvegarde
Public Function CreerNomFichier(numeroEnlevement As String) As String
    Dim nomFichier As String
    Dim dateStr As String
    
    dateStr = Format(Now, "yyyymmdd")
    nomFichier = "SAV_" & numeroEnlevement & "_" & dateStr & ".txt"
    
    CreerNomFichier = App.Path & "\Sauvegardes\" & nomFichier
End Function

' Fonction pour créer le répertoire de sauvegarde s'il n'existe pas
Public Sub CreerRepertoireSauvegarde()
    Dim chemin As String
    chemin = App.Path & "\Sauvegardes"
    
    ' Vérifier si le répertoire existe
    If Dir(chemin, vbDirectory) = "" Then
        ' Créer le répertoire
        MkDir chemin
    End If
End Sub

' Fonction pour exporter les données au format CSV
Public Function ExporterCSV(donnees As TypeSAV, nomFichier As String) As Boolean
    On Error GoTo GestionErreur
    
    Dim numeroFichier As Integer
    Dim ligne As String
    
    numeroFichier = FreeFile
    
    ' Créer l'en-tête CSV s'il s'agit d'un nouveau fichier
    If Dir(nomFichier) = "" Then
        Open nomFichier For Output As #numeroFichier
        Print #numeroFichier, "NumeroEnlevement;NumeroReception;DateRetour;ReferenceProduit;MotifRetour;CoherenceBoutique;DiagnosticPiece;DiagnosticTechnique;DiagnosticRayures;DateCreation;Statut"
    Else
        Open nomFichier For Append As #numeroFichier
    End If
    
    ' Créer la ligne de données
    ligne = donnees.numeroEnlevement & ";" & _
            donnees.NumeroReception & ";" & _
            donnees.DateRetour & ";" & _
            donnees.ReferenceProduit & ";" & _
            donnees.MotifRetour & ";" & _
            IIf(donnees.CoherenceBoutique, "OUI", "NON") & ";" & _
            IIf(donnees.DiagnosticPiece, "OUI", "NON") & ";" & _
            IIf(donnees.DiagnosticTechnique, "OUI", "NON") & ";" & _
            IIf(donnees.DiagnosticRayures, "OUI", "NON") & ";" & _
            Format(donnees.DateCreation, "dd/mm/yyyy hh:nn:ss") & ";" & _
            donnees.Statut
    
    Print #numeroFichier, ligne
    Close #numeroFichier
    
    ExporterCSV = True
    Exit Function
    
GestionErreur:
    If numeroFichier > 0 Then Close #numeroFichier
    MsgBox "Erreur lors de l'export CSV : " & Err.Description, vbCritical, "Erreur"
    ExporterCSV = False
End Function

' Fonction pour charger les données depuis un fichier
Public Function ChargerDonnees(nomFichier As String) As TypeSAV
    On Error GoTo GestionErreur
    
    Dim donnees As TypeSAV
    Dim numeroFichier As Integer
    Dim ligne As String
    
    numeroFichier = FreeFile
    Open nomFichier For Input As #numeroFichier
    
    ' Lire le fichier ligne par ligne
    Do While Not EOF(numeroFichier)
        Line Input #numeroFichier, ligne
        ' Traiter chaque ligne selon le format
        If InStr(ligne, "N° Enlèvement:") > 0 Then
            donnees.numeroEnlevement = Trim(Mid(ligne, InStr(ligne, ":") + 1))
        ElseIf InStr(ligne, "N° Réception:") > 0 Then
            donnees.NumeroReception = Trim(Mid(ligne, InStr(ligne, ":") + 1))
        ElseIf InStr(ligne, "Date:") > 0 And InStr(ligne, "Date de création") = 0 Then
            donnees.DateRetour = Trim(Mid(ligne, InStr(ligne, ":") + 1))
        ElseIf InStr(ligne, "Référence produit:") > 0 Then
            donnees.ReferenceProduit = Trim(Mid(ligne, InStr(ligne, ":") + 1))
        End If
    Loop
    
    Close #numeroFichier
    ChargerDonnees = donnees
    Exit Function
    
GestionErreur:
    If numeroFichier > 0 Then Close #numeroFichier
    MsgBox "Erreur lors du chargement : " & Err.Description, vbCritical, "Erreur"
End Function

' Fonction pour nettoyer les fichiers temporaires
Public Sub NettoyerFichiersTemp()
    Dim fichier As String
    Dim chemin As String
    
    chemin = App.Path & "\Temp\"
    
    If Dir(chemin, vbDirectory) <> "" Then
        fichier = Dir(chemin & "*.tmp")
        Do While fichier <> ""
            Kill chemin & fichier
            fichier = Dir
        Loop
    End If
End Sub

' Fonction pour créer une sauvegarde automatique
Public Sub SauvegardeAutomatique()
    Dim cheminSauvegarde As String
    Dim dateStr As String
    
    dateStr = Format(Now, "yyyymmdd_hhnnss")
    cheminSauvegarde = App.Path & "\Sauvegardes\Auto_" & dateStr & ".bak"
    
    ' Cette fonction peut être étendue pour sauvegarder
    ' automatiquement les données importantes
End Sub

