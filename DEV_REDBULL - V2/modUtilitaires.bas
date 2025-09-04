Attribute VB_Name = "modUtilitaires"
' === modUtilitaires.bas - MODULE COMPLET ===
Option Explicit

' === FONCTIONS D'INITIALISATION ===

' Fonction d'initialisation appel�e au d�marrage
Public Sub InitialiserApplication()
    ' Cr�er les r�pertoires n�cessaires
    CreerRepertoires
    
    ' Initialiser les fichiers de stock
    InitialiserStockPieces
    InitialiserStockReparable
    
    ' Nettoyer les fichiers temporaires anciens
    NettoyerFichiersTemporaires
    
    ' �tablir la connexion BDD (SANS synchronisation automatique pour �viter la lenteur)
    If ConnecterBDD() Then
        Debug.Print "Application initialis�e avec BDD : " & ObtenirDateTimeFormatee()
    Else
        Debug.Print "Application d�marr�e sans connexion BDD - mode d�grad�"
    End If
End Sub

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

' Fonction pour initialiser le fichier stock pi�ces
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

' === FONCTIONS DE DATE ET HEURE ===

' Fonction pour obtenir la date/heure format�e
Public Function ObtenirDateTimeFormatee() As String
    ObtenirDateTimeFormatee = Format(Now, "dd/mm/yyyy hh:nn:ss")
End Function

' Fonction pour formater une date en fran�ais
Public Function FormaterDateFrancaise(laDate As Date) As String
    FormaterDateFrancaise = Format(laDate, "dd/mm/yyyy")
End Function

' === FONCTIONS DE G�N�RATION ===

' Fonction pour g�n�rer un num�ro de s�rie SAV
Public Function GenererNumeroSerie() As String
    GenererNumeroSerie = "SAV" & Format(Now, "yyyymmddhhnnss")
End Function

' Fonction pour cr�er un nom de fichier
Public Function CreerNomFichier(numeroEnlevement As String) As String
    CreerNomFichier = App.Path & "\Sauvegardes\SAV_" & numeroEnlevement & "_" & Format(Now, "yyyymmdd") & ".txt"
End Function

' === FONCTIONS D'INFORMATION SYST�ME ===

' Fonction pour obtenir des informations syst�me
Public Function ObtenirInfosSysteme() As String
    Dim infos As String
    Dim statutBDD As String
    
    If VerifierConnexionBDD() Then
        statutBDD = "CONNECT�E"
    Else
        statutBDD = "D�CONNECT�E"
    End If
    
    infos = "=== INFORMATIONS SYST�ME ===" & vbCrLf
    infos = infos & "Application: " & NOM_APP & " " & VERSION_APP & vbCrLf
    infos = infos & "Chemin: " & App.Path & vbCrLf
    infos = infos & "Date syst�me: " & ObtenirDateTimeFormatee() & vbCrLf
    infos = infos & "Utilisateur: " & Environ("USERNAME") & vbCrLf
    infos = infos & "Ordinateur: " & Environ("COMPUTERNAME") & vbCrLf
    infos = infos & "Base de donn�es: " & statutBDD & vbCrLf
    infos = infos & "Serveur BDD: " & SERVER_NAME & vbCrLf
    infos = infos & "Base: " & DATABASE_NAME & vbCrLf
    
    ObtenirInfosSysteme = infos
End Function

' === FONCTIONS DE FORMATAGE ===

' Fonction pour nettoyer une cha�ne
Public Function NettoyerChaine(chaine As String) As String
    chaine = Trim(UCase(chaine))
    chaine = Replace(chaine, vbCr, "")
    chaine = Replace(chaine, vbLf, "")
    chaine = Replace(chaine, Chr(0), "")
    NettoyerChaine = chaine
End Function

' Fonction pour valider un email (basique)
Public Function ValiderEmail(email As String) As Boolean
    If InStr(email, "@") > 0 And InStr(email, ".") > 0 Then
        ValiderEmail = True
    Else
        ValiderEmail = False
    End If
End Function

' === FONCTIONS DE CONVERSION ===

' Fonction pour convertir un bool�en en texte
Public Function BooleanVersTexte(valeur As Boolean) As String
    If valeur Then
        BooleanVersTexte = "OUI"
    Else
        BooleanVersTexte = "NON"
    End If
End Function

' Fonction pour formater un prix
Public Function FormaterPrix(prix As Double) As String
    FormaterPrix = Format(prix, "0.00") & "�"
End Function
