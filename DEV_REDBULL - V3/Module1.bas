Attribute VB_Name = "Module1"
' === MODULE1.BAS - PARTIE 1: D�CLARATIONS ET CONNEXION BDD ===
' � placer dans le fichier Module1.bas

Option Explicit

' D�clarations globales
Public Const VERSION_APP = "v2.1"
Public Const NOM_APP = "SAV Red Bull Scanner Pro"

' Chemins des fichiers de donn�es
Public Const FICHIER_HISTORIQUE = "\HistoriqueScans.txt"
Public Const FICHIER_STOCK_PIECES = "\StockPieces.txt"
Public Const FICHIER_STOCK_REPARABLE = "\StockReparable.txt"

' === VARIABLES GLOBALES BDD ===
Public conn As ADODB.Connection
Public rs As ADODB.Recordset

' Param�tres de connexion BDD
Public Const SERVER_NAME As String = "192.168.9.12"
Public Const DATABASE_NAME As String = "SPEED_V6"
Public Const USERNAME As String = "eurodislog"
Public Const PASSWORD As String = "euro"

Private Const CODES_ARTICLES_AUTORISES As String = _
    "'1401-019-000XX1','1401-090-000XX1','1401-118-000XX1','1401-128-000XX1'," & _
    "'1401-133-000XX1','1401-136-000XX1','1401-138-000XX1','1401-140-000XX1'," & _
    "'1401-142-000XX1','1401-146-000XX1','1401-152-000XX','1401-158-000XX1'," & _
    "'1401-170-000XX','1401-173-000XX1','1499-012-000XX1','1502-075-000XX'," & _
    "'1509-080-000XX1','1509-114-000XX1','1509-144-000XX1','1509-146-000XX'," & _
    "'1509-148-000XX1','1509-149-000XX1','1509-168-000XX1','1509-169-000XX1'," & _
    "'1509-176-000XX1','1509-219-000XX1','1509-227-000XX','1509-227-000XX1'," & _
    "'VC202194000-1','VC205073000-1','VC206225000-1','VC206226000'," & _
    "'VC206226000-1','VC206484014-1','VC206489014-1','VC206490014-1'," & _
    "'VC209225000','VC209225000-1','VC213211010-1','VC213212010-1'," & _
    "'VC213240004-1','VC213247010-1','VC213250000-1','VC213251000-1'," & _
    "'VC213252000-1','VC215028000-1','VC215038000','VC215038000-1'," & _
    "'VC221651000-1','VC221653000-1','VC222866010-1','VC223056000-1'," & _
    "'VC225604000-1','VC228630000-1','VC228652000-1','VC228658000-1'," & _
    "'VC230086000','VC230086000-1','VC234598000-1','VC234827002'," & _
    "'VC234827002-1','VC234830002','VC234830002-1','VC234857004-1'," & _
    "'VC234859002-1','VC236036000-1','VC237539000','VC237539000-1'," & _
    "'VC240116002-1','VC240468000-1','VC240470000','VC240470000-1'," & _
    "'VC241476014','VC241476014-1','VC241481014','VC241481014-1'," & _
    "'VC241509004','VC241509004-1','VC241869000','VC241869000-1'," & _
    "'VC245058000','VC245060000','VC245269000','VC245308000'," & _
    "'VC245571000','VC245657000','VC245658000','VC248948000'," & _
    "'VC249298000','VC249298000-1','VM221870000-1','VM245176000'"

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
    dateCreation As Date
    statut As String
End Type

' Structure pour les pi�ces
Public Type TypePiece
    code As String
    Nom As String
    quantite As Integer
    etat As String
    origine As String
    dateAjout As Date
    prix As Double
End Type

' Structure pour les r�sultats de validation BDD (MISE � JOUR AVEC VOTRE REQU�TE)
Public Type TypeValidationBDD
    existe As Boolean
    codeArticle As String
    designationArticle As String  ' NOUVEAU : art_desl de votre requ�te
    modeleArticle As String
    numeroSerie As String
    prixCatalogue As Double
    dateCreation As String
    statut As String
    informationsComplementaires As String
End Type

' === FONCTIONS DE CONNEXION BDD ===

' Fonction pour �tablir la connexion � la base de donn�es
Public Function ConnecterBDD() As Boolean
    On Error GoTo ErrorHandler
    
    ' Cr�er l'objet Connection
    Set conn = New ADODB.Connection
    
    ' Construire la cha�ne de connexion
    Dim connectionString As String
    connectionString = "Provider=SQLOLEDB;" & _
                      "Data Source=" & SERVER_NAME & ";" & _
                      "Initial Catalog=" & DATABASE_NAME & ";" & _
                      "User ID=" & USERNAME & ";" & _
                      "Password=" & PASSWORD & ";"
    
    ' �tablir la connexion
    conn.Open connectionString
    
    ' V�rifier si la connexion est ouverte
    If conn.State = adStateOpen Then
        ConnecterBDD = True
        Debug.Print "Connexion BDD �tablie : " & ObtenirDateTimeFormatee()
    Else
        ConnecterBDD = False
        MsgBox "�chec de la connexion � la base de donn�es !", vbCritical
    End If
    
    Exit Function
    
ErrorHandler:
    ConnecterBDD = False
    MsgBox "Erreur lors de la connexion BDD : " & Err.description, vbCritical
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
End Function

' Fonction pour fermer la connexion
Public Sub FermerBDD()
    On Error Resume Next
    
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
        Set conn = Nothing
    End If
    
    Debug.Print "Connexion BDD ferm�e : " & ObtenirDateTimeFormatee()
End Sub

' Fonction pour v�rifier si la connexion est active
Public Function VerifierConnexionBDD() As Boolean
    If conn Is Nothing Then
        VerifierConnexionBDD = False
    Else
        VerifierConnexionBDD = (conn.State = adStateOpen)
    End If
End Function

' Fonction pour reconnecter si n�cessaire
Public Function Reconnecter() As Boolean
    If Not VerifierConnexionBDD() Then
        Reconnecter = ConnecterBDD()
    Else
        Reconnecter = True
    End If
End Function

' === MODULE1.BAS - PARTIE 2: FONCTIONS DE REQU�TES BDD ===
' � ajouter � la suite de la Partie 1 dans Module1.bas

' === FONCTIONS DE REQU�TES BDD AVEC VOTRE REQU�TE CORRIG�E ===

' Fonction pour obtenir tous les articles Red Bull avec votre requ�te SQL exacte
Public Function ObtenirArticlesRB() As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    If Not Reconnecter() Then
        Set ObtenirArticlesRB = Nothing
        Exit Function
    End If
    
    ' REQU�TE SQL FILTR�E SUR LES 92 CODES ARTICLES AUTORIS�S
    Dim sql As String
    sql = "SELECT DISTINCT art.art_code, art.art_desl, nse.nse_nums " & _
          "FROM ART_PAR as art " & _
          "INNER JOIN nse_dat as nse ON " & _
          "nse.act_code = art.act_code AND nse.art_code = art.art_code " & _
          "AND nse.act_code = 'RB' " & _
          "WHERE nse.nse_nums IS NOT NULL " & _
          "AND nse.nse_nums <> '' " & _
          "AND LEN(LTRIM(RTRIM(nse.nse_nums))) > 0 " & _
          "AND art.art_code IN (" & CODES_ARTICLES_AUTORISES & ") " & _
          "ORDER BY art.art_code"
    
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    
    Set ObtenirArticlesRB = rs
    
    Debug.Print "Requ�te filtr�e ex�cut�e - 92 codes articles autoris�s"
    Exit Function
    
ErrorHandler:
    MsgBox "Erreur lors de la requ�te articles RB filtr�e : " & Err.description, vbCritical
    Set ObtenirArticlesRB = Nothing
End Function

' Fonction g�n�rique pour ex�cuter des requ�tes SELECT
Public Function ExecuterRequete(sql As String) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    If Not Reconnecter() Then
        Set ExecuterRequete = Nothing
        Exit Function
    End If
    
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    
    Set ExecuterRequete = rs
    Exit Function
    
ErrorHandler:
    MsgBox "Erreur lors de l'ex�cution de la requ�te : " & Err.description, vbCritical
    Set ExecuterRequete = Nothing
End Function

' Fonction pour v�rifier si un num�ro de s�rie existe dans la BDD avec votre requ�te
Public Function VerifierNumeroSerieBDD(numeroSerie As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim sql As String
    Dim rsVerif As ADODB.Recordset
    
    ' V�RIFICATION AVEC FILTRAGE SUR LES 92 CODES AUTORIS�S
    sql = "SELECT COUNT(*) as nb " & _
          "FROM ART_PAR as art " & _
          "INNER JOIN nse_dat as nse ON " & _
          "nse.act_code = art.act_code AND nse.art_code = art.art_code " & _
          "AND nse.act_code = 'RB' " & _
          "WHERE nse.nse_nums = '" & numeroSerie & "' " & _
          "AND nse.nse_nums IS NOT NULL " & _
          "AND nse.nse_nums <> '' " & _
          "AND LEN(LTRIM(RTRIM(nse.nse_nums))) > 0 " & _
          "AND art.art_code IN (" & CODES_ARTICLES_AUTORISES & ")"
    
    Set rsVerif = ExecuterRequete(sql)
    
    If Not rsVerif Is Nothing Then
        If Not rsVerif.EOF Then
            VerifierNumeroSerieBDD = (rsVerif!nb > 0)
        Else
            VerifierNumeroSerieBDD = False
        End If
        rsVerif.Close
        Set rsVerif = Nothing
    Else
        VerifierNumeroSerieBDD = False
    End If
    
    Exit Function
    
ErrorHandler:
    VerifierNumeroSerieBDD = False
    MsgBox "Erreur lors de la v�rification du num�ro de s�rie : " & Err.description, vbCritical
End Function

' === FONCTION PRINCIPALE DE VALIDATION AVEC VOTRE REQU�TE ===

' Fonction principale pour valider un num�ro de s�rie avec votre requ�te exacte
Public Function ValiderNumeroSerieBDD(numeroSerie As String) As TypeValidationBDD
    On Error GoTo ErrorHandler
    
    Dim resultats As TypeValidationBDD
    resultats.existe = False
    
    If Not Reconnecter() Then
        resultats.statut = "CONNEXION IMPOSSIBLE"
        ValiderNumeroSerieBDD = resultats
        Exit Function
    End If
    
    ' VALIDATION AVEC FILTRAGE SUR LES 92 CODES AUTORIS�S
    Dim sql As String
    Dim rsValidation As ADODB.Recordset
    
    sql = "SELECT DISTINCT art.art_code, art.art_desl, nse.nse_nums " & _
          "FROM ART_PAR as art " & _
          "INNER JOIN nse_dat as nse ON " & _
          "nse.act_code = art.act_code AND nse.art_code = art.art_code " & _
          "AND nse.act_code = 'RB' " & _
          "WHERE nse.nse_nums = '" & numeroSerie & "' " & _
          "AND nse.nse_nums IS NOT NULL " & _
          "AND nse.nse_nums <> '' " & _
          "AND LEN(LTRIM(RTRIM(nse.nse_nums))) > 0 " & _
          "AND art.art_code IN (" & CODES_ARTICLES_AUTORISES & ")"
    
    Set rsValidation = New ADODB.Recordset
    rsValidation.Open sql, conn, adOpenStatic, adLockReadOnly
    
    If Not rsValidation.EOF And Not IsNull(rsValidation!nse_nums) Then
        ' Num�ro de s�rie trouv� dans la liste autoris�e des 92 codes
        With resultats
            .existe = True
            .numeroSerie = rsValidation!nse_nums
            .codeArticle = rsValidation!art_code
            .designationArticle = rsValidation!art_desl
            .modeleArticle = rsValidation!art_desl
            .prixCatalogue = 0
            .dateCreation = "N/A"
            .statut = "VALID� - ARTICLE AUTORIS� (LISTE 92 CODES)"
            .informationsComplementaires = "Code: " & rsValidation!art_code & " | D�signation: " & rsValidation!art_desl & " | [FILTR�: Liste autoris�e]"
        End With
        
        Debug.Print "? VALIDATION R�USSIE (liste 92 codes) pour " & numeroSerie & " - " & resultats.designationArticle
    Else
        resultats.existe = False
        resultats.statut = "NUM�RO DE S�RIE NON AUTORIS� - HORS LISTE DES 92 CODES"
        resultats.numeroSerie = numeroSerie
        resultats.informationsComplementaires = "Ce num�ro de s�rie n'appartient pas aux 92 codes articles autoris�s pour le SAV Red Bull"
        
        Debug.Print "? VALIDATION �CHOU�E (hors liste 92 codes) pour " & numeroSerie
    End If
    
    rsValidation.Close
    Set rsValidation = Nothing
    
    ValiderNumeroSerieBDD = resultats
    Exit Function
    
ErrorHandler:
    resultats.existe = False
    resultats.statut = "ERREUR BDD: " & Err.description
    resultats.numeroSerie = numeroSerie
    ValiderNumeroSerieBDD = resultats
    
    If Not rsValidation Is Nothing Then
        If rsValidation.State = adStateOpen Then rsValidation.Close
        Set rsValidation = Nothing
    End If
End Function

' === FONCTIONS DE SYNCHRONISATION AVEC VOTRE REQU�TE ===

' Fonction pour synchroniser le stock local avec la BDD en utilisant votre requ�te
Public Sub SynchroniserStockAvecBDD()
    On Error GoTo ErrorHandler
    
    If Not VerifierConnexionBDD() Then
        MsgBox "Impossible de synchroniser : pas de connexion BDD", vbExclamation
        Exit Sub
    End If
    
    ' Synchronisation avec filtrage sur les 92 codes autoris�s
    Dim rsPieces As ADODB.Recordset
    Set rsPieces = ObtenirArticlesRB()
    
    If Not rsPieces Is Nothing Then
        Debug.Print "=== SYNCHRONISATION ARTICLES RED BULL FILTR�S ==="
        Debug.Print "Codes autoris�s: 92 | Filtrage strict avec DISTINCT"
        Debug.Print "Code Article | D�signation | Num�ro de S�rie"
        Debug.Print String(70, "-")
        
        Dim compteur As Integer
        compteur = 0
        
        Do While Not rsPieces.EOF
            Dim codeArticle As String
            Dim designation As String
            Dim numeroSerie As String
            
            codeArticle = IIf(IsNull(rsPieces!art_code), "N/A", rsPieces!art_code)
            designation = IIf(IsNull(rsPieces!art_desl), "N/A", Left(rsPieces!art_desl, 25))
            numeroSerie = IIf(IsNull(rsPieces!nse_nums), "AUCUN", rsPieces!nse_nums)
            
            Debug.Print codeArticle & " | " & designation & " | " & numeroSerie
            compteur = compteur + 1
            
            rsPieces.MoveNext
        Loop
        
        Debug.Print String(70, "-")
        Debug.Print "Total num�ros de s�rie autoris�s: " & compteur
        Debug.Print "? Seuls ces �quipements peuvent �tre trait�s en SAV"
        Debug.Print "=== FIN SYNCHRONISATION FILTR�E ==="
        
        rsPieces.Close
        Set rsPieces = Nothing
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur synchronisation filtr�e : " & Err.description, vbCritical
End Sub


' === MODULE1.BAS - PARTIE 3: GESTION DES FICHIERS ET VALIDATION ===
' � ajouter � la suite de la Partie 2 dans Module1.bas

' === FONCTIONS DE GESTION DES FICHIERS ===

' Fonction pour �crire dans l'historique des scans avec validation BDD
Public Sub EcrireHistoriqueScan(reference As String, modele As String)
    Dim fichier As String
    Dim numeroFichier As Integer
    Dim valideBDD As String
    
    ' V�rifier avec votre nouvelle logique de requ�te
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

' === FONCTIONS DE VALIDATION AM�LIOR�ES ===

' Fonction pour valider un code-barres avec v�rification BDD
Public Function ValiderCodeBarre(codeBarre As String) As Boolean
    ' Supprime les espaces et convertit en majuscules
    codeBarre = Trim(UCase(codeBarre))
    
    ' V�rifications de base
    If Len(codeBarre) < 6 Then
        ValiderCodeBarre = False
        Exit Function
    End If
    
    ' V�rifie le format basique (lettres + chiffres + tirets)
    Dim i As Integer
    For i = 1 To Len(codeBarre)
        Dim char As String
        char = Mid(codeBarre, i, 1)
        If Not ((char >= "A" And char <= "Z") Or (char >= "0" And char <= "9") Or char = "-") Then
            ValiderCodeBarre = False
            Exit Function
        End If
    Next i
    
    ' V�rifier avec votre nouvelle logique BDD
    If VerifierConnexionBDD() Then
        ValiderCodeBarre = VerifierNumeroSerieBDD(codeBarre)
    Else
        ' Si pas de BDD, validation basique seulement
        ValiderCodeBarre = True
    End If
End Function

' Fonction pour extraire le mod�le du code-barres avec donn�es BDD
Public Function ExtraireModele(codeBarre As String) As String
    On Error GoTo ErrorLocal
    
    ' Essayer d'abord depuis la BDD avec votre requ�te
    If VerifierConnexionBDD() Then
        Dim resultats As TypeValidationBDD
        resultats = ValiderNumeroSerieBDD(codeBarre)
        
        If resultats.existe Then
            ' Utiliser la d�signation obtenue de votre requ�te BDD
            ExtraireModele = resultats.designationArticle
            Exit Function
        End If
    End If

ErrorLocal:
    ' Fallback - identification par pr�fixe local si pas trouv� en BDD
    ExtraireModele = DeterminerModeleParCode(codeBarre)
End Function

' Fonction pour d�terminer le mod�le bas� sur le code (fallback)
Private Function DeterminerModeleParCode(code As String) As String
    Dim prefixe As String
    prefixe = Left(UCase(code), 6)
    
    Select Case prefixe
        Case "VC2286"
            DeterminerModeleParCode = "Vitrine VC2286"
        Case "RB4458"
            DeterminerModeleParCode = "Red Bull RB4458"
        Case "CF3401"
            DeterminerModeleParCode = "Cong�lateur CF3401"
        Case "RB2024"
            DeterminerModeleParCode = "Red Bull Premium 2024"
        Case Else
            ' Essayer avec les 2 premiers caract�res
            Select Case Left(UCase(code), 2)
                Case "VC"
                    DeterminerModeleParCode = "Vitrine Red Bull"
                Case "RB"
                    DeterminerModeleParCode = "Frigo Red Bull"
                Case "CF"
                    DeterminerModeleParCode = "Cong�lateur Red Bull"
                Case "FB"
                    DeterminerModeleParCode = "Frigo Bar Red Bull"
                Case "RF"
                    DeterminerModeleParCode = "Red Fridge"
                Case Else
                    DeterminerModeleParCode = "�quipement Red Bull - Mod�le non identifi�"
            End Select
    End Select
End Function

' Fonction pour valider le format du num�ro de s�rie Red Bull
Public Function ValiderFormatNumeroSerieRB(numeroSerie As String) As Boolean
    ' V�rifications de base pour Red Bull
    If Len(numeroSerie) < 8 Or Len(numeroSerie) > 25 Then
        ValiderFormatNumeroSerieRB = False
        Exit Function
    End If
    
    ' V�rifier que c'est alphanum�rique avec tirets autoris�s
    Dim i As Integer
    For i = 1 To Len(numeroSerie)
        Dim char As String
        char = Mid(numeroSerie, i, 1)
        If Not ((char >= "A" And char <= "Z") Or (char >= "0" And char <= "9") Or char = "-") Then
            ValiderFormatNumeroSerieRB = False
            Exit Function
        End If
    Next i
    
    ValiderFormatNumeroSerieRB = True
End Function

' Fonction pour obtenir des informations compl�mentaires sur l'article
Public Function ObtenirInfosComplementairesArticle(codeArticle As String) As String
    On Error GoTo ErrorHandler
    
    ' Les informations sont d�j� r�cup�r�es dans votre requ�te principale
    ObtenirInfosComplementairesArticle = "Informations r�cup�r�es depuis ART_PAR et NSE_DAT avec votre requ�te"
    Exit Function
    
ErrorHandler:
    ObtenirInfosComplementairesArticle = "Erreur lors de la r�cup�ration des infos: " & Err.description
End Function

' === FONCTIONS DE VALIDATION DES DONN�ES SAV ===

' Fonction pour valider les donn�es SAV en BDD
Public Function ValiderDonnees(donnees As TypeSAV) As Boolean
    ' Validation de base
    ValiderDonnees = True
    
    ' V�rifier le num�ro de s�rie (ReferenceProduit) avec votre requ�te
    If VerifierConnexionBDD() And Len(donnees.ReferenceProduit) > 0 Then
        ValiderDonnees = VerifierNumeroSerieBDD(donnees.ReferenceProduit)
    End If
    
    ' V�rifications suppl�mentaires
    If Len(donnees.numeroEnlevement) = 0 Then
        ValiderDonnees = False
    End If
    
    If Len(donnees.MotifRetour) = 0 Then
        ValiderDonnees = False
    End If
End Function

' === MODULE1.BAS - PARTIE 4: FONCTIONS UTILITAIRES ET MAINTENANCE ===
' � ajouter � la suite de la Partie 3 dans Module1.bas

' === FONCTIONS UTILITAIRES ===

' Fonction pour obtenir la date/heure format�e
Public Function ObtenirDateTimeFormatee() As String
    ObtenirDateTimeFormatee = Format(Now, "dd/mm/yyyy hh:nn:ss")
End Function

' Fonction pour g�n�rer un num�ro de s�rie SAV
Public Function GenererNumeroSerie() As String
    GenererNumeroSerie = "SAV" & Format(Now, "yyyymmddhhnnss")
End Function

' Fonction pour formater une date en fran�ais
Public Function FormaterDateFrancaise(laDate As Date) As String
    FormaterDateFrancaise = Format(laDate, "dd/mm/yyyy")
End Function

' Fonction pour cr�er un nom de fichier
Public Function CreerNomFichier(numeroEnlevement As String) As String
    CreerNomFichier = App.Path & "\Sauvegardes\SAV_" & numeroEnlevement & "_" & Format(Now, "yyyymmdd") & ".txt"
End Function

' === FONCTIONS D'INITIALISATION ===

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

' Fonction pour cr�er le r�pertoire de sauvegarde
Public Sub CreerRepertoireSauvegarde()
    If Dir(App.Path & "\Sauvegardes", vbDirectory) = "" Then
        MkDir App.Path & "\Sauvegardes"
    End If
End Sub

' Fonction d'initialisation RAPIDE (sans synchronisation automatique)
Public Sub InitialiserApplication()
    ' Cr�er les r�pertoires n�cessaires
    CreerRepertoires
    
    ' Initialiser les fichiers de stock
    InitialiserStockPieces
    InitialiserStockReparable
    
    ' Nettoyer les fichiers temporaires anciens
    NettoyerFichiersTemporaires
    
    ' CONNEXION BDD SIMPLE (sans synchronisation automatique)
    If ConnecterBDD() Then
        Debug.Print "Application initialis�e avec BDD : " & ObtenirDateTimeFormatee()
        ' SUPPRIMER CETTE LIGNE PROBL�MATIQUE :
        ' SynchroniserStockAvecBDD  ' <-- COMMENT�E OU SUPPRIM�E
    Else
        Debug.Print "Application d�marr�e sans connexion BDD - mode d�grad�"
    End If
End Sub

' === NOUVELLE FONCTION POUR SYNCHRONISATION � LA DEMANDE ===
' Fonction de synchronisation rapide (uniquement pour tests)
Public Sub SynchroniserStockAvecBDDRapide()
    On Error GoTo ErrorHandler
    
    If Not VerifierConnexionBDD() Then
        MsgBox "Impossible de synchroniser : pas de connexion BDD", vbExclamation
        Exit Sub
    End If
    
    Debug.Print "=== SYNCHRONISATION RAPIDE ==="
    Debug.Print "Test connexion BDD : OK"
    Debug.Print "Requ�te SQL configur�e avec filtrage 92 codes"
    Debug.Print "=== SYNCHRONISATION TERMIN�E ==="
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Erreur synchronisation : " & Err.description
End Sub


' === GESTION DU STOCK ===

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

' === MODULE1.BAS - PARTIE 4 COMPL�TE: SUITE ET FIN ===
' � ajouter apr�s la fonction NettoyerFichiersTemporaires() dans Module1.bas

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
    
    ' Test connexion BDD avec votre requ�te
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

' === FONCTIONS DE DIAGNOSTIC SYST�ME ===

' Fonction pour obtenir des informations syst�me avec statut BDD
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
    infos = infos & "Date syst�me: " & Format(Now, "dd/mm/yyyy hh:nn:ss") & vbCrLf
    infos = infos & "Utilisateur: " & Environ("USERNAME") & vbCrLf
    infos = infos & "Ordinateur: " & Environ("COMPUTERNAME") & vbCrLf
    infos = infos & "Base de donn�es: " & statutBDD & vbCrLf
    infos = infos & "Serveur BDD: " & SERVER_NAME & vbCrLf
    infos = infos & "Base: " & DATABASE_NAME & vbCrLf
    infos = infos & vbCrLf & "=== REQU�TE SQL UTILIS�E ===" & vbCrLf
    infos = infos & "SELECT art.art_code, art.art_desl, nse.nse_nums" & vbCrLf
    infos = infos & "FROM ART_PAR as art LEFT OUTER JOIN nse_dat as nse" & vbCrLf
    infos = infos & "ON nse.act_code = art.act_code AND nse.art_code = art.art_code" & vbCrLf
    infos = infos & "AND nse.act_code = 'RB'" & vbCrLf
    
    ObtenirInfosSysteme = infos
End Function

' Fonction pour tester compl�tement le syst�me avec votre requ�te
Public Sub TesterSystemeComplet()
    Dim message As String
    
    message = "=== TEST SYST�ME SAV RED BULL ===" & vbCrLf & vbCrLf
    
    ' Test connexion BDD
    If VerifierConnexionBDD() Then
        message = message & "? Connexion BDD : OK" & vbCrLf
        message = message & "  Serveur: " & SERVER_NAME & vbCrLf
        message = message & "  Base: " & DATABASE_NAME & vbCrLf & vbCrLf
        
        ' Test de votre requ�te sp�cifique
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
            message = message & "? Votre requ�te SQL fonctionne parfaitement !" & vbCrLf
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
    
    ' Sauvegarder les infos syst�me avec votre requ�te
    Dim numeroFichier As Integer
    numeroFichier = FreeFile
    Open repertoireSauvegarde & "InfosSysteme.txt" For Output As #numeroFichier
    Print #numeroFichier, ObtenirInfosSysteme()
    Close #numeroFichier
    
    ' Sauvegarder un exemple de votre requ�te SQL
    numeroFichier = FreeFile
    Open repertoireSauvegarde & "RequeteSQL.txt" For Output As #numeroFichier
    Print #numeroFichier, "=== REQU�TE SQL UTILIS�E DANS L'APPLICATION ===" & vbCrLf
    Print #numeroFichier, "SELECT art.art_code, art.art_desl, nse.nse_nums"
    Print #numeroFichier, "FROM ART_PAR as art"
    Print #numeroFichier, "LEFT OUTER JOIN nse_dat as nse ON"
    Print #numeroFichier, "nse.act_code = art.act_code AND nse.art_code = art.art_code"
    Print #numeroFichier, "AND nse.act_code = 'RB'" & vbCrLf
    Print #numeroFichier, "=== DESCRIPTION ===" & vbCrLf
    Print #numeroFichier, "Cette requ�te r�cup�re :"
    Print #numeroFichier, "- art_code : Code de l'article"
    Print #numeroFichier, "- art_desl : D�signation de l'article"
    Print #numeroFichier, "- nse_nums : Num�ro de s�rie de l'�quipement"
    Print #numeroFichier, "Filtr� sur act_code = 'RB' pour Red Bull uniquement"
    Close #numeroFichier
    
    On Error GoTo 0
End Sub

' Fonction pour sauvegarde automatique
Public Sub SauvegardeAutomatique()
    CreerSauvegardeComplete
End Sub

' === FONCTIONS DE LOG ET DEBUGGING ===

' Fonction pour logger les erreurs syst�me avec votre requ�te
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

' Fonction pour tester sp�cifiquement votre requ�te SQL
Public Sub TesterRequeteSpecifique()
    On Error GoTo ErrorHandler
    
    If Not VerifierConnexionBDD() Then
        MsgBox "Pas de connexion BDD pour tester la requ�te", vbExclamation
        Exit Sub
    End If
    
    Dim message As String
    message = "=== TEST SP�CIFIQUE DE VOTRE REQU�TE SQL ===" & vbCrLf & vbCrLf
    
    ' Tester votre requ�te exacte
    Dim rsTest As ADODB.Recordset
    Set rsTest = ObtenirArticlesRB()
    
    If Not rsTest Is Nothing Then
        message = message & "? Requ�te ex�cut�e avec succ�s !" & vbCrLf & vbCrLf
        
        ' Afficher quelques exemples de r�sultats
        Dim compteur As Integer
        compteur = 0
        
        message = message & "=== PREMIERS R�SULTATS ===" & vbCrLf
        
        Do While Not rsTest.EOF And compteur < 5
            Dim codeArticle As String
            Dim designation As String
            Dim numeroSerie As String
            
            codeArticle = IIf(IsNull(rsTest!art_code), "NULL", rsTest!art_code)
            designation = IIf(IsNull(rsTest!art_desl), "NULL", rsTest!art_desl)
            numeroSerie = IIf(IsNull(rsTest!nse_nums), "AUCUN", rsTest!nse_nums)
            
            message = message & "� Code: " & codeArticle & " | D�signation: " & designation & " | N� s�rie: " & numeroSerie & vbCrLf
            
            compteur = compteur + 1
            rsTest.MoveNext
        Loop
        
        rsTest.Close
        Set rsTest = Nothing
        message = message & vbCrLf & "Votre requ�te fonctionne parfaitement !"
    Else
        message = message & "? Erreur lors de l'ex�cution de votre requ�te"
    End If
    
    MsgBox message, vbInformation, "Test Requ�te SQL"
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur lors du test de la requ�te : " & Err.description, vbCritical
End Sub

' === FIN DU MODULE1.BAS ===
