Attribute VB_Name = "modRequetesBDD"
' === modRequetesBDD.bas - MODULE COMPLET ===
Option Explicit

' === FONCTIONS DE REQUÊTES BDD ===

' Fonction pour obtenir tous les articles Red Bull avec filtrage 92 codes
Public Function ObtenirArticlesRB() As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    If Not Reconnecter() Then
        Set ObtenirArticlesRB = Nothing
        Exit Function
    End If
    
    ' REQUÊTE SQL FILTRÉE SUR LES 92 CODES ARTICLES AUTORISÉS
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
    
    Debug.Print "Requête filtrée exécutée - 92 codes articles autorisés"
    Exit Function
    
ErrorHandler:
    MsgBox "Erreur lors de la requête articles RB filtrée : " & Err.description, vbCritical
    Set ObtenirArticlesRB = Nothing
End Function

' Fonction pour vérifier si un numéro de série existe dans la BDD
Public Function VerifierNumeroSerieBDD(numeroSerie As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim sql As String
    Dim rsVerif As ADODB.Recordset
    
    ' VÉRIFICATION AVEC FILTRAGE SUR LES 92 CODES AUTORISÉS
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
    MsgBox "Erreur lors de la vérification du numéro de série : " & Err.description, vbCritical
End Function

' Fonction principale pour valider un numéro de série
Public Function ValiderNumeroSerieBDD(numeroSerie As String) As TypeValidationBDD
    On Error GoTo ErrorHandler
    
    Dim resultats As TypeValidationBDD
    resultats.existe = False
    
    If Not Reconnecter() Then
        resultats.statut = "CONNEXION IMPOSSIBLE"
        ValiderNumeroSerieBDD = resultats
        Exit Function
    End If
    
    ' VALIDATION AVEC FILTRAGE SUR LES 92 CODES AUTORISÉS
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
        ' Numéro de série trouvé dans la liste autorisée
        With resultats
            .existe = True
            .numeroSerie = rsValidation!nse_nums
            .codeArticle = rsValidation!ART_CODE
            .designationArticle = rsValidation!art_desl
            .modeleArticle = rsValidation!art_desl
            .prixCatalogue = 0
            .dateCreation = "N/A"
            .statut = "VALIDÉ - ARTICLE AUTORISÉ (LISTE 92 CODES)"
            .informationsComplementaires = "Code: " & rsValidation!ART_CODE & " | [FILTRÉ]"
        End With
        
        Debug.Print "Validation réussie pour " & numeroSerie
    Else
        resultats.existe = False
        resultats.statut = "NUMÉRO DE SÉRIE NON AUTORISÉ - HORS LISTE DES 92 CODES"
        resultats.numeroSerie = numeroSerie
        
        Debug.Print "Validation échouée pour " & numeroSerie
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

' Fonction pour récupérer les données REE_DAT par numéro de série (MODIFIÉE)
Public Function RecupererDonneesREE(numeroSerie As String) As TypeDonneesREE
    On Error GoTo ErrorHandler
    
    Dim resultats As TypeDonneesREE
    resultats.trouve = False
    resultats.numeroReception = ""
    resultats.numeroEnlevement = ""  ' On garde la structure mais on ne l'utilise plus
    resultats.messageErreur = ""
    
    If Not Reconnecter() Then
        resultats.messageErreur = "CONNEXION BDD IMPOSSIBLE"
        RecupererDonneesREE = resultats
        Exit Function
    End If
    
    ' Requête pour récupérer SEULEMENT REE_Nore depuis REE_DAT
    Dim sql As String
    Dim rsREE As ADODB.Recordset
    
    sql = "SELECT REE_Nore " & _
          "FROM REE_DAT " & _
          "WHERE nse_nums = '" & numeroSerie & "' " & _
          "AND REE_Nore IS NOT NULL " & _
          "AND REE_Nore <> ''"
    
    Set rsREE = New ADODB.Recordset
    rsREE.Open sql, conn, adOpenStatic, adLockReadOnly
    
    If Not rsREE.EOF Then
        ' Données trouvées
        With resultats
            .trouve = True
            .numeroReception = IIf(IsNull(rsREE!REE_NORE), "", Trim(rsREE!REE_NORE))
            .numeroEnlevement = ""  ' Pas utilisé
            .messageErreur = ""
        End With
        
        Debug.Print "Numéro de réception REE trouvé pour " & numeroSerie & " : " & resultats.numeroReception
    Else
        ' Aucune donnée trouvée
        resultats.trouve = False
        resultats.messageErreur = "Aucun numéro de réception REE_Nore trouvé pour ce numéro de série"
        
        Debug.Print "Aucun numéro de réception REE trouvé pour " & numeroSerie
    End If
    
    rsREE.Close
    Set rsREE = Nothing
    
    RecupererDonneesREE = resultats
    Exit Function
    
ErrorHandler:
    resultats.trouve = False
    resultats.messageErreur = "ERREUR BDD REE_DAT: " & Err.description
    RecupererDonneesREE = resultats
    
    If Not rsREE Is Nothing Then
        If rsREE.State = adStateOpen Then rsREE.Close
        Set rsREE = Nothing
    End If
End Function

' Fonction pour récupérer le numéro de réception par numéro de série (AVEC VOTRE REQUÊTE)
Public Function RecupererNumeroReception(numeroSerie As String) As TypeDonneesREE
    On Error GoTo ErrorHandler
    
    Dim resultats As TypeDonneesREE
    resultats.trouve = False
    resultats.numeroReception = ""
    resultats.numeroEnlevement = ""
    resultats.messageErreur = ""
    
    If Not Reconnecter() Then
        resultats.messageErreur = "CONNEXION BDD IMPOSSIBLE"
        RecupererNumeroReception = resultats
        Exit Function
    End If
    
    ' Votre requête adaptée pour récupérer REE_NORE par numéro de série
    Dim sql As String
    Dim rsREE As ADODB.Recordset
    
    sql = "SELECT REL.REE_NORE " & _
          "FROM REL_DAT REL " & _
          "JOIN NSE_DAT NSE ON NSE.ART_CODE = REL.ART_CODE AND REL.REL_NoSU = NSE.STK_NoSU " & _
          "WHERE REL.ACT_CODE = 'RB' " & _
          "AND NSE.NSE_NUMS = '" & numeroSerie & "' " & _
          "AND REL.REE_NORE IS NOT NULL " & _
          "AND REL.REE_NORE <> ''"
    
    Set rsREE = New ADODB.Recordset
    rsREE.Open sql, conn, adOpenStatic, adLockReadOnly
    
    If Not rsREE.EOF Then
        ' Numéro de réception trouvé
        With resultats
            .trouve = True
            .numeroReception = IIf(IsNull(rsREE!REE_NORE), "", Trim(rsREE!REE_NORE))
            .numeroEnlevement = ""  ' Pas utilisé
            .messageErreur = ""
        End With
        
        Debug.Print "Numéro de réception trouvé pour " & numeroSerie & " : " & resultats.numeroReception
    Else
        ' Aucune donnée trouvée
        resultats.trouve = False
        resultats.messageErreur = "Aucun numéro de réception trouvé pour ce numéro de série"
        
        Debug.Print "Aucun numéro de réception trouvé pour " & numeroSerie
    End If
    
    rsREE.Close
    Set rsREE = Nothing
    
    RecupererNumeroReception = resultats
    Exit Function
    
ErrorHandler:
    resultats.trouve = False
    resultats.messageErreur = "ERREUR BDD REL_DAT/NSE_DAT: " & Err.description
    RecupererNumeroReception = resultats
    
    If Not rsREE Is Nothing Then
        If rsREE.State = adStateOpen Then rsREE.Close
        Set rsREE = Nothing
    End If
End Function


' Fonction corrigée avec gestion des types de données
Public Function RecupererNumeroReceptionAvecArtCode(numeroSerie As String) As TypeDonneesREE
    On Error GoTo ErrorHandler
    
    Dim resultats As TypeDonneesREE
    resultats.trouve = False
    resultats.numeroReception = ""
    resultats.numeroEnlevement = ""
    resultats.messageErreur = ""
    
    If Not Reconnecter() Then
        resultats.messageErreur = "CONNEXION BDD IMPOSSIBLE"
        RecupererNumeroReceptionAvecArtCode = resultats
        Exit Function
    End If
    
    ' ÉTAPE 1 : Récupérer l'ART_CODE et STK_NoSU à partir du numéro de série
    Dim sqlEtape1 As String
    Dim rsEtape1 As ADODB.Recordset
    Dim artCode As String
    Dim stkNoSU As String
    
    sqlEtape1 = "SELECT ART_CODE, STK_NoSU " & _
                "FROM NSE_DAT " & _
                "WHERE NSE_NUMS = '" & numeroSerie & "' " & _
                "AND ACT_CODE = 'RB'"
    
    Set rsEtape1 = New ADODB.Recordset
    rsEtape1.Open sqlEtape1, conn, adOpenStatic, adLockReadOnly
    
    If rsEtape1.EOF Then
        resultats.messageErreur = "Numéro de série non trouvé dans NSE_DAT"
        rsEtape1.Close
        Set rsEtape1 = Nothing
        RecupererNumeroReceptionAvecArtCode = resultats
        Exit Function
    End If
    
    artCode = IIf(IsNull(rsEtape1!ART_CODE), "", Trim(rsEtape1!ART_CODE))
    stkNoSU = IIf(IsNull(rsEtape1!STK_NoSU), "", Trim(rsEtape1!STK_NoSU))
    rsEtape1.Close
    Set rsEtape1 = Nothing
    
    Debug.Print "ART_CODE: " & artCode & ", STK_NoSU: " & stkNoSU
    
    ' ÉTAPE 2 : Chercher dans REL_DAT avec les valeurs exactes
    Dim sqlEtape2 As String
    Dim rsEtape2 As ADODB.Recordset
    
    ' CORRECTION : Utiliser les mêmes types de données et éviter les conversions
    sqlEtape2 = "SELECT REE_NORE " & _
                "FROM REL_DAT " & _
                "WHERE ACT_CODE = 'RB' " & _
                "AND ART_CODE = '" & artCode & "' " & _
                "AND LTRIM(RTRIM(STR(REL_NoSU))) = '" & stkNoSU & "' " & _
                "AND REE_NORE IS NOT NULL " & _
                "AND REE_NORE <> ''"
    
    Set rsEtape2 = New ADODB.Recordset
    rsEtape2.Open sqlEtape2, conn, adOpenStatic, adLockReadOnly
    
    If Not rsEtape2.EOF Then
        With resultats
            .trouve = True
            .numeroReception = IIf(IsNull(rsEtape2!REE_NORE), "", Trim(rsEtape2!REE_NORE))
            .numeroEnlevement = ""
            .messageErreur = ""
        End With
        
        Debug.Print "REE_NORE trouvé : " & resultats.numeroReception
    Else
        resultats.messageErreur = "Aucun REE_NORE trouvé pour ART_CODE: " & artCode & " et STK_NoSU: " & stkNoSU
    End If
    
    rsEtape2.Close
    Set rsEtape2 = Nothing
    
    RecupererNumeroReceptionAvecArtCode = resultats
    Exit Function
    
ErrorHandler:
    resultats.trouve = False
    resultats.messageErreur = "ERREUR BDD: " & Err.description
    RecupererNumeroReceptionAvecArtCode = resultats
    
    ' Nettoyage des objets
    If Not rsEtape1 Is Nothing Then
        If rsEtape1.State = adStateOpen Then rsEtape1.Close
        Set rsEtape1 = Nothing
    End If
    
    If Not rsEtape2 Is Nothing Then
        If rsEtape2.State = adStateOpen Then rsEtape2.Close
        Set rsEtape2 = Nothing
    End If
End Function
' Fonction pour synchroniser le stock avec la BDD (optionnelle)
Public Sub SynchroniserStockAvecBDD()
    On Error GoTo ErrorHandler
    
    If Not VerifierConnexionBDD() Then
        MsgBox "Impossible de synchroniser : pas de connexion BDD", vbExclamation
        Exit Sub
    End If
    
    ' Synchronisation avec filtrage sur les 92 codes autorisés
    Dim rsPieces As ADODB.Recordset
    Set rsPieces = ObtenirArticlesRB()
    
    If Not rsPieces Is Nothing Then
        Debug.Print "=== SYNCHRONISATION ARTICLES RED BULL FILTRÉS ==="
        Debug.Print "Codes autorisés: 92 | Filtrage strict avec DISTINCT"
        
        Dim compteur As Integer
        compteur = 0
        
        Do While Not rsPieces.EOF
            compteur = compteur + 1
            rsPieces.MoveNext
        Loop
        
        Debug.Print "Total numéros de série autorisés: " & compteur
        
        rsPieces.Close
        Set rsPieces = Nothing
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur synchronisation filtrée : " & Err.description, vbCritical
End Sub

' Version ultra-simplifiée sans jointures complexes
Public Function RecupererNumeroReceptionDirect(numeroSerie As String) As TypeDonneesREE
    On Error GoTo ErrorHandler
    
    Dim resultats As TypeDonneesREE
    resultats.trouve = False
    resultats.numeroReception = ""
    resultats.messageErreur = ""
    
    If Not Reconnecter() Then
        resultats.messageErreur = "CONNEXION BDD IMPOSSIBLE"
        RecupererNumeroReceptionDirect = resultats
        Exit Function
    End If
    
    ' Approche directe : chercher dans REE_DAT d'abord
    Dim sql As String
    Dim rs As ADODB.Recordset
    
    sql = "SELECT TOP 1 REE_NORE " & _
          "FROM REE_DAT " & _
          "WHERE nse_nums = '" & numeroSerie & "' " & _
          "AND ACT_CODE = 'RB' " & _
          "AND REE_NORE IS NOT NULL " & _
          "AND REE_NORE <> ''"
    
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then
        ' Trouvé dans REE_DAT
        With resultats
            .trouve = True
            .numeroReception = IIf(IsNull(rs!REE_NORE), "", Trim(rs!REE_NORE))
            .messageErreur = ""
        End With
        
        Debug.Print "REE_NORE trouvé directement : " & resultats.numeroReception
    Else
        ' Pas trouvé, essayer avec REL_DAT sans jointure
        rs.Close
        Set rs = Nothing
        
        sql = "SELECT TOP 1 REE_NORE " & _
              "FROM REL_DAT " & _
              "WHERE ACT_CODE = 'RB' " & _
              "AND REE_NORE IS NOT NULL " & _
              "AND REE_NORE <> '' " & _
              "AND EXISTS ( " & _
              "    SELECT 1 FROM NSE_DAT " & _
              "    WHERE NSE_DAT.NSE_NUMS = '" & numeroSerie & "' " & _
              "    AND NSE_DAT.ACT_CODE = 'RB' " & _
              "    AND NSE_DAT.ART_CODE = REL_DAT.ART_CODE " & _
              ")"
        
        Set rs = New ADODB.Recordset
        rs.Open sql, conn, adOpenStatic, adLockReadOnly
        
        If Not rs.EOF Then
            With resultats
                .trouve = True
                .numeroReception = IIf(IsNull(rs!REE_NORE), "", Trim(rs!REE_NORE))
                .messageErreur = ""
            End With
            
            Debug.Print "REE_NORE trouvé via EXISTS : " & resultats.numeroReception
        Else
            resultats.messageErreur = "Aucun numéro de réception trouvé"
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    
    RecupererNumeroReceptionDirect = resultats
    Exit Function
    
ErrorHandler:
    resultats.trouve = False
    resultats.messageErreur = "ERREUR BDD: " & Err.description
    RecupererNumeroReceptionDirect = resultats
    
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
End Function

