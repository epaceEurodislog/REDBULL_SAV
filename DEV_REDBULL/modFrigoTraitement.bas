Attribute VB_Name = "modFrigoTraitement"
Attribute VB_Name = "modFrigoTraitement"

' ========== RÉCEPTION DES ÉQUIPEMENTS ==========
Private Sub LireFichierReception(sFichier As String)
    Dim NFile As Integer
    Dim sLigne As String
    Dim compteurLigne As Long
    
    On Error GoTo ErrorHandler
    
    NFile = FreeFile
    Open sFichier For Input As #NFile
    
    compteurLigne = 0
    While Not EOF(NFile)
        Line Input #NFile, sLigne
        compteurLigne = compteurLigne + 1
        
        ' Ignorer l'en-tête
        If compteurLigne > 1 And Len(Trim(sLigne)) > 0 Then
            Call TraiterLigneReception(sLigne, compteurLigne)
        End If
    Wend
    
    Close #NFile
    Exit Sub
    
ErrorHandler:
    If NFile > 0 Then Close #NFile
    DDE_SEND "Erreur lecture fichier: " & Err.Description
End Sub

Private Sub TraiterLigneReception(sLigne As String, numeroLigne As Long)
    Dim sSerial As String
    Dim sMarque As String
    Dim sModele As String
    Dim sDescription As String
    Dim sSQL As String
    
    On Error GoTo ErrorHandler
    
    ' Parsing de la ligne (format CSV avec séparateur ;)
    sSerial = CSVZone(sLigne, 1, ";")
    sMarque = CSVZone(sLigne, 2, ";")
    sModele = CSVZone(sLigne, 3, ";")
    sDescription = CSVZone(sLigne, 4, ";")
    
    ' Vérification des données obligatoires
    If Len(Trim(sSerial)) = 0 Then
        DDE_SEND "Attention ligne " & CStr(numeroLigne) & ": numéro de série vide"
        Exit Sub
    End If
    
    ' Utilisation de votre classe cSQL_Query avec les fonctions Insert
    Query.SQL_Insert_Init
    Query.SQL_Insert_Add "SerialNumber", sSerial, Format_Texte
    Query.SQL_Insert_Add "Brand", sMarque, Format_Texte
    Query.SQL_Insert_Add "Model", sModele, Format_Texte
    Query.SQL_Insert_Add "Description", sDescription, Format_Texte
    Query.SQL_Insert_Add "Status", "0", Format_Number ' 0 = Réception
    Query.SQL_Insert_Add "EntryDate", Now, Format_Date
    Query.SQL_Insert_Add "CreationUser", Environ("USERNAME"), Format_Texte
    
    Query.SQL_Insert_Exc "FRIGO_EQUIPMENT", "REDBULL_FRIGOS", Insert_Into
    
    Exit Sub
    
ErrorHandler:
    DDE_SEND "Erreur traitement ligne " & CStr(numeroLigne) & ": " & Err.Description
End Sub

' ========== DIAGNOSTIC ==========
Private Sub TraiterDiagnostics(sFichier As String)
    Dim NFile As Integer
    Dim sLigne As String
    Dim compteurLigne As Long
    
    On Error GoTo ErrorHandler
    
    NFile = FreeFile
    Open sFichier For Input As #NFile
    
    compteurLigne = 0
    While Not EOF(NFile)
        Line Input #NFile, sLigne
        compteurLigne = compteurLigne + 1
        
        If compteurLigne > 1 And Len(Trim(sLigne)) > 0 Then
            Call TraiterLigneDiagnostic(sLigne, compteurLigne)
        End If
    Wend
    
    Close #NFile
    Exit Sub
    
ErrorHandler:
    If NFile > 0 Then Close #NFile
    DDE_SEND "Erreur lecture fichier diagnostic: " & Err.Description
End Sub

Private Sub TraiterLigneDiagnostic(sLigne As String, numeroLigne As Long)
    Dim sSerial As String
    Dim sDiagnostic As String
    Dim sEtat As String
    Dim sTechnicien As String
    Dim nouveauStatut As String
    
    On Error GoTo ErrorHandler
    
    sSerial = CSVZone(sLigne, 1, ";")
    sDiagnostic = CSVZone(sLigne, 2, ";")
    sEtat = CSVZone(sLigne, 3, ";") ' "REPARABLE", "PIECES", "DESTRUCTION"
    sTechnicien = CSVZone(sLigne, 4, ";")
    
    ' Déterminer le nouveau statut selon votre diagramme
    Select Case UCase(Trim(sEtat))
        Case "REPARABLE"
            nouveauStatut = "6" ' StatusRepairable
        Case "PIECES", "DONNEUR"
            nouveauStatut = "7" ' StatusPartsProvider
        Case "DESTRUCTION"
            nouveauStatut = "11" ' StatusDestruction
        Case Else
            nouveauStatut = "5" ' StatusAwaitingDiagnosis
    End Select
    
    ' Mise à jour de l'équipement
    Query.SQL_Update_Init
    Query.SQL_Update_Add "Status", nouveauStatut, Format_Number
    Query.SQL_Update_Add "DiagnosticDate", Now, Format_Date
    Query.SQL_Update_Add "DiagnosticNotes", sDiagnostic, Format_Texte
    Query.SQL_Update_Add "TechnicianName", sTechnicien, Format_Texte
    Query.SQL_Update_Add "LastUpdateDate", Now, Format_Date
    
    Query.SQL_Table = "FRIGO_EQUIPMENT"
    Query.SQL_Where = "SerialNumber = '" & sSerial & "'"
    Query.SQL_Update_Exc
    
    Exit Sub
    
ErrorHandler:
    DDE_SEND "Erreur diagnostic ligne " & CStr(numeroLigne) & ": " & Err.Description
End Sub

' ========== GÉNÉRATION RAPPORTS EXCEL ==========
Private Sub GenererRapportReception(sFichier As String)
    Dim sSQL As String
    Dim Table As Variant
    Dim Temp() As Variant
    Dim appExcel As Object
    Dim wbExcel As Object
    Dim wsExcel As Object
    Dim i As Long, j As Long
    Dim NLf As Long
    Dim ColF As String
    
    On Error GoTo ErrorHandler
    
    ' Requête pour les réceptions du jour
    sSQL = "SELECT SerialNumber, Brand, Model, Description, EntryDate, CreationUser " & _
           "FROM FRIGO_EQUIPMENT " & _
           "WHERE CONVERT(date, EntryDate) = CONVERT(date, GETDATE()) " & _
           "ORDER BY EntryDate DESC"
    
    Table = Query.SQL_Get_Query(sSQL)
    
    If Query.SQL_Count = 0 Then
        DDE_SEND "Aucune réception aujourd'hui"
        Exit Sub
    End If
    
    ' Création du tableau avec en-têtes
    ReDim Temp(0 To UBound(Table, 2) + 1, 0 To UBound(Table, 1) + 1) As Variant
    
    ' En-têtes
    Temp(0, 0) = "Numéro de série"
    Temp(0, 1) = "Marque"
    Temp(0, 2) = "Modèle"
    Temp(0, 3) = "Description"
    Temp(0, 4) = "Date de réception"
    Temp(0, 5) = "Utilisateur"
    
    ' Données
    For i = 0 To UBound(Table, 1)
        For j = 0 To UBound(Table, 2)
            Temp(j + 1, i + 1) = Table(i, j)
        Next j
    Next i
    
    ' Création du fichier Excel avec votre système existant
    Set appExcel = CreateObject("Excel.Application")
    Set wbExcel = appExcel.Workbooks.Add
    Set wsExcel = wbExcel.ActiveSheet
    
    ' Configuration de la feuille
    wbExcel.Sheets("Feuil2").Delete
    wbExcel.Sheets("Feuil3").Delete
    wsExcel.name = "Réceptions du " & Format(Date, "dd-mm-yyyy")
    
    NLf = UBound(Temp, 2) + 1
    ColF = Base_26(UBound(Temp, 1))
    
    ' Insertion des données
    wsExcel.Range("A1").Resize(NLf, UBound(Temp, 1) + 1).Value = appExcel.Transpose(appExcel.Transpose(Temp))
    
    ' Formatage selon votre style existant
    wsExcel.Range("A1:" & ColF & "1").Font.Bold = True
    wsExcel.Range("A1:" & ColF & "1").Interior.ThemeColor = 1
    wsExcel.Range("A1:" & ColF & "1").Interior.TintAndShade = -0.149998474074526
    wsExcel.Range("A1:" & ColF & "1").HorizontalAlignment = xlCenter
    
    ' Formatage des colonnes
    wsExcel.Range("E2:E" & CStr(NLf)).NumberFormat = "dd/mm/yyyy hh:mm"
    wsExcel.Columns("A:" & ColF).EntireColumn.AutoFit
    
    ' Création du cadre avec votre fonction existante
    Call Creation_Cadre(wbExcel, wsExcel.name, "A1:" & ColF & CStr(NLf))
    
    ' Titre du rapport
    wsExcel.Range("A" & CStr(NLf + 2)).Value = "Rapport réceptions frigos Red Bull - " & Format(Now, "dd/mm/yyyy hh:mm") & " - " & CStr(Query.SQL_Count) & " équipements"
    
    ' Sauvegarde
    wbExcel.SaveAs sFichier, 51
    wbExcel.Close
    appExcel.Quit
    
    Set wsExcel = Nothing
    Set wbExcel = Nothing
    Set appExcel = Nothing
    
    DDE_SEND "Rapport réception généré: " & sFichier
    Exit Sub
    
ErrorHandler:
    DDE_SEND "Erreur génération rapport: " & Err.Description
    On Error Resume Next
    If Not wbExcel Is Nothing Then wbExcel.Close
    If Not appExcel Is Nothing Then appExcel.Quit
    Set wsExcel = Nothing
    Set wbExcel = Nothing
    Set appExcel = Nothing
End Sub

' ========== STUBS POUR LES AUTRES FONCTIONS ==========
Private Sub TraiterReparations(sFichier As String)
    ' À implémenter selon vos besoins
    DDE_SEND "Fonction TraiterReparations en cours de développement"
End Sub

Private Sub GenererRapportReparation(sFichier As String)
    ' À implémenter selon vos besoins
    DDE_SEND "Fonction GenererRapportReparation en cours de développement"
End Sub

Private Sub MettreAJourStockPieces(sFichier As String)
    ' À implémenter selon vos besoins
    DDE_SEND "Fonction MettreAJourStockPieces en cours de développement"
End Sub

Private Sub GenererRapportStockPieces(sFichier As String)
    ' À implémenter selon vos besoins
    DDE_SEND "Fonction GenererRapportStockPieces en cours de développement"
End Sub

Private Sub TraiterDemontage(sFichier As String)
    ' À implémenter selon vos besoins
    DDE_SEND "Fonction TraiterDemontage en cours de développement"
End Sub

Private Sub GenererRapportDemontage(sFichier As String)
    ' À implémenter selon vos besoins
    DDE_SEND "Fonction GenererRapportDemontage en cours de développement"
End Sub

Private Sub TraiterExpedition(sFichier As String)
    ' À implémenter selon vos besoins
    DDE_SEND "Fonction TraiterExpedition en cours de développement"
End Sub

Private Sub GenererRapportExpedition(sFichier As String)
    ' À implémenter selon vos besoins
    DDE_SEND "Fonction GenererRapportExpedition en cours de développement"
End Sub

Private Sub TraiterRetours(sFichier As String)
    ' À implémenter selon vos besoins
    DDE_SEND "Fonction TraiterRetours en cours de développement"
End Sub

Private Sub GenererRapportRetours(sFichier As String)
    ' À implémenter selon vos besoins
    DDE_SEND "Fonction GenererRapportRetours en cours de développement"
End Sub

Private Sub GenererSuiviGlobal(sFichier As String)
    ' À implémenter selon vos besoins - similaire à votre REP_SUIVI existant
    DDE_SEND "Fonction GenererSuiviGlobal en cours de développement"
End Sub

Private Sub GenererRapportStock(sFichier As String)
    ' À implémenter selon vos besoins
    DDE_SEND "Fonction GenererRapportStock en cours de développement"
End Sub

Private Sub GenererRapportPieces(sFichier As String)
    ' À implémenter selon vos besoins
    DDE_SEND "Fonction GenererRapportPieces en cours de développement"
End Sub

ErrorHandler:
    DDE_SEND "Erreur dans modFrigoTraitement: " & Err.Description
End Sub

