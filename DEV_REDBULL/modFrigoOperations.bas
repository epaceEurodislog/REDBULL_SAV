Attribute VB_Name = "modFrigoOperations"
Attribute VB_Name = "modFrigoOperations"

' Variables globales héritées du formulaire principal
Public Declare Conv As Conversion
Public Declare Query As cSQL_Query

' ========== RÉCEPTION DES FRIGOS ==========
Public Sub FRIGO_RECEPTION()
    DDE_SEND "Début traitement réception frigos"
    
    ' Configuration de la base de données
    Set Query = New cSQL_Query
    Query.SQL_User = "sa"
    Query.SQL_Pswd = "sadmin"
    Query.SQL_Host = "192.168.9.12"
    Query.SQL_Base = "REDBULL_FRIGOS"
    Query.SQL_Type = SQL_Type_MsSQL
    
    ' Lecture du fichier d'entrée
    Call LireFichierReception(WRK_REP & WRK_FIC & ".txt")
    
    ' Génération du rapport Excel
    Call GenererRapportReception(WRK_REP & "SORTIE_" & WRK_FIC & ".xlsx")
    
    DDE_SEND "Fin traitement réception frigos"
End Sub

' ========== DIAGNOSTIC FRIGORISTE ==========
Public Sub FRIGO_DIAGNOSTIC()
    DDE_SEND "Début traitement diagnostic frigos"
    
    Set Query = New cSQL_Query
    Query.SQL_User = "sa"
    Query.SQL_Pswd = "sadmin"
    Query.SQL_Host = "192.168.9.12"
    Query.SQL_Base = "REDBULL_FRIGOS"
    Query.SQL_Type = SQL_Type_MsSQL
    
    Call TraiterDiagnostics(WRK_REP & WRK_FIC & ".txt")
    Call GenererRapportDiagnostic(WRK_REP & "SORTIE_" & WRK_FIC & ".xlsx")
    
    DDE_SEND "Fin traitement diagnostic frigos"
End Sub

' ========== RÉPARATION ==========
Public Sub FRIGO_REPARATION()
    DDE_SEND "Début traitement réparation frigos"
    
    Set Query = New cSQL_Query
    Query.SQL_User = "sa"
    Query.SQL_Pswd = "sadmin"
    Query.SQL_Host = "192.168.9.12"
    Query.SQL_Base = "REDBULL_FRIGOS"
    Query.SQL_Type = SQL_Type_MsSQL
    
    Call TraiterReparations(WRK_REP & WRK_FIC & ".txt")
    Call GenererRapportReparation(WRK_REP & "SORTIE_" & WRK_FIC & ".xlsx")
    
    DDE_SEND "Fin traitement réparation frigos"
End Sub

' ========== GESTION STOCK PIÈCES ==========
Public Sub FRIGO_PIECES_STOCK()
    DDE_SEND "Début traitement stock pièces"
    
    Set Query = New cSQL_Query
    Query.SQL_User = "sa"
    Query.SQL_Pswd = "sadmin"
    Query.SQL_Host = "192.168.9.12"
    Query.SQL_Base = "REDBULL_FRIGOS"
    Query.SQL_Type = SQL_Type_MsSQL
    
    Call MettreAJourStockPieces(WRK_REP & WRK_FIC & ".txt")
    Call GenererRapportStockPieces(WRK_REP & "SORTIE_" & WRK_FIC & ".xlsx")
    
    DDE_SEND "Fin traitement stock pièces"
End Sub

' ========== DÉMONTAGE POUR PIÈCES ==========
Public Sub FRIGO_DEMONTAGE()
    DDE_SEND "Début traitement démontage frigos"
    
    Set Query = New cSQL_Query
    Query.SQL_User = "sa"
    Query.SQL_Pswd = "sadmin"
    Query.SQL_Host = "192.168.9.12"
    Query.SQL_Base = "REDBULL_FRIGOS"
    Query.SQL_Type = SQL_Type_MsSQL
    
    Call TraiterDemontage(WRK_REP & WRK_FIC & ".txt")
    Call GenererRapportDemontage(WRK_REP & "SORTIE_" & WRK_FIC & ".xlsx")
    
    DDE_SEND "Fin traitement démontage frigos"
End Sub

' ========== EXPÉDITION ==========
Public Sub FRIGO_EXPEDITION()
    DDE_SEND "Début traitement expédition frigos"
    
    Set Query = New cSQL_Query
    Query.SQL_User = "sa"
    Query.SQL_Pswd = "sadmin"
    Query.SQL_Host = "192.168.9.12"
    Query.SQL_Base = "REDBULL_FRIGOS"
    Query.SQL_Type = SQL_Type_MsSQL
    
    Call TraiterExpedition(WRK_REP & WRK_FIC & ".txt")
    Call GenererRapportExpedition(WRK_REP & "SORTIE_" & WRK_FIC & ".xlsx")
    
    DDE_SEND "Fin traitement expédition frigos"
End Sub

' ========== RETOURS ==========
Public Sub FRIGO_RETOUR()
    DDE_SEND "Début traitement retours frigos"
    
    Set Query = New cSQL_Query
    Query.SQL_User = "sa"
    Query.SQL_Pswd = "sadmin"
    Query.SQL_Host = "192.168.9.12"
    Query.SQL_Base = "REDBULL_FRIGOS"
    Query.SQL_Type = SQL_Type_MsSQL
    
    Call TraiterRetours(WRK_REP & WRK_FIC & ".txt")
    Call GenererRapportRetours(WRK_REP & "SORTIE_" & WRK_FIC & ".xlsx")
    
    DDE_SEND "Fin traitement retours frigos"
End Sub

' ========== SUIVI GLOBAL ==========
Public Sub FRIGO_SUIVI_GLOBAL()
    DDE_SEND "Début génération suivi global frigos"
    
    Set Query = New cSQL_Query
    Query.SQL_User = "sa"
    Query.SQL_Pswd = "sadmin"
    Query.SQL_Host = "192.168.9.12"
    Query.SQL_Base = "REDBULL_FRIGOS"
    Query.SQL_Type = SQL_Type_MsSQL
    
    Call GenererSuiviGlobal(WRK_REP & "REDBULL_SUIVI_GLOBAL.xlsx")
    
    DDE_SEND "Fin génération suivi global frigos"
End Sub

' ========== RAPPORTS ==========
Public Sub FRIGO_RAPPORT_STOCK()
    DDE_SEND "Début génération rapport stock"
    
    Set Query = New cSQL_Query
    Query.SQL_User = "sa"
    Query.SQL_Pswd = "sadmin"
    Query.SQL_Host = "192.168.9.12"
    Query.SQL_Base = "REDBULL_FRIGOS"
    Query.SQL_Type = SQL_Type_MsSQL
    
    Call GenererRapportStock(WRK_REP & "RAPPORT_STOCK.xlsx")
    
    DDE_SEND "Fin génération rapport stock"
End Sub

Public Sub FRIGO_RAPPORT_PIECES()
    DDE_SEND "Début génération rapport pièces"
    
    Set Query = New cSQL_Query
    Query.SQL_User = "sa"
    Query.SQL_Pswd = "sadmin"
    Query.SQL_Host = "192.168.9.12"
    Query.SQL_Base = "REDBULL_FRIGOS"
    Query.SQL_Type = SQL_Type_MsSQL
    
    Call GenererRapportPieces(WRK_REP & "RAPPORT_PIECES.xlsx")
    
    DDE_SEND "Fin génération rapport pièces"
End Sub
