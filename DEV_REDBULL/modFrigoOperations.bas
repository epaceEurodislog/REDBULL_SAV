Attribute VB_Name = "modFrigoOperations"
Attribute VB_Name = "modFrigoOperations"

' Variables globales h�rit�es du formulaire principal
Public Declare Conv As Conversion
Public Declare Query As cSQL_Query

' ========== R�CEPTION DES FRIGOS ==========
Public Sub FRIGO_RECEPTION()
    DDE_SEND "D�but traitement r�ception frigos"
    
    ' Configuration de la base de donn�es
    Set Query = New cSQL_Query
    Query.SQL_User = "sa"
    Query.SQL_Pswd = "sadmin"
    Query.SQL_Host = "192.168.9.12"
    Query.SQL_Base = "REDBULL_FRIGOS"
    Query.SQL_Type = SQL_Type_MsSQL
    
    ' Lecture du fichier d'entr�e
    Call LireFichierReception(WRK_REP & WRK_FIC & ".txt")
    
    ' G�n�ration du rapport Excel
    Call GenererRapportReception(WRK_REP & "SORTIE_" & WRK_FIC & ".xlsx")
    
    DDE_SEND "Fin traitement r�ception frigos"
End Sub

' ========== DIAGNOSTIC FRIGORISTE ==========
Public Sub FRIGO_DIAGNOSTIC()
    DDE_SEND "D�but traitement diagnostic frigos"
    
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

' ========== R�PARATION ==========
Public Sub FRIGO_REPARATION()
    DDE_SEND "D�but traitement r�paration frigos"
    
    Set Query = New cSQL_Query
    Query.SQL_User = "sa"
    Query.SQL_Pswd = "sadmin"
    Query.SQL_Host = "192.168.9.12"
    Query.SQL_Base = "REDBULL_FRIGOS"
    Query.SQL_Type = SQL_Type_MsSQL
    
    Call TraiterReparations(WRK_REP & WRK_FIC & ".txt")
    Call GenererRapportReparation(WRK_REP & "SORTIE_" & WRK_FIC & ".xlsx")
    
    DDE_SEND "Fin traitement r�paration frigos"
End Sub

' ========== GESTION STOCK PI�CES ==========
Public Sub FRIGO_PIECES_STOCK()
    DDE_SEND "D�but traitement stock pi�ces"
    
    Set Query = New cSQL_Query
    Query.SQL_User = "sa"
    Query.SQL_Pswd = "sadmin"
    Query.SQL_Host = "192.168.9.12"
    Query.SQL_Base = "REDBULL_FRIGOS"
    Query.SQL_Type = SQL_Type_MsSQL
    
    Call MettreAJourStockPieces(WRK_REP & WRK_FIC & ".txt")
    Call GenererRapportStockPieces(WRK_REP & "SORTIE_" & WRK_FIC & ".xlsx")
    
    DDE_SEND "Fin traitement stock pi�ces"
End Sub

' ========== D�MONTAGE POUR PI�CES ==========
Public Sub FRIGO_DEMONTAGE()
    DDE_SEND "D�but traitement d�montage frigos"
    
    Set Query = New cSQL_Query
    Query.SQL_User = "sa"
    Query.SQL_Pswd = "sadmin"
    Query.SQL_Host = "192.168.9.12"
    Query.SQL_Base = "REDBULL_FRIGOS"
    Query.SQL_Type = SQL_Type_MsSQL
    
    Call TraiterDemontage(WRK_REP & WRK_FIC & ".txt")
    Call GenererRapportDemontage(WRK_REP & "SORTIE_" & WRK_FIC & ".xlsx")
    
    DDE_SEND "Fin traitement d�montage frigos"
End Sub

' ========== EXP�DITION ==========
Public Sub FRIGO_EXPEDITION()
    DDE_SEND "D�but traitement exp�dition frigos"
    
    Set Query = New cSQL_Query
    Query.SQL_User = "sa"
    Query.SQL_Pswd = "sadmin"
    Query.SQL_Host = "192.168.9.12"
    Query.SQL_Base = "REDBULL_FRIGOS"
    Query.SQL_Type = SQL_Type_MsSQL
    
    Call TraiterExpedition(WRK_REP & WRK_FIC & ".txt")
    Call GenererRapportExpedition(WRK_REP & "SORTIE_" & WRK_FIC & ".xlsx")
    
    DDE_SEND "Fin traitement exp�dition frigos"
End Sub

' ========== RETOURS ==========
Public Sub FRIGO_RETOUR()
    DDE_SEND "D�but traitement retours frigos"
    
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
    DDE_SEND "D�but g�n�ration suivi global frigos"
    
    Set Query = New cSQL_Query
    Query.SQL_User = "sa"
    Query.SQL_Pswd = "sadmin"
    Query.SQL_Host = "192.168.9.12"
    Query.SQL_Base = "REDBULL_FRIGOS"
    Query.SQL_Type = SQL_Type_MsSQL
    
    Call GenererSuiviGlobal(WRK_REP & "REDBULL_SUIVI_GLOBAL.xlsx")
    
    DDE_SEND "Fin g�n�ration suivi global frigos"
End Sub

' ========== RAPPORTS ==========
Public Sub FRIGO_RAPPORT_STOCK()
    DDE_SEND "D�but g�n�ration rapport stock"
    
    Set Query = New cSQL_Query
    Query.SQL_User = "sa"
    Query.SQL_Pswd = "sadmin"
    Query.SQL_Host = "192.168.9.12"
    Query.SQL_Base = "REDBULL_FRIGOS"
    Query.SQL_Type = SQL_Type_MsSQL
    
    Call GenererRapportStock(WRK_REP & "RAPPORT_STOCK.xlsx")
    
    DDE_SEND "Fin g�n�ration rapport stock"
End Sub

Public Sub FRIGO_RAPPORT_PIECES()
    DDE_SEND "D�but g�n�ration rapport pi�ces"
    
    Set Query = New cSQL_Query
    Query.SQL_User = "sa"
    Query.SQL_Pswd = "sadmin"
    Query.SQL_Host = "192.168.9.12"
    Query.SQL_Base = "REDBULL_FRIGOS"
    Query.SQL_Type = SQL_Type_MsSQL
    
    Call GenererRapportPieces(WRK_REP & "RAPPORT_PIECES.xlsx")
    
    DDE_SEND "Fin g�n�ration rapport pi�ces"
End Sub
