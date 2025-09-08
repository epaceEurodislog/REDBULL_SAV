Attribute VB_Name = "modConstantes"
' === modConstantes.bas ===
' Module contenant toutes les constantes et structures de données

Option Explicit

' === INFORMATIONS APPLICATION ===
Public Const VERSION_APP = "v2.1"
Public Const NOM_APP = "SAV Red Bull Scanner Pro"

' === CHEMINS FICHIERS ===
Public Const FICHIER_HISTORIQUE = "\HistoriqueScans.txt"
Public Const FICHIER_STOCK_PIECES = "\StockPieces.txt"
Public Const FICHIER_STOCK_REPARABLE = "\StockReparable.txt"

' === PARAMÈTRES BDD ===
Public Const SERVER_NAME As String = "192.168.9.12"
Public Const DATABASE_NAME As String = "SPEED_V6"
Public Const USERNAME As String = "eurodislog"
Public Const PASSWORD As String = "euro"

' === CODES ARTICLES AUTORISÉS ===
Public Const CODES_ARTICLES_AUTORISES As String = _
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

' === STRUCTURES DE DONNÉES ===

' Structure pour les données SAV
Public Type TypeSAV
    numeroEnlevement As String
    numeroReception As String
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

' Structure pour les pièces
Public Type TypePiece
    code As String
    Nom As String
    quantite As Integer
    etat As String
    origine As String
    dateAjout As Date
    prix As Double
End Type

' Structure pour les résultats de validation BDD
Public Type TypeValidationBDD
    existe As Boolean
    codeArticle As String
    designationArticle As String
    modeleArticle As String
    numeroSerie As String
    prixCatalogue As Double
    dateCreation As String
    statut As String
    informationsComplementaires As String
End Type

' Structure pour les données REE_DAT
Public Type TypeDonneesREE
    numeroReception As String  ' REE_Nore
    numeroEnlevement As String ' REE_Nofo
    trouve As Boolean
    messageErreur As String
End Type

' === ÉNUMÉRATIONS ===
Public Enum TypeStatutFrigo
    Reparable = 1
    HorsService = 2
    BonEtat = 3
    Obsolete = 4
End Enum

Public Enum TypeMotifRetour
    Mecanique = 1
    Esthetique = 2
    Mixte = 3
End Enum

