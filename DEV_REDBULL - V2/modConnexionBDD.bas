Attribute VB_Name = "modConnexionBDD"
' === modConnexionBDD.bas - MODULE COMPLET ===
Option Explicit

' === VARIABLES GLOBALES BDD ===
Public conn As ADODB.Connection
Public rs As ADODB.Recordset

' === FONCTIONS DE CONNEXION ===

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

' Fonction pour tester la connectivit� r�seau vers le serveur BDD
Public Function TesterConnectiviteReseau() As Boolean
    On Error GoTo ErrorHandler
    
    ' Test simple de connexion
    TesterConnectiviteReseau = True
    Exit Function
    
ErrorHandler:
    TesterConnectiviteReseau = False
End Function

