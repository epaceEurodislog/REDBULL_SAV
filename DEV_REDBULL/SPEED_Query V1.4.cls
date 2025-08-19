VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SPEED_Query"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Version 1.3

Private SPEED_Cn As ADODB.Connection
Private SPEED_Rc As ADODB.Recordset

Private m_Host As String
Private m_Base As String
Private m_Table As String

Private m_User As String
Private m_Pswd As String

Private m_Select As String
Private m_Where As String
Private m_Having As String
Private m_GroupBy As String
Private m_OrderBy As String

Private m_Count As Long
Private m_Filename As String

Public Enum ENUM_SQL_CONNECT
    SQL_Cnx_Auto = 0
    SQL_Cnx_Manuel = 1
End Enum


Private m_AppName As String
Private m_LocName As String

Private m_SQL_Connect As ENUM_SQL_CONNECT

Event Download(Count As Long)
Event QueryState(Texte As String)











' code à placer dans un module
Private Type WSADATA
        wVersion As Integer
        wHighVersion As Integer
        szDescription As String * 257
        szSystemStatus As String * 129
        iMaxSockets As Long
        iMaxUdpDg As Long
        lpVendorInfo As Long
End Type

' API
Private Declare Function gethostname Lib "wsock32.dll" (ByVal name As String, ByVal namelen As Long) As Long

Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequested As Integer, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long

'Public HostName As String

Private Function MAKEWORD(ByVal bLow As Byte, ByVal bHigh As Byte) As Integer
        MAKEWORD = Val("&H" & Right("00" & Hex(bHigh), 2) & Right("00" & Hex(bLow), 2))
End Function


Private Function Get_HostName() As String

    Dim slocalhostName As String
    Dim retval As Long
    Dim sockinfo As WSADATA

    ' ouvrir une Winsock session
    retval = WSACleanup()
    retval = WSAStartup(MAKEWORD(2, 2), sockinfo)
    If retval <> 0 Then
            MsgBox "Erreur sur chargement Winsock !"
            Exit Function
    End If

    slocalhostName = Space(256)
    retval = gethostname(slocalhostName, 256)
    Get_HostName = Left$(slocalhostName, InStr(slocalhostName, vbNullChar) - 1)

End Function

Public Property Get AppName() As String
    AppName = m_AppName
End Property

Public Property Let AppName(ByVal Value As String)
    m_AppName = Value
End Property

Public Property Get LocName() As String
    LocName = m_LocName
End Property

Public Property Let LocName(ByVal Value As String)
    m_LocName = Value
End Property

Public Property Get SQL_Connect() As ENUM_SQL_CONNECT
    SQL_Connect = m_SQL_Connect
End Property

Public Property Let SQL_Connect(ByVal Value As ENUM_SQL_CONNECT)
    m_SQL_Connect = Value
End Property

Public Property Get SQL_Count() As Long
    SQL_Count = m_Count
End Property

Public Property Get SPEED_Host() As String
    SPEED_Host = m_Host
End Property

Public Property Let SPEED_Host(ByVal Value As String)
    m_Host = Value
End Property

Public Property Get SPEED_Base() As String
    SPEED_Base = m_Base
End Property

Public Property Let SPEED_Base(ByVal Value As String)
    m_Base = Value
End Property

Public Property Get SPEED_User() As String
    SPEED_User = m_User
End Property

Public Property Let SPEED_User(ByVal Value As String)
    m_User = Value
End Property

Public Property Get SPEED_Pswd() As String
    SPEED_Pswd = m_Pswd
End Property

Public Property Let SPEED_Pswd(ByVal Value As String)
    m_Pswd = Value
End Property

Public Property Get SQL_Select() As String
    SQL_Select = m_Select
End Property

Public Property Let SQL_Select(ByVal Value As String)
    m_Select = Value
End Property

Public Property Get SQL_Where() As String
    SQL_Where = m_Where
End Property

Public Property Let SQL_Where(ByVal Value As String)
    m_Where = Value
End Property

Public Property Get SQL_Having() As String
    SQL_Having = m_Having
End Property

Public Property Let SQL_Having(ByVal Value As String)
    m_Having = Value
End Property

Public Property Get SQL_OrderBy() As String
    SQL_OrderBy = m_OrderBy
End Property

Public Property Let SQL_OrderBy(ByVal Value As String)
    m_OrderBy = Value
End Property

Public Property Get SQL_GroupBy() As String
    SQL_GroupBy = m_GroupBy
End Property

Public Property Let SQL_GroupBy(ByVal Value As String)
    m_GroupBy = Value
End Property

Public Property Get Filename() As String
    Filename = m_Filename
End Property

Public Property Let Filename(ByVal Value As String)
    m_Filename = Value
End Property

Public Property Get SPEED_Table() As String
    SPEED_Table = m_Table
End Property

Public Property Let SPEED_Table(ByVal Value As String)
    m_Table = Value
End Property

Private Function Check() As Boolean

Check = True

If App.EXEName = "Project1" Then
    m_AppName = "Project1"
End If

If m_AppName = "" Then
    MsgBox "Le nom d'application est vide, Veuillez le renseigner", vbOKOnly + vbCritical
    Check = False
End If

If m_LocName = "" Then
    MsgBox "Le nom local est vide, Veuillez le renseigner", vbOKOnly + vbCritical
    Check = False
End If

End Function





'Public Sub SPEED_Get_Query()
'
'On Error GoTo x
'Dim CA_Dlr As New cwbx.DatabaseDownloadRequest      'dlr
'Dim CA_PCo As cwbx.DatabaseDownloadPCOptions        'pcOpts
'Dim CA_DDA As cwbx.DatabaseDownloadASCIIOptions     'AsciiOpts
'Dim TextLine1, TextLine2
'Dim Fichier
'Dim P, i, Fil1, Fil2
'Dim Mig As Boolean
'
'Mig = False
'Set CA_Dlr.System = CA_Sys
'    CA_Dlr.SPEEDFile = m_Table
'
'If CExt(m_Filename) = ".TXT" Then
'    CA_Dlr.pcFile.FileType = cwbdtASCIIText
'    Set CA_DDA = CA_Dlr.pcFile.Options
'    'If Option1.Value = False Then
'        CA_DDA.ExistingFileOption = cwbdtOverwriteOutFile
'        CA_DDA.CodePage = cwbnlCodePageClientANSI
'        CA_DDA.TruncateSpaces = True
'    'Else
'    '    CA_DDA.ExistingFileOption = cwbdtAppendOutFile
'    'End If
'ElseIf CExt(m_Filename) = ".XLS" Then
'    CA_Dlr.pcFile.FileType = cwbdtDIF
'    Set CA_PCo = CA_Dlr.pcFile.Options
'        CA_PCo.ExistingFileOption = cwbdtOverwriteOutFile
'ElseIf CExt(m_Filename) = ".CSV" Then
'    Mig = True
'    CA_Dlr.pcFile.FileType = cwbdtCSV
'    Set CA_PCo = CA_Dlr.pcFile.Options
'        CA_PCo.ExistingFileOption = cwbdtOverwriteOutFile
'End If
'
'CA_Dlr.pcFile = m_Filename
'CA_Dlr.Query.Select = m_Select
'CA_Dlr.Query.Where = m_Where
'CA_Dlr.Query.GroupBy = m_GroupBy
'CA_Dlr.Query.Having = m_Having
'CA_Dlr.Query.OrderBy = m_OrderBy
'
'CA_Dlr.Download
'
'''migration
''If Mig = True Then
''    Fil1 = Len(CommonDialog1.Filename) - Len(CommonDialog1.FileTitle)
''    Fil2 = Len(CommonDialog1.FileTitle) - 4
''    Fichier = Left(CommonDialog1.Filename, Fil1) & Left(CommonDialog1.FileTitle, Fil2) & ".txt"
''    Open m_Filename For Input As #1
''    Open Fichier For Output As #2
''        Seek #2, 1
''
''        Do While Not EOF(1)
''            TextLine2 = ""
''            Line Input #1, TextLine1
''            'Debug.Print TextLine
''            LenLine = Len(TextLine1)
''
''            For i = 1 To LenLine + 1
''                P = Mid(TextLine1, i, 1)
''                If P = "," Then: P = ";"
''                If P = "." Then: P = ","
''                If P = Chr(34) Then: P = ""
''                TextLine2 = TextLine2 & P
''            Next
''            Print #2, TextLine2
''        Loop
''    Close #1
''    Close #2
''End If
'
'MsgBox "Transfert Terminer avec succès"
'
'x:
'    Select Case Err.Number
'        Case 32755: MsgBox "Transfert Annuler", vbCritical, "Erreur N° : " & Err.Number
'        Case -2147467259: MsgBox Err.Description, vbCritical, "Erreur N° : " & Err.Number
'        Case Is <> 0: MsgBox Err.Description, vbCritical, Err.Number
'    End Select
'
'End Sub


Public Sub TRANS_BEGIN()

SPEED_Connection
SPEED_Cn.BeginTrans

End Sub

Public Function TRANS_EXECUTE(Commande) As Recordset

Set TRANS_EXECUTE = SPEED_Cn.Execute(Commande)

End Function

Public Sub TRANS_ROLLBACK()

SPEED_Cn.RollbackTrans
SPEED_Deconnection

End Sub

Public Sub TRANS_COMMIT()

SPEED_Cn.CommitTrans
SPEED_Deconnection

End Sub

Public Function SPEED_Get_Query(Optional Requete = "") As Variant

If Check = False Then GoTo R

Dim Table() As String

Dim i As Integer
Dim RstSelect As String
Dim RstWhere As String
Dim RstGroupBy As String
Dim RstHaving As String
Dim RstOrderBy As String
Err.Clear

If m_SQL_Connect = SQL_Cnx_Auto Then
    RaiseEvent QueryState("Connexion à SPEED...")
    DoEvents
    
    SPEED_Connection
End If

RaiseEvent QueryState("Envoi de la requête...")
DoEvents

If Requete = "" Then
    RstSelect = "Select " & m_Select
    RstWhere = ""
    RstGroupBy = ""
    RstHaving = ""
    RstOrderBy = ""
    m_Count = 0
    
    If m_Where <> "" Then RstWhere = "Where " & m_Where
    If m_GroupBy <> "" Then RstGroupBy = "GroupBy " & m_GroupBy
    If m_Having <> "" Then RstHaving = "Having " & m_Having
    If m_OrderBy <> "" Then RstOrderBy = "Order By " & m_OrderBy

    Debug.Print RstSelect & "  FROM " & Replace(m_Table, "/", ".") & " " & RstWhere & " " & RstGroupBy & " " & RstHaving & " " & RstOrderBy
    SPEED_Rc.Open RstSelect & "  FROM " & Replace(m_Table, "/", ".") & " " & RstWhere & " " & RstGroupBy & " " & RstHaving & " " & RstOrderBy
Else

    m_Count = 0
    
    Debug.Print Requete
    SPEED_Rc.Open Requete

End If




'SPEED_Rc.e
Do While SPEED_Rc.EOF = False
    ReDim Preserve Table(0 To SPEED_Rc.Fields.Count - 1, 0 To m_Count + 1)
    
    For i = 0 To SPEED_Rc.Fields.Count - 1
        If IsNull(SPEED_Rc(i)) = False Then
            
            Table(i, m_Count) = SPEED_Rc(i)
        Else
            Table(i, m_Count) = ""
        End If
        
'        Table(i, m_Count) = SPEED_Rc(i)
    Next i
    m_Count = m_Count + 1
    
'    RaiseEvent Download(m_Count)
'    DoEvents
    
    RaiseEvent QueryState(CStr(m_Count) & " lignes transférées")
    DoEvents
    
    SPEED_Rc.MoveNext
'    DoEvents
Loop

If m_Count >= 1 Then
    ReDim Preserve Table(0 To SPEED_Rc.Fields.Count - 1, 0 To m_Count - 1)
Else
    ReDim Preserve Table(0 To SPEED_Rc.Fields.Count - 1, 0 To 0)
End If

SPEED_Rc.Close
SPEED_Get_Query = Table

R:
Select Case Err
     Case Is <> 0: MsgBox Err.Description, vbCritical, Err.Number
End Select

If m_SQL_Connect = SQL_Cnx_Auto Then
    RaiseEvent QueryState("Déconnexion en cours...")
    DoEvents
    
    SPEED_Deconnection
    
    RaiseEvent QueryState("Déconnecté")
    DoEvents
End If



''    '
''    ' Connect to the database
''    '
''    Dim cn As ADODB.Connection
''    Set cn = New Connection
''    cn.ConnectionString = "Provider=SQLNCLI.1;Integrated Security=SSPI;" & _
''        "Persist Security Info=False;" & _
''        "AttachDBFileName=" & App.Path & "\northwnd.mdf;Data Source=server1\sqlexpress"
''    cn.Open
''
''    '
''    ' Get all the orders
''    '
''    Dim rsOrders As Recordset
''    Set rsOrders = New Recordset
''    rsOrders.Open "Select * from orders", cn
''    Do While Not rsOrders.EOF
''        '
''        ' If the order matches some custom business logic then get the details for
''        ' that order, without opening a new connection.
''        '
''        If SomeBusinessLogic(rsOrders("CustomerID")) Then
''            Dim rsDetails As Recordset
''            Set rsDetails = New Recordset
''            '
''            ' Open a new recordset using the same connection. Normally it's not
''            ' possible to have two recordsets simultaniously using the same
''            ' connection, but MARS makes this possible
''            '
''            rsDetails.Open "Select sum(quantity * unitprice) as total " & _
''                "from [order details] " & _
''                "where OrderID=" & rsOrders("OrderID"), _
''                cn
''            grandTotal = grandTotal + rsDetails("total")
''        End If
''        rsOrders.MoveNext
''    Loop
''
''    lblTotalOrders = grandTotal

End Function

Public Function SPEED_Connection() As Integer

'Set CA_Sys = New cwbx.SPEEDSystem
'    CA_Sys.Define m_Host
'
'    CA_Sys.UserID = m_User
'    CA_Sys.Password = m_Pswd
'
'    'CA_Sys.Connect cwbcoServiceAll
'
'
'Set CA_Dtq.System = CA_Sys
'Set CA_Prg.System = CA_Sys
'Set CA_Cmd.System = CA_Sys

Set SPEED_Cn = CreateObject("ADODB.Connection")
'    SPEED_Cn.Open "provider=IBMDA400;data source=" & CA_Sys.SystemName & "", "" & m_User & "", "" & m_Pswd & ""
'SPEED_Cn.ConnectionString = "Provider=SQLNCLI.1;Integrated Security=SSPI;" & _
        "Persist Security Info=False;" & _
        "AttachDBFileName=\\192.168.9.11\e$\SQLServer\DB\SPEEDTEST\speed_data.mdf;Data Source=192.168.9.11\speedtest"

'SPEED_Cn.ConnectionString = "Provider=SQLOLEDB;Data Source=192.168.9.12;Initial Catalog=Speed; User ID=BKS;Password=eurobks;"
'''''SPEED_Cn.ConnectionString = "Provider=SQLOLEDB;SERVER=192.168.9.8;Initial Catalog=SPEED1;UID=eurodislog;PWD=euro1;APP=Microsoft Office 2003;WSID=SMINOTTI;Trusted_Connection=No"
'SPEED_Cn.ConnectionString = "Provider=SQLOLEDB;SERVER=192.168.9.8;Initial Catalog=speed1;UID=eurodislog;PWD=euro1;APP=Microsoft Office 2003;WSID=SMINOTTI;Trusted_Connection=Yes"


'& " from " & HostName
'SPEED_Cn.ConnectionString = "Provider=SQLOLEDB;SERVER=" & m_Host & ";Initial Catalog=" & m_Base & ";UID=" & m_User & ";PWD=" & m_Pswd & ";APP=Microsoft Office 2003;WSID=SMINOTTI;Trusted_Connection=No"
SPEED_Cn.ConnectionString = "Provider=SQLOLEDB;SERVER=" & m_Host & ";Initial Catalog=" & m_Base & ";UID=" & m_User & ";PWD=" & m_Pswd & ";APP=" & m_LocName & "." & m_AppName & ";WSID=SMINOTTI;Trusted_Connection=No"

''''SPEED_Cn.ConnectionString = "Provider=SQLNCLI10;SERVER=" & m_Host & ";Database=" & m_Base & ";UID=" & m_User & ";PWD=" & m_Pswd & ";"

SPEED_Cn.Open
Set SPEED_Rc = CreateObject("ADODB.recordset")
Set SPEED_Rc.ActiveConnection = SPEED_Cn

SPEED_Cn.CommandTimeout = 240
SPEED_Connection = 1

'''    Dim rsOrders As Recordset
'''    Set rsOrders = New Recordset
''    SPEED_Rc.Open "Select * from ACT_PAR", SPEED_Cn
''    Do While Not SPEED_Rc.EOF
''        '
''        ' If the order matches some custom business logic then get the details for
''        ' that order, without opening a new connection.
''        '
'''        If SomeBusinessLogic(rsOrders("CustomerID")) Then
'''            Dim rsDetails As Recordset
'''            Set rsDetails = New Recordset
'''            '
'''            ' Open a new recordset using the same connection. Normally it's not
'''            ' possible to have two recordsets simultaniously using the same
'''            ' connection, but MARS makes this possible
'''            '
'''            rsDetails.Open "Select sum(quantity * unitprice) as total " & _
'''                "from [order details] " & _
'''                "where OrderID=" & rsOrders("OrderID"), _
'''                cn
'''            grandTotal = grandTotal + rsDetails("total")
'''        End If
''        SPEED_Rc.MoveNext
''    Loop
''
''    aaa = SPEED_Rc.RecordCount
''
''    'lblTotalOrders = grandTotal

End Function

Public Sub SPEED_Deconnection()

On Error Resume Next

SPEED_Rc.Close
SPEED_Cn.Close

Set SPEED_Rc = Nothing
Set SPEED_Cn = Nothing

'Set CA_Dtq = Nothing
'Set CA_Prg = Nothing
'Set CA_Cmd = Nothing
'
'CA_Sys.Disconnect cwbcoServiceAll
'
'Set CA_Sys = Nothing
    
On Error GoTo 0
    
End Sub

Private Function CExt(File) As String

MyPos = 1

Do While MyPos > 0
MyPos = InStr(MyPos, File, ".", 1)
If MyPos > 0 Then
    LsPos = MyPos
    MyPos = MyPos + 1
End If
Loop

CExt = Left(File, LsPos - 1)
End Function

Private Sub Class_Initialize()

m_SQL_Connect = SQL_Cnx_Auto
LocName = Get_HostName

On Error Resume Next

AppName = ActiveWorkbook.name
AppName = App.ProductName

On Error Resume Next

End Sub

Private Sub Class_Terminate()
    SPEED_Deconnection
End Sub

Public Function SPEED_Set_Query(Commande_SQL) As Variant

Dim Table() As String

Dim i As Integer
Dim RstSelect As String
Dim RstWhere As String
Dim RstGroupBy As String
Dim RstHaving As String
Dim RstOrderBy As String
Err.Clear

If m_SQL_Connect = SQL_Cnx_Auto Then
    RaiseEvent QueryState("Connexion à SPEED...")
    DoEvents
    
    SPEED_Connection
End If

RstSelect = "Select " & m_Select
RstWhere = ""
RstGroupBy = ""
RstHaving = ""
RstOrderBy = ""
m_Count = 0

If m_Where <> "" Then RstWhere = "Where " & m_Where
If m_GroupBy <> "" Then RstGroupBy = "GroupBy " & m_GroupBy
If m_Having <> "" Then RstHaving = "Having " & m_Having
If m_OrderBy <> "" Then RstOrderBy = "Order By " & m_OrderBy

RaiseEvent QueryState("Envoi de la requête...")
DoEvents

Debug.Print Commande_SQL
SPEED_Cn.Execute Commande_SQL, n '"UPDATE TABLE SET CHAMP1='" & VALEUR1 & "', CHAMP2='" & VALEUR2 & "' WHERE CLE="


'SPEED_Rc.Open RstSelect & "  FROM " & Replace(m_Table, "/", ".") & " " & RstWhere & " " & RstGroupBy & " " & RstHaving & " " & RstOrderBy
'
'Do While SPEED_Rc.EOF = False
'    ReDim Preserve Table(0 To SPEED_Rc.Fields.Count - 1, 0 To m_Count + 1)
'
'    For i = 0 To SPEED_Rc.Fields.Count - 1
'        If IsNull(SPEED_Rc(i)) = False Then
'
'            Table(i, m_Count) = SPEED_Rc(i)
'        Else
'            Table(i, m_Count) = "NULL"
'        End If
'
''        Table(i, m_Count) = SPEED_Rc(i)
'    Next i
'    m_Count = m_Count + 1
'
''    RaiseEvent Download(m_Count)
''    DoEvents
'
'    RaiseEvent QueryState(CStr(m_Count) & " lignes transférées")
'    DoEvents
'
'    SPEED_Rc.MoveNext
''    DoEvents
'Loop
'
'If m_Count >= 1 Then
'    ReDim Preserve Table(0 To SPEED_Rc.Fields.Count - 1, 0 To m_Count - 1)
'Else
'    ReDim Preserve Table(0 To SPEED_Rc.Fields.Count - 1, 0 To 0)
'End If
qqq = n
'SPEED_Rc.Close
'SPEED_Get_Query = Table

R:
Select Case Err
     Case Is <> 0: MsgBox Err.Description, vbCritical, Err.Number
End Select

If m_SQL_Connect = SQL_Cnx_Auto Then
    RaiseEvent QueryState("Déconnexion en cours...")
    DoEvents
    
    SPEED_Deconnection

    RaiseEvent QueryState("Déconnecté")
    DoEvents
End If


End Function
