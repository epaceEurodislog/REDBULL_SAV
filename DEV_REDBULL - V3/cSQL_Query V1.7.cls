VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSQL_Query"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Version 1.6

Private SQL_Cn As ADODB.Connection
Private SQL_Rc As ADODB.Recordset

Private m_Host As String
Private m_Base As String
Private m_Table As String

Private m_Comma_Accepted As Boolean

Private m_Insert_Base As String
Private m_Insert_Table As String
Private m_Insert_Data() As String


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

Public Enum ENUM_SQL_TYPE
    SQL_Type_MsSQL = 0
    SQL_Type_MySQL = 1
End Enum

Public Enum SQL_Format
    Format_Texte = 0
    Format_Date = 1
    Format_Number = 2
End Enum

Public Enum ENUM_SQL_INSERT
    Insert_Into = 0
    Insert_Ignore = 1
End Enum

Private m_AppName As String
Private m_LocName As String

Private m_SQL_Connect As ENUM_SQL_CONNECT
Private m_Type As ENUM_SQL_TYPE

Event Download(Count As Long)
Event QueryState(Texte As String)


Private l_Update() As String
Private l_Insert() As String


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

Public Property Get SQL_Type() As ENUM_SQL_TYPE
    SQL_Type = m_Type
End Property

Public Property Let SQL_Type(ByVal Value As ENUM_SQL_TYPE)
    m_Type = Value
End Property

Public Property Get SQL_Host() As String
    SQL_Host = m_Host
End Property

Public Property Let SQL_Host(ByVal Value As String)
    m_Host = Value
End Property

Public Property Get SQL_Base() As String
    SQL_Base = m_Base
End Property

Public Property Let SQL_Base(ByVal Value As String)
    m_Base = Value
End Property

Public Property Get SQL_User() As String
    SQL_User = m_User
End Property

Public Property Let SQL_User(ByVal Value As String)
    m_User = Value
End Property

Public Property Get SQL_Pswd() As String
    SQL_Pswd = m_Pswd
End Property

Public Property Let SQL_Pswd(ByVal Value As String)
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

Public Property Get SQL_Table() As String
    SQL_Table = m_Table
End Property

Public Property Let SQL_Table(ByVal Value As String)
    m_Table = Value
End Property


Public Property Get SQL_Comma_Accepted() As Boolean
    SQL_Comma_Accepted = m_Comma_Accepted
End Property

Public Property Let SQL_Comma_Accepted(ByVal Value As Boolean)
    m_Comma_Accepted = Value
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
'    CA_Dlr.sqlFile = m_Table
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

SQL_Connection
SQL_Cn.BeginTrans

End Sub

Public Function TRANS_EXECUTE(Commande) As Recordset

Set TRANS_EXECUTE = SQL_Cn.Execute(Commande)

End Function

Public Sub TRANS_ROLLBACK()

SQL_Cn.RollbackTrans
SQL_Deconnection

End Sub

Public Sub TRANS_COMMIT()

SQL_Cn.CommitTrans
SQL_Deconnection

End Sub

Public Function SQL_Get_Query(Optional Commande = "") As Variant

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
    RaiseEvent QueryState("Connexion à SQL...")
    DoEvents
    
    SQL_Connection
End If

If Commande = "" Then

    RstSelect = "Select " & m_Select
    RstWhere = ""
    RstGroupBy = ""
    RstHaving = ""
    RstOrderBy = ""
    
    If m_Where <> "" Then RstWhere = "Where " & m_Where
    If m_GroupBy <> "" Then RstGroupBy = "Group By " & m_GroupBy
    If m_Having <> "" Then RstHaving = "Having " & m_Having
    If m_OrderBy <> "" Then RstOrderBy = "Order By " & m_OrderBy

    Commande = RstSelect & "  FROM " & Replace(m_Table, "/", ".") & " " & RstWhere & " " & RstGroupBy & " " & RstHaving & " " & RstOrderBy
End If

RaiseEvent QueryState("Envoi de la requête...")
DoEvents


m_Count = 0

If m_Count = 123456789 Then
    Exit Function
End If

Debug.Print Commande

ddd = Commande

SQL_Rc.Open Commande

t = Timer

Do While SQL_Rc.EOF = False
    ReDim Preserve Table(0 To SQL_Rc.Fields.Count - 1, 0 To m_Count + 1)
    
    For i = 0 To SQL_Rc.Fields.Count - 1
        If IsNull(SQL_Rc(i)) = False Then
            
            Table(i, m_Count) = SQL_Rc(i)
        Else
            Table(i, m_Count) = ""
        End If
        
'        Table(i, m_Count) = SPEED_Rc(i)
    Next i
    m_Count = m_Count + 1
    
'    RaiseEvent Download(m_Count)
'    DoEvents
    
    If Timer - t > 0.5 Then
        RaiseEvent QueryState(CStr(m_Count) & " lignes transférées")
        DoEvents
        
        t = Timer
    End If
    
    SQL_Rc.MoveNext
'    DoEvents
Loop

If m_Count >= 1 Then
    ReDim Preserve Table(0 To SQL_Rc.Fields.Count - 1, 0 To m_Count - 1)
Else
    ReDim Preserve Table(0 To SQL_Rc.Fields.Count - 1, 0 To 0)
End If

SQL_Rc.Close
SQL_Get_Query = Table

R:
Select Case Err
     Case Is <> 0: MsgBox Err.Description, vbCritical, Err.Number
End Select

If m_SQL_Connect = SQL_Cnx_Auto Then
    RaiseEvent QueryState("Déconnexion en cours...")
    DoEvents
    
    SQL_Deconnection
    
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

Public Function SQL_Connection() As Integer

'Set CA_Sys = New cwbx.sqlSystem
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

Set SQL_Cn = CreateObject("ADODB.Connection")
'    SPEED_Cn.Open "provider=IBMDA400;data source=" & CA_Sys.SystemName & "", "" & m_User & "", "" & m_Pswd & ""
'SPEED_Cn.ConnectionString = "Provider=SQLNCLI.1;Integrated Security=SSPI;" & _
        "Persist Security Info=False;" & _
        "AttachDBFileName=\\192.168.9.11\e$\SQLServer\DB\SPEEDTEST\speed_data.mdf;Data Source=192.168.9.11\speedtest"

'SPEED_Cn.ConnectionString = "Provider=SQLOLEDB;Data Source=192.168.9.17;Initial Catalog=Speed; User ID=BKS;Password=eurobks;"
'''''SPEED_Cn.ConnectionString = "Provider=SQLOLEDB;SERVER=192.168.9.8;Initial Catalog=SPEED1;UID=eurodislog;PWD=euro1;APP=Microsoft Office 2003;WSID=SMINOTTI;Trusted_Connection=No"
'SPEED_Cn.ConnectionString = "Provider=SQLOLEDB;SERVER=192.168.9.8;Initial Catalog=speed1;UID=eurodislog;PWD=euro1;APP=Microsoft Office 2003;WSID=SMINOTTI;Trusted_Connection=Yes"


'& " from " & HostName
'SPEED_Cn.ConnectionString = "Provider=SQLOLEDB;SERVER=" & m_Host & ";Initial Catalog=" & m_Base & ";UID=" & m_User & ";PWD=" & m_Pswd & ";APP=Microsoft Office 2003;WSID=SMINOTTI;Trusted_Connection=No"

If m_Type = SQL_Type_MsSQL Then
    SQL_Cn.ConnectionString = "Provider=SQLOLEDB;SERVER=" & m_Host & ";Initial Catalog=" & m_Base & ";UID=" & m_User & ";PWD=" & m_Pswd & ";APP=" & m_LocName & "." & m_AppName & ";WSID=SMINOTTI;Trusted_Connection=No"
ElseIf m_Type = SQL_Type_MySQL Then
    SQL_Cn.ConnectionString = "Driver={MySQL ODBC 3.51 Driver};Server=" & m_Host & ";Port=3306;Database=" & m_Base & ";User=" & m_User & ";Password=" & m_Pswd & ";APP=Requête MySQL;"
End If

SQL_Cn.Open
Set SQL_Rc = CreateObject("ADODB.recordset")
Set SQL_Rc.ActiveConnection = SQL_Cn

SQL_Cn.CommandTimeout = 240
SQL_Connection = 1

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

Public Sub SQL_Deconnection()

On Error Resume Next

SQL_Rc.Close
SQL_Cn.Close

Set SQL_Rc = Nothing
Set SQL_Cn = Nothing

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

m_Comma_Accepted = False

End Sub

Private Sub Class_Terminate()
    SQL_Deconnection
End Sub

Public Function SQL_Insert(Insert_Table, Insert_Var, Insert_Val)



Q = "INSERT INTO " & Insert_Table & "(" & Insert_Var & ") VALUES(" & Insert_Val & ")"













''''Dim Table() As String
'''
''''Si la liste descriptive de table n'est pas à jour : MAJ
'''If m_Insert_Table <> m_Table Or m_Insert_Base <> m_Base Then
'''
'''    If m_SQL_Connect = SQL_Cnx_Auto Then
'''        RaiseEvent QueryState("Connexion à SQL...")
'''        DoEvents
'''
'''        SQL_Connection
'''    End If
'''
'''    qTxt = "SELECT column_name 'Column Name',data_type 'Data Type',character_maximum_length 'Maximum Length'"
'''    qTxt = qTxt & " FROM information_schema.Columns"
'''    qTxt = qTxt & " WHERE table_name = '" & m_Table & "' and (table_schema='gestion_personnel'"
'''    qTxt = qTxt & " or table_catalog='" & m_Base & "')"
'''
'''    SQL_Rc.Open qTxt
'''
'''
'''    t = Timer
'''    Do While SQL_Rc.EOF = False
'''        ReDim Preserve m_Insert_Data(0 To SQL_Rc.Fields.Count - 1, 0 To m_Count + 1)
'''
'''        For i = 0 To SQL_Rc.Fields.Count - 1
'''            If IsNull(SQL_Rc(i)) = False Then
'''
'''                m_Insert_Data(i, m_Count) = SQL_Rc(i)
'''            Else
'''                m_Insert_Data(i, m_Count) = ""
'''            End If
'''        Next i
'''        m_Count = m_Count + 1
'''
'''    '    RaiseEvent Download(m_Count)
'''    '    DoEvents
'''
'''        If Timer - t > 1 Then
'''            RaiseEvent QueryState(CStr(m_Count) & " lignes transférées")
'''            DoEvents
'''
'''            t = Timer
'''        End If
'''
'''        SQL_Rc.MoveNext
'''    Loop
'''
'''    If m_Count >= 1 Then
'''        ReDim Preserve m_Insert_Data(0 To SQL_Rc.Fields.Count - 1, 0 To m_Count - 1)
'''
'''        m_Insert_Table = m_Table
'''        m_Insert_Base = m_Base
'''    Else
'''        ReDim Preserve m_Insert_Data(0 To SQL_Rc.Fields.Count - 1, 0 To 0)
'''    End If
'''
'''    SQL_Rc.Close
'''
'''    If m_SQL_Connect = SQL_Cnx_Auto Then
'''        RaiseEvent QueryState("Déconnexion en cours...")
'''        DoEvents
'''
'''        SQL_Deconnection
'''
'''        RaiseEvent QueryState("Déconnecté")
'''        DoEvents
'''    End If
'''End If
'''
'''
'''t = ""
'''
'''i = 1
'''Do While i <= UBound(m_Insert_Data, 2)
'''    t = t & m_Insert_Data(0, i) & ","
'''
'''    i = i + 1
'''Loop


'Make = "INSERT INTO " & m_Table & "(" & tChamps & ") VALUES(" & tValeurs & ")"


End Function


Public Function SQL_Get_Columns(Optional q_Base = "", Optional q_Table = "") As Variant


If q_Base = "" Then q_Base = m_Base
If q_Table = "" Then q_Table = m_Table


m_Insert_Table = ""

'Si la liste descriptive de table n'est pas à jour : MAJ
If m_Insert_Table <> q_Table Or m_Insert_Base <> q_Base Then

    If m_SQL_Connect = SQL_Cnx_Auto Then
        RaiseEvent QueryState("Connexion à SQL...")
        DoEvents

        SQL_Connection
    End If

    qTxt = "SELECT column_name 'Column Name',data_type 'Data Type',character_maximum_length 'Maximum Length',column_comment 'Commentaires'"
    qTxt = qTxt & " FROM information_schema.Columns"
    qTxt = qTxt & " WHERE table_name = '" & q_Table & "' and (table_schema='" & MySQL_Base & "'"
    qTxt = qTxt & " or table_catalog='" & q_Base & "')"

    Debug.Print qTxt

    m_Count = 0
    SQL_Rc.Open qTxt

    t = Timer
    Do While SQL_Rc.EOF = False
        ReDim Preserve m_Insert_Data(0 To SQL_Rc.Fields.Count - 1, 0 To m_Count + 1)

        For i = 0 To SQL_Rc.Fields.Count - 1
            If IsNull(SQL_Rc(i)) = False Then

                m_Insert_Data(i, m_Count) = SQL_Rc(i)
            Else
                m_Insert_Data(i, m_Count) = ""
            End If
        Next i
        m_Count = m_Count + 1

    '    RaiseEvent Download(m_Count)
    '    DoEvents

        If Timer - t > 1 Then
            RaiseEvent QueryState(CStr(m_Count) & " lignes transférées")
            DoEvents

            t = Timer
        End If

        SQL_Rc.MoveNext
    Loop

    If m_Count >= 1 Then
        ReDim Preserve m_Insert_Data(0 To SQL_Rc.Fields.Count - 1, 0 To m_Count - 1)

        m_Insert_Table = q_Table
        m_Insert_Base = q_Base
    Else
        ReDim Preserve m_Insert_Data(0 To SQL_Rc.Fields.Count - 1, 0 To 0)
    End If

    SQL_Rc.Close

    If m_SQL_Connect = SQL_Cnx_Auto Then
        RaiseEvent QueryState("Déconnexion en cours...")
        DoEvents

        SQL_Deconnection

        RaiseEvent QueryState("Déconnecté")
        DoEvents
    End If
End If


SQL_Get_Columns = m_Insert_Data

t = ""

i = 1
Do While i <= UBound(m_Insert_Data, 2)
    t = t & m_Insert_Data(0, i) & ","

    i = i + 1
Loop


End Function



Public Function SQL_Set_Query(Commande_SQL) As Variant

Dim Table() As String

Dim i As Integer
Dim RstSelect As String
Dim RstWhere As String
Dim RstGroupBy As String
Dim RstHaving As String
Dim RstOrderBy As String
Err.Clear

If m_SQL_Connect = SQL_Cnx_Auto Then
    RaiseEvent QueryState("Connexion à SQL...")
    DoEvents
    
    SQL_Connection
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

SQL_Cn.Execute Commande_SQL, n

SQL_Set_Query = n

R:
Select Case Err
     Case Is <> 0: MsgBox Err.Description, vbCritical, Err.Number
End Select

If m_SQL_Connect = SQL_Cnx_Auto Then
    RaiseEvent QueryState("Déconnexion en cours...")
    DoEvents
    
    SQL_Deconnection

    RaiseEvent QueryState("Déconnecté")
    DoEvents
End If


End Function

Public Function SQL_Update_Init()

'm_Table = Insert_Table
ReDim l_Update(0 To 1, 0 To 0)

End Function

Public Function SQL_Update_Add(Champ, Valeur, Format As SQL_Format)

If l_Update(0, 0) = "" Then
    Idx = 0
Else
    Idx = UBound(l_Update, 2) + 1
    ReDim Preserve l_Update(0 To 1, 0 To Idx)
End If

l_Update(0, Idx) = Champ

If Format = Format_Date Then
    v = SQL_Date(Valeur)
ElseIf Format = Format_Number Then
    v = SQL_Number(Valeur)
ElseIf Format = Format_Texte Then
    v = SQL_Texte(Valeur)
End If

l_Update(1, Idx) = v


End Function

Public Function SQL_Update_Exc(Optional Update_Table = "", Optional Update_Where = "") As String

If Update_Table = "" And m_Table <> "" Then
    Update_Table = m_Table
ElseIf Update_Table = "" And m_Table = "" Then
    Exit Function
End If

If Update_Where = "" And m_Where <> "" Then
    Update_Where = m_Where
ElseIf Update_Where = "" And m_Where = "" Then
    Exit Function
End If

tChamps = ""

n = 0
Do While n <= UBound(l_Update, 2)
    tChamps = tChamps & l_Update(0, n) & "=" & l_Update(1, n) & ","
    n = n + 1
Loop

tChamps = Left(tChamps, Len(tChamps) - 1)

SQL_Cmd = "UPDATE " & Update_Table & " SET " & tChamps & " WHERE " & Update_Where
SQL_Set_Query SQL_Cmd

End Function


Public Function SQL_Insert_Init()

'm_Table = Insert_Table
ReDim l_Insert(0 To 1, 0 To 0)

End Function

Public Function SQL_Insert_Add(Champ, Valeur, Format As SQL_Format)

If l_Insert(0, 0) = "" Then
    Idx = 0
Else
    Idx = UBound(l_Insert, 2) + 1
    ReDim Preserve l_Insert(0 To 1, 0 To Idx)
End If

l_Insert(0, Idx) = Champ

If Format = Format_Date Then
    v = SQL_Date(Valeur)
ElseIf Format = Format_Number Then
    v = SQL_Number(Valeur)
ElseIf Format = Format_Texte Then
    v = SQL_Texte(Valeur)
End If

l_Insert(1, Idx) = v


End Function


Public Function SQL_Insert_Exc(Optional Insert_Table = "", Optional Insert_Base = "", Optional Insert_Option As ENUM_SQL_INSERT = Insert_Into) As String


If Insert_Table = "" And m_Table <> "" Then
    Insert_Table = m_Table
ElseIf Insert_Table = "" And m_Table = "" Then
    Exit Function
End If


tChamps = ""
tValeurs = ""

n = 0
Do While n <= UBound(l_Insert, 2)
    tChamps = tChamps & l_Insert(0, n) & ","
    tValeurs = tValeurs & l_Insert(1, n) & ","

    n = n + 1
Loop

tChamps = Left(tChamps, Len(tChamps) - 1)
tValeurs = Left(tValeurs, Len(tValeurs) - 1)

If Insert_Option = Insert_Into Then
    tInsert = "INSERT INTO"
ElseIf Insert_Option = Insert_Ignore Then
    tInsert = "INSERT IGNORE INTO"
End If

SQL_Cmd = tInsert & " " & Insert_Table & "(" & tChamps & ") VALUES(" & tValeurs & ")"

'Debug.Print SQL_Cmd
SQL_Set_Query SQL_Cmd


End Function

Private Function SQL_Date(Valeur) As String

vTemp = " " & Valeur
If Trim(vTemp) = "" Then
    SQL_Date = "Null"
Else
    If IsNumeric(Valeur) = True Then
        vTemp = Right(Valeur, 2) & "/" & Mid(Valeur, 5, 2) & "/" & Left(Valeur, 4)
    Else
        vTemp = Valeur
    End If


    t = Format(Year(vTemp), "0000")
    t = t & Format(Month(vTemp), "00")
    t = t & Format(Day(vTemp), "00")
    
    SQL_Date = "'" & t & "'"
End If

End Function

Private Function SQL_Texte(Valeur) As String

v = " " & Valeur
v = Trim(v)

If m_Comma_Accepted = False Then
    v = Replace(v, "'", "`")
Else
    v = Replace(v, "'", "''")
End If

SQL_Texte = "'" & v & "'"

End Function

Private Function SQL_Number(Valeur) As String

vTemp = " " & Valeur


If Trim(vTemp) = "" Then
    SQL_Number = 0
Else
    SQL_Number = Replace(Valeur, ",", ".")
End If

End Function
