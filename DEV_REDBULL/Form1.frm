VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9930
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkMode        =   1  'Source
   LinkTopic       =   "System"
   ScaleHeight     =   6945
   ScaleWidth      =   9930
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer DDE_Timer 
      Interval        =   1000
      Left            =   1020
      Top             =   1530
   End
   Begin VB.TextBox TXT_REC 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1050
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1020
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Envoyer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3300
      TabIndex        =   1
      Top             =   570
      Width           =   1305
   End
   Begin VB.TextBox TXT_SND 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1050
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   540
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Query As SPEED_Query
Dim WithEvents Query As cSQL_Query
Attribute Query.VB_VarHelpID = -1
'Dim Conv As Conversion

Dim Sto_POA() As String

Dim WRK_REP As String
Dim WRK_PAR As String
Dim WRK_FIC As String

Dim Temp() As String
Dim Stock() As String

Dim appExcel As Object
Dim wbExcel As Object
Dim wsExcel As Object


Private Sub Command1_Click()

TXT_SND.LinkMode = vbLinkNone
TXT_SND.LinkTopic = "EAI_Eurodislog|System"
TXT_SND.LinkMode = vbLinkManual
TXT_SND.LinkExecute TXT_SND.Text

End Sub

Private Sub DDE_SEND(Texte)

Err.Clear
On Error Resume Next

If App.EXEName <> "Project1" Then
    TXT_SND = Texte
    
    TXT_SND.LinkMode = vbLinkNone
    TXT_SND.LinkTopic = "EAI_Eurodislog|System"
    TXT_SND.LinkMode = vbLinkManual
    TXT_SND.LinkExecute TXT_SND.Text
    
    Do While Err.Number <> 0
        Err.Clear
        TXT_SND.LinkMode = vbLinkNone
        TXT_SND.LinkTopic = "EAI_Eurodislog|System"
        TXT_SND.LinkMode = vbLinkManual
        TXT_SND.LinkExecute TXT_SND.Text
    Loop
End If

On Error GoTo 0

End Sub

Private Sub DDE_Timer_Timer()

DDE_SEND "PING"

End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)

TXT_REC = CmdStr
Cancel = False

End Sub

Private Sub Form_Load()

Dim dCal As Date

Set Query = New cSQL_Query

'************** REQUETE BASE PRODUCTION **************

'Query.SPEED_User = "eurodislog"
'Query.SPEED_Pswd = "euro"
'Query.SPEED_Host = "192.168.9.11"
'Query.SPEED_Base = "speedtest"
'
'Query.SPEED_Set_Query "UPDATE REE_DAT SET REE_TOP11=0 WHERE ACT_CODE='PMU' AND REE_TOP11=1" ' AND REE_NORE LIKE '%2%'"
'
'End
Dim Sh As SHFILEOPSTRUCT
Dim Cmd As String
Dim NFileI As Long
Dim NFileO As Long

Set Conv = New Conversion

Cmd = Command

WRK_REP = CSVZone(Cmd, 1, "|")
WRK_PAR = CSVZone(Cmd, 2, "|")
WRK_FIC = CSVZone(Cmd, 3, "|")

'MsgBox Cmd
'MsgBox WRK_REP
'MsgBox WRK_PAR
'MsgBox WRK_FIC


'
''WRK_REP = "C:\Visual Basic\Visual Basic 6.0\Projets VB\EAI Logidis\EAI\WORK\"
''WRK_REP = App.Path & "\"
''WRK_PAR = "REP_RECEP_PRESTAGO"
''WRK_FIC = "REP_RECEP_PRESTAGO"

''WRK_REP = "C:\Visual Basic\Visual Basic 6.0\Projets VB\EAI Logidis\EAI\WORK\"
''WRK_REP = App.Path & "\"
''WRK_PAR = "REP_SUIVI"
''WRK_FIC = "REP_SUIVI"

WRK_REP = "C:\Visual Basic\Visual Basic 6.0\Projets VB\EAI Logidis\EAI\WORK\"
WRK_REP = App.Path & "\"
WRK_PAR = "REDBULL_SUIVI"
WRK_FIC = "REDBULL_SUIVI"


'Concaténation fichiers / remplacement compteur ligne


If WRK_PAR = "NOVA_STOCK" Then

    NOVA_STOCK

ElseIf WRK_PAR = "NOVA_EXPE" Then

    NOVA_EXPE

ElseIf WRK_PAR = "REP_CDE_SAISIE" Then

    REP_CDE_SAISIE
    
ElseIf WRK_PAR = "CHECK_SPEEDTASK" Then

    CHECK_SPEEDTASK
    
ElseIf WRK_PAR = "REP_STK_PRESTAGO" Then
    REP_STK_PRESTAGO

ElseIf WRK_PAR = "REP_EXP_MENS_PRESTAGO" Then
    REP_EXP_MENS_PRESTAGO


ElseIf WRK_PAR = "REP_SUIVI" Then
    REP_SUIVI
    
ElseIf WRK_PAR = "REDBULL_SUIVI" Then
    REDBULL_SUIVI

Else
    DDE_SEND "Traitement " & WRK_PAR & " non défini"
End If

DDE_SEND "Fin de traitement"

End

End Sub

Public Function REP_RECEP_PRESTAGO()

'd = Date
d = CDate("01/01/2025")
'd1 = Date - 31

Query_Date = Format(Year(d - 1), "0000") & Format(Month(d - 1), "00") & Format(Day(d - 1), "00")
'Query_Date_Fin = Format(Year(d1), "0000") & Format(Month(d1), "00") & Format(Day(d1), "00")

Set Query = New cSQL_Query


Query.SQL_User = "Eurodislog"
Query.SQL_Pswd = "euro"
Query.SQL_Host = "192.168.9.12"
Query.SQL_Base = "SPEED_V6"


req = " select convert(date,ree.REE_DARE), ree.REE_NoRE, rel.ART_CODE, art.ART_DESL, rel.REL_LOT1, nse.NSE_NUMS,"
req = req & " case when art.art_num1>1 then art.art_num1 else case when art.art_qtes=999999999 then 1 else art.art_qtes end end,"
req = req & " sum( case when nse.NSE_NUMS is not null then 1 Else rel.REL_QTRE end)"

req = req & " from REE_DAT ree"

req = req & " left outer join REL_DAT rel on ree.ACT_CODE = rel.ACT_CODE And ree.REE_NoRE = rel.REE_NoRE"
req = req & " left outer join ART_PAR art on rel.ACT_CODE = art.ACT_CODE And rel.ART_CODE = art.ART_CODE"
req = req & " left outer join NSE_DAT nse on rel.ACT_CODE = nse.ACT_CODE And rel.REE_NoRE = nse.REE_NoRE And rel.REL_NORL = nse.REL_NORL"

req = req & " where ree.ACT_CODE = 'PRESTAGO' and ree.REE_DARE = '" & Query_Date & "'"
'Req = Req & " where ree.ACT_CODE = 'PRESTAGO' and ree.REE_DARE >= '" & Query_Date_Fin & "' and ree.REE_DARE >= '" & Query_Date_Deb & "'"

'**************** ADD RD 02/01/2025 *********************************************'
req = req & " and coalesce(rel.ART_CODE,'') <> ''"

req = req & " group by ree.REE_DARE, ree.REE_NoRE, rel.ART_CODE, art.ART_DESL, rel.REL_LOT1 , nse.NSE_NUMS, art.art_num1, art.art_qtes"

req = req & " order by ree.REE_DARE, ree.REE_NoRE, rel.ART_CODE, art.ART_DESL, rel.REL_LOT1 , nse.NSE_NUMS"

Table = Query.SQL_Get_Query(req)

If Query.SQL_Count > 0 Then

ReDim Temp(0 To (UBound(Table, 2) + 2), 0 To (UBound(Table, 1))) As String

Temp(0, 0) = "Date de réception"
Temp(0, 1) = "N° de réception"
Temp(0, 2) = "Code article"
Temp(0, 3) = "Désignation"
Temp(0, 4) = "N° de lot"
Temp(0, 5) = "N° de série"
Temp(0, 6) = "Conditionnement"
Temp(0, 7) = "Quantité"

i = 0
    Do While i <= UBound(Table, 1)
        j = 0
        Do While j <= UBound(Table, 2)
            Temp(j + 1, i) = Table(i, j)
            j = j + 1
        Loop
        i = i + 1
    Loop

n = 1
Do While n < UBound(Temp, 1)
    If Temp(n, 6) = 0 Then
        Temp(n, 6) = " "
    End If
        
    n = n + 1
Loop

    Set appExcel = CreateObject("Excel.application")
    Set wbExcel = appExcel.Workbooks.Add
    Set wsExcel = wbExcel.ActiveSheet

    sName = " Feuille1"

    On Error Resume Next
        wbExcel.Sheets("Feuil2").Delete
        wbExcel.Sheets("Feuil3").Delete
    On Error GoTo 0

    wbExcel.Sheets(1).name = sName

    NLf = UBound(Temp, 1) '+ 1
''
    wbExcel.Sheets(sName).Range("A2:A" & CStr(NLf)).NumberFormat = "dd/mm/yyyy;@"
    wbExcel.Sheets(sName).Range("C2:F" & CStr(NLf)).NumberFormat = "@"
    'wbExcel.Sheets(sName).Range("H2:I" & CStr(NLf)).NumberFormat = "0"

   
            ColF = Base_26(UBound(Temp, 2))
            NLf = UBound(Temp, 1)

            wbExcel.Sheets(sName).Range("A" & CStr(1)).Resize(UBound(Temp, 1) + 1, UBound(Temp, 2) + 1).Value = appExcel.transpose(appExcel.transpose(Temp))
    
            wbExcel.Sheets(sName).Range("A1:" & ColF & "1").Font.Bold = True
            wbExcel.Sheets(sName).Range("A1:" & ColF & "1").Interior.ThemeColor = 1
            wbExcel.Sheets(sName).Range("A1:" & ColF & "1").Interior.TintAndShade = -0.149998474074526
            wbExcel.Sheets(sName).Range("A1:" & ColF & "1").HorizontalAlignment = xlCenter

            wbExcel.Sheets(sName).Columns("A:" & ColF).EntireColumn.AutoFit
          '  Creation_Cadre wbExcel, sName, "A1:" & ColF & CStr(NLf)
            XLS_Creation_Cadre sName, "A1:" & ColF & CStr(NLf)

            wbExcel.Sheets(sName).Range("A" & CStr(NLf + 2)).Value = "Prestago_Etat quotidien des réceptions détaillés -" & Replace(CStr(Date), "/", "-")

        s = Split(appExcel.Version, ".")
        XLS_Version = CLng(s(0))
        XLS_Path = WRK_REP

        If Right(XLS_Path, 1) <> "\" Then
            XLS_Path = XLS_Path & "\"
        End If


        DateOK = Replace(CStr(Now), "/", "-")
        DateOK = Replace(CStr(DateOK), ":", "-")

        XLS_File = "Prestago_Etat quotidien des réceptions détaillés -" & DateOK
        XLS_File = XLS_Path & XLS_File & ".xlsx"
        XLS_Format = 51

        'On Error Resume Next

        wbExcel.SaveAs XLS_File, XLS_Format '56 ' 51
        wbExcel.Close ''True, Xls_File

        On Error GoTo 0

        appExcel.Quit
        Set wsExcel = Nothing
        Set wbExcel = Nothing
        Set appExcel = Nothing
    
End If


End Function

'Public Function REP_EXP_QUOT_PRESTAGO()
'
'd = Date
''d = CDate("27/10/2022")
'
'Query_Date = Format(Year(d - 1), "0000") & Format(Month(d - 1), "00") & Format(Day(d - 1), "00")
'
'Set Query = New SPEED_Query
'
'
'Query.SPEED_User = "Eurodislog"
'Query.SPEED_Pswd = "euro"
'Query.SPEED_Host = "192.168.9.12"
'Query.SPEED_Base = "SPEED_V6"
'
'
'req = " select mie.MIE_DAVL , art.ART_CODE, art.ART_DESL,case WHEN nse.nse_nums is not null THEN 1 Else Sum (mil.MIL_QTTP) end,"
'req = req & " nse.NSE_NUMS, stk.STK_LOT1,"
'req = req & " case when ope.ope_alpha8 not like ':CZ%' and ope.ope_alpha8 <> '' then substring(ope.ope_alpha8,charindex(':',ope.ope_alpha8,1)+1,"
'req = req & " charindex(':',ope.ope_alpha8,charindex(':',ope.ope_alpha8,1)+1)-(charindex(':',ope.ope_alpha8,1)+1))"
'req = req & " when charindex(':', ope.ope_alpha8, 2) > 0 then substring(ope.OPE_ALPHA8, 2, charindex(':', ope.ope_alpha8, 2)-2)"
'req = req & " Else '' end,ope.OPE_RTIE,"
'req = req & " case when charindex(';',ope.ope_alpha8,charindex(':',ope.ope_alpha8,charindex(':',ope.ope_alpha8)+1)+1)>0"
'req = req & " and charindex(':',ope.ope_alpha8,charindex(':',ope.ope_alpha8)+1)>0 then"
'req = req & " substring(ope.ope_alpha8,charindex(':',ope.ope_alpha8,charindex(':',ope.ope_alpha8)+1)+1,charindex(';',ope.ope_alpha8,charindex(':',ope.ope_alpha8,"
'req = req & " charindex(':',ope.ope_alpha8)+1)+1)-charindex(':',ope.ope_alpha8,charindex(':',ope.ope_alpha8)+1)-1)Else '' end,"
'req = req & " ope.OPE_ALPHA18, ope.TIE_NOM,"
'req = req & " case when art.art_num1>1 then art.art_num1 Else case when art.art_qtes=999999999 then 1 Else art.art_qtes End end"
'
'req = req & " from MIL_DAT mil"
'req = req & " left outer join MIE_DAT mie on mil.ACT_CODE = mie.ACT_CODE and mil.MIE_NOMI = mie.MIE_NOMI"
'req = req & " left outer join ART_PAR art on mil.ACT_CODE = art.ACT_CODE and mil.ART_CODE = art.ART_CODE"
'req = req & " left outer join OPE_DAT ope on mil.ACT_CODE = ope.ACT_CODE and mil.OPE_NoOE = ope.OPE_NoOE"
'req = req & " left outer join NSE_DAT nse on mil.ACT_CODE = nse.ACT_CODE and art.ART_CODE = nse.ART_CODE and ope.OPE_NoOE = nse.OPE_NoOE and mil.MIL_NoLM = nse.MIL_NoLM"
'req = req & " left outer join STK_DAT stk on mil.ACT_CODE = stk.ACT_CODE and art.ART_CODE = stk.ART_CODE"
'req = req & " Outer Apply (select top 1 sex_dtexp, sex_nooe from sex_dat where sex_nooe=OPE.OPE_NOOE and sex_act=OPE.ACT_CODE) as sex"
'
''Req = Req & " where art.ACT_CODE = 'Prestago' and sex.SEX_DTEXP >= '20220801' and sex.SEX_DTEXP <= '20220831'"
'req = req & " where art.ACT_CODE = 'Prestago' and mie.MIE_DAVL = '" & Query_Date & "' "
''Req = Req & " and coalesce(nse.nse_nums, '')<> ''"
'
''Req = Req & " group by art.ART_CODE, art.ART_DESL, mil.MIL_LOT1P, nse.NSE_NUMS, stk.STK_LOT1, ope.OPE_ALPHA8, sex.SEX_DTEXP, ope.OPE_RTIE, ope.OPE_NoVA, ope.OPE_ALPHA18, ope.TIE_NOM, art.art_num1, art.art_qtes"
'req = req & " group by art.ART_CODE, art.ART_DESL, mil.MIL_LOT1P, nse.NSE_NUMS, stk.STK_LOT1, ope.OPE_ALPHA8 , mie.MIE_DAVL, "
'req = req & " ope.OPE_RTIE , ope.OPE_NoVA, ope.OPE_ALPHA18, ope.TIE_NOM, art.art_num1, art.art_qtes"
''Req = Req & " order by sex.SEX_DTEXP;"
'req = req & " order by mie.MIE_DAVL asc;"
'
'Table = Query.SPEED_Get_Query(req)
'
'If Query.SQL_Count > 0 Then
'
'ReDim Temp(0 To (UBound(Table, 2) + 2), 0 To (UBound(Table, 1))) As String
'
'Temp(0, 0) = "Date d'expédition"
'Temp(0, 1) = "Code article"
'Temp(0, 2) = "Désignation"
'Temp(0, 3) = "Quantité"
'Temp(0, 4) = "Numéro de série"
'Temp(0, 5) = "Numéro de lot"
'Temp(0, 6) = "Centre de coût"
'Temp(0, 7) = "CMD PMU"
'Temp(0, 8) = "N° de vague"
'Temp(0, 9) = "Nom destinataire PDV"
'Temp(0, 10) = "Nom Installateur"
'Temp(0, 11) = "Conditionnement"
'
'i = 0
'    Do While i <= UBound(Table, 1)
'        j = 0
'        Do While j <= UBound(Table, 2)
'            Temp(j + 1, i) = Table(i, j)
'            j = j + 1
'        Loop
'        i = i + 1
'    Loop
'
'    Set appExcel = CreateObject("Excel.application")
'    Set wbExcel = appExcel.Workbooks.Add
'    Set wsExcel = wbExcel.ActiveSheet
'
'    sName = " Feuille1"
'
'    On Error Resume Next
'        wbExcel.sheets("Feuil2").Delete
'        wbExcel.sheets("Feuil3").Delete
'    On Error GoTo 0
'
'    wbExcel.sheets(1).name = sName
'
'    NLf = UBound(Temp, 1) '+ 1
'''
'    wbExcel.sheets(sName).Range("A2:A" & CStr(NLf)).NumberFormat = "dd/mm/yyyy;@"
'    wbExcel.sheets(sName).Range("C2:C" & CStr(NLf)).NumberFormat = "0"
'
'
'            ColF = Base_26(UBound(Temp, 2))
'            NLf = UBound(Temp, 1)
'
'            wbExcel.sheets(sName).Range("A" & CStr(1)).Resize(UBound(Temp, 1) + 1, UBound(Temp, 2) + 1).Value = appExcel.transpose(appExcel.transpose(Temp))
'
'            wbExcel.sheets(sName).Range("A1:" & ColF & "1").Font.Bold = True
'            wbExcel.sheets(sName).Range("A1:" & ColF & "1").Interior.ThemeColor = 1
'            wbExcel.sheets(sName).Range("A1:" & ColF & "1").Interior.TintAndShade = -0.149998474074526
'            wbExcel.sheets(sName).Range("A1:" & ColF & "1").HorizontalAlignment = xlCenter
'
'            wbExcel.sheets(sName).Columns("A:" & ColF).EntireColumn.AutoFit
'          '  Creation_Cadre wbExcel, sName, "A1:" & ColF & CStr(NLf)
'            XLS_Creation_Cadre sName, "A1:" & ColF & CStr(NLf)
'
'            wbExcel.sheets(sName).Range("A" & CStr(NLf + 2)).Value = "Prestago_Etat quotidien des expéditions détaillées -" & Replace(CStr(Date), "/", "-")
'
'        s = Split(appExcel.Version, ".")
'        XLS_Version = CLng(s(0))
'        XLS_Path = WRK_REP
'
'        If Right(XLS_Path, 1) <> "\" Then
'            XLS_Path = XLS_Path & "\"
'        End If
'
'
'        DateOK = Replace(CStr(Now), "/", "-")
'        DateOK = Replace(CStr(DateOK), ":", "-")
'
'        XLS_File = "Prestago_Etat quotidien des expéditions détaillées -" & DateOK
'        XLS_File = XLS_Path & XLS_File & ".xlsx"
'        XLS_Format = 51
'
'
'
'
'        'On Error Resume Next
'
'        wbExcel.SaveAs XLS_File, XLS_Format '56 ' 51
'        wbExcel.Close ''True, Xls_File
'
'        On Error GoTo 0
'
'        appExcel.Quit
'        Set wsExcel = Nothing
'        Set wbExcel = Nothing
'        Set appExcel = Nothing
'
'End If
'
'
'End Function


' Formatage des données (de la ligne 2 jusqu'à la dernière ligne)
If NLf > 0 Then  ' S'assurer qu'il y a des données
    wsExcel.Range("A2:A" & CStr(NLf + 1)).NumberFormat = "dd/mm/yyyy"  ' Date création
    wsExcel.Range("B2:D" & CStr(NLf + 1)).NumberFormat = "@"  ' Texte
    ' FORMATAGE CONFORME AU FICHIER DE RÉFÉRENCE
    wsExcel.Range("E2:E" & CStr(NLf + 1)).NumberFormat = "@"  ' N° interne (compteur séquentiel)
    wsExcel.Range("F2:F" & CStr(NLf + 1)).NumberFormat = "@"  ' N° externe
    wsExcel.Range("G2:J" & CStr(NLf + 1)).NumberFormat = "@"  ' Texte
    wsExcel.Range("K2:K" & CStr(NLf + 1)).NumberFormat = "@"  ' Code postal en TEXTE
    wsExcel.Range("L2:L" & CStr(NLf + 1)).NumberFormat = "dd/mm/yyyy"  ' Date livraison
    wsExcel.Range("M2:" & ColF & CStr(NLf + 1)).NumberFormat = "@"  ' Reste en texte
End If

Private Function REP_SUIVI()

Dim d As Date
Dim h As Date
Dim dt As String
Dim ht As String
Dim i As Long
Dim j As Long
Dim req As String
Dim Table As Variant
Dim Temp() As Variant
Dim appExcel As Object
Dim wbExcel As Object
Dim wsExcel As Object
Dim sName As String
Dim NLf As Long
Dim ColF As String
Dim s As Variant
Dim XLS_Version As Long
Dim XLS_Path As String
Dim XLS_File As String
Dim XLS_Format As Long
Dim nbCols As Long
Dim nbRows As Long
Dim maxSample As Long

' Vérification que Query est initialisé
If Query Is Nothing Then
    Set Query = New cSQL_Query
End If

' Configuration de la connexion SQL
Query.SQL_User = "sa"
Query.SQL_Pswd = "sadmin"
Query.SQL_Host = "192.168.9.12"
Query.SQL_Base = "SPEED_V6"

' Nouvelle requête SQL pour Bouygues
req = "SELECT OPE.OPE_CRDA, ope.ope_ccli, ope.ope_natu, "
req = req & "CASE "
req = req & "WHEN ope.ope_stat = '010' THEN 'EN SAISIE' "
req = req & "WHEN ope.ope_stat = '020' THEN 'EN VAGUE' "
req = req & "WHEN ope.ope_stat = '030' THEN 'EN PREPARATION' "
req = req & "WHEN ope.ope_stat = '040' THEN 'VALIDEE' "
req = req & "WHEN ope.ope_stat = '050' THEN 'ANNULER' "
req = req & "WHEN ope.ope_stat = '060' THEN 'MISE EN EXPEDITION' "
req = req & "WHEN ope.ope_stat = '070' THEN 'EXPEDIEE' "
req = req & "END AS Statut, "
req = req & "ope.ope_nooe, "
req = req & "ope.ope_redo, "
req = req & "ope.ope_rtie, OPE.TIE_NOM, "
req = req & "ope_adr1, ope_adr2, ope_adcp, ope_advl, "
req = req & "ope.ope_dali, ope.ope_alpha1, ope.ope_alpha2, "
req = req & "ope.ope_tel, ope.OPE_IMEL, ope.ope_ctra, "
req = req & "ope.ope_alpha13, tie.tie_nom, sex.sex_supe, "
req = req & "sex.sex_trak, sex.sex_urlt "
req = req & "FROM ope_dat AS ope "
req = req & "LEFT OUTER JOIN sex_dat AS sex ON ope.act_code=sex.sex_act AND ope.ope_nooe=sex.sex_nooe "
req = req & "LEFT OUTER JOIN tie_par AS tie ON ope.act_code = tie.act_code AND ope.ope_ctra = tie.tie_code "
req = req & "WHERE ope.act_code='CCESP' "
req = req & "ORDER BY ope.ope_nooe, sex.sex_supe ASC"

' Exécution de la requête
On Error GoTo ErrorHandler
Table = Query.SQL_Get_Query(req)
On Error GoTo 0

' Vérification que des données ont été retournées
If Query.SQL_Count = 0 Then
    MsgBox "Aucune donnée trouvée pour Bouygues"
    Exit Function
End If

' === DIAGNOSTIC ET CORRECTION FORCÉE POUR LES VRAIES VALEURS BDD ===
Debug.Print "=== DIAGNOSTIC TABLEAU SQL ==="
Debug.Print "UBound(Table, 1) = " & CStr(UBound(Table, 1))
Debug.Print "UBound(Table, 2) = " & CStr(UBound(Table, 2))

' Test des premières valeurs pour comprendre l'organisation ET VOIR LES VRAIES DONNÉES
Debug.Print "=== TEST ORIENTATION ET CONTENU RÉEL ==="
For i = 0 To IIf(UBound(Table, 1) > 2, 2, UBound(Table, 1))
    Debug.Print "=== LIGNE " & CStr(i) & " ==="
    For j = 0 To IIf(UBound(Table, 2) > 6, 6, UBound(Table, 2))  ' Afficher plus de colonnes
        On Error Resume Next
        Debug.Print "  Col " & CStr(j) & " = [" & CStr(Table(i, j)) & "] (type=" & TypeName(Table(i, j)) & ")"
        ' Spécialement pour ope_nooe (colonne 4 normalement)
        If j = 4 Then
            Debug.Print "    *** C'EST ope_nooe ! Valeur brute = " & CStr(Table(i, j))
            If IsNumeric(Table(i, j)) Then
                Debug.Print "    *** Converti en Double = " & CStr(CDbl(Table(i, j)))
            End If
        End If
        On Error GoTo 0
    Next j
Next i

' FORCER LA BONNE ORIENTATION : Si Table a plus de colonnes que de lignes,
' c'est probablement que les données sont transposées
If UBound(Table, 2) > UBound(Table, 1) Then
    Debug.Print "DÉTECTION: Les données sont transposées - Correction en cours..."
    ' Inverser les dimensions
    nbRows = UBound(Table, 2)  ' Les "colonnes" deviennent les lignes
    nbCols = UBound(Table, 1)  ' Les "lignes" deviennent les colonnes
    Debug.Print "Après correction: " & CStr(nbRows + 1) & " lignes, " & CStr(nbCols + 1) & " colonnes"
Else
    nbRows = UBound(Table, 1)
    nbCols = UBound(Table, 2)
    Debug.Print "Orientation normale: " & CStr(nbRows + 1) & " lignes, " & CStr(nbCols + 1) & " colonnes"
End If

' Vérification de cohérence
If nbCols < 22 Then
    MsgBox "Attention : La requête retourne seulement " & CStr(nbCols + 1) & " colonnes au lieu de 23 attendues"
End If

' === CRÉATION DYNAMIQUE DU TABLEAU FINAL ===
' Créer le tableau final avec en-têtes (23 colonnes) + nombre variable de lignes
ReDim Temp(0 To nbRows + 1, 0 To 22) As Variant
Debug.Print "Tableau Temp créé avec " & CStr(nbRows + 2) & " lignes (incluant en-tête) et 23 colonnes"

' En-têtes des colonnes
Temp(0, 0) = "Date de création de la commande dans Speed"
Temp(0, 1) = "Compte Client"
Temp(0, 2) = "Nature de la commande"
Temp(0, 3) = "Statut de la commande"
Temp(0, 4) = "N° interne"
Temp(0, 5) = "N° externe"
Temp(0, 6) = "ope_rtie"
Temp(0, 7) = "Nom du destinataire"
Temp(0, 8) = "Adresse 1"
Temp(0, 9) = "Adresse 2"
Temp(0, 10) = "Code postal"
Temp(0, 11) = "Ville"
Temp(0, 12) = "Date de livraison demandée"
Temp(0, 13) = "ope_alpha1"
Temp(0, 14) = "ope_alpha2"
Temp(0, 15) = "Téléphone"
Temp(0, 16) = "Mail"
Temp(0, 17) = "Transporteur Prévu"
Temp(0, 18) = "Transporteur Expédiée"
Temp(0, 19) = "Libellé du transporteur"
Temp(0, 20) = "Support"
Temp(0, 21) = "Tracking"
Temp(0, 22) = "Url de suivi"

' === REMPLISSAGE AVEC GESTION DE TRANSPOSITION ET FORMATAGE SPÉCIAL ===
Debug.Print "Début du remplissage des données..."

' REMPLISSAGE EN GÉRANT LES DEUX ORIENTATIONS POSSIBLES
For i = 0 To nbRows  ' Chaque ligne de données
    If i + 1 <= UBound(Temp, 1) Then
        For j = 0 To IIf(nbCols > 22, 22, nbCols)  ' Chaque colonne (max 23)
            On Error Resume Next
            
            Dim Valeur As Variant
            
            ' Si les données sont transposées dans Table
            If UBound(Table, 2) > UBound(Table, 1) Then
                ' Lire Table(colonne, ligne) pour écrire en Temp(ligne, colonne)
                Valeur = Table(j, i)
            Else
                ' Orientation normale Table(ligne, colonne)
                Valeur = Table(i, j)
            End If
            
            ' TRAITEMENT POUR RÉCUPÉRER LES VRAIES VALEURS DE LA BDD
            If IsNull(Valeur) Or IsEmpty(Valeur) Then
                Temp(i + 1, j) = ""
            Else
                ' Colonne 4 = N° interne (ope_nooe) - GARDER LA VRAIE VALEUR DE LA BDD
                If j = 4 Then
                    ' Forcer la conversion correcte pour les gros numéros
                    If IsNumeric(Valeur) Then
                        ' Convertir en Long puis en String pour éviter la troncature
                        Temp(i + 1, j) = CStr(CDbl(Valeur))  ' CDbl pour gérer les gros nombres
                        Debug.Print "VRAIE valeur ope_nooe ligne " & CStr(i + 1) & " = " & CStr(CDbl(Valeur))
                    Else
                        Temp(i + 1, j) = CStr(Valeur)
                        Debug.Print "ope_nooe non numérique ligne " & CStr(i + 1) & " = " & CStr(Valeur)
                    End If
                ' Colonne 5 = N° externe (ope_redo) - GARDER LA VRAIE VALEUR
                ElseIf j = 5 Then
                    If IsNumeric(Valeur) Then
                        Temp(i + 1, j) = CStr(CDbl(Valeur))  ' Pour les gros numéros aussi
                    Else
                        Temp(i + 1, j) = CStr(Valeur)
                    End If
                    Debug.Print "VRAIE valeur ope_redo ligne " & CStr(i + 1) & " = " & CStr(Valeur)
                ' Colonne 10 = Code postal (préserver format)
                ElseIf j = 10 Then
                    Temp(i + 1, j) = CStr(Valeur)
                ' Colonnes de dates (0 et 12)
                ElseIf j = 0 Or j = 12 Then
                    If IsDate(Valeur) Then
                        Temp(i + 1, j) = Format(CDate(Valeur), "dd/mm/yyyy")
                    Else
                        Temp(i + 1, j) = CStr(Valeur)
                    End If
                ' Autres colonnes - traiter les NULL
                Else
                    If CStr(Valeur) = "" Then
                        Temp(i + 1, j) = "NULL"
                    Else
                        Temp(i + 1, j) = CStr(Valeur)
                    End If
                End If
            End If
            On Error GoTo 0
        Next j
    End If
    
    ' Afficher progression
    If (i + 1) Mod 100 = 0 Then
        Debug.Print "Traitement ligne " & CStr(i + 1) & "/" & CStr(nbRows + 1)
    End If
Next i

Debug.Print "Remplissage terminé : " & CStr(nbRows + 1) & " lignes de données traitées"

' Génération des informations de date et heure
d = Date
h = Time

dt = SetLen(CStr(Day(d)), 2, "0", Droite) & "/"
dt = dt & SetLen(CStr(Month(d)), 2, "0", Droite) & "/"
dt = dt & CStr(Year(d))

ht = SetLen(CStr(Hour(h)), 2, "0", Droite) & ":"
ht = ht & SetLen(CStr(Minute(h)), 2, "0", Droite) & ":"
ht = ht & SetLen(CStr(Second(h)), 2, "0", Droite)

' === GÉNÉRATION DYNAMIQUE DU FICHIER EXCEL ===
On Error GoTo ErrorHandler

' Création de l'application Excel
Set appExcel = CreateObject("Excel.Application")
appExcel.Visible = False ' Masquer Excel pendant le traitement
Set wbExcel = appExcel.Workbooks.Add
Set wsExcel = wbExcel.ActiveSheet

sName = "Suivi Transport Bouygues"
' En VB6, il faut faire attention à la syntaxe pour renommer la feuille
wbExcel.Worksheets("Feuil1").name = sName

' Configuration du format des colonnes en fonction du nombre réel de lignes
NLf = UBound(Temp, 1)  ' Nombre total de lignes (en-tête + données)
ColF = Base_26(22)  ' 23 colonnes (A à W)

Debug.Print "=== CRÉATION DU FICHIER EXCEL ==="
Debug.Print "Nombre total de lignes dans Excel : " & CStr(NLf + 1)
Debug.Print "Plage de données : A1:" & ColF & CStr(NLf + 1)

' Insertion des données dans Excel (taille dynamique)
' En VB6, il faut être explicite avec les méthodes
Set wsExcel = wbExcel.Worksheets(sName)
wsExcel.Range("A1").Resize(NLf + 1, 23).Value = Temp

' Formatage de l'en-tête (toujours ligne 1)
With wsExcel.Range("A1:" & ColF & "1")
    .Font.Bold = True
    .Interior.ThemeColor = 1
    .Interior.TintAndShade = -0.149998474074526
    .HorizontalAlignment = -4108  ' xlCenter en VB6
End With

' Formatage des données (de la ligne 2 jusqu'à la dernière ligne)
If NLf > 0 Then  ' S'assurer qu'il y a des données
    wsExcel.Range("A2:A" & CStr(NLf + 1)).NumberFormat = "dd/mm/yyyy"  ' Date création
    wsExcel.Range("B2:D" & CStr(NLf + 1)).NumberFormat = "@"  ' Texte
    ' CORRECTION FORMATAGE POUR LES NUMÉROS
    wsExcel.Range("E2:E" & CStr(NLf + 1)).NumberFormat = "@"  ' N° interne en TEXTE
    wsExcel.Range("F2:F" & CStr(NLf + 1)).NumberFormat = "@"  ' N° externe en TEXTE
    wsExcel.Range("G2:J" & CStr(NLf + 1)).NumberFormat = "@"  ' Texte
    wsExcel.Range("K2:K" & CStr(NLf + 1)).NumberFormat = "@"  ' Code postal en TEXTE
    wsExcel.Range("L2:L" & CStr(NLf + 1)).NumberFormat = "dd/mm/yyyy"  ' Date livraison
    wsExcel.Range("M2:" & ColF & CStr(NLf + 1)).NumberFormat = "@"  ' Reste en texte
End If

' Auto-ajustement des colonnes
wsExcel.Columns("A:" & ColF).EntireColumn.AutoFit
Call Creation_Cadre(wbExcel, sName, "A1:" & ColF & CStr(NLf + 1))

' Ajout du titre du rapport (après la dernière ligne de données + 2 lignes)
wsExcel.Range("A" & CStr(NLf + 3)).Value = "Rapport de suivi transport Bouygues généré le " & CStr(dt) & " à " & CStr(ht) & " - " & CStr(nbRows + 1) & " enregistrements"

' Sauvegarde du fichier Excel
s = Split(appExcel.Version, ".")
XLS_Version = CLng(s(0))
XLS_Path = WRK_REP

If Right$(XLS_Path, 1) <> "\" Then
    XLS_Path = XLS_Path & "\"
End If

XLS_File = "Bouygues - Suivi Transport - " & Replace(dt, "/", "-") & " à " & Replace(ht, ":", "-")
XLS_File = XLS_Path & XLS_File & ".xlsx"
XLS_Format = 51

wbExcel.SaveAs XLS_File, XLS_Format
wbExcel.Close
appExcel.Quit

Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing

MsgBox "Fichier Excel généré avec succès : " & XLS_File
Exit Function

ErrorHandler:
    MsgBox "Erreur : " & Err.Description & " (Code: " & CStr(Err.Number) & ")" & vbCrLf & "Ligne: " & CStr(Erl)
    
    ' Afficher des informations de debug en cas d'erreur
    Debug.Print "Erreur lors du traitement"
    If Not IsEmpty(Temp) Then
        Debug.Print "Dimensions Temp: " & CStr(UBound(Temp, 1)) & " x " & CStr(UBound(Temp, 2))
    End If
    
    ' Nettoyage des objets Excel en cas d'erreur
    On Error Resume Next
    If Not wbExcel Is Nothing Then wbExcel.Close
    If Not appExcel Is Nothing Then appExcel.Quit
    Set wsExcel = Nothing
    Set wbExcel = Nothing
    Set appExcel = Nothing
    On Error GoTo 0

End Function
