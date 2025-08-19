Attribute VB_Name = "Publique"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long ' only used if FOF_SIMPLEPROGRESS, sets dialog title
End Type

Public Const FO_COPY = &H2 ' Copy File/Folder
Public Const FO_DELETE = &H3 ' Delete File/Folder
Public Const FO_MOVE = &H1 ' Move File/Folder
Public Const FO_RENAME = &H4 ' Rename File/Folder
Public Const FOF_ALLOWUNDO = &H40 ' Allow to undo rename, delete ie sends to recycle bin
Public Const FOF_FILESONLY = &H80  ' Only allow files
Public Const FOF_NOCONFIRMATION = &H10  ' No File Delete or Overwrite Confirmation Dialog
Public Const FOF_SILENT = &H4 ' No copy/move dialog
Public Const FOF_SIMPLEPROGRESS = &H100 ' Does not display file names

Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
                        (lpFileOp As SHFILEOPSTRUCT) As Long

Public TAB_FTP() As String
'0 : actif/inactif (0/1)
'1 : Identifiant compte
'2 : Type (FTP/LAN)
'3 : Adresse IP
'4 : Login
'5 : Mot de passe

Public TAB_FIC() As String
'0 : actif/inactif
'1 : Connexion
'2 : Repertoire
'3 : Fichier
'4 : Programme
'5 : Paramètre


Public CSVZoneL As Boolean

Public Enum AlignType
    Droite = 1
    Gauche = 2
End Enum

Public Function SetLen(Champ, Longeur, Caractere, Alignement As AlignType)

Champ_Org = Champ
LC = Len(Champ)

If LC > Longeur Then
    SetLen = Left(Champ, Longeur)
Else
    If Alignement = 1 Then
        Do While Len(Champ) < Longeur
            Champ = Caractere + Champ
        Loop
        
        SetLen = Champ
    ElseIf Alignement = 2 Then
        Do While Len(Champ) < Longeur
            Champ = Champ + Caractere
        Loop
        
        SetLen = Champ
    End If
End If

Champ = Champ_Org

End Function

Public Function CSVZone(Ligne As String, Position As Long, Optional Separateur As String)
i = 1
MP = 0

If Separateur = "" Then Separateur = ";"

Do While i < Position
    MP = InStr(MP + 1, Ligne, Separateur, 1)
    i = i + 1
Loop

P1 = MP
P2 = InStr(P1 + 1, Ligne, Separateur, 1)

On Error Resume Next

CSVZone = Mid(Ligne, P1 + 1, P2 - P1 - 1)

If Err.Number = 5 Then
    CSVZone = ""
    On Error GoTo 0
    Exit Function
End If

On Error GoTo 0

i = 1
MP = 0

Do While i < 100
    MP = InStr(MP + 1, Ligne, Separateur, 1)
    
    If MP = 0 Then Exit Do
    
    i = i + 1
Loop

CSVCount = i - 1

If Position = i - 1 Then
    CSVZoneL = True
Else
    CSVZoneL = False
End If

End Function

'This will copy c:\backup to c:\backup2 and will not show filenames:
'
'Dim op As SHFILEOPSTRUCT
'With op
'    .wFunc = FO_COPY ' Set function
'    .pTo = "C:\backup2" ' Set new path
'    .pFrom = "C:\backup" ' Set current path
'    .fFlags = FOF_SIMPLEPROGRESS
'End With
'' Perform operation
'SHFileOperation op
'
'Not all functions require all the parameters. When you delete a file you do not need to specify the pTo parameter. This example sends the file c:\temp.txt to the recycle bin:
'
'Dim op As SHFILEOPSTRUCT
'With op
'    .wFunc = FO_DELETE ' Set function
'    .pFrom = "C:\temp.txt" ' Set File to delete
'    .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION ' Set Flags
'End With
'' Perform operation
'SHFileOperation op


Public Sub QuickSort(ByRef Arr() As String, _
    Optional ByVal lngLeft As Long = -2, _
    Optional ByVal lngRight As Long = -2, _
    Optional ByVal lngChamp As Long = -2, _
    Optional ShowProgress As Boolean = False)

    Dim i, j, lngMid As Long
    Dim vntTestVal As Variant
    
    t = 0
    
    If lngLeft = -2 Then lngLeft = LBound(Arr)
    If lngRight = -2 Then lngRight = UBound(Arr)

    If lngLeft < lngRight Then
        lngMid = (lngLeft + lngRight) \ 2
        
        If lngChamp = -2 Then
            vntTestVal = Arr(lngMid)
        Else
            vntTestVal = Arr(lngChamp, lngMid)
        End If
        
        i = lngLeft
        j = lngRight
        Do
            If lngChamp = -2 Then
                Do While Arr(i) < vntTestVal
                    i = i + 1
                Loop
                Do While Arr(j) > vntTestVal
                    j = j - 1
                Loop
            Else
                Do While Arr(lngChamp, i) < vntTestVal
                    i = i + 1
                Loop
                Do While Arr(lngChamp, j) > vntTestVal
                    j = j - 1
                Loop
            End If
            
            If i <= j Then
                Call SwapElements(Arr, i, j, lngChamp)
                i = i + 1
                j = j - 1
                
                 '--------------------------
                If ShowProgress = True Then
                    frmMain.DrawBackground
                    frmMain.Count_Quick_Sort = frmMain.Count_Quick_Sort + 1
                    frmMain.Count_Quick_Sort.Refresh
                    Sleep (1)
                End If
                '--------------------------
'                InfoMacro.Label2.Caption = StrN(t)
'                t = t + 1
'                DoEvents

            End If
        Loop Until i > j

        ' Optimize sort by sorting smaller segment first
        If j <= lngMid Then
            Call QuickSort(Arr, lngLeft, j, lngChamp, ShowProgress)
            Call QuickSort(Arr, i, lngRight, lngChamp, ShowProgress)
        Else
            Call QuickSort(Arr, i, lngRight, lngChamp, ShowProgress)
            Call QuickSort(Arr, lngLeft, j, lngChamp, ShowProgress)
        End If
    End If
    
End Sub


' Used in QuickSort function
Private Sub SwapElements(ByRef vntItems As Variant, ByVal lngItem1 As Long, ByVal lngItem2 As Long, ByVal lngChp As Long)

    Dim vntTemp As Variant

    If lngChp = -2 Then
        vntTemp = vntItems(lngItem2)
        vntItems(lngItem2) = vntItems(lngItem1)
        vntItems(lngItem1) = vntTemp
    Else
        LB = LBound(vntItems, 1)
        UB = UBound(vntItems, 1)
    
        i = LB
        Do While i <= UB
            vntTemp = vntItems(i, lngItem2)
            vntItems(i, lngItem2) = vntItems(i, lngItem1)
            vntItems(i, lngItem1) = vntTemp
        
            i = i + 1
        Loop
    End If
    
End Sub

Public Function GetISOWeek(ByVal vdInput As Date) As Long
    GetISOWeek = DatePart("ww", vdInput, vbMonday, vbFirstFourDays)
    If GetISOWeek >= 52 And DatePart("ww", vdInput + 7, vbMonday, vbFirstFourDays) = 2 Then
        GetISOWeek = 1
    End If
End Function

Public Function GetISOYear(ByVal vdInput As Date) As Long
    
'nWeek = DatePart("ww", vdInput, vbMonday, vbFirstFourDays)
'If nWeek >= 52 And DatePart("ww", vdInput + 7, vbMonday, vbFirstFourDays) = 2 Then
'    nWeek = 1
'End If
'
'Cond_Next = (nWeek = 1 And vdInput <= CDate("31/12/" & Format(Year(vdInput), "0000")))
'Cond_Next = Cond_Next And (vdInput >= CDate("25/12/" & Format(Year(vdInput), "0000")))
'
'If Cond_Next = True Then
'    GetISOYear = Year(vdInput) + 1
'Else
'    GetISOYear = Year(vdInput)
'End If
    
    
nWeek = GetISOWeek(vdInput)
    
If nWeek > 50 And Month(vdInput) = 1 Then
    GetISOYear = Year(vdInput) - 1
ElseIf nWeek = 1 And Month(vdInput) = 12 Then
    GetISOYear = Year(vdInput) + 1
Else
    GetISOYear = Year(vdInput)
End If

    
End Function




Public Function Week_to_Date(nWeek, nYear, iDate, oDate) As Boolean

d = CDate("01/01/" & CStr(nYear))
n = GetISOWeek(d)

Do While n > 1
    d = d + 1
    n = GetISOWeek(d)
Loop

j = Weekday(d, vbMonday)

d1 = d - (j - 1)
d2 = d1 + 6

iDate = d1 + 7 * (nWeek - 1)
oDate = d2 + 7 * (nWeek - 1)


End Function


Public Function Cal_Delai(d1 As Date, d2 As Date) As Long

Dim d As Date

Cal_Delai = 0
d = d1 + 1

Do While d <= d2
    If Jour_Ouvrable(d) = True Then
        Cal_Delai = Cal_Delai + 1
    End If
    
    d = d + 1
Loop

End Function

Public Function Cal_Date_Delai(d1 As Date, n) As Date

Dim d As Date

old = n
d = d1

If n > 0 Then
    Do While n > 0
        d = d + 1
    
        If Jour_Ouvrable(d) = True Then
            n = n - 1
        End If
    Loop
ElseIf n < 0 Then
    Do While n < 0
        d = d - 1
    
        If Jour_Ouvrable(d) = True Then
            n = n + 1
        End If
    Loop
End If

Cal_Date_Delai = d
n = old

End Function

Public Function Paques(Annee) ' détermination de la date de Pâques en fonction de l'année

    Dim var1, var2, var3, var4, var5, var6, var7
    var1 = Annee Mod 19 + 1
    var2 = (Annee \ 100) + 1
    var3 = ((3 * var2) \ 4) - 12
    var4 = (((8 * var2) + 5) \ 25) - 5
    var5 = ((5 * Annee) \ 4) - var3 - 10
    var6 = (11 * var1 + 20 + var4 - var3) Mod 30
    
    If (var6 = 25 And var1 > 11) Or (var6 = 24) Then
        var6 = var6 + 1
    End If
    
    var7 = 44 - var6
    
    If var7 < 21 Then
        var7 = var7 + 30
    End If
    
    var7 = var7 + 7
    var7 = var7 - (var5 + var7) Mod 7
    
    
    If var7 <= 31 Then
        Paques = DateValue(CStr(var7) & "/3/" & CStr(Annee))
    Else
        Paques = DateValue(CStr(var7 - 31) & "/4/" & CStr(Annee))
    End If

End Function

Public Function Jour_Ouvrable(Jour As Date) As Boolean

Dim DD As Date
Dim DF As Date
Dim DA As Date

Dim JO As Long

DD = Jour
DA = DD

DP1 = Paques(Year(DD)) + 1
DP2 = Paques(Year(DD)) + 39
DP3 = Paques(Year(DD)) + 50
Jour_Ouvrable = False

'JO = 0

If Weekday(DA, vbMonday) < 6 Then
    If Day(DA) = 1 And Month(DA) = 1 Then
    ElseIf DA = DP1 Then
    ElseIf Day(DA) = 1 And Month(DA) = 5 Then
    ElseIf Day(DA) = 8 And Month(DA) = 5 Then
    ElseIf DA = DP2 Then
    'ElseIf DA = DP3 Then
    ElseIf Day(DA) = 14 And Month(DA) = 7 Then
    ElseIf Day(DA) = 15 And Month(DA) = 8 Then
    ElseIf Day(DA) = 1 And Month(DA) = 11 Then
    ElseIf Day(DA) = 11 And Month(DA) = 11 Then
    ElseIf Day(DA) = 25 And Month(DA) = 12 Then
    Else
        Jour_Ouvrable = True
    End If
End If
    
End Function
