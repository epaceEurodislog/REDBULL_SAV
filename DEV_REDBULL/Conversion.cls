VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Conversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const CP_UTF8 = 65001
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

Public Function Date_to_SQL(d As Date) As String

Date_to_SQL = Format(Year(d), "0000") & Format(Month(d), "00") & Format(Day(d), "00")

End Function


Public Function SQL_to_Date(t As Variant) As Date

On Error Resume Next
If Len(t) = 8 Then
    If IsNumeric(t) Then
        SQL_to_Date = CDate(Right(t, 2) & "/" & Mid(t, 5, 2) & "/" & Left(t, 4))
    End If
End If
On Error GoTo 0
End Function

Public Function ASCII_To_UTF8(ByVal wText As String) As String
    Dim vNeeded As Long
    Dim vSize   As Long
    vSize = Len(wText)
    vNeeded = WideCharToMultiByte(CP_UTF8, 0, StrPtr(wText), vSize, "", 0, 0, 0)
    ASCII_To_UTF8 = String(vNeeded, 0)
    WideCharToMultiByte CP_UTF8, 0, StrPtr(wText), vSize, ASCII_To_UTF8, vNeeded, 0, 0
End Function

Public Function UTF8_To_ASCII(ByVal wText As String) As String
    Dim vNeeded As Long
    Dim vSize   As Long
    vSize = Len(wText)
    vNeeded = MultiByteToWideChar(CP_UTF8, 0, wText, vSize, 0, 0)
    UTF8_To_ASCII = String(vNeeded, 0)
    MultiByteToWideChar CP_UTF8, 0, wText, vSize, StrPtr(UTF8_To_ASCII), vNeeded
End Function

