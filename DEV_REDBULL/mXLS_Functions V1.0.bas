Attribute VB_Name = "mXLS_Functions"

Const xlDiagonalDown = 5
Const xlDiagonalUp = 6
Const xlEdgeBottom = 9
Const xlEdgeLeft = 7
Const xlEdgeRight = 10
Const xlEdgeTop = 8
Const xlInsideHorizontal = 12
Const xlInsideVertical = 11

Const xlContinuous = 1

'Membre de Excel.XlBorderWeight
Const xlThin = 2
Const xlThick = 4
Const xlMedium = -4138
Const xlHairline = 1

'Membre de Excel.Constants
Const xlAutomatic = -4105

Public Const xlCenter = -4108
Public Const xlLeft = -4131
Public Const xlRight = -4152

Public Function Creation_Cadre(Classeur As Object, Feuille, Cellules)

On Error Resume Next

Classeur.Sheets(Feuille).Range(Cellules).Borders(xlDiagonalDown).LineStyle = xlNone
Classeur.Sheets(Feuille).Range(Cellules).Borders(xlDiagonalUp).LineStyle = xlNone
With Classeur.Sheets(Feuille).Range(Cellules).Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .Weight = xlThin
    .ColorIndex = xlAutomatic
End With
With Classeur.Sheets(Feuille).Range(Cellules).Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Weight = xlThin
    .ColorIndex = xlAutomatic
End With
With Classeur.Sheets(Feuille).Range(Cellules).Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Weight = xlThin
    .ColorIndex = xlAutomatic
End With
With Classeur.Sheets(Feuille).Range(Cellules).Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .Weight = xlThin
    .ColorIndex = xlAutomatic
End With
With Classeur.Sheets(Feuille).Range(Cellules).Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .Weight = xlThin
    .ColorIndex = xlAutomatic
End With
With Classeur.Sheets(Feuille).Range(Cellules).Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .Weight = xlThin
    .ColorIndex = xlAutomatic
End With

On Error GoTo 0

End Function

Public Function Creation_Cadre_Large(Classeur As Object, Feuille, Cellules)

On Error Resume Next

Classeur.Sheets(Feuille).Range(Cellules).Borders(xlDiagonalDown).LineStyle = xlNone
Classeur.Sheets(Feuille).Range(Cellules).Borders(xlDiagonalUp).LineStyle = xlNone
With Classeur.Sheets(Feuille).Range(Cellules).Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .Weight = xlMedium
    .ColorIndex = xlAutomatic
End With
With Classeur.Sheets(Feuille).Range(Cellules).Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Weight = xlMedium
    .ColorIndex = xlAutomatic
End With
With Classeur.Sheets(Feuille).Range(Cellules).Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Weight = xlMedium
    .ColorIndex = xlAutomatic
End With
With Classeur.Sheets(Feuille).Range(Cellules).Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .Weight = xlMedium
    .ColorIndex = xlAutomatic
End With
'With Sheets(Feuille).Range(Cellules).Borders(xlInsideVertical)
'    .LineStyle = xlContinuous
'    .Weight = xlThin
'    .ColorIndex = xlAutomatic
'End With
'With Sheets(Feuille).Range(Cellules).Borders(xlInsideHorizontal)
'    .LineStyle = xlContinuous
'    .Weight = xlThin
'    .ColorIndex = xlAutomatic
'End With

On Error GoTo 0

End Function

Public Function Base_26(Nombre) As String

N1 = Int(Nombre / 26)
N2 = Nombre Mod 26

'A = 65, Nombre = 0

If N1 = 0 Then
    Base_26 = Chr(Nombre + 65)
Else
    Base_26 = Chr(N2 + 65)
    Base_26 = Chr(N1 + 64) & Base_26
End If

End Function

Public Function XLS_Creation_Cadre(Excel_Sheet, Cellules)

On Error Resume Next

Excel_Sheet.Range(Cellules).Borders(xlDiagonalDown).LineStyle = xlNone
Excel_Sheet.Range(Cellules).Borders(xlDiagonalUp).LineStyle = xlNone
With Excel_Sheet.Range(Cellules).Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .Weight = xlThin
    .ColorIndex = xlAutomatic
End With
With Excel_Sheet.Range(Cellules).Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Weight = xlThin
    .ColorIndex = xlAutomatic
End With
With Excel_Sheet.Range(Cellules).Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Weight = xlThin
    .ColorIndex = xlAutomatic
End With
With Excel_Sheet.Range(Cellules).Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .Weight = xlThin
    .ColorIndex = xlAutomatic
End With
With Excel_Sheet.Range(Cellules).Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .Weight = xlThin
    .ColorIndex = xlAutomatic
End With
With Excel_Sheet.Range(Cellules).Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .Weight = xlThin
    .ColorIndex = xlAutomatic
End With

On Error GoTo 0

End Function

Public Function XLS_Creation_Cadre_Large(Excel_Sheet, Cellules, Optional Couleur As Long = xlAutomatic)

On Error Resume Next

Excel_Sheet.Range(Cellules).Borders(xlDiagonalDown).LineStyle = xlNone
Excel_Sheet.Range(Cellules).Borders(xlDiagonalUp).LineStyle = xlNone
With Excel_Sheet.Range(Cellules).Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .Weight = xlMedium
    .ColorIndex = Couleur
    '.Color = Couleur
End With
With Excel_Sheet.Range(Cellules).Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Weight = xlMedium
    .ColorIndex = Couleur
End With
With Excel_Sheet.Range(Cellules).Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Weight = xlMedium
    .ColorIndex = Couleur
End With
With Excel_Sheet.Range(Cellules).Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .Weight = xlMedium
    .ColorIndex = Couleur
End With
'With Sheets(Feuille).Range(Cellules).Borders(xlInsideVertical)
'    .LineStyle = xlContinuous
'    .Weight = xlThin
'    .ColorIndex = xlAutomatic
'End With
'With Sheets(Feuille).Range(Cellules).Borders(xlInsideHorizontal)
'    .LineStyle = xlContinuous
'    .Weight = xlThin
'    .ColorIndex = xlAutomatic
'End With

On Error GoTo 0

End Function
