Attribute VB_Name = "XML_Mod"
Private Const CP_UTF8 = 65001
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

'<XML Style>
'Font
'Size
'Bold
'Italic
'Format (texte, numérique...)
'Horizontal alignment
'Vertical alignment
'Fore Color
'Back Color
'T taille trait
'B taille trait
'L taille trait
'R taille trait

'<XML Cellule>
'Merge à droite
'Merge en bas

Public XML_Columns()
' 0 : Largeur colonne

Public XML_Styles()
' 0 : Nom style
' 1 : Font
' 2 : Size
' 3 : Bold
' 4 : Italic
' 5 : Format (texte, numérique...)
' 6 : Horizontal alignment
' 7 : Vertical alignment
' 8 : Fore Color
' 9 : Back Color
'10 : Wrap Texte
'11 : T taille trait
'12 : B taille trait
'13 : L taille trait
'14 : R taille trait

Public XML_Styles_Tri() As String

Public XML_Data()
' 0 to 18
' 0 : Données
' 1 : Merge à droite
' 2 : Merge en bas

' 3 : Font
' 4 : Size
' 5 : Bold
' 6 : Italic
' 7 : Format (texte, numérique...)
' 8 : Horizontal alignment
' 9 : Vertical alignment
'10 : Fore Color
'11 : Back Color
'12 : Wrap Texte
'13 : T taille trait
'14 : B taille trait
'15 : L taille trait
'16 : R taille trait

'18 : Index de style XML

' 1 to i : Colonnes
' 1 to j : Lignes


'Public Enum XML_HAlignment
'    XML_Left = 0
'    XML_Right = 1
'    XML_Center = 2
'End Enum

'Public Enum XML_VAlignment
'    XML_Top = 0
'    XML_Bottom = 1
'    XML_Center = 2
'End Enum

Public Const XML_Left = 0
Public Const XML_Right = 1
Public Const XML_Center = 2
Public Const XML_Top = 3
Public Const XML_Bottom = 4

Public XML_SHEET_NAME As String

Public Enum XML_Style_Key
    XML_Police = 0
    XML_Taille = 1
    XML_Gras = 2
    XML_Italic = 3
    XML_Format = 4
    XML_HAlignment = 5
    XML_VAlignment = 6
    XML_ForeColor = 7
    XML_BackColor = 8
    XML_WrapText = 9
    XML_Contour = 10
    XML_Quadrillage = 11
    XML_Row_Height = 12
'    XML_Line_T = 9
'    XML_Line_B = 10
'    XML_Line_L = 11
'    XML_Line_R = 12
End Enum

Public XML_NColonnes As Long

Public Function XML_GENERATE(Optional XML_Fichier As String = "")

ReDim XML_Styles_Tri(0 To 0, 0 To 0)

l = 1
Do While l <= UBound(XML_Data, 3)
    c = 1
    Do While c <= UBound(XML_Data, 2)
        Cle = ""
        
        i = 3
        Do While i <= 16
            Cle = Cle & XML_Data(i, c, l) & ";"
        
            i = i + 1
        Loop

        Cle = "L" & SetLen(CStr(Len(Cle)), 3, "0", Droite) & ";" & Cle

        Cle = Cle & CStr(c) & ";"
        Cle = Cle & CStr(l) & ";"

        If XML_Styles_Tri(0, 0) = "" Then
            XML_Styles_Tri(0, 0) = Cle
        Else
            ReDim Preserve XML_Styles_Tri(0 To 0, 0 To (UBound(XML_Styles_Tri, 2) + 1))
            XML_Styles_Tri(0, UBound(XML_Styles_Tri, 2)) = Cle
        End If
                
        c = c + 1
    Loop

    l = l + 1
Loop

QuickSort XML_Styles_Tri, 0, UBound(XML_Styles_Tri, 2), 0

Style_Idx = 0
Style_Pre = ""

i = 0
Do While i <= UBound(XML_Styles_Tri, 2)
    t = XML_Styles_Tri(0, i)
    P = Mid(t, 2, 3)

    Style_Cel = Left(t, CInt(P) + 5)
    XML_Range = Right(t, Len(t) - CInt(P) - 5)
    
    MP_c = InStr(1, XML_Range, ";", 1)
    MP_l = InStr(MP_c + 1, XML_Range, ";", 1)
    
    c = CLng(Mid(XML_Range, 1, MP_c - 1))
    l = CLng(Mid(XML_Range, MP_c + 1, MP_l - MP_c - 1))
    
    'Ajout du style
    If Style_Cel <> Style_Pre Then
        n = (UBound(XML_Styles, 2) + 1)
        ReDim Preserve XML_Styles(0 To 14, 0 To n)
        
        Style_Idx = Style_Idx + 1
        Style_Nom = "S" & SetLen(CStr(Style_Idx), 6, "0", Droite)
        
        XML_Styles(0, n) = Style_Nom
        XML_Data(18, c, l) = Style_Nom
        
        j = 1
        Do While j <= 14
            XML_Styles(j, n) = XML_Data(j + 2, c, l)
            j = j + 1
        Loop
    
        Style_Pre = Style_Cel
    Else
        XML_Data(18, c, l) = Style_Nom
    End If

    i = i + 1
Loop

Dim Bt As String
Dim Bb As String
Dim Bl As String
Dim Br As String

If XML_Fichier = "" Then
    XML_Fichier = App.Path & "\TEST.XLS"
End If

Dim XFile As Long

XFile = FreeFile
Open XML_Fichier For Output As #XFile

'XML HEADER
Print #XFile, XML_HEADER

'XML STYLES
Print #XFile, " <Styles>"

i = 0
Do While i <= UBound(XML_Styles, 2)
    Print #XFile, "  <Style ss:ID=" & Chr(34) & XML_Styles(0, i) & Chr(34) & ">"

    HAlign = ""
    VAlign = ""
    WTexte = ""

    If XML_Styles(6, i) = XML_Center Then
        HAlign = "Center"
    ElseIf XML_Styles(6, i) = XML_Left Then
        HAlign = "Left"
    ElseIf XML_Styles(6, i) = XML_Right Then
        HAlign = "Right"
    End If

    If XML_Styles(7, i) = XML_Center Then
        VAlign = "Center"
    ElseIf XML_Styles(7, i) = XML_Top Then
        VAlign = "Top"
    ElseIf XML_Styles(7, i) = XML_Bottom Then
        VAlign = "Bottom"
    End If
    
    If XML_Styles(10, i) = "1" Then
        WTexte = "1"
    End If
    
    If VAlign <> "" Or HAlign <> "" Or WTexte <> "" Then
        If VAlign <> "" Then
            VAlign = " ss:Vertical=" & Chr(34) & VAlign & Chr(34)
        End If
    
        If HAlign <> "" Then
            HAlign = " ss:Horizontal=" & Chr(34) & HAlign & Chr(34)
        End If
    
        If WTexte <> "" Then 'ss:WrapText="1"/>
            WTexte = " ss:WrapText=" & Chr(34) & WTexte & Chr(34)
        End If
    
        Print #XFile, "   <Alignment" & HAlign & VAlign & WTexte & "/>"
    End If

    Bt = XML_Styles(11, i)
    Bb = XML_Styles(12, i)
    Bl = XML_Styles(13, i)
    Br = XML_Styles(14, i)

    If IsNumeric(Bt) Then
        Bt = CInt(Bt)
    Else
        Bt = 0
    End If

    If IsNumeric(Bb) Then
        Bb = CInt(Bb)
    Else
        Bb = 0
    End If

    If IsNumeric(Bl) Then
        Bl = CInt(Bl)
    Else
        Bl = 0
    End If

    If IsNumeric(Br) Then
        Br = CInt(Br)
    Else
        Br = 0
    End If

    'Au moins une bordure
    If CInt(Bt) + CInt(Bb) + CInt(Bl) + CInt(Br) > 0 Then
        Print #XFile, "   <Borders>"

        If Bt > 0 Then
            Print #XFile, "    <Border ss:Position=" & Chr(34) & "Top" & Chr(34) & " ss:LineStyle=" & Chr(34) & "Continuous" & Chr(34) & " ss:Weight=" & Chr(34) & CStr(Bt) & Chr(34) & "/>"
        End If

        If Bb > 0 Then
            Print #XFile, "    <Border ss:Position=" & Chr(34) & "Bottom" & Chr(34) & " ss:LineStyle=" & Chr(34) & "Continuous" & Chr(34) & " ss:Weight=" & Chr(34) & CStr(Bb) & Chr(34) & "/>"
        End If

        If Bl > 0 Then
            Print #XFile, "    <Border ss:Position=" & Chr(34) & "Left" & Chr(34) & " ss:LineStyle=" & Chr(34) & "Continuous" & Chr(34) & " ss:Weight=" & Chr(34) & CStr(Bl) & Chr(34) & "/>"
        End If

        If Br > 0 Then
            Print #XFile, "    <Border ss:Position=" & Chr(34) & "Right" & Chr(34) & " ss:LineStyle=" & Chr(34) & "Continuous" & Chr(34) & " ss:Weight=" & Chr(34) & CStr(Br) & Chr(34) & "/>"
        End If

        Print #XFile, "   </Borders>"
    End If

    'Font
    Font_Print = False
    Font_Name = ""
    Font_Size = ""
    Font_Bold = ""
    Font_Italic = ""
    Font_Color = ""
    
    If XML_Styles(1, i) <> "" Then
        Font_Print = True
        Font_Name = " ss:FontName=" & Chr(34) & XML_Styles(1, i) & Chr(34)
    End If

    If XML_Styles(2, i) <> "" Then
        Font_Print = True
        Font_Size = " ss:Size=" & Chr(34) & XML_Styles(2, i) & Chr(34)
    End If

    If XML_Styles(3, i) <> "" Then
        Font_Print = True
        Font_Bold = " ss:Bold=" & Chr(34) & "1" & Chr(34)
    End If

    If XML_Styles(4, i) <> "" Then
        Font_Print = True
        Font_Bold = " ss:Italic=" & Chr(34) & "1" & Chr(34)
    End If

    If XML_Styles(8, i) <> "" Then
        Font_Print = True
        Font_Color = " ss:Color=" & Chr(34) & Couleur_HEXA(XML_Styles(8, i)) & Chr(34)
    End If

    If Font_Print = True Then
        Print #XFile, "   <Font" & Font_Name & " x:Family=" & Chr(34) & "Swiss" & Chr(34) & Font_Size & Font_Color & Font_Bold & "/>"
    End If

    'Couleur intérieur

    If XML_Styles(9, i) <> "" Then
        Print #XFile, "   <Interior ss:Color=" & Chr(34) & Couleur_HEXA(XML_Styles(9, i)) & Chr(34) & " ss:Pattern=" & Chr(34) & "Solid" & Chr(34) & "/>"
    End If

    'Format
    If XML_Styles(5, i) <> "" Then
        Print #XFile, "   <NumberFormat ss:Format=" & Chr(34) & XML_Styles(5, i) & Chr(34) & "/>"
    End If

    Print #XFile, "  </Style>"

    i = i + 1
Loop

' 0 : Nom style
'OK  1 : Font
'OK  2 : Size
'OK  3 : Bold
' 4 : Italic
'OK  5 : Format (texte, numérique...)
'OK  6 : Horizontal alignment
'OK  7 : Vertical alignment
'OK  8 : Fore Color
'OK  9 : Back Color
'OK 10 : T taille trait
'OK 11 : B taille trait
'OK 12 : L taille trait
'OK 13 : R taille trait

Print #XFile, " </Styles>"

'XML SHEET
Print #XFile, " <Worksheet ss:Name=" & Chr(34) & XML_SHEET_NAME & Chr(34) & ">"
'Print #XFile, "  <Names>"
'Print #XFile, "   <NamedRange ss:Name=" & Chr(34) & "_FilterDatabase" & Chr(34) & " ss:RefersTo=" & Chr(34) & "=F1!R7C1:R7C12" & Chr(34) & " ss:Hidden=" & Chr(34) & "1" & Chr(34) & "/>"
'Print #XFile, "   <NamedRange ss:Name=" & Chr(34) & "Print_Area" & Chr(34) & " ss:RefersTo=" & Chr(34) & "=BL!R1C1:R" & CStr(NLignes) & "C12" & Chr(34) & "/>"
'Print #XFile, "  </Names>"

'Print #XFile, "  <Table ss:ExpandedColumnCount=" & Chr(34) & CStr(UBound(XML_Data, 2)) & Chr(34) & " ss:ExpandedRowCount=" & Chr(34) & CStr(UBound(XML_Data, 3)) & Chr(34) & ">"
'Print #XFile, "   x:FullRows=" & Chr(34) & "1" & Chr(34) & " ss:StyleID=" & Chr(34) & "Default" & Chr(34)  '& ">"
'Print #XFile, "   ss:DefaultRowHeight=" & Chr(34) & "15.75" & Chr(34) & ">"

Print #XFile, "  <Table ss:ExpandedColumnCount=" & Chr(34) & CStr(UBound(XML_Data, 2)) & Chr(34) & " ss:ExpandedRowCount=" & Chr(34) & CStr(UBound(XML_Data, 3)) & Chr(34) & " x:FullColumns=" & Chr(34) & "1" & Chr(34)
Print #XFile, "   x:FullRows=" & Chr(34) & "1" & Chr(34) & " ss:StyleID=" & Chr(34) & "Default" & Chr(34) & " ss:DefaultColumnWidth=" & Chr(34) & "60" & Chr(34) '& ">"
Print #XFile, "   ss:DefaultRowHeight=" & Chr(34) & "15.75" & Chr(34) & ">"

'XML COLONNES
'Print #XFile, "   <Column ss:AutoFitWidth=" & Chr(34) & "1" & Chr(34) & "/>"
'Print #XFile, "   <Column ss:AutoFitWidth=" & Chr(34) & "1" & Chr(34) & "/>"
'Print #XFile, "   <Column ss:AutoFitWidth=" & Chr(34) & "1" & Chr(34) & "/>"
'Print #XFile, "   <Column ss:AutoFitWidth=" & Chr(34) & "1" & Chr(34) & "/>"
'Print #XFile, "   <Column ss:AutoFitWidth=" & Chr(34) & "1" & Chr(34) & "/>"

'XLS width = XML width * 4/3

c = 1
Do While c <= UBound(XML_Columns, 2)
    Print #XFile, "   <Column ss:StyleID=" & Chr(34) & "Default" & Chr(34) & " ss:AutoFitWidth=" & Chr(34) & "0" & Chr(34) & " ss:Width=" & Chr(34) & Replace(CStr(XML_Columns(0, c) * 3 / 4), ",", ".") & Chr(34) & "/>"
    c = c + 1
Loop

'Print #XFile, "   <Column ss:StyleID=" & Chr(34) & "Default" & Chr(34) & " ss:AutoFitWidth=" & Chr(34) & "0" & Chr(34) & " ss:Width=" & Chr(34) & Replace(CStr(103 * 3 / 4), ",", ".") & Chr(34) & "/>"
'Print #XFile, "   <Column ss:StyleID=" & Chr(34) & "Default" & Chr(34) & " ss:AutoFitWidth=" & Chr(34) & "0" & Chr(34) & " ss:Width=" & Chr(34) & Replace(CStr(263 * 3 / 4), ",", ".") & Chr(34) & "/>"
'Print #XFile, "   <Column ss:StyleID=" & Chr(34) & "Default" & Chr(34) & " ss:AutoFitWidth=" & Chr(34) & "0" & Chr(34) & " ss:Width=" & Chr(34) & Replace(CStr(35 * 3 / 4), ",", ".") & Chr(34) & "/>"
'Print #XFile, "   <Column ss:StyleID=" & Chr(34) & "Default" & Chr(34) & " ss:AutoFitWidth=" & Chr(34) & "0" & Chr(34) & " ss:Width=" & Chr(34) & Replace(CStr(145 * 3 / 4), ",", ".") & Chr(34) & "/>"
'Print #XFile, "   <Column ss:StyleID=" & Chr(34) & "Default" & Chr(34) & " ss:AutoFitWidth=" & Chr(34) & "0" & Chr(34) & " ss:Width=" & Chr(34) & Replace(CStr(62 * 3 / 4), ",", ".") & Chr(34) & "/>"
'Print #XFile, "   <Column ss:StyleID=" & Chr(34) & "Default" & Chr(34) & " ss:AutoFitWidth=" & Chr(34) & "0" & Chr(34) & " ss:Width=" & Chr(34) & Replace(CStr(62 * 3 / 4), ",", ".") & Chr(34) & "/>"
'Print #XFile, "   <Column ss:StyleID=" & Chr(34) & "Default" & Chr(34) & " ss:AutoFitWidth=" & Chr(34) & "0" & Chr(34) & " ss:Width=" & Chr(34) & Replace(CStr(62 * 3 / 4), ",", ".") & Chr(34) & "/>"
'Print #XFile, "   <Column ss:StyleID=" & Chr(34) & "Default" & Chr(34) & " ss:AutoFitWidth=" & Chr(34) & "0" & Chr(34) & " ss:Width=" & Chr(34) & Replace(CStr(62 * 3 / 4), ",", ".") & Chr(34) & "/>"
'Print #XFile, "   <Column ss:StyleID=" & Chr(34) & "Default" & Chr(34) & " ss:AutoFitWidth=" & Chr(34) & "0" & Chr(34) & " ss:Width=" & Chr(34) & Replace(CStr(62 * 3 / 4), ",", ".") & Chr(34) & "/>"

'XML CELLULES
l = 1
Do While l <= UBound(XML_Data, 3)
    h = ""
    
    If XML_Data(17, 1, l) = vbEmpty Then
    Else
        If IsNumeric(XML_Data(17, 1, l)) Then
            h = CStr(XML_Data(17, 1, l))
        End If
    End If

    If h = "" Then
        Print #XFile, XML_ROW_DEB(l)
    Else
        Print #XFile, XML_ROW_DEB(l, h)
    End If
    
    XML_ROW_TXT = ""
    
    c = 1
    Do While c <= UBound(XML_Data, 2)
        If XML_Data(1, c, l) <> "-1" Then
            XML_ROW_TXT = XML_ROW_TXT & XML_CELL_ADD(CInt(l), CInt(c)) 'XML_Data(0, c, l), XML_Data(18, c, l), c, CInt(XML_Data(1, c, l)), CInt(XML_Data(2, c, l)))
        End If

        c = c + 1
    Loop

    Print #XFile, XML_ROW_TXT
    Print #XFile, XML_ROW_FIN
    
    l = l + 1
Loop

Print #XFile, "  </Table>"
Print #XFile, "  <WorksheetOptions xmlns=" & Chr(34) & "urn:schemas-microsoft-com:office:excel" & Chr(34) & ">"
Print #XFile, "   <PageSetup>"
Print #XFile, "    <Layout x:CenterHorizontal=" & Chr(34) & "1" & Chr(34) & "/>"
Print #XFile, "    <Header x:Margin=" & Chr(34) & "0.42" & Chr(34) & "/>"
Print #XFile, "    <Footer x:Margin=" & Chr(34) & "0.3" & Chr(34) & "/>"
Print #XFile, "    <PageMargins x:Bottom=" & Chr(34) & "0.74" & Chr(34) & " x:Left=" & Chr(34) & "0.25" & Chr(34) & " x:Right=" & Chr(34) & "0.25" & Chr(34) & " x:Top=" & Chr(34) & "0.46" & Chr(34) & "/>"
Print #XFile, "   </PageSetup>"
Print #XFile, "   <NoSummaryRowsBelowDetail/>"
Print #XFile, "   <Print>"
Print #XFile, "    <FitHeight>2</FitHeight>"
Print #XFile, "    <ValidPrinterInfo/>"
Print #XFile, "    <PaperSizeIndex>9</PaperSizeIndex>"
Print #XFile, "    <Scale>83</Scale>"
Print #XFile, "    <HorizontalResolution>600</HorizontalResolution>"
Print #XFile, "    <VerticalResolution>600</VerticalResolution>"
Print #XFile, "   </Print>"
Print #XFile, "   <PageBreakZoom>50</PageBreakZoom>"
Print #XFile, "   <Selected/>"
'Print #XFile, "   <TopRowVisible>168</TopRowVisible>"
'Print #XFile, "   <Panes>"
'Print #XFile, "    <Pane>"
'Print #XFile, "     <Number>3</Number>"
'Print #XFile, "     <ActiveRow>1</ActiveRow>"
'Print #XFile, "     <ActiveCol>10</ActiveCol>"
'Print #XFile, "     <RangeSelection>R2C11:R2C12</RangeSelection>"
'Print #XFile, "    </Pane>"
'Print #XFile, "   </Panes>"
Print #XFile, "   <ProtectObjects>False</ProtectObjects>"
Print #XFile, "   <ProtectScenarios>False</ProtectScenarios>"
Print #XFile, "  </WorksheetOptions>"
'Print #XFile, "  <ConditionalFormatting xmlns=" & Chr(34) & "urn:schemas-microsoft-com:office:excel" & Chr(34) & ">"
'Print #XFile, "   <Range>R116C9</Range>"
'Print #XFile, "   <Condition>"
'Print #XFile, "    <Qualifier>NotEqual</Qualifier>"
'Print #XFile, "    <Value1>0</Value1>"
'Print #XFile, "    <Format Style='background:#FFCC99'/>"
'Print #XFile, "   </Condition>"
'Print #XFile, "  </ConditionalFormatting>"
'Print #XFile, "  <ConditionalFormatting xmlns=" & Chr(34) & "urn:schemas-microsoft-com:office:excel" & Chr(34) & ">"
'Print #XFile, "   <Range>R116C1:R116C8,R116C10</Range>"
'Print #XFile, "   <Condition>"
'Print #XFile, "    <Value1>R58C9</Value1>"
'Print #XFile, "    <Format Style='background:#FFCC99'/>"
'Print #XFile, "   </Condition>"
'Print #XFile, "  </ConditionalFormatting>"
Print #XFile, " </Worksheet>"
Print #XFile, "</Workbook>"

Close #XFile


End Function

Public Function XML_ROW_DEB(Optional Index = -1, Optional Hauteur = 12.75) As String


t = ""
h = Replace(CStr(Hauteur), ",", ".")
If Index >= 0 Then
    t = t & "   <Row ss:Index=" & Chr(34) & CStr(Index) & Chr(34) & " ss:Height=" & Chr(34) & h & Chr(34) & " ss:AutoFitHeight=" & Chr(34) & "0" & Chr(34) & ">" & vbCrLf
Else
    t = t & "   <Row ss:AutoFitHeight=" & Chr(34) & "0" & Chr(34) & " ss:Height=" & Chr(34) & h & Chr(34) & ">" & vbCrLf
End If

XML_ROW_DEB = t
End Function

Public Function XML_ROW_FIN(Optional Index) As String

t = ""
t = t & "   </Row>" & vbCrLf

XML_ROW_FIN = t
End Function

Public Function XML_CELL_ADD(XML_Ligne As Integer, XML_Colonne As Integer) As String 'Texte, Style, Colonne, Optional Merge_Across As Integer = 0, Optional Merge_Down As Integer = 0) As String

'XML_CELL_ADD(XML_Data(0, c, l), XML_Data(18, c, l), c, CInt(XML_Data(1, c, l)), CInt(XML_Data(2, c, l)))

Dim c As Integer
Dim l As Integer

c = XML_Colonne
l = XML_Ligne

Texte = XML_Data(0, c, l)
Style = XML_Data(18, c, l)
Colonne = c

Merge_Across = CInt(XML_Data(1, c, l))
Merge_Down = CInt(XML_Data(2, c, l))

t = ""
m = ""
d = ""

If Merge_Across <> 0 Then
    ma = "ss:MergeAcross=" & Chr(34) & CStr(Merge_Across) & Chr(34) & " "
End If

If Merge_Down <> 0 Then
    md = "ss:MergeDown=" & Chr(34) & CStr(Merge_Down) & Chr(34) & " "
End If

'Texte = "0.5"

If XML_Data(7, c, l) = "@" Then
    Data_Type = "String"
ElseIf IsNumeric(Replace(XML_Data(7, c, l), ".", ",")) And IsNumeric(Replace(Texte, ".", ",")) Then
    Texte = Replace(Texte, ",", ".")
    Data_Type = "Number"
Else
    Data_Type = "String"
End If

If Texte <> "" Then
    If Colonne <= 12 Then
        d = "><Data ss:Type=" & Chr(34) & Data_Type & Chr(34) & ">" & Texte & "</Data><NamedCell ss:Name=" & Chr(34) & "Print_Area" & Chr(34) & "/></Cell>"
    Else
        d = "><Data ss:Type=" & Chr(34) & Data_Type & Chr(34) & ">" & Texte & "</Data></Cell>"
    End If
Else
    d = "/>"
End If

    c_idx = "ss:Index=" & Chr(34) & CStr(Colonne) & Chr(34) & " "

If Style <> "" Then
    s = "ss:StyleID=" & Chr(34) & Style & Chr(34)
Else
    s = "ss:StyleID=" & Chr(34) & "Default" & Chr(34)
End If

'If m <> "" Then
    t = t & "    <Cell " & c_idx & ma & md & s & d
'End If

XML_CELL_ADD = t & vbCrLf
't = t & "    <Cell ss:MergeAcross=" & Chr(34) & "4" & Chr(34) & " ss:StyleID=" & Chr(34) & "b03" & Chr(34) & "><Data ss:Type=" & Chr(34) & "String" & Chr(34) & "></Data></Cell>" & vbCrLf

End Function

Public Function XML_CELLULE(Ligne As Long, Colonne As Long, Valeur As String) As String ', Optional Merge_R = 0, Optional Merge_B = 0) As String

c = Colonne
l = Ligne

Cond = (Colonne + Merge_R <= UBound(XML_Data, 2))
'Cond = Cond And (Ligne <= UBound(XML_Data, 3))

If Cond = False Then
    XML_CELLULE = "Erreur de colonne : la dimension de la table est inférieure"
    Exit Function
End If

If (Ligne + Merge_B) > UBound(XML_Data, 3) Then
    ReDim Preserve XML_Data(0 To 18, 1 To XML_NColonnes, 1 To (Ligne + Merge_B))
End If

'Convertion des accents
Valeur = AToUTF8(Valeur)

XML_Data(0, c, l) = Valeur


End Function

Public Function XML_MERGE(XML_Range) As String

'XML_Range : "R1C1:R10C2" ou "R10C3" ...

MP = InStr(1, XML_Range, ":", 1)
If MP = 0 Then
    R = XML_Range & ":" & XML_Range
Else
    R = XML_Range
End If

'Row initiale
MP_Ri = InStr(1, R, "R", 1)
MP_Ci = InStr(MP_Ri + 1, R, "C", 1)
MP_To = InStr(MP_Ci + 1, R, ":", 1)
MP_Rf = InStr(MP_To + 1, R, "R", 1)
MP_Cf = InStr(MP_Rf + 1, R, "C", 1)

Cond = IsNumeric(MP_Ri)
Cond = Cond And IsNumeric(MP_Ci)
Cond = Cond And IsNumeric(MP_To)
Cond = Cond And IsNumeric(MP_Rf)
Cond = Cond And IsNumeric(MP_Cf)

If Cond = False Then
    XML_MERGE = "Erreur de Syntax : Range non valide"
    Exit Function
End If

Cond = (MP_Ri > 0)
Cond = Cond And (MP_Ci > 0)
Cond = Cond And (MP_To > 0)
Cond = Cond And (MP_Rf > 0)
Cond = Cond And (MP_Cf > 0)

If Cond = False Then
    XML_MERGE = "Erreur de Syntax : Range non valide"
    Exit Function
End If

Ri = Mid(R, MP_Ri + 1, MP_Ci - MP_Ri - 1)
Ci = Mid(R, MP_Ci + 1, MP_To - MP_Ci - 1)

Rf = Mid(R, MP_Rf + 1, MP_Cf - MP_Rf - 1)
Cf = Mid(R, MP_Cf + 1, Len(R) - MP_Cf)

Cond = IsNumeric(Ri)
Cond = Cond And IsNumeric(Ci)
Cond = Cond And IsNumeric(Rf)
Cond = Cond And IsNumeric(Cf)

If Cond = False Then
    XML_MERGE = "Erreur de Syntax : Range non valide"
    Exit Function
End If

Ci = CLng(Ci)
Ri = CLng(Ri)
Cf = CLng(Cf)
Rf = CLng(Rf)

Cond = (Ri > 0)
Cond = Cond And (Ci > 0)
Cond = Cond And (Rf > 0)
Cond = Cond And (Cf > 0)
Cond = Cond And (Rf >= Ri)
Cond = Cond And (Cf >= Ci)
Cond = Cond And (Cf <= UBound(XML_Data, 2))
Cond = Cond And (Rf <= UBound(XML_Data, 3))

If Cond = False Then
    XML_MERGE = "Erreur de Syntax : Range non valide"
    Exit Function
End If

c = Ci
l = Ri

Merge_R = Cf - Ci
Merge_B = Rf - Ri

XML_Data(1, Ci, Ri) = Merge_R
XML_Data(2, Ci, Ri) = Merge_B

If Merge_R > 0 Then
    i = c + 1
    Do While i <= c + Merge_R
        XML_Data(1, i, l) = -1
        XML_Data(2, i, l) = -1
        i = i + 1
    Loop

    If Merge_B > 0 Then
        i = c
        Do While i <= c + Merge_R
            
            j = l + 1
            Do While j <= l + Merge_B
                XML_Data(1, i, j) = -1
                XML_Data(2, i, j) = -1
                j = j + 1
            Loop
            
            i = i + 1
        Loop
    End If
ElseIf Merge_B > 0 Then
    j = l + 1
    Do While j <= l + Merge_B
        XML_Data(1, c, j) = -1
        XML_Data(2, c, j) = -1
        j = j + 1
    Loop
End If


End Function


Public Function XML_INITIATE(Nom_Feuille, Nb_Colonne)

ReDim XML_Data(0 To 18, 1 To Nb_Colonne, 1 To 1)
XML_NColonnes = Nb_Colonne

ReDim XML_Styles(0 To 14, 0 To 0)
ReDim XML_Columns(0 To 0, 1 To Nb_Colonne)
'ReDim XML_Columns(0 To 0, 1 To (Nb_Colonne - 1))

'Style par défaut
XML_Styles(0, 0) = "Default"
XML_SHEET_NAME = Nom_Feuille

End Function


Public Function XML_STYLE(XML_Range, Style_Key As XML_Style_Key, Valeur) As String

'XML_Range : "R1C1:R10C2" ou "R10C3" ...

MP = InStr(1, XML_Range, ":", 1)
If MP = 0 Then
    R = XML_Range & ":" & XML_Range
Else
    R = XML_Range
End If

'Row initiale
MP_Ri = InStr(1, R, "R", 1)
MP_Ci = InStr(MP_Ri + 1, R, "C", 1)
MP_To = InStr(MP_Ci + 1, R, ":", 1)
MP_Rf = InStr(MP_To + 1, R, "R", 1)
MP_Cf = InStr(MP_Rf + 1, R, "C", 1)

Cond = IsNumeric(MP_Ri)
Cond = Cond And IsNumeric(MP_Ci)
Cond = Cond And IsNumeric(MP_To)
Cond = Cond And IsNumeric(MP_Rf)
Cond = Cond And IsNumeric(MP_Cf)

If Cond = False Then
    XML_STYLE = "Erreur de Syntax : Range non valide"
    Exit Function
End If

Cond = (MP_Ri > 0)
Cond = Cond And (MP_Ci > 0)
Cond = Cond And (MP_To > 0)
Cond = Cond And (MP_Rf > 0)
Cond = Cond And (MP_Cf > 0)

If Cond = False Then
    XML_STYLE = "Erreur de Syntax : Range non valide"
    Exit Function
End If

Ri = Mid(R, MP_Ri + 1, MP_Ci - MP_Ri - 1)
Ci = Mid(R, MP_Ci + 1, MP_To - MP_Ci - 1)

Rf = Mid(R, MP_Rf + 1, MP_Cf - MP_Rf - 1)
Cf = Mid(R, MP_Cf + 1, Len(R) - MP_Cf)

Cond = IsNumeric(Ri)
Cond = Cond And IsNumeric(Ci)
Cond = Cond And IsNumeric(Rf)
Cond = Cond And IsNumeric(Cf)

If Cond = False Then
    XML_STYLE = "Erreur de Syntax : Range non valide"
    Exit Function
End If

Ci = CLng(Ci)
Ri = CLng(Ri)
Cf = CLng(Cf)
Rf = CLng(Rf)

Cond = (Ri > 0)
Cond = Cond And (Ci > 0)
Cond = Cond And (Rf > 0)
Cond = Cond And (Cf > 0)
Cond = Cond And (Rf >= Ri)
Cond = Cond And (Cf >= Ci)
Cond = Cond And (Cf <= UBound(XML_Data, 2))
Cond = Cond And (Rf <= UBound(XML_Data, 3))

If Cond = False Then
    XML_STYLE = "Erreur de Syntax : Range non valide"
    Exit Function
End If

If Style_Key <= 9 Then
    Idx = Style_Key + 3

    c = Ci
    Do While c <= Cf
        l = Ri
        Do While l <= Rf
            XML_Data(Idx, c, l) = Valeur
            l = l + 1
        Loop
    
        c = c + 1
    Loop
ElseIf Style_Key <= 11 Then
    '10 : Contour
    '11: Quadrillage
    
'13 : T taille trait
'14 : B taille trait
'15 : L taille trait
'16 : R taille trait
    
    c = Ci
    Do While c <= Cf
        l = Ri
        Do While l <= Rf
            If Style_Key = XML_Contour Then
                If l = Ri Then
                    XML_Data(13, c, l) = Valeur
                End If
                
                If l = Rf Then
                    XML_Data(14, c, l) = Valeur
                End If
                
                If c = Ci Then
                    XML_Data(15, c, l) = Valeur
                End If
                
                If c = Cf Then
                    XML_Data(16, c, l) = Valeur
                End If
            Else
                XML_Data(13, c, l) = Valeur
                XML_Data(14, c, l) = Valeur
                XML_Data(15, c, l) = Valeur
                XML_Data(16, c, l) = Valeur
            End If
            
            l = l + 1
        Loop
    
        c = c + 1
    Loop
Else
    l = Ri
    Do While l <= Rf
        XML_Data(17, 1, l) = Valeur

        l = l + 1
    Loop
End If

End Function

Public Function XML_HEADER() As String

t = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & "?>" & vbCrLf
t = t & "<?mso-application progid=" & Chr(34) & "Excel.Sheet" & Chr(34) & "?>" & vbCrLf
t = t & "<Workbook xmlns=" & Chr(34) & "urn:schemas-microsoft-com:office:spreadsheet" & Chr(34) & vbCrLf
t = t & " xmlns:o=" & Chr(34) & "urn:schemas-microsoft-com:office:office" & Chr(34) & vbCrLf
t = t & " xmlns:x=" & Chr(34) & "urn:schemas-microsoft-com:office:excel" & Chr(34) & vbCrLf
t = t & " xmlns:ss=" & Chr(34) & "urn:schemas-microsoft-com:office:spreadsheet" & Chr(34) & vbCrLf
t = t & " xmlns:html=" & Chr(34) & "http://www.w3.org/TR/REC-html40" & Chr(34) & ">" & vbCrLf
t = t & " <DocumentProperties xmlns=" & Chr(34) & "urn:schemas-microsoft-com:office:office" & Chr(34) & ">" & vbCrLf
t = t & "  <Author>Jean-Michel COLLIN</Author>" & vbCrLf
t = t & "  <LastAuthor>Jean-Michel COLLIN</LastAuthor>" & vbCrLf
t = t & "  <LastPrinted>2008-03-05T11:25:33Z</LastPrinted>" & vbCrLf
t = t & "  <Created>2008-03-05T09:49:20Z</Created>" & vbCrLf
t = t & "  <LastSaved>2008-03-05T11:25:36Z</LastSaved>" & vbCrLf
t = t & "  <Version>11.6360</Version>" & vbCrLf
t = t & " </DocumentProperties>" & vbCrLf
t = t & " <ExcelWorkbook xmlns=" & Chr(34) & "urn:schemas-microsoft-com:office:excel" & Chr(34) & ">" & vbCrLf
t = t & "  <WindowHeight>8445</WindowHeight>" & vbCrLf
t = t & "  <WindowWidth>18795</WindowWidth>" & vbCrLf
t = t & "  <WindowTopX>120</WindowTopX>" & vbCrLf
t = t & "  <WindowTopY>90</WindowTopY>" & vbCrLf
t = t & "  <ProtectStructure>False</ProtectStructure>" & vbCrLf
t = t & "  <ProtectWindows>False</ProtectWindows>" & vbCrLf
t = t & " </ExcelWorkbook>" & vbCrLf

XML_HEADER = t

End Function

Private Function Couleur_HEXA(n) As String

If IsNumeric(n) = False Then
    Couleur_HEXA = n
Else
    h = Hex(n)
    
    Do While Len(h) < 6
        h = "0" + h
    Loop
    
    Couleur_HEXA = Right(h, 2) & Mid(h, 3, 2) & Left(h, 2)
    Couleur_HEXA = "#" & Couleur_HEXA
End If

End Function

Public Function AToUTF8(ByVal wText As String) As String
    Dim vNeeded As Long
    Dim vSize   As Long
    vSize = Len(wText)
    vNeeded = WideCharToMultiByte(CP_UTF8, 0, StrPtr(wText), vSize, "", 0, 0, 0)
    AToUTF8 = String(vNeeded, 0)
    WideCharToMultiByte CP_UTF8, 0, StrPtr(wText), vSize, AToUTF8, vNeeded, 0, 0
End Function

Public Function UTF8ToA(ByVal wText As String) As String
    Dim vNeeded As Long
    Dim vSize   As Long
    vSize = Len(wText)
    vNeeded = MultiByteToWideChar(CP_UTF8, 0, wText, vSize, 0, 0)
    UTF8ToA = String(vNeeded, 0)
    MultiByteToWideChar CP_UTF8, 0, wText, vSize, StrPtr(UTF8ToA), vNeeded
End Function

