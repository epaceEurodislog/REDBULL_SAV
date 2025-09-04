Attribute VB_Name = "modValidation"
' === modValidation.bas - MODULE COMPLET ===
Option Explicit

' === FONCTIONS DE VALIDATION ===

' Fonction pour valider le format du num�ro de s�rie Red Bull
Public Function ValiderFormatNumeroSerieRB(numeroSerie As String) As Boolean
    ' V�rifications de base pour Red Bull
    If Len(numeroSerie) < 8 Or Len(numeroSerie) > 25 Then
        ValiderFormatNumeroSerieRB = False
        Exit Function
    End If
    
    ' V�rifier que c'est alphanum�rique avec tirets autoris�s
    Dim i As Integer
    For i = 1 To Len(numeroSerie)
        Dim char As String
        char = Mid(numeroSerie, i, 1)
        If Not ((char >= "A" And char <= "Z") Or (char >= "0" And char <= "9") Or char = "-") Then
            ValiderFormatNumeroSerieRB = False
            Exit Function
        End If
    Next i
    
    ValiderFormatNumeroSerieRB = True
End Function

' Fonction pour obtenir des informations compl�mentaires sur l'article
Public Function ObtenirInfosComplementairesArticle(codeArticle As String) As String
    On Error GoTo ErrorHandler
    
    ' Les informations sont r�cup�r�es dans la requ�te principale
    ObtenirInfosComplementairesArticle = "Informations r�cup�r�es depuis ART_PAR et NSE_DAT"
    Exit Function
    
ErrorHandler:
    ObtenirInfosComplementairesArticle = "Erreur lors de la r�cup�ration des infos: " & Err.description
End Function

' Fonction pour valider les donn�es SAV
Public Function ValiderDonnees(donnees As TypeSAV) As Boolean
    ' Validation de base
    ValiderDonnees = True
    
    ' V�rifier le num�ro de s�rie avec la requ�te BDD
    If VerifierConnexionBDD() And Len(donnees.ReferenceProduit) > 0 Then
        ValiderDonnees = VerifierNumeroSerieBDD(donnees.ReferenceProduit)
    End If
    
    ' V�rifications suppl�mentaires
    If Len(donnees.numeroEnlevement) = 0 Then
        ValiderDonnees = False
    End If
    
    If Len(donnees.MotifRetour) = 0 Then
        ValiderDonnees = False
    End If
End Function

' Fonction pour valider un code-barres
Public Function ValiderCodeBarre(codeBarre As String) As Boolean
    ' Supprime les espaces et convertit en majuscules
    codeBarre = Trim(UCase(codeBarre))
    
    ' V�rifications de base
    If Len(codeBarre) < 6 Then
        ValiderCodeBarre = False
        Exit Function
    End If
    
    ' V�rifie le format basique (lettres + chiffres + tirets)
    Dim i As Integer
    For i = 1 To Len(codeBarre)
        Dim char As String
        char = Mid(codeBarre, i, 1)
        If Not ((char >= "A" And char <= "Z") Or (char >= "0" And char <= "9") Or char = "-") Then
            ValiderCodeBarre = False
            Exit Function
        End If
    Next i
    
    ' V�rifier avec la logique BDD
    If VerifierConnexionBDD() Then
        ValiderCodeBarre = VerifierNumeroSerieBDD(codeBarre)
    Else
        ' Si pas de BDD, validation basique seulement
        ValiderCodeBarre = True
    End If
End Function

' Fonction pour extraire le mod�le du code-barres
Public Function ExtraireModele(codeBarre As String) As String
    On Error GoTo ErrorLocal
    
    ' Essayer d'abord depuis la BDD
    If VerifierConnexionBDD() Then
        Dim resultats As TypeValidationBDD
        resultats = ValiderNumeroSerieBDD(codeBarre)
        
        If resultats.existe Then
            ' Utiliser la d�signation obtenue de la requ�te BDD
            ExtraireModele = resultats.designationArticle
            Exit Function
        End If
    End If

ErrorLocal:
    ' Fallback - identification par pr�fixe local si pas trouv� en BDD
    ExtraireModele = DeterminerModeleParCode(codeBarre)
End Function

' Fonction pour d�terminer le mod�le bas� sur le code (fallback)
Private Function DeterminerModeleParCode(code As String) As String
    Dim prefixe As String
    prefixe = Left(UCase(code), 6)
    
    Select Case prefixe
        Case "VC2286"
            DeterminerModeleParCode = "Vitrine VC2286"
        Case "RB4458"
            DeterminerModeleParCode = "Red Bull RB4458"
        Case "CF3401"
            DeterminerModeleParCode = "Cong�lateur CF3401"
        Case "RB2024"
            DeterminerModeleParCode = "Red Bull Premium 2024"
        Case Else
            ' Essayer avec les 2 premiers caract�res
            Select Case Left(UCase(code), 2)
                Case "VC"
                    DeterminerModeleParCode = "Vitrine Red Bull"
                Case "RB"
                    DeterminerModeleParCode = "Frigo Red Bull"
                Case "CF"
                    DeterminerModeleParCode = "Cong�lateur Red Bull"
                Case "FB"
                    DeterminerModeleParCode = "Frigo Bar Red Bull"
                Case "RF"
                    DeterminerModeleParCode = "Red Fridge"
                Case Else
                    DeterminerModeleParCode = "�quipement Red Bull - Mod�le non identifi�"
            End Select
    End Select
End Function

' === FONCTIONS DE VALIDATION SP�CIALIS�ES ===

' Fonction pour valider une date
Public Function ValiderDate(dateStr As String) As Boolean
    On Error GoTo ErrorHandler
    
    If IsDate(dateStr) Then
        ValiderDate = True
    Else
        ValiderDate = False
    End If
    
    Exit Function
    
ErrorHandler:
    ValiderDate = False
End Function

' Fonction pour valider un num�ro (uniquement chiffres)
Public Function ValiderNumerique(valeur As String) As Boolean
    Dim i As Integer
    
    If Len(valeur) = 0 Then
        ValiderNumerique = False
        Exit Function
    End If
    
    For i = 1 To Len(valeur)
        If Not (Mid(valeur, i, 1) >= "0" And Mid(valeur, i, 1) <= "9") Then
            ValiderNumerique = False
            Exit Function
        End If
    Next i
    
    ValiderNumerique = True
End Function

' Fonction pour valider un prix
Public Function ValiderPrix(prix As String) As Boolean
    On Error GoTo ErrorHandler
    
    If IsNumeric(prix) Then
        Dim valeurPrix As Double
        valeurPrix = CDbl(prix)
        ValiderPrix = (valeurPrix >= 0)
    Else
        ValiderPrix = False
    End If
    
    Exit Function
    
ErrorHandler:
    ValiderPrix = False
End Function
