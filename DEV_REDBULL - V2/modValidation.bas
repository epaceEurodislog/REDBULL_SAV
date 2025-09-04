Attribute VB_Name = "modValidation"
' === modValidation.bas - MODULE COMPLET ===
Option Explicit

' === FONCTIONS DE VALIDATION ===

' Fonction pour valider le format du numéro de série Red Bull
Public Function ValiderFormatNumeroSerieRB(numeroSerie As String) As Boolean
    ' Vérifications de base pour Red Bull
    If Len(numeroSerie) < 8 Or Len(numeroSerie) > 25 Then
        ValiderFormatNumeroSerieRB = False
        Exit Function
    End If
    
    ' Vérifier que c'est alphanumérique avec tirets autorisés
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

' Fonction pour obtenir des informations complémentaires sur l'article
Public Function ObtenirInfosComplementairesArticle(codeArticle As String) As String
    On Error GoTo ErrorHandler
    
    ' Les informations sont récupérées dans la requête principale
    ObtenirInfosComplementairesArticle = "Informations récupérées depuis ART_PAR et NSE_DAT"
    Exit Function
    
ErrorHandler:
    ObtenirInfosComplementairesArticle = "Erreur lors de la récupération des infos: " & Err.description
End Function

' Fonction pour valider les données SAV
Public Function ValiderDonnees(donnees As TypeSAV) As Boolean
    ' Validation de base
    ValiderDonnees = True
    
    ' Vérifier le numéro de série avec la requête BDD
    If VerifierConnexionBDD() And Len(donnees.ReferenceProduit) > 0 Then
        ValiderDonnees = VerifierNumeroSerieBDD(donnees.ReferenceProduit)
    End If
    
    ' Vérifications supplémentaires
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
    
    ' Vérifications de base
    If Len(codeBarre) < 6 Then
        ValiderCodeBarre = False
        Exit Function
    End If
    
    ' Vérifie le format basique (lettres + chiffres + tirets)
    Dim i As Integer
    For i = 1 To Len(codeBarre)
        Dim char As String
        char = Mid(codeBarre, i, 1)
        If Not ((char >= "A" And char <= "Z") Or (char >= "0" And char <= "9") Or char = "-") Then
            ValiderCodeBarre = False
            Exit Function
        End If
    Next i
    
    ' Vérifier avec la logique BDD
    If VerifierConnexionBDD() Then
        ValiderCodeBarre = VerifierNumeroSerieBDD(codeBarre)
    Else
        ' Si pas de BDD, validation basique seulement
        ValiderCodeBarre = True
    End If
End Function

' Fonction pour extraire le modèle du code-barres
Public Function ExtraireModele(codeBarre As String) As String
    On Error GoTo ErrorLocal
    
    ' Essayer d'abord depuis la BDD
    If VerifierConnexionBDD() Then
        Dim resultats As TypeValidationBDD
        resultats = ValiderNumeroSerieBDD(codeBarre)
        
        If resultats.existe Then
            ' Utiliser la désignation obtenue de la requête BDD
            ExtraireModele = resultats.designationArticle
            Exit Function
        End If
    End If

ErrorLocal:
    ' Fallback - identification par préfixe local si pas trouvé en BDD
    ExtraireModele = DeterminerModeleParCode(codeBarre)
End Function

' Fonction pour déterminer le modèle basé sur le code (fallback)
Private Function DeterminerModeleParCode(code As String) As String
    Dim prefixe As String
    prefixe = Left(UCase(code), 6)
    
    Select Case prefixe
        Case "VC2286"
            DeterminerModeleParCode = "Vitrine VC2286"
        Case "RB4458"
            DeterminerModeleParCode = "Red Bull RB4458"
        Case "CF3401"
            DeterminerModeleParCode = "Congélateur CF3401"
        Case "RB2024"
            DeterminerModeleParCode = "Red Bull Premium 2024"
        Case Else
            ' Essayer avec les 2 premiers caractères
            Select Case Left(UCase(code), 2)
                Case "VC"
                    DeterminerModeleParCode = "Vitrine Red Bull"
                Case "RB"
                    DeterminerModeleParCode = "Frigo Red Bull"
                Case "CF"
                    DeterminerModeleParCode = "Congélateur Red Bull"
                Case "FB"
                    DeterminerModeleParCode = "Frigo Bar Red Bull"
                Case "RF"
                    DeterminerModeleParCode = "Red Fridge"
                Case Else
                    DeterminerModeleParCode = "Équipement Red Bull - Modèle non identifié"
            End Select
    End Select
End Function

' === FONCTIONS DE VALIDATION SPÉCIALISÉES ===

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

' Fonction pour valider un numéro (uniquement chiffres)
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
