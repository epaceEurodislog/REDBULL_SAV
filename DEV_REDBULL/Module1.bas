Attribute VB_Name = "Module1"
Attribute VB_Name = "Module1"
Option Explicit

' Type pour définir un équipement
Public Type Equipement
    ID As Long
    typeEq As String
    Modele As String
    statut As String
    DateOperation As Date
    Destination As String
    Remarques As String
    priorite As String
    Technicien As String
End Type

' Constantes pour les statuts
Public Const STATUT_RECEPTION = "Réception"
Public Const STATUT_STOCK = "Stock"
Public Const STATUT_PREPARATION = "Préparation"
Public Const STATUT_EXPEDITION = "Expédition"
Public Const STATUT_RETOUR = "Retour"

' Statuts de réparation
Public Const STATUT_DIAGNOSTIC = "Diagnostic"
Public Const STATUT_ATTENTE_PIECES = "Attente Pièces"
Public Const STATUT_REPARABLE = "Réparable"
Public Const STATUT_DONNEUR_PIECES = "Donneur Pièces"
Public Const STATUT_ATELIER = "Atelier"
Public Const STATUT_STOCK_PRET = "Stock Prêt"

' Types d'équipements
Public Const TYPE_FRIGO = "Frigo"
Public Const TYPE_DISTRIBUTEUR = "Distributeur"
Public Const TYPE_PRESENTOIR = "Présentoir"

' Priorités
Public Const PRIORITE_HAUTE = "Haute"
Public Const PRIORITE_NORMALE = "Normale"
Public Const PRIORITE_BASSE = "Basse"

' Fonction pour valider un équipement
Public Function ValiderEquipement(eq As Equipement) As String
    Dim erreurs As String
    
    ' Vérifier les champs obligatoires
    If Trim(eq.typeEq) = "" Then
        erreurs = erreurs & "- Le type d'équipement est obligatoire" & vbCrLf
    End If
    
    If Trim(eq.Modele) = "" Then
        erreurs = erreurs & "- Le modèle est obligatoire" & vbCrLf
    End If
    
    If Trim(eq.statut) = "" Then
        erreurs = erreurs & "- Le statut est obligatoire" & vbCrLf
    End If
    
    If Trim(eq.Destination) = "" Then
        erreurs = erreurs & "- La destination est obligatoire" & vbCrLf
    End If
    
    ' Vérifier la validité des valeurs
    If Not StatutValide(eq.statut) Then
        erreurs = erreurs & "- Statut non valide: " & eq.statut & vbCrLf
    End If
    
    If Not TypeValide(eq.typeEq) Then
        erreurs = erreurs & "- Type d'équipement non valide: " & eq.typeEq & vbCrLf
    End If
    
    ' Vérifications spécifiques pour les équipements en réparation
    If EstStatutReparation(eq.statut) Then
        If Trim(eq.Technicien) = "" Then
            erreurs = erreurs & "- Un technicien doit être assigné pour les réparations" & vbCrLf
        End If
        
        If Trim(eq.priorite) = "" Then
            erreurs = erreurs & "- Une priorité doit être définie pour les réparations" & vbCrLf
        ElseIf Not PrioriteValide(eq.priorite) Then
            erreurs = erreurs & "- Priorité non valide: " & eq.priorite & vbCrLf
        End If
    End If
    
    ValiderEquipement = erreurs
End Function

' Fonction pour vérifier si un statut est valide
Public Function StatutValide(statut As String) As Boolean
    Select Case statut
        Case STATUT_RECEPTION, STATUT_STOCK, STATUT_PREPARATION, STATUT_EXPEDITION, STATUT_RETOUR, _
             STATUT_DIAGNOSTIC, STATUT_ATTENTE_PIECES, STATUT_REPARABLE, STATUT_DONNEUR_PIECES, _
             STATUT_ATELIER, STATUT_STOCK_PRET
            StatutValide = True
        Case Else
            StatutValide = False
    End Select
End Function

' Fonction pour vérifier si un type d'équipement est valide
Public Function TypeValide(typeEq As String) As Boolean
    Select Case typeEq
        Case TYPE_FRIGO, TYPE_DISTRIBUTEUR, TYPE_PRESENTOIR
            TypeValide = True
        Case Else
            TypeValide = False
    End Select
End Function

' Fonction pour vérifier si une priorité est valide
Public Function PrioriteValide(priorite As String) As Boolean
    Select Case priorite
        Case PRIORITE_HAUTE, PRIORITE_NORMALE, PRIORITE_BASSE
            PrioriteValide = True
        Case Else
            PrioriteValide = False
    End Select
End Function

' Fonction pour vérifier si un statut correspond à une réparation
Public Function EstStatutReparation(statut As String) As Boolean
    Select Case statut
        Case STATUT_DIAGNOSTIC, STATUT_ATTENTE_PIECES, STATUT_REPARABLE, _
             STATUT_DONNEUR_PIECES, STATUT_ATELIER, STATUT_STOCK_PRET
            EstStatutReparation = True
        Case Else
            EstStatutReparation = False
    End Select
End Function

' Fonction pour obtenir la couleur d'un statut
Public Function CouleurStatut(statut As String) As Long
    Select Case statut
        Case STATUT_RECEPTION
            CouleurStatut = &H80FF80   ' Vert clair
        Case STATUT_STOCK
            CouleurStatut = &H8080FF   ' Bleu
        Case STATUT_PREPARATION
            CouleurStatut = &HFFFF80   ' Jaune
        Case STATUT_EXPEDITION
            CouleurStatut = &HFF8080   ' Rouge clair
        Case STATUT_RETOUR
            CouleurStatut = &HFF8040   ' Orange clair
        Case STATUT_DIAGNOSTIC
            CouleurStatut = &HFF8000   ' Orange
        Case STATUT_ATTENTE_PIECES
            CouleurStatut = &H8000FF   ' Violet
        Case STATUT_REPARABLE
            CouleurStatut = &H80FFFF   ' Cyan
        Case STATUT_DONNEUR_PIECES
            CouleurStatut = &H400040   ' Violet foncé
        Case STATUT_ATELIER
            CouleurStatut = &H4080FF   ' Bleu orange
        Case STATUT_STOCK_PRET
            CouleurStatut = &H40FF40   ' Vert
        Case Else
            CouleurStatut = &H808080   ' Gris
    End Select
End Function

' Fonction pour générer un ID unique
Public Function GenererID() As String
    GenererID = "RB" & Format(Date, "yyyymmdd") & Format(Time, "hhmmss")
End Function

' Fonction pour formater une date
Public Function FormaterDate(dateVal As Date) As String
    FormaterDate = Format(dateVal, "dd/mm/yyyy")
End Function

' Fonction pour formater une date avec heure
Public Function FormaterDateHeure(dateVal As Date) As String
    FormaterDateHeure = Format(dateVal, "dd/mm/yyyy hh:mm")
End Function

' Fonction pour obtenir le prochain statut logique
Public Function ProchainStatut(statutActuel As String) As String
    Select Case statutActuel
        Case STATUT_RECEPTION
            ProchainStatut = STATUT_STOCK
        Case STATUT_STOCK
            ProchainStatut = STATUT_PREPARATION
        Case STATUT_PREPARATION
            ProchainStatut = STATUT_EXPEDITION
        Case STATUT_EXPEDITION
            ProchainStatut = STATUT_RETOUR
        Case STATUT_DIAGNOSTIC
            ProchainStatut = STATUT_REPARABLE
        Case STATUT_REPARABLE
            ProchainStatut = STATUT_ATELIER
        Case STATUT_ATELIER
            ProchainStatut = STATUT_STOCK_PRET
        Case Else
            ProchainStatut = statutActuel
    End Select
End Function

' Fonction pour obtenir la description d'un statut
Public Function DescriptionStatut(statut As String) As String
    Select Case statut
        Case STATUT_RECEPTION
            DescriptionStatut = "Équipement en cours de réception"
        Case STATUT_STOCK
            DescriptionStatut = "Équipement stocké et disponible"
        Case STATUT_PREPARATION
            DescriptionStatut = "Préparation pour expédition"
        Case STATUT_EXPEDITION
            DescriptionStatut = "En cours de livraison"
        Case STATUT_RETOUR
            DescriptionStatut = "Retour client"
        Case STATUT_DIAGNOSTIC
            DescriptionStatut = "Diagnostic technique en cours"
        Case STATUT_ATTENTE_PIECES
            DescriptionStatut = "En attente de pièces détachées"
        Case STATUT_REPARABLE
            DescriptionStatut = "Équipement réparable identifié"
        Case STATUT_DONNEUR_PIECES
            DescriptionStatut = "Utilisé comme donneur de pièces"
        Case STATUT_ATELIER
            DescriptionStatut = "Réparation en cours"
        Case STATUT_STOCK_PRET
            DescriptionStatut = "Réparé et prêt à expédier"
        Case Else
            DescriptionStatut = "Statut inconnu"
    End Select
End Function

' Fonction pour créer un rapport de statut
Public Function CreerRapportStatut(equipements() As Equipement) As String
    Dim rapport As String
    Dim i As Integer
    Dim compteurs(10) As Integer ' Tableau pour compter chaque statut
    
    rapport = "=== RAPPORT DE STATUT ===" & vbCrLf
    rapport = rapport & "Généré le: " & FormaterDateHeure(Now) & vbCrLf & vbCrLf
    
    ' Compter les équipements par statut
    For i = 0 To UBound(equipements)
        Select Case equipements(i).statut
            Case STATUT_RECEPTION: compteurs(0) = compteurs(0) + 1
            Case STATUT_STOCK: compteurs(1) = compteurs(1) + 1
            Case STATUT_PREPARATION: compteurs(2) = compteurs(2) + 1
            Case STATUT_EXPEDITION: compteurs(3) = compteurs(3) + 1
            Case STATUT_RETOUR: compteurs(4) = compteurs(4) + 1
            Case STATUT_DIAGNOSTIC: compteurs(5) = compteurs(5) + 1
            Case STATUT_ATTENTE_PIECES: compteurs(6) = compteurs(6) + 1
            Case STATUT_REPARABLE: compteurs(7) = compteurs(7) + 1
            Case STATUT_DONNEUR_PIECES: compteurs(8) = compteurs(8) + 1
            Case STATUT_ATELIER: compteurs(9) = compteurs(9) + 1
            Case STATUT_STOCK_PRET: compteurs(10) = compteurs(10) + 1
        End Select
    Next i
    
    ' Génération du rapport
    rapport = rapport & "PROCESSUS PRINCIPAL:" & vbCrLf
    rapport = rapport & "- Réception: " & compteurs(0) & " équipements" & vbCrLf
    rapport = rapport & "- Stock: " & compteurs(1) & " équipements" & vbCrLf
    rapport = rapport & "- Préparation: " & compteurs(2) & " équipements" & vbCrLf
    rapport = rapport & "- Expédition: " & compteurs(3) & " équipements" & vbCrLf
    rapport = rapport & "- Retour: " & compteurs(4) & " équipements" & vbCrLf & vbCrLf
    
    rapport = rapport & "SERVICE RÉPARATION:" & vbCrLf
    rapport = rapport & "- Diagnostic: " & compteurs(5) & " équipements" & vbCrLf
    rapport = rapport & "- Attente pièces: " & compteurs(6) & " équipements" & vbCrLf
    rapport = rapport & "- Réparable: " & compteurs(7) & " équipements" & vbCrLf
    rapport = rapport & "- Donneur pièces: " & compteurs(8) & " équipements" & vbCrLf
    rapport = rapport & "- Atelier: " & compteurs(9) & " équipements" & vbCrLf
    rapport = rapport & "- Stock prêt: " & compteurs(10) & " équipements" & vbCrLf & vbCrLf
    
    rapport = rapport & "TOTAL: " & (UBound(equipements) + 1) & " équipements"
    
    CreerRapportStatut = rapport
End Function

' Fonction utilitaire pour nettoyer une chaîne
Public Function NettoyerChaine(texte As String) As String
    NettoyerChaine = Trim(Replace(Replace(texte, vbCrLf, " "), vbTab, " "))
End Function

' Fonction pour exporter vers CSV (simulation)
Public Function ExporterCSV(equipements() As Equipement) As String
    Dim csv As String
    Dim i As Integer
    
    ' En-têtes
    csv = "ID,Type,Modele,Statut,Date,Destination,Remarques,Technicien,Priorite" & vbCrLf
    
    ' Données
    For i = 0 To UBound(equipements)
        With equipements(i)
            csv = csv & .ID & ","
            csv = csv & .typeEq & ","
            csv = csv & .Modele & ","
            csv = csv & .statut & ","
            csv = csv & FormaterDate(.DateOperation) & ","
            csv = csv & .Destination & ","
            csv = csv & """" & NettoyerChaine(.Remarques) & ""","
            csv = csv & .Technicien & ","
            csv = csv & .priorite & vbCrLf
        End With
    Next i
    
    ExporterCSV = csv
End Function

