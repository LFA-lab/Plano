Private Sub Button_Calculer_Click()
    Dim qtRestante As Double
    Dim rendement As Double
    Dim personnes As Double
    Dim heuresJour As Double
    Dim travailTotal As Double
    Dim joursFeries As Variant
    
    ' Jours f?ri?s fran?ais 2025
    joursFeries = Array(CDate("01/01/2025"), CDate("21/04/2025"), CDate("22/04/2025"), _
                       CDate("01/05/2025"), CDate("08/05/2025"), CDate("29/05/2025"), _
                       CDate("09/06/2025"), CDate("14/07/2025"), CDate("15/08/2025"), _
                       CDate("01/11/2025"), CDate("11/11/2025"), CDate("25/12/2025"))
    
    ' V?rification des champs obligatoires
    If Not IsNumeric(TextBox_QuantitePosee.Value) Or _
       Not IsNumeric(TextBox_Rendement.Value) Or _
       Not IsNumeric(TextBox_P.Value) Or _
       Not IsNumeric(TextBox_Heures.Value) Then
       
        MsgBox "Veuillez remplir tous les champs avec des valeurs valides.", vbExclamation
        Exit Sub
    End If
    
    ' R?cup?ration des valeurs
    qtRestante = CDbl(TextBox_QuantitePosee.Value)
    rendement = CDbl(TextBox_Rendement.Value)
    personnes = CDbl(TextBox_P.Value)
    heuresJour = CDbl(TextBox_Heures.Value)
    
    If qtRestante <= 0 Or rendement <= 0 Or personnes <= 0 Or heuresJour <= 0 Then
        MsgBox "Toutes les valeurs doivent ?tre positives", vbCritical
        Exit Sub
    End If
    
    ' Calcul : travail total ?quipe
    travailTotal = (qtRestante / rendement) * heuresJour * personnes
    Label_Resultat.Caption = "Travail estim? : " & Format(travailTotal, "0.0") & " heures"
    
    ' Date estim?e de fin
    If IsDate(TextBox_DD.Value) Then
        Dim joursRestants As Double
        Dim dateDebut As Date
        Dim dateFin As Date
        joursRestants = (qtRestante / rendement) * 7 / 5
        dateDebut = CDate(TextBox_DD.Value)
        
        ' Calcul de la date de fin en excluant week-ends et jours f?ri?s
        On Error Resume Next
        dateFin = Application.WorksheetFunction.WorkDay(dateDebut, Int(joursRestants), joursFeries)
        On Error GoTo 0
        
        ' Si WorkDay ?choue, calcul simple en fallback
        If dateFin = 0 Then
            dateFin = dateDebut + joursRestants
        End If
        
        Label_DateFinEstimee.Caption = "Fin estim?e : " & Format(dateFin, "dd/mm/yyyy")
    Else
        Label_DateFinEstimee.Caption = ""
    End If
    
    ' CHEMIN D : Back-calculation avec NETWORKDAYS
    If IsDate(TextBox_DFS.Value) And IsDate(TextBox_DD.Value) Then
        Dim dateFinSouhaitee As Date
        Dim joursOuvres As Double
        Dim itmJourNecessaire As Double
        Dim ratioPersonnes As Double
        Dim personnesNecessaires As Double
        
        dateFinSouhaitee = CDate(TextBox_DFS.Value)
        dateDebut = CDate(TextBox_DD.Value)
        
        ' NETWORKDAYS : calcul des jours ouvr?s r?els
        On Error Resume Next
        joursOuvres = Application.WorksheetFunction.NetworkDays(dateDebut, dateFinSouhaitee, joursFeries)
        On Error GoTo 0
        
        ' Si NETWORKDAYS ?choue, calcul simple
        If joursOuvres = 0 Then
            joursOuvres = dateFinSouhaitee - dateDebut
            joursOuvres = joursOuvres * 5 / 7
        End If
        
        If joursOuvres > 0 Then
            ' Calcul direct : combien de personnes pour finir ? temps
            itmJourNecessaire = qtRestante / joursOuvres
            ratioPersonnes = itmJourNecessaire / rendement
            Dim brut As Double
            brut = ratioPersonnes * personnes
            personnesNecessaires = Int(brut)
            If brut > personnesNecessaires Then personnesNecessaires = personnesNecessaires + 1
            
            ' Calcul du rendement cible
            Dim rendementCible As Double
            rendementCible = qtRestante / joursOuvres
            
            Label_PersosNecessaires.Caption = "Pour finir avant le " & Format(dateFinSouhaitee, "dd/mm/yyyy") & _
            " (" & Format(joursOuvres, "0") & " jours ouvr?s)" & vbCrLf & _
            "? ?? " & personnesNecessaires & " personnes" & vbCrLf & _
            "? ?? Rendement attendu : " & Format(rendementCible, "0.0") & " itm/jour"
        Else
            Label_PersosNecessaires.Caption = "Date de fin souhait?e invalide"
        End If
    Else
        Label_PersosNecessaires.Caption = ""
    End If
End Sub

Private Sub Button_Fermer_Click()
    Dim texteComplet As String
    Dim valeurNumerique As String
    
    ' V?rifier uniquement Label_Resultat et extraire la valeur num?rique avant "heures"
    If Label_Resultat.Caption <> "" And InStr(Label_Resultat.Caption, "heures") > 0 Then
        texteComplet = Label_Resultat.Caption
        
        ' Extraire la valeur num?rique avant le mot "heures"
        Dim posHeures As Integer
        posHeures = InStr(texteComplet, "heures")
        
        If posHeures > 0 Then
            ' Extraire la partie avant "heures" et chercher le dernier nombre
            Dim partieAvantHeures As String
            partieAvantHeures = Left(texteComplet, posHeures - 1)
            partieAvantHeures = Trim(partieAvantHeures)
            
            ' Diviser par espaces et prendre le dernier ?l?ment
            Dim parties() As String
            parties = Split(partieAvantHeures, " ")
            If UBound(parties) >= 0 Then
                valeurNumerique = parties(UBound(parties))
            End If
        End If
    Else
        MsgBox "Aucun r?sultat ? copier. Veuillez d'abord calculer.", vbExclamation
        Exit Sub
    End If
    
    ' Copier dans le presse-papiers
    If valeurNumerique <> "" Then
        Dim dataObj As Object
        Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        dataObj.SetText valeurNumerique
        dataObj.PutInClipboard
        
        ' Fermer le UserForm
        Unload Me
    Else
        MsgBox "Impossible de trouver la valeur num?rique.", vbExclamation
    End If
End Sub







