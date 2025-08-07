VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} datecalcul 
   Caption         =   "UserForm1"
   ClientHeight    =   8316.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10284
   OleObjectBlob   =   "reste_a_fairev2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "datecalcul"
Attribute VB_GlobalNamespace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Button_Calculer_Click()
    Dim qtRestante As Double
    Dim rendement As Double
    Dim personnes As Double
    Dim heuresJour As Double
    Dim travailTotal As Double
    
    If Not IsNumeric(TextBox_QuantitePosee.Value) Or _
       Not IsNumeric(TextBox_Rendement.Value) Or _
       Not IsNumeric(TextBox_P.Value) Or _
       Not IsNumeric(TextBox_Heures.Value) Then
        MsgBox "Veuillez remplir tous les champs avec des valeurs valides.", vbExclamation
        Exit Sub
    End If
    
    qtRestante = CDbl(TextBox_QuantitePosee.Value)
    rendement = CDbl(TextBox_Rendement.Value)
    personnes = CDbl(TextBox_P.Value)
    heuresJour = CDbl(TextBox_Heures.Value)
    
    If qtRestante <= 0 Or rendement <= 0 Or personnes <= 0 Or heuresJour <= 0 Then
        MsgBox "Toutes les valeurs doivent être positives", vbCritical
        Exit Sub
    End If
    
    travailTotal = (qtRestante / rendement) * heuresJour * personnes
    Label_Resultat.Caption = "📊 Travail estimé : " & Format(travailTotal, "0.0") & " heures"
    
    ' Jours fériés France 2025
    Dim joursFeries As Variant
    joursFeries = Array(CDate("01/01/2025"), CDate("21/04/2025"), CDate("22/04/2025"), _
                        CDate("01/05/2025"), CDate("08/05/2025"), CDate("29/05/2025"), _
                        CDate("09/06/2025"), CDate("14/07/2025"), CDate("15/08/2025"), _
                        CDate("01/11/2025"), CDate("11/11/2025"), CDate("25/12/2025"))
    
    If IsDate(TextBox_DD.Value) Then
        Dim joursRestants As Double
        Dim dateDebut As Date
        Dim dateFin As Date
        
        joursRestants = (qtRestante / rendement)
        dateDebut = CDate(TextBox_DD.Value)
        dateFin = CalculeDateFinOuvree(dateDebut, RoundUp(joursRestants, 0), joursFeries)
        
        Label_DateFinEstimee.Caption = "📅 Fin estimée : " & Format(dateFin, "dd/mm/yyyy")
    Else
        Label_DateFinEstimee.Caption = ""
    End If

    If IsDate(TextBox_DFS.Value) And IsDate(TextBox_DD.Value) Then
        Dim dateFinSouhaitee As Date
        Dim joursOuvres As Double
        Dim nbPerso As Double
        dateFinSouhaitee = CDate(TextBox_DFS.Value)
        dateDebut = CDate(TextBox_DD.Value)

        joursOuvres = NbJoursOuvres(dateDebut, dateFinSouhaitee, joursFeries)
        
        If joursOuvres > 0 Then
            nbPerso = travailTotal / (joursOuvres * heuresJour)
            nbPerso = Int(nbPerso) + IIf(nbPerso > Int(nbPerso), 1, 0)
            Label_PersosNecessaires.Caption = "👥 Pour tenir jusqu'au " & Format(dateFinSouhaitee, "dd/mm/yyyy") & _
            " (" & Format(joursOuvres, "0") & " jours ouvrés) il faut : " & nbPerso & " personnes"
        Else
            Label_PersosNecessaires.Caption = "❌ Date de fin souhaitée invalide"
        End If
    Else
        Label_PersosNecessaires.Caption = ""
    End If
End Sub

Private Sub Button_Fermer_Click()
    Dim texteComplet As String
    Dim valeurNumerique As String

    If Label_PersosNecessaires.Caption <> "" And InStr(Label_PersosNecessaires.Caption, "personnes") > 0 Then
        texteComplet = Label_PersosNecessaires.Caption
        Dim parties() As String
        parties = Split(texteComplet, " ")
        If UBound(parties) >= 1 Then
            valeurNumerique = parties(UBound(parties) - 1)
        End If
    ElseIf Label_Resultat.Caption <> "" And InStr(Label_Resultat.Caption, "heures") > 0 Then
        texteComplet = Label_Resultat.Caption
        Dim parties2() As String
        parties2 = Split(texteComplet, " ")
        If UBound(parties2) >= 1 Then
            valeurNumerique = parties2(UBound(parties2) - 1)
        End If
    Else
        MsgBox "Aucun résultat à copier. Veuillez d'abord calculer.", vbExclamation
        Exit Sub
    End If

    If valeurNumerique <> "" Then
        Dim dataObj As Object
        Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        dataObj.SetText valeurNumerique
        dataObj.PutInClipboard
        Unload Me
    Else
        MsgBox "Impossible de trouver la valeur numérique.", vbExclamation
    End If
End Sub

Private Function RoundUp(ByVal val As Double, ByVal decimals As Integer) As Double
    RoundUp = -Int(-val * 10 ^ decimals) / 10 ^ decimals
End Function

Private Function CalculeDateFinOuvree(ByVal dateDebut As Date, ByVal jours As Long, ByVal joursFeries As Variant) As Date
    Dim i As Long
    Dim dateCourante As Date
    dateCourante = dateDebut
    i = 0
    Do While i < jours
        dateCourante = dateCourante + 1
        If Weekday(dateCourante, vbMonday) <= 5 Then
            Dim estFerie As Boolean: estFerie = False
            Dim j As Long
            For j = LBound(joursFeries) To UBound(joursFeries)
                If joursFeries(j) = dateCourante Then
                    estFerie = True
                    Exit For
                End If
            Next j
            If Not estFerie Then
                i = i + 1
            End If
        End If
    Loop
    CalculeDateFinOuvree = dateCourante
End Function

Private Function NbJoursOuvres(date1 As Date, date2 As Date, joursFeries As Variant) As Long
    Dim d As Date
    Dim cpt As Long
    For d = date1 To date2
        If Weekday(d, vbMonday) <= 5 Then
            Dim estFerie As Boolean: estFerie = False
            Dim j As Long
            For j = LBound(joursFeries) To UBound(joursFeries)
                If joursFeries(j) = d Then
                    estFerie = True
                    Exit For
                End If
            Next j
            If Not estFerie Then
                cpt = cpt + 1
            End If
        End If
    Next d
    NbJoursOuvres = cpt
End Function
