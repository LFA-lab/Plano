'========================================================================
' MODULE: ExportHeuresSapin - Version Hybride Confirme/Previsionnel
' DESCRIPTION: Export heures avec distinction gains confirmes vs previsionnels
' AUTHOR: Copilot
' DATE: 18/08/2025
' VERSION: 4.0 - Analyse hybride pour pilotage temps reel optimal
'========================================================================

' Variables de configuration
Private Const SEUIL_REL As Double = 0.03  ' 3%
Private Const SEUIL_ABS_H As Double = 2   ' 2 heures
Private Const SEUIL_MIN_AVANCEMENT As Double = 20  ' 20% minimum pour analyse previsionnelle
Private Const PERCENT_COMPLETE As Double = 100  ' 100% pour taches terminees

'========================================================================
' MACRO PRINCIPALE
'========================================================================
Sub ExportHeuresSapin()
    ' Variables principales
    Dim xlApp As Object, xlWb As Object
    Dim wsResume As Object, wsSapin As Object, wsLots As Object
    Dim DateEtat As Date
    Dim nomProjet As String, anneeSemaine As String, cheminExport As String, nomFichier As String
    
    ' Recuperation de la date d'etat
    If IsDate(ActiveProject.StatusDate) And ActiveProject.StatusDate > #1/1/1900# Then
        DateEtat = ActiveProject.StatusDate
    Else
        DateEtat = Date
    End If
    
    ' Creation Excel et feuilles
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    xlApp.DisplayAlerts = False
    Set xlWb = xlApp.Workbooks.Add
    
    ' Creation des 3 feuilles
    Do While xlWb.Worksheets.Count < 3
        xlWb.Worksheets.Add After:=xlWb.Worksheets(xlWb.Worksheets.Count)
    Loop
    Do While xlWb.Worksheets.Count > 3
        xlWb.Worksheets(1).Delete
    Loop
    
    ' Nommage des feuilles
    xlWb.Worksheets(1).Name = "RESUME_DIRIGEANT"
    xlWb.Worksheets(2).Name = "SAPIN"
    xlWb.Worksheets(3).Name = "RECAP_LOTS"
    
    Set wsResume = xlWb.Worksheets("RESUME_DIRIGEANT")
    Set wsSapin = xlWb.Worksheets("SAPIN")
    Set wsLots = xlWb.Worksheets("RECAP_LOTS")
    
    ' Appel des 3 fonctions principales
    CollecteEtExportTaches wsSapin
    CalculEtExportLots wsLots
    ResumeDirigeant wsResume, DateEtat
    
    ' Application du formatage
    FormaterSapin wsSapin
    FormaterLots wsLots
    FormaterResume wsResume
    
    ' Sauvegarde
    nomProjet = ActiveProject.Name
    If InStr(nomProjet, ".") > 0 Then
        nomProjet = Left(nomProjet, InStrRev(nomProjet, ".") - 1)
    End If
    
    Dim numeroSemaine As Integer
    numeroSemaine = Format(DateEtat, "ww")
    anneeSemaine = Year(DateEtat) & "-W" & Format(numeroSemaine, "00")
    
    cheminExport = ActiveProject.Path & "\Optimisation\Exports\"
    On Error Resume Next
    CreateObject("Scripting.FileSystemObject").CreateFolder ActiveProject.Path & "\Optimisation"
    CreateObject("Scripting.FileSystemObject").CreateFolder cheminExport
    On Error GoTo 0
    
    nomFichier = cheminExport & nomProjet & "_" & anneeSemaine & ".xlsx"
    xlWb.SaveAs nomFichier
    
    ' Selection feuille resume
    wsResume.Activate
    wsResume.Range("A1").Select
    
    MsgBox "Export heures sapin termine !" & vbCrLf & _
           "Fichier: " & nomFichier, vbInformation, "Export Heures Sapin"
End Sub
'========================================================================
' FONCTIONS PRINCIPALES
'========================================================================

'------------------------------------------------------------------------
' Collecte toutes les taches et les exporte dans la feuille SAPIN avec analyse hybride
'------------------------------------------------------------------------
Sub CollecteEtExportTaches(ByRef wsSapin As Object)
    ' En-tetes enrichies pour analyse hybride
    wsSapin.Range("A1:P1").Value = Array("ID", "WBS", "Niveau", "Tache", "%C", "Baseline", "Actual", "Remaining", "Ecart_h", "Gain_Confirme", "Perte_Confirmee", "Gain_Previsionnel", "Perte_Previsionnelle", "Performance", "Flag", "Taux_Conso")
    
    Dim t As Task
    Dim ligne As Long
    Dim BaselineH As Double, ActualH As Double, RemainingH As Double
    Dim EcartH As Double, GainConfirme As Double, PerteConfirmee As Double
    Dim GainPrevisionnel As Double, PertePrevisionnelle As Double
    Dim Flag As String, Performance As String, TauxConso As Double
    
    ligne = 2
    
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            ' Recuperation des valeurs (conversion minutes -> heures)
            BaselineH = 0
            ActualH = 0
            RemainingH = 0
            
            If Not IsNull(t.BaselineWork) Then BaselineH = t.BaselineWork / 60
            If Not IsNull(t.ActualWork) Then ActualH = t.ActualWork / 60
            If Not IsNull(t.RemainingWork) Then RemainingH = t.RemainingWork / 60
            
            ' Calculs de base
            EcartH = ActualH - BaselineH
            GainConfirme = 0
            PerteConfirmee = 0
            GainPrevisionnel = 0
            PertePrevisionnelle = 0
            Flag = ""
            Performance = ""
            TauxConso = 0
            
            ' ANALYSE POUR TACHES TERMINEES (100%) - CONFIRME
            If BaselineH > 0 And t.PercentComplete >= PERCENT_COMPLETE Then
                If (BaselineH - ActualH) >= SEUIL_ABS_H And (BaselineH - ActualH) / BaselineH >= SEUIL_REL Then
                    GainConfirme = BaselineH - ActualH
                    Flag = "OK"
                    Performance = "Efficace"
                End If
                If (ActualH - BaselineH) >= SEUIL_ABS_H And (ActualH - BaselineH) / BaselineH >= SEUIL_REL Then
                    PerteConfirmee = ActualH - BaselineH
                    Flag = "ALERTE"
                    Performance = "Derive"
                End If
                If GainConfirme = 0 And PerteConfirmee = 0 Then
                    Performance = "Nominal"
                End If
            
            ' ANALYSE POUR TACHES EN COURS (20%-99%) - PREVISIONNEL
            ElseIf BaselineH > 0 And t.PercentComplete >= SEUIL_MIN_AVANCEMENT And t.PercentComplete < PERCENT_COMPLETE Then
                ' Calcul du taux de consommation
                Dim AvancementDecimal As Double
                AvancementDecimal = t.PercentComplete / 100
                
                If AvancementDecimal > 0 Then
                    TauxConso = ActualH / (BaselineH * AvancementDecimal)
                    
                    ' Projection du total final estime
                    Dim TotalEstime As Double
                    TotalEstime = BaselineH * TauxConso
                    
                    ' Ecart previsionnel
                    Dim EcartPrevisionnel As Double
                    EcartPrevisionnel = TotalEstime - BaselineH
                    
                    ' Classification previsionnelle
                    If Abs(EcartPrevisionnel) >= SEUIL_ABS_H And Abs(EcartPrevisionnel) / BaselineH >= SEUIL_REL Then
                        If EcartPrevisionnel < 0 Then ' Gain previsionnel
                            GainPrevisionnel = Abs(EcartPrevisionnel)
                            Flag = "TEND+"
                            Performance = "Tend. Efficace"
                        Else ' Perte previsionnelle
                            PertePrevisionnelle = EcartPrevisionnel
                            Flag = "TEND-"
                            Performance = "Tend. Derive"
                        End If
                    Else
                        Performance = "Tend. Nominal"
                    End If
                End If
            
            ' TACHES PAS ENCORE ANALYSABLES
            ElseIf t.PercentComplete < SEUIL_MIN_AVANCEMENT Then
                Performance = "Trop tot"
            End If
            
            ' Ecriture ligne par ligne avec nouvelles colonnes
            wsSapin.Cells(ligne, 1).Value = t.ID
            wsSapin.Cells(ligne, 2).Value = t.WBS
            wsSapin.Cells(ligne, 3).Value = t.OutlineLevel
            wsSapin.Cells(ligne, 4).Value = t.Name
            wsSapin.Cells(ligne, 4).IndentLevel = t.OutlineLevel - 1  ' Indentation hierarchique
            wsSapin.Cells(ligne, 5).Value = t.PercentComplete / 100
            wsSapin.Cells(ligne, 6).Value = BaselineH
            wsSapin.Cells(ligne, 7).Value = ActualH
            wsSapin.Cells(ligne, 8).Value = RemainingH
            wsSapin.Cells(ligne, 9).Value = EcartH
            wsSapin.Cells(ligne, 10).Value = GainConfirme
            wsSapin.Cells(ligne, 11).Value = PerteConfirmee
            wsSapin.Cells(ligne, 12).Value = GainPrevisionnel
            wsSapin.Cells(ligne, 13).Value = PertePrevisionnelle
            wsSapin.Cells(ligne, 14).Value = Performance
            wsSapin.Cells(ligne, 15).Value = Flag
            wsSapin.Cells(ligne, 16).Value = TauxConso
            
            ligne = ligne + 1
        End If
    Next t
End Sub

'------------------------------------------------------------------------
' Calcule les agregations par lot avec analyse hybride et exporte dans RECAP_LOTS
'------------------------------------------------------------------------
Sub CalculEtExportLots(ByRef wsLots As Object)
    ' En-tetes enrichies
    wsLots.Range("A1:Q1").Value = Array("WBS", "Lot/Phase", "Baseline h", "PV_h", "EW", "Actual", "Remaining", "Ecart_h", "Gain_Confirme", "Perte_Confirmee", "Gain_Previsionnel", "Perte_Previsionnelle", "SPI_h", "CPI_h", "Fiabilite_%", "Performance", "Flag")
    
    ' Dictionnaire pour agregation par WBS niveau 1
    Dim dictLots As Object
    Set dictLots = CreateObject("Scripting.Dictionary")
    
    Dim t As Task
    Dim WBS1 As String
    Dim BaselineH As Double, ActualH As Double, RemainingH As Double
    Dim EcartH As Double, GainConfirme As Double, PerteConfirmee As Double
    Dim GainPrevisionnel As Double, PertePrevisionnelle As Double
    Dim PVH As Double, EW As Double
    Dim DateEtat As Date
    
    ' Date d'etat pour calcul PV
    If IsDate(ActiveProject.StatusDate) And ActiveProject.StatusDate > #1/1/1900# Then
        DateEtat = ActiveProject.StatusDate
    Else
        DateEtat = Date
    End If
    
    ' Collecte et agregation des donnees par lot
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing And Not t.Summary Then  ' Seulement taches non-recap
            ' Determination WBS niveau 1
            If InStr(t.WBS, ".") > 0 Then
                WBS1 = Left(t.WBS, InStr(t.WBS, ".") - 1)
            Else
                WBS1 = t.WBS
            End If
            
            ' Recuperation des valeurs
            BaselineH = 0
            ActualH = 0
            RemainingH = 0
            PVH = 0
            EW = 0
            GainConfirme = 0
            PerteConfirmee = 0
            GainPrevisionnel = 0
            PertePrevisionnelle = 0
            
            If Not IsNull(t.BaselineWork) Then BaselineH = t.BaselineWork / 60
            If Not IsNull(t.ActualWork) Then ActualH = t.ActualWork / 60
            If Not IsNull(t.RemainingWork) Then RemainingH = t.RemainingWork / 60
            
            ' Calcul PV et EW
            If BaselineH > 0 And Not IsNull(t.BaselineFinish) Then
                If t.BaselineFinish <= DateEtat Then PVH = BaselineH
            End If
            EW = BaselineH * t.PercentComplete / 100
            
            ' Calculs de base
            EcartH = ActualH - BaselineH
            
            ' ANALYSE HYBRIDE PAR TACHE
            ' Gains/Pertes confirmes (taches 100% terminees)
            If BaselineH > 0 And t.PercentComplete >= PERCENT_COMPLETE Then
                If (BaselineH - ActualH) >= SEUIL_ABS_H And (BaselineH - ActualH) / BaselineH >= SEUIL_REL Then
                    GainConfirme = BaselineH - ActualH
                End If
                If (ActualH - BaselineH) >= SEUIL_ABS_H And (ActualH - BaselineH) / BaselineH >= SEUIL_REL Then
                    PerteConfirmee = ActualH - BaselineH
                End If
            
            ' Gains/Pertes previsionnels (taches en cours 20%-99%)
            ElseIf BaselineH > 0 And t.PercentComplete >= SEUIL_MIN_AVANCEMENT And t.PercentComplete < PERCENT_COMPLETE Then
                Dim AvancementDecimal As Double
                AvancementDecimal = t.PercentComplete / 100
                
                If AvancementDecimal > 0 Then
                    Dim TauxConso As Double
                    TauxConso = ActualH / (BaselineH * AvancementDecimal)
                    
                    Dim TotalEstime As Double
                    TotalEstime = BaselineH * TauxConso
                    
                    Dim EcartPrevisionnel As Double
                    EcartPrevisionnel = TotalEstime - BaselineH
                    
                    If Abs(EcartPrevisionnel) >= SEUIL_ABS_H And Abs(EcartPrevisionnel) / BaselineH >= SEUIL_REL Then
                        If EcartPrevisionnel < 0 Then
                            GainPrevisionnel = Abs(EcartPrevisionnel)
                        Else
                            PertePrevisionnelle = EcartPrevisionnel
                        End If
                    End If
                End If
            End If
            
            ' Agregation dans le dictionnaire avec nouvelles metriques
            If dictLots.Exists(WBS1) Then
                Dim arrExistant As Variant
                arrExistant = dictLots(WBS1)
                arrExistant(2) = arrExistant(2) + BaselineH     ' Baseline
                arrExistant(3) = arrExistant(3) + PVH          ' PV
                arrExistant(4) = arrExistant(4) + EW           ' EW
                arrExistant(5) = arrExistant(5) + ActualH      ' Actual
                arrExistant(6) = arrExistant(6) + RemainingH   ' Remaining
                arrExistant(7) = arrExistant(7) + EcartH       ' Ecart
                arrExistant(8) = arrExistant(8) + GainConfirme       ' Gain confirme
                arrExistant(9) = arrExistant(9) + PerteConfirmee     ' Perte confirmee
                arrExistant(10) = arrExistant(10) + GainPrevisionnel  ' Gain previsionnel
                arrExistant(11) = arrExistant(11) + PertePrevisionnelle ' Perte previsionnelle
                ' Compteurs pour fiabilite
                arrExistant(16) = arrExistant(16) + 1  ' Total taches
                If t.PercentComplete >= PERCENT_COMPLETE Then arrExistant(17) = arrExistant(17) + 1  ' Taches terminees
                dictLots(WBS1) = arrExistant
            Else
                Dim arrNouveau(17) As Variant
                arrNouveau(0) = WBS1
                arrNouveau(1) = TrouverNomLot(WBS1)
                arrNouveau(2) = BaselineH
                arrNouveau(3) = PVH
                arrNouveau(4) = EW
                arrNouveau(5) = ActualH
                arrNouveau(6) = RemainingH
                arrNouveau(7) = EcartH
                arrNouveau(8) = GainConfirme
                arrNouveau(9) = PerteConfirmee
                arrNouveau(10) = GainPrevisionnel
                arrNouveau(11) = PertePrevisionnelle
                arrNouveau(16) = 1  ' Total taches
                If t.PercentComplete >= PERCENT_COMPLETE Then arrNouveau(17) = 1 Else arrNouveau(17) = 0  ' Taches terminees
                dictLots(WBS1) = arrNouveau
            End If
        End If
    Next t
    
    ' Ecriture des lots tries par impact total decroissant (pertes + pertes previsionnelles)
    Dim arrLots() As Variant
    Dim nbLots As Long, i As Long, j As Long
    Dim cle As Variant
    
    nbLots = dictLots.Count
    If nbLots = 0 Then Exit Sub
    
    ReDim arrLots(nbLots - 1, 16)
    i = 0
    
    ' Remplissage du tableau et calcul des metriques finales
    For Each cle In dictLots.Keys
        Dim lotData As Variant
        lotData = dictLots(cle)
        
        ' Calcul SPI et CPI
        If lotData(3) > 0 Then  ' PV > 0
            lotData(12) = lotData(4) / lotData(3)  ' SPI = EW / PV
        Else
            lotData(12) = "N/A"
        End If
        
        If lotData(5) > 0 Then  ' Actual > 0
            lotData(13) = lotData(4) / lotData(5)  ' CPI = EW / Actual
        Else
            lotData(13) = "N/A"
        End If
        
        ' Calcul fiabilite (% taches terminees)
        If lotData(16) > 0 Then
            lotData(14) = lotData(17) / lotData(16)  ' % fiabilite
        Else
            lotData(14) = 0
        End If
        
        ' Performance globale du lot
        Dim GainTotal As Double, PerteTotal As Double
        GainTotal = lotData(8) + lotData(10)  ' Confirme + Previsionnel
        PerteTotal = lotData(9) + lotData(11)  ' Confirme + Previsionnel
        
        If GainTotal > PerteTotal Then
            lotData(15) = "Performant"
        ElseIf PerteTotal > GainTotal Then
            If lotData(14) > 0.8 Then  ' Fiabilite elevee
                lotData(15) = "Derive Confirmee"
            Else
                lotData(15) = "Tend. Derive"
            End If
        Else
            lotData(15) = "Equilibre"
        End If
        
        ' Flag global
        If GainTotal > PerteTotal * 1.5 Then
            lotData(16) = "OK"
        ElseIf PerteTotal > GainTotal * 1.5 Then
            lotData(16) = "ALERTE"
        Else
            lotData(16) = "INFO"
        End If
        
        For j = 0 To 16
            arrLots(i, j) = lotData(j)
        Next j
        i = i + 1
    Next cle
    
    ' Tri par impact total decroissant (pertes confirmees + previsionnelles)
    For i = 0 To nbLots - 2
        For j = i + 1 To nbLots - 1
            Dim impactI As Double, impactJ As Double
            impactI = arrLots(i, 9) + arrLots(i, 11)  ' Pertes totales
            impactJ = arrLots(j, 9) + arrLots(j, 11)  ' Pertes totales
            
            If impactJ > impactI Then
                Dim temp As Variant
                For k = 0 To 16
                    temp = arrLots(i, k)
                    arrLots(i, k) = arrLots(j, k)
                    arrLots(j, k) = temp
                Next k
            End If
        Next j
    Next i
    
    ' Ecriture dans Excel
    For i = 0 To nbLots - 1
        For j = 0 To 16
            wsLots.Cells(i + 2, j + 1).Value = arrLots(i, j)
        Next j
    Next i
End Sub

'------------------------------------------------------------------------
' Cree le resume dirigeant optimise pour lecture en 1 minute
'------------------------------------------------------------------------
Sub ResumeDirigeant(ByRef wsResume As Object, ByVal DateEtat As Date)
    ' Verification et activation de la feuille
    On Error Resume Next
    wsResume.Activate
    On Error GoTo 0
    
    ' Variables pour collecte des donnees
    Dim dictLots As Object
    Set dictLots = CreateObject("Scripting.Dictionary")
    
    ' Variables globales pour calculs agreges
    Dim TotalPV As Double, TotalEW As Double, TotalAC As Double
    Dim TotalGainsConfirmes As Double, TotalPertesConfirmees As Double
    Dim ArrAnomalies() As String
    Dim nbAnomalies As Long
    nbAnomalies = 0
    
    ' Collections pour TOP 3
    Dim arrLotsSPI() As String, arrLotsCPI() As String
    Dim nbLotsSPI As Long, nbLotsCPI As Long
    nbLotsSPI = 0: nbLotsCPI = 0
    
    ' === COLLECTE DES DONNEES PAR LOT ===
    Dim t As Task
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing And Not t.Summary Then
            ' Determination WBS niveau 1 (lot)
            Dim WBS1 As String
            If InStr(t.WBS, ".") > 0 Then
                WBS1 = Left(t.WBS, InStr(t.WBS, ".") - 1)
            Else
                WBS1 = t.WBS
            End If
            
            ' Calculs de base (conversion minutes -> heures)
            Dim BaselineH As Double, ActualH As Double, EW As Double, PVH As Double
            BaselineH = 0: ActualH = 0: EW = 0: PVH = 0
            
            If Not IsNull(t.BaselineWork) Then BaselineH = t.BaselineWork / 60
            If Not IsNull(t.ActualWork) Then ActualH = t.ActualWork / 60
            EW = BaselineH * t.PercentComplete / 100
            
            ' Calcul PV
            If BaselineH > 0 And Not IsNull(t.BaselineFinish) Then
                If t.BaselineFinish <= DateEtat Then PVH = BaselineH
            End If
            
            ' Calculs gains/pertes confirmes (seulement taches 100% terminees)
            Dim GainConfirme As Double, PerteConfirmee As Double
            GainConfirme = 0: PerteConfirmee = 0
            
            If BaselineH > 0 And t.PercentComplete >= 100 Then ' Seulement taches terminees
                Dim EcartH As Double
                EcartH = ActualH - BaselineH
                If (BaselineH - ActualH) >= SEUIL_ABS_H And (BaselineH - ActualH) / BaselineH >= SEUIL_REL Then
                    GainConfirme = BaselineH - ActualH
                End If
                If (ActualH - BaselineH) >= SEUIL_ABS_H And (ActualH - BaselineH) / BaselineH >= SEUIL_REL Then
                    PerteConfirmee = ActualH - BaselineH
                End If
                
                ' Detection anomalies : Ecart_h < -50 ET Perte_Confirmee = 0
                If EcartH < -50 And PerteConfirmee = 0 Then
                    ReDim Preserve ArrAnomalies(nbAnomalies)
                    ArrAnomalies(nbAnomalies) = WBS1 & "|" & t.Name & "|" & Format(EcartH, "0") & " h"
                    nbAnomalies = nbAnomalies + 1
                End If
            End If
            
            ' Agregation par lot
            If dictLots.Exists(WBS1) Then
                Dim arrExistant As Variant
                arrExistant = dictLots(WBS1)
                arrExistant(1) = arrExistant(1) + PVH        ' Total PV
                arrExistant(2) = arrExistant(2) + EW         ' Total EW
                arrExistant(3) = arrExistant(3) + ActualH    ' Total AC
                arrExistant(4) = arrExistant(4) + GainConfirme    ' Gains confirmes
                arrExistant(5) = arrExistant(5) + PerteConfirmee  ' Pertes confirmees
                arrExistant(6) = arrExistant(6) + BaselineH  ' Total Baseline
                ' Compteurs pour fiabilite
                arrExistant(9) = arrExistant(9) + 1  ' Total taches
                If t.PercentComplete >= 100 Then arrExistant(10) = arrExistant(10) + 1  ' Taches terminees
                dictLots(WBS1) = arrExistant
            Else
                Dim arrNouveau(10) As Variant
                arrNouveau(0) = TrouverNomLot(WBS1)  ' Nom du lot
                arrNouveau(1) = PVH                  ' PV
                arrNouveau(2) = EW                   ' EW
                arrNouveau(3) = ActualH              ' AC
                arrNouveau(4) = GainConfirme         ' Gains confirmes
                arrNouveau(5) = PerteConfirmee       ' Pertes confirmees
                arrNouveau(6) = BaselineH            ' Baseline
                arrNouveau(7) = 0                    ' SPI (calcule apres)
                arrNouveau(8) = 0                    ' CPI (calcule apres)
                arrNouveau(9) = 1                    ' Total taches
                If t.PercentComplete >= 100 Then arrNouveau(10) = 1 Else arrNouveau(10) = 0  ' Taches terminees
                dictLots(WBS1) = arrNouveau
            End If
            
            ' Agregation globale
            TotalPV = TotalPV + PVH
            TotalEW = TotalEW + EW
            TotalAC = TotalAC + ActualH
            TotalGainsConfirmes = TotalGainsConfirmes + GainConfirme
            TotalPertesConfirmees = TotalPertesConfirmees + PerteConfirmee
        End If
    Next t
    
    ' === CALCUL SPI/CPI GLOBAUX ===
    Dim SPIGlobal As Variant, CPIGlobal As Variant
    If TotalPV > 0 Then SPIGlobal = TotalEW / TotalPV Else SPIGlobal = "N/A"
    If TotalAC > 0 Then CPIGlobal = TotalEW / TotalAC Else CPIGlobal = "N/A"
    
    ' === CALCUL SPI/CPI PAR LOT ET PREPARATION TOPS ===
    ReDim arrLotsSPI(0): ReDim arrLotsCPI(0)
    Dim cle As Variant
    For Each cle In dictLots.Keys
        Dim lotData As Variant
        lotData = dictLots(cle)
        
        ' Calcul fiabilite du lot
        Dim FiabiliteLot As Double
        If lotData(9) > 0 Then FiabiliteLot = lotData(10) / lotData(9) Else FiabiliteLot = 0
        
        ' Calcul SPI/CPI seulement si fiabilite >= 50%
        If FiabiliteLot >= 0.5 Then
            If lotData(1) > 0 Then ' PV > 0
                lotData(7) = lotData(2) / lotData(1)  ' SPI = EW / PV
                If lotData(7) < 1 Then ' SPI < 1 = retard
                    ReDim Preserve arrLotsSPI(nbLotsSPI)
                    arrLotsSPI(nbLotsSPI) = cle & "|" & lotData(0) & "|" & Format(lotData(7), "0.00") & "|" & Format(FiabiliteLot, "0%")
                    nbLotsSPI = nbLotsSPI + 1
                End If
            End If
            
            If lotData(3) > 0 Then ' AC > 0
                lotData(8) = lotData(2) / lotData(3)  ' CPI = EW / AC
                If lotData(8) < 1 Then ' CPI < 1 = surconsommation
                    ReDim Preserve arrLotsCPI(nbLotsCPI)
                    arrLotsCPI(nbLotsCPI) = cle & "|" & lotData(0) & "|" & Format(lotData(8), "0.00") & "|" & Format(FiabiliteLot, "0%")
                    nbLotsCPI = nbLotsCPI + 1
                End If
            End If
        End If
        
        dictLots(cle) = lotData
    Next cle
    
    ' === TRI DES TOPS ===
    If nbLotsSPI > 1 Then TrierTableauCroissant arrLotsSPI, nbLotsSPI - 1, 2 ' Tri par SPI croissant
    If nbLotsCPI > 1 Then TrierTableauCroissant arrLotsCPI, nbLotsCPI - 1, 2 ' Tri par CPI croissant
    
    ' === AFFICHAGE DU RESUME DIRIGEANT ===
    Dim ligne As Long
    ligne = 1
    
    ' En-tete principal
    wsResume.Cells(ligne, 1).Value = "RESUME DIRIGEANT - PROJET " & ActiveProject.Name
    wsResume.Cells(ligne, 3).Value = "Date: " & Format(DateEtat, "dd/mm/yyyy")
    ligne = ligne + 2
    
    ' === TABLEAU PRINCIPAL ===
    ' Titre section
    wsResume.Cells(ligne, 1).Value = "INDICATEURS GLOBAUX"
    wsResume.Cells(ligne, 1).Font.Bold = True
    wsResume.Cells(ligne, 1).Font.Size = 14
    ligne = ligne + 1
    
    ' KPI Globaux
    wsResume.Cells(ligne, 1).Value = "SPI Global (Planning)"
    wsResume.Cells(ligne, 2).Value = SPIGlobal
    If IsNumeric(SPIGlobal) Then
        If SPIGlobal < 0.8 Then
            wsResume.Cells(ligne, 3).Value = "🔴 CRITIQUE"
        ElseIf SPIGlobal < 0.9 Then
            wsResume.Cells(ligne, 3).Value = "🟠 ATTENTION"
        Else
            wsResume.Cells(ligne, 3).Value = "🟢 OK"
        End If
    End If
    ligne = ligne + 1
    
    wsResume.Cells(ligne, 1).Value = "CPI Global (Couts)"
    wsResume.Cells(ligne, 2).Value = CPIGlobal
    If IsNumeric(CPIGlobal) Then
        If CPIGlobal < 0.8 Then
            wsResume.Cells(ligne, 3).Value = "🔴 CRITIQUE"
        ElseIf CPIGlobal < 0.9 Then
            wsResume.Cells(ligne, 3).Value = "🟠 ATTENTION"
        Else
            wsResume.Cells(ligne, 3).Value = "🟢 OK"
        End If
    End If
    ligne = ligne + 1
    
    wsResume.Cells(ligne, 1).Value = "Gains Confirmes"
    wsResume.Cells(ligne, 2).Value = TotalGainsConfirmes & " h"
    wsResume.Cells(ligne, 3).Value = IIf(TotalGainsConfirmes > 0, "🟢 POSITIF", "")
    ligne = ligne + 1

    wsResume.Cells(ligne, 1).Value = "Pertes Confirmees"
    wsResume.Cells(ligne, 2).Value = TotalPertesConfirmees & " h"
    wsResume.Cells(ligne, 3).Value = IIf(TotalPertesConfirmees > 0, "🔴 NEGATIF", "")
    ligne = ligne + 2
    
    ' === TOP 3 RETARDS ===
    wsResume.Cells(ligne, 1).Value = "TOP 3 LOTS EN RETARD (SPI LE PLUS BAS)"
    wsResume.Cells(ligne, 1).Font.Bold = True
    wsResume.Cells(ligne, 1).Font.Size = 12
    ligne = ligne + 1
    
    If nbLotsSPI > 0 Then
        wsResume.Cells(ligne, 1).Value = "Lot"
        wsResume.Cells(ligne, 2).Value = "SPI"
        wsResume.Cells(ligne, 3).Value = "Fiabilité"
        wsResume.Cells(ligne, 4).Value = "Status"
        ligne = ligne + 1
        
        Dim maxRetards As Long
        maxRetards = IIf(nbLotsSPI > 3, 3, nbLotsSPI)
        For i = 0 To maxRetards - 1
            Dim arrRetard As Variant
            arrRetard = Split(arrLotsSPI(i), "|")
            wsResume.Cells(ligne + i, 1).Value = arrRetard(1) ' Nom lot
            wsResume.Cells(ligne + i, 2).Value = CDbl(arrRetard(2)) ' SPI
            wsResume.Cells(ligne + i, 3).Value = arrRetard(3) ' Fiabilite
            
            Dim spiVal As Double
            spiVal = CDbl(arrRetard(2))
            If spiVal < 0.8 Then
                wsResume.Cells(ligne + i, 4).Value = "🔴 CRITIQUE"
            ElseIf spiVal < 0.9 Then
                wsResume.Cells(ligne + i, 4).Value = "🟠 ATTENTION"
            End If
        Next i
        ligne = ligne + maxRetards + 1
    Else
        wsResume.Cells(ligne, 1).Value = "🟢 Aucun lot en retard significatif"
        ligne = ligne + 2
    End If
    
    ' === TOP 3 SURCONSOMMATIONS ===
    wsResume.Cells(ligne, 1).Value = "TOP 3 LOTS EN SURCONSOMMATION (CPI LE PLUS BAS)"
    wsResume.Cells(ligne, 1).Font.Bold = True
    wsResume.Cells(ligne, 1).Font.Size = 12
    ligne = ligne + 1
    
    If nbLotsCPI > 0 Then
        wsResume.Cells(ligne, 1).Value = "Lot"
        wsResume.Cells(ligne, 2).Value = "CPI"
        wsResume.Cells(ligne, 3).Value = "Fiabilite"
        wsResume.Cells(ligne, 4).Value = "Status"
        ligne = ligne + 1
        
        Dim maxSurcons As Long
        maxSurcons = IIf(nbLotsCPI > 3, 3, nbLotsCPI)
        For i = 0 To maxSurcons - 1
            Dim arrSurcons As Variant
            arrSurcons = Split(arrLotsCPI(i), "|")
            wsResume.Cells(ligne + i, 1).Value = arrSurcons(1) ' Nom lot
            wsResume.Cells(ligne + i, 2).Value = CDbl(arrSurcons(2)) ' CPI
            wsResume.Cells(ligne + i, 3).Value = arrSurcons(3) ' Fiabilite
            
            Dim cpiVal As Double
            cpiVal = CDbl(arrSurcons(2))
            If cpiVal < 0.8 Then
                wsResume.Cells(ligne + i, 4).Value = "🔴 CRITIQUE"
            ElseIf cpiVal < 0.9 Then
                wsResume.Cells(ligne + i, 4).Value = "🟠 ATTENTION"
            End If
        Next i
        ligne = ligne + maxSurcons + 1
    Else
        wsResume.Cells(ligne, 1).Value = "🟢 Aucun lot en surconsommation significative"
        ligne = ligne + 2
    End If
    
    ' === ANOMALIES DE COHERENCE ===
    wsResume.Cells(ligne, 1).Value = "ANOMALIES DE COHERENCE"
    wsResume.Cells(ligne, 1).Font.Bold = True
    wsResume.Cells(ligne, 1).Font.Size = 12
    wsResume.Cells(ligne, 1).Font.Color = RGB(255, 0, 0)
    ligne = ligne + 1
    
    If nbAnomalies > 0 Then
        wsResume.Cells(ligne, 1).Value = "Lot"
        wsResume.Cells(ligne, 2).Value = "Tâche"
        wsResume.Cells(ligne, 3).Value = "Écart"
        wsResume.Cells(ligne, 4).Value = "Status"
        ligne = ligne + 1
        
        For i = 0 To nbAnomalies - 1
            Dim arrAnomalie As Variant
            arrAnomalie = Split(ArrAnomalies(i), "|")
            wsResume.Cells(ligne + i, 1).Value = arrAnomalie(0) ' WBS
            wsResume.Cells(ligne + i, 2).Value = arrAnomalie(1) ' Nom tache
            wsResume.Cells(ligne + i, 3).Value = arrAnomalie(2) ' Ecart
            wsResume.Cells(ligne + i, 4).Value = "🔴 À VERIFIER"
            ' Mise en forme rouge
            wsResume.Range("A" & (ligne + i) & ":D" & (ligne + i)).Interior.Color = RGB(255, 182, 193)
            wsResume.Range("A" & (ligne + i) & ":D" & (ligne + i)).Font.Color = RGB(139, 0, 0)
        Next i
        ligne = ligne + nbAnomalies + 1
    Else
        wsResume.Cells(ligne, 1).Value = "🟢 Aucune anomalie détectée"
        ligne = ligne + 2
    End If
    
    ' === DONNEES NON FIABLES ===
    Dim nbLotsNonFiables As Long
    nbLotsNonFiables = 0
    For Each cle In dictLots.Keys
        Dim lotDataFinal As Variant
        lotDataFinal = dictLots(cle)
        Dim FiabiliteFinal As Double
        If lotDataFinal(9) > 0 Then FiabiliteFinal = lotDataFinal(10) / lotDataFinal(9) Else FiabiliteFinal = 0
        If FiabiliteFinal < 0.5 Then nbLotsNonFiables = nbLotsNonFiables + 1
    Next cle
    
    If nbLotsNonFiables > 0 Then
        wsResume.Cells(ligne, 1).Value = "⚠ DONNÉES NON FIABLES"
        wsResume.Cells(ligne, 1).Font.Bold = True
        wsResume.Cells(ligne, 1).Font.Color = RGB(255, 0, 0)
        wsResume.Cells(ligne, 2).Value = nbLotsNonFiables & " lots avec fiabilite < 50%"
        ligne = ligne + 1
    End If
End Sub

'========================================================================
' FONCTIONS UTILITAIRES
'========================================================================

'------------------------------------------------------------------------
' Trouve le nom d'un lot a partir de son WBS1
'------------------------------------------------------------------------
Function TrouverNomLot(ByVal WBS1 As String) As String
    Dim t As Task
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.Summary And t.OutlineLevel = 1 And t.WBS = WBS1 Then
                TrouverNomLot = t.Name
                Exit Function
            End If
        End If
    Next t
    TrouverNomLot = "Lot " & WBS1
End Function

'------------------------------------------------------------------------
' Tri un tableau de chaînes par ordre croissant sur un élément spécifique
'------------------------------------------------------------------------
Sub TrierTableauCroissant(ByRef arr() As String, ByVal max As Long, ByVal indexTri As Long)
    Dim i As Long, j As Long
    For i = 0 To max - 1
        For j = i + 1 To max
            Dim valeursI As Variant, valeursJ As Variant
            valeursI = Split(arr(i), "|")
            valeursJ = Split(arr(j), "|")
            
            If CDbl(valeursI(indexTri)) > CDbl(valeursJ(indexTri)) Then
                Dim temp As String
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

'------------------------------------------------------------------------
' Tri un tableau de chaînes par ordre décroissant sur un élément spécifique
'------------------------------------------------------------------------
Sub TrierTableauDecroissant(ByRef arr() As String, ByVal max As Long)
    Dim i As Long, j As Long
    For i = 0 To max - 1
        For j = i + 1 To max
            Dim valeursI As Variant, valeursJ As Variant
            valeursI = Split(arr(i), "|")
            valeursJ = Split(arr(j), "|")
            
            ' Tri par ordre décroissant sur l'élément 2 (valeur numérique)
            If CDbl(valeursI(2)) < CDbl(valeursJ(2)) Then
                Dim temp As String
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

'========================================================================
' FONCTIONS DE FORMATAGE
'========================================================================

'------------------------------------------------------------------------
' Formate la feuille SAPIN avec formatage conditionnel "sapin de Noel" hybride
'------------------------------------------------------------------------
Sub FormaterSapin(ByRef wsSapin As Object)
    ' Formats numeriques
    wsSapin.Range("E:E").NumberFormat = "0%" ' %C
    wsSapin.Range("F:H").NumberFormat = "0" ' Heures entieres
    wsSapin.Range("I:M").NumberFormat = "0" ' Ecart, Gains/Pertes
    wsSapin.Range("P:P").NumberFormat = "0.00" ' Taux consommation
    
    ' AutoFilter et FreezePanes
    wsSapin.Range("A1:P1").AutoFilter
    wsSapin.Activate
    wsSapin.Application.ActiveWindow.SplitRow = 1
    wsSapin.Application.ActiveWindow.FreezePanes = True
    
    ' Formatage conditionnel "sapin de Noel" sur colonne Ecart_h
    Dim derniereColonneUsed As Long
    derniereColonneUsed = wsSapin.Cells(wsSapin.Rows.Count, "A").End(-4162).Row ' xlUp = -4162
    
    If derniereColonneUsed > 1 Then
        ' Formatage colonne Ecart_h (I)
        Dim rngEcart As Object
        Set rngEcart = wsSapin.Range("I2:I" & derniereColonneUsed)
        
        rngEcart.FormatConditions.Delete
        
        ' Vert si negatif (gain) - Ecart < 0
        rngEcart.FormatConditions.Add Type:=1, Operator:=6, Formula1:="0"
        rngEcart.FormatConditions(1).Interior.Color = RGB(144, 238, 144) ' Vert clair
        rngEcart.FormatConditions(1).Font.Color = RGB(0, 100, 0) ' Vert fonce
        
        ' Rouge si positif (perte) - Ecart > 0
        rngEcart.FormatConditions.Add Type:=1, Operator:=4, Formula1:="0"
        rngEcart.FormatConditions(2).Interior.Color = RGB(255, 182, 193) ' Rouge clair
        rngEcart.FormatConditions(2).Font.Color = RGB(139, 0, 0) ' Rouge fonce
        
        ' Barres de donnees pour visualisation type "sapin"
        On Error Resume Next
        Dim dbCondition As Object
        Set dbCondition = rngEcart.FormatConditions.AddDatabar
        If Not dbCondition Is Nothing Then
            With dbCondition
                .MinPoint.Modify -1, , 1 ' xlConditionValueAutomaticMin
                .MaxPoint.Modify -1, , 2 ' xlConditionValueAutomaticMax
                .BarColor.Color = RGB(79, 129, 189) ' Bleu sapin
                .ShowValue = True
            End With
        End If
        On Error GoTo 0
        
        ' Formatage conditionnel sur colonne Performance (N)
        Dim rngPerformance As Object
        Set rngPerformance = wsSapin.Range("N2:N" & derniereColonneUsed)
        
        rngPerformance.FormatConditions.Delete
        
        ' Vert pour performances efficaces (xlExpression = 2)
        On Error Resume Next
        Dim fcVert As Object
        Set fcVert = rngPerformance.FormatConditions.Add(2, , "=OR(LEFT(N2,8)=""Efficace"",LEFT(N2,4)=""Tend"")")
        If Not fcVert Is Nothing Then
            fcVert.Interior.Color = RGB(144, 238, 144)
            fcVert.Font.Color = RGB(0, 100, 0)
        End If
        
        ' Orange pour alertes
        Dim fcOrange As Object
        Set fcOrange = rngPerformance.FormatConditions.Add(2, , "=LEFT(N2,6)=""Derive""")
        If Not fcOrange Is Nothing Then
            fcOrange.Interior.Color = RGB(255, 165, 0)
            fcOrange.Font.Color = RGB(139, 69, 19)
        End If
        
        ' Gris pour trop tot
        Dim fcGris As Object
        Set fcGris = rngPerformance.FormatConditions.Add(2, , "=LEFT(N2,4)=""Trop""")
        If Not fcGris Is Nothing Then
            fcGris.Interior.Color = RGB(211, 211, 211)
            fcGris.Font.Color = RGB(105, 105, 105)
        End If
        On Error GoTo 0
    End If
    
    ' Mise en forme des en-tetes
    wsSapin.Range("A1:P1").Font.Bold = True
    wsSapin.Range("A1:P1").Interior.Color = RGB(79, 129, 189)
    wsSapin.Range("A1:P1").Font.Color = RGB(255, 255, 255)
    
    ' Mise en evidence des colonnes hybrides
    wsSapin.Range("J1:M1").Interior.Color = RGB(255, 215, 0) ' Or pour colonnes hybrides
    wsSapin.Range("J1:M1").Font.Color = RGB(0, 0, 0)
    
    wsSapin.Columns.AutoFit
End Sub

'------------------------------------------------------------------------
' Formate la feuille RECAP_LOTS avec analyse hybride
'------------------------------------------------------------------------
Sub FormaterLots(ByRef wsLots As Object)
    ' Formats numeriques
    wsLots.Range("C:L").NumberFormat = "0" ' Heures entieres
    wsLots.Range("M:N").NumberFormat = "0.00" ' SPI/CPI 2 decimales
    wsLots.Range("O:O").NumberFormat = "0%" ' Fiabilite en %
    
    ' AutoFilter
    wsLots.Range("A1:Q1").AutoFilter
    
    ' Formatage conditionnel sur colonne Ecart_h
    Dim derniereColonneUsed As Long
    derniereColonneUsed = wsLots.Cells(wsLots.Rows.Count, "A").End(-4162).Row
    
    If derniereColonneUsed > 1 Then
        Dim rngEcart As Object
        Set rngEcart = wsLots.Range("H2:H" & derniereColonneUsed)
        
        rngEcart.FormatConditions.Delete
        
        ' Vert si negatif (gain)
        rngEcart.FormatConditions.Add Type:=1, Operator:=6, Formula1:="0"
        rngEcart.FormatConditions(1).Interior.Color = RGB(144, 238, 144)
        rngEcart.FormatConditions(1).Font.Color = RGB(0, 100, 0)
        
        ' Rouge si positif (perte)
        rngEcart.FormatConditions.Add Type:=1, Operator:=4, Formula1:="0"
        rngEcart.FormatConditions(2).Interior.Color = RGB(255, 182, 193)
        rngEcart.FormatConditions(2).Font.Color = RGB(139, 0, 0)
        
        ' Formatage conditionnel sur colonne Performance (P)
        Dim rngPerformance As Object
        Set rngPerformance = wsLots.Range("P2:P" & derniereColonneUsed)
        
        rngPerformance.FormatConditions.Delete
        
        ' Vert pour performances
        On Error Resume Next
        Dim fcVert As Object
        Set fcVert = rngPerformance.FormatConditions.Add(2, , "=LEFT(P2,2)=""✅""")
        If Not fcVert Is Nothing Then
            fcVert.Interior.Color = RGB(144, 238, 144)
            fcVert.Font.Color = RGB(0, 100, 0)
        End If
        
        ' Orange pour derives
        Dim fcOrange As Object
        Set fcOrange = rngPerformance.FormatConditions.Add(2, , "=OR(LEFT(P2,2)=""⚠"",LEFT(P2,2)=""📉"")")
        If Not fcOrange Is Nothing Then
            fcOrange.Interior.Color = RGB(255, 165, 0)
            fcOrange.Font.Color = RGB(139, 69, 19)
        End If
        On Error GoTo 0
        
        ' Formatage conditionnel sur fiabilite (O)
        Dim rngFiabilite As Object
        Set rngFiabilite = wsLots.Range("O2:O" & derniereColonneUsed)
        
        rngFiabilite.FormatConditions.Delete
        
        ' Vert si fiabilite elevee (>80%)
        rngFiabilite.FormatConditions.Add Type:=1, Operator:=4, Formula1:="0.8"
        rngFiabilite.FormatConditions(1).Interior.Color = RGB(144, 238, 144)
        
        ' Orange si fiabilite moyenne (50%-80%)
        rngFiabilite.FormatConditions.Add Type:=1, Operator:=7, Formula1:="0.5", Formula2:="0.8"
        rngFiabilite.FormatConditions(2).Interior.Color = RGB(255, 215, 0)
        
        ' Rouge si fiabilite faible (<50%)
        rngFiabilite.FormatConditions.Add Type:=1, Operator:=6, Formula1:="0.5"
        rngFiabilite.FormatConditions(3).Interior.Color = RGB(255, 182, 193)
    End If
    
    ' Mise en forme des en-tetes
    wsLots.Range("A1:Q1").Font.Bold = True
    wsLots.Range("A1:Q1").Interior.Color = RGB(79, 129, 189)
    wsLots.Range("A1:Q1").Font.Color = RGB(255, 255, 255)
    
    ' Mise en evidence des colonnes hybrides
    wsLots.Range("I1:L1").Interior.Color = RGB(255, 215, 0) ' Or pour colonnes hybrides
    wsLots.Range("I1:L1").Font.Color = RGB(0, 0, 0)
    
    wsLots.Columns.AutoFit
End Sub

'------------------------------------------------------------------------
' Formate la feuille RESUME_DIRIGEANT pour lecture rapide (1 minute)
'------------------------------------------------------------------------
Sub FormaterResume(ByRef wsResume As Object)
    ' Recherche de la dernière ligne utilisée
    Dim derniereColonneUsed As Long
    derniereColonneUsed = wsResume.Cells(wsResume.Rows.Count, "A").End(-4162).Row
    
    Dim i As Long
    For i = 1 To derniereColonneUsed
        Dim cellValue As String
        cellValue = CStr(wsResume.Cells(i, 1).Value)
        
        ' Formatage du titre principal
        If InStr(cellValue, "RESUME DIRIGEANT") > 0 Then
            wsResume.Cells(i, 1).Font.Size = 16
            wsResume.Cells(i, 1).Font.Bold = True
            wsResume.Cells(i, 1).Interior.Color = RGB(0, 70, 132) ' Bleu Vinci
            wsResume.Cells(i, 1).Font.Color = RGB(255, 255, 255)
            ' Fusion sur 4 colonnes pour le titre
            wsResume.Range("A" & i & ":D" & i).Merge
        End If
        
        ' Formatage des sections principales
        If InStr(cellValue, "INDICATEURS GLOBAUX") > 0 Or InStr(cellValue, "TOP 3") > 0 Or InStr(cellValue, "ANOMALIES") > 0 Then
            wsResume.Cells(i, 1).Font.Size = 14
            wsResume.Cells(i, 1).Font.Bold = True
            
            If InStr(cellValue, "INDICATEURS") > 0 Then
                wsResume.Cells(i, 1).Interior.Color = RGB(79, 129, 189) ' Bleu foncé
                wsResume.Cells(i, 1).Font.Color = RGB(255, 255, 255)
            ElseIf InStr(cellValue, "RETARD") > 0 Then
                wsResume.Cells(i, 1).Interior.Color = RGB(255, 165, 0) ' Orange
                wsResume.Cells(i, 1).Font.Color = RGB(139, 69, 19)
            ElseIf InStr(cellValue, "SURCONSOMMATION") > 0 Then
                wsResume.Cells(i, 1).Interior.Color = RGB(255, 99, 71) ' Rouge orangé
                wsResume.Cells(i, 1).Font.Color = RGB(255, 255, 255)
            ElseIf InStr(cellValue, "ANOMALIES") > 0 Then
                wsResume.Cells(i, 1).Interior.Color = RGB(220, 20, 60) ' Rouge foncé
                wsResume.Cells(i, 1).Font.Color = RGB(255, 255, 255)
            End If
        End If
        
        ' Formatage des indicateurs SPI/CPI
        If InStr(cellValue, "SPI Global") > 0 Or InStr(cellValue, "CPI Global") > 0 Then
            Dim valeurIndicateur As Variant
            valeurIndicateur = wsResume.Cells(i, 2).Value
            
            If IsNumeric(valeurIndicateur) Then
                ' Formatage conditionnel basé sur les seuils
                If valeurIndicateur < 0.8 Then
                    wsResume.Cells(i, 2).Interior.Color = RGB(255, 0, 0) ' Rouge
                    wsResume.Cells(i, 2).Font.Color = RGB(255, 255, 255)
                ElseIf valeurIndicateur < 0.9 Then
                    wsResume.Cells(i, 2).Interior.Color = RGB(255, 165, 0) ' Orange
                    wsResume.Cells(i, 2).Font.Color = RGB(255, 255, 255)
                Else
                    wsResume.Cells(i, 2).Interior.Color = RGB(0, 128, 0) ' Vert
                    wsResume.Cells(i, 2).Font.Color = RGB(255, 255, 255)
                End If
                wsResume.Cells(i, 2).Font.Bold = True
                wsResume.Cells(i, 2).NumberFormat = "0.00"
            End If
        End If
        
        ' Formatage des gains/pertes confirmés
        If InStr(cellValue, "Gains Confirmés") > 0 Then
            wsResume.Cells(i, 2).Interior.Color = RGB(144, 238, 144) ' Vert clair
            wsResume.Cells(i, 2).Font.Color = RGB(0, 100, 0)
            wsResume.Cells(i, 2).Font.Bold = True
        End If
        
        If InStr(cellValue, "Pertes Confirmées") > 0 Then
            wsResume.Cells(i, 2).Interior.Color = RGB(255, 182, 193) ' Rouge clair
            wsResume.Cells(i, 2).Font.Color = RGB(139, 0, 0)
            wsResume.Cells(i, 2).Font.Bold = True
        End If
        
        ' Formatage des en-têtes de tableaux
        If cellValue = "Lot" And wsResume.Cells(i, 2).Value = "SPI" Then
            wsResume.Range("A" & i & ":D" & i).Font.Bold = True
            wsResume.Range("A" & i & ":D" & i).Interior.Color = RGB(220, 220, 220)
            wsResume.Range("A" & i & ":D" & i).Font.Color = RGB(0, 0, 0)
        End If
        
        If cellValue = "Lot" And wsResume.Cells(i, 2).Value = "CPI" Then
            wsResume.Range("A" & i & ":D" & i).Font.Bold = True
            wsResume.Range("A" & i & ":D" & i).Interior.Color = RGB(220, 220, 220)
            wsResume.Range("A" & i & ":D" & i).Font.Color = RGB(0, 0, 0)
        End If
        
        If cellValue = "Lot" And wsResume.Cells(i, 2).Value = "Tâche" Then
            wsResume.Range("A" & i & ":D" & i).Font.Bold = True
            wsResume.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 182, 193) ' Rouge pour anomalies
            wsResume.Range("A" & i & ":D" & i).Font.Color = RGB(139, 0, 0)
        End If
        
        ' Formatage spécial pour fiabilité < 50%
        If InStr(cellValue, "DONNÉES NON FIABLES") > 0 Then
            wsResume.Cells(i, 1).Font.Bold = True
            wsResume.Cells(i, 1).Interior.Color = RGB(255, 0, 0)
            wsResume.Cells(i, 1).Font.Color = RGB(255, 255, 255)
            wsResume.Cells(i, 1).Font.Strikethrough = True
            wsResume.Cells(i, 2).Font.Bold = True
            wsResume.Cells(i, 2).Interior.Color = RGB(255, 182, 193)
            wsResume.Cells(i, 2).Font.Color = RGB(139, 0, 0)
        End If
        
        ' Formatage des colonnes de fiabilité
        If i > 1 Then ' Éviter la première ligne
            Dim fiabiliteStr As String
            fiabiliteStr = CStr(wsResume.Cells(i, 3).Value)
            If InStr(fiabiliteStr, "%") > 0 Then
                Dim fiabiliteVal As Double
                fiabiliteVal = Val(Replace(fiabiliteStr, "%", ""))
                If fiabiliteVal < 50 Then
                    wsResume.Cells(i, 3).Interior.Color = RGB(255, 0, 0) ' Rouge
                    wsResume.Cells(i, 3).Font.Color = RGB(255, 255, 255)
                    wsResume.Cells(i, 3).Font.Strikethrough = True
                ElseIf fiabiliteVal < 80 Then
                    wsResume.Cells(i, 3).Interior.Color = RGB(255, 165, 0) ' Orange
                    wsResume.Cells(i, 3).Font.Color = RGB(255, 255, 255)
                Else
                    wsResume.Cells(i, 3).Interior.Color = RGB(0, 128, 0) ' Vert
                    wsResume.Cells(i, 3).Font.Color = RGB(255, 255, 255)
                End If
                wsResume.Cells(i, 3).Font.Bold = True
            End If
        End If
        
        ' Formatage des messages d'état (emojis)
        Dim statusValue As String
        statusValue = CStr(wsResume.Cells(i, 4).Value)
        If InStr(statusValue, "🔴") > 0 Then
            wsResume.Cells(i, 4).Interior.Color = RGB(255, 182, 193)
            wsResume.Cells(i, 4).Font.Color = RGB(139, 0, 0)
            wsResume.Cells(i, 4).Font.Bold = True
        ElseIf InStr(statusValue, "🟠") > 0 Then
            wsResume.Cells(i, 4).Interior.Color = RGB(255, 218, 185)
            wsResume.Cells(i, 4).Font.Color = RGB(139, 69, 19)
            wsResume.Cells(i, 4).Font.Bold = True
        ElseIf InStr(statusValue, "🟢") > 0 Then
            wsResume.Cells(i, 4).Interior.Color = RGB(144, 238, 144)
            wsResume.Cells(i, 4).Font.Color = RGB(0, 100, 0)
            wsResume.Cells(i, 4).Font.Bold = True
        End If
    Next i
    
    ' Ajustement des largeurs de colonnes pour lisibilité optimale
    wsResume.Columns("A:A").ColumnWidth = 35 ' Descriptions
    wsResume.Columns("B:B").ColumnWidth = 15 ' Valeurs
    wsResume.Columns("C:C").ColumnWidth = 12 ' Fiabilité/SPI/CPI
    wsResume.Columns("D:D").ColumnWidth = 15 ' Status
    
    ' Ajustement de la hauteur des lignes pour meilleure lisibilité
    For i = 1 To derniereColonneUsed
        wsResume.Rows(i).RowHeight = 20
    Next i
    
    ' Ajout de bordures pour séparer les sections
    Dim rng As Object
    Set rng = wsResume.Range("A1:D" & derniereColonneUsed)
    
    ' Bordures fines
    On Error Resume Next
    With rng.Borders
        .LineStyle = 1 ' xlContinuous
        .Weight = 2    ' xlThin
        .ColorIndex = 1 ' xlAutomatic (noir)
    End With
    On Error GoTo 0
    
    ' Mise en forme générale
    wsResume.Range("A1:D" & derniereColonneUsed).Font.Size = 11
    wsResume.Range("A1:D" & derniereColonneUsed).Font.Name = "Calibri"
    
    ' Centrage vertical pour toutes les cellules
    wsResume.Range("A1:D" & derniereColonneUsed).VerticalAlignment = -4108 ' xlCenter
    
    ' Alignement spécifique par colonne
    wsResume.Columns("A:A").HorizontalAlignment = -4131 ' xlLeft
    wsResume.Columns("B:B").HorizontalAlignment = -4108 ' xlCenter
    wsResume.Columns("C:C").HorizontalAlignment = -4108 ' xlCenter
    wsResume.Columns("D:D").HorizontalAlignment = -4108 ' xlCenter
End Sub
