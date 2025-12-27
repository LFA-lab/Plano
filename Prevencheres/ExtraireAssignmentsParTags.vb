' ======================================================================
' MACRO DE VERIFICATION ET EXTRACTION DES ASSIGNMENTS PAR TAGS
' ======================================================================
' Cette macro extrait et affiche les assignments (affectations) filtrés
' par Tranche et Type, démontrant que les tags métier ont bien été
' propagés depuis les tâches vers les assignments.
'
' Utilisation:
' 1. Ouvrir le projet MS Project après l'import
' 2. Exécuter cette macro
' 3. Saisir la Tranche et le Type souhaités
' 4. Un fichier Excel sera généré avec les résultats
' ======================================================================

Sub ExtraireAssignmentsParTranche()
    
    Dim pjApp As MSProject.Application
    Dim pjProj As MSProject.Project
    Dim t As Task
    Dim a As Assignment
    
    ' Saisie utilisateur
    Dim trancheRecherche As String
    Dim typeRecherche As String
    
    trancheRecherche = InputBox("Entrez la Tranche à rechercher (ex: T1, T2...)", "Filtrage par Tranche", "T1")
    If trancheRecherche = "" Then
        MsgBox "Opération annulée.", vbExclamation
        Exit Sub
    End If
    
    typeRecherche = InputBox("Entrez le Type/Métier à rechercher (ex: Génie Civil, Electrique, CQ...)", "Filtrage par Type", "")
    ' typeRecherche peut être vide = pas de filtre sur type
    
    ' Récupération du projet actif
    Set pjApp = MSProject.Application
    Set pjProj = pjApp.ActiveProject
    
    If pjProj Is Nothing Then
        MsgBox "Aucun projet MS Project ouvert.", vbCritical
        Exit Sub
    End If
    
    ' Création d'un fichier Excel pour export
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)
    
    ' En-têtes
    xlSheet.Cells(1, 1).Value = "Tâche"
    xlSheet.Cells(1, 2).Value = "Ressource"
    xlSheet.Cells(1, 3).Value = "Type Ressource"
    xlSheet.Cells(1, 4).Value = "Unités/Work"
    xlSheet.Cells(1, 5).Value = "Tranche"
    xlSheet.Cells(1, 6).Value = "Zone"
    xlSheet.Cells(1, 7).Value = "Sous-Zone"
    xlSheet.Cells(1, 8).Value = "Type/Métier"
    xlSheet.Cells(1, 9).Value = "Entreprise"
    
    ' Mise en forme des en-têtes
    With xlSheet.Range("A1:I1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    
    Dim ligneExcel As Long
    ligneExcel = 2
    Dim compteur As Long
    compteur = 0
    
    ' Parcours de toutes les tâches
    For Each t In pjProj.Tasks
        If Not t Is Nothing Then
            If Not t.Summary Then ' Ignorer les tâches récapitulatives
                
                ' Parcours des assignments de cette tâche
                For Each a In t.Assignments
                    If Not a Is Nothing Then
                        
                        ' Filtrage par Tranche (obligatoire)
                        Dim matchTranche As Boolean
                        matchTranche = (Trim(UCase(a.Text1)) = Trim(UCase(trancheRecherche)))
                        
                        ' Filtrage par Type (optionnel)
                        Dim matchType As Boolean
                        If typeRecherche = "" Then
                            matchType = True ' Pas de filtre type
                        Else
                            matchType = (Trim(UCase(a.Text4)) = Trim(UCase(typeRecherche)))
                        End If
                        
                        ' Si les 2 critères matchent, on exporte
                        If matchTranche And matchType Then
                            
                            xlSheet.Cells(ligneExcel, 1).Value = t.Name
                            xlSheet.Cells(ligneExcel, 2).Value = a.ResourceName
                            
                            ' Type de ressource
                            Dim typeRessource As String
                            Select Case a.Resource.Type
                                Case pjResourceTypeWork
                                    typeRessource = "Travail"
                                Case pjResourceTypeMaterial
                                    typeRessource = "Matériel"
                                Case pjResourceTypeCost
                                    typeRessource = "Coût"
                                Case Else
                                    typeRessource = "Inconnu"
                            End Select
                            xlSheet.Cells(ligneExcel, 3).Value = typeRessource
                            
                            ' Unités ou Work
                            Dim valeurRessource As String
                            If a.Resource.Type = pjResourceTypeWork Then
                                valeurRessource = Format(a.Work / 60, "0.00") & " h (" & a.Units & "%)"
                            Else
                                valeurRessource = a.Units
                            End If
                            xlSheet.Cells(ligneExcel, 4).Value = valeurRessource
                            
                            ' Tags métier (hérités de la tâche)
                            xlSheet.Cells(ligneExcel, 5).Value = a.Text1 ' Tranche
                            xlSheet.Cells(ligneExcel, 6).Value = a.Text2 ' Zone
                            xlSheet.Cells(ligneExcel, 7).Value = a.Text3 ' Sous-Zone
                            xlSheet.Cells(ligneExcel, 8).Value = a.Text4 ' Type
                            xlSheet.Cells(ligneExcel, 9).Value = a.Text5 ' Entreprise
                            
                            ligneExcel = ligneExcel + 1
                            compteur = compteur + 1
                            
                        End If
                    End If
                Next a
                
            End If
        End If
    Next t
    
    ' Ajustement automatique des colonnes
    xlSheet.Columns("A:I").AutoFit
    
    ' Message final
    Dim message As String
    message = "Extraction terminée !" & vbCrLf & vbCrLf
    message = message & "Critères de recherche:" & vbCrLf
    message = message & "  - Tranche: " & trancheRecherche & vbCrLf
    If typeRecherche <> "" Then
        message = message & "  - Type/Métier: " & typeRecherche & vbCrLf
    Else
        message = message & "  - Type/Métier: (tous)" & vbCrLf
    End If
    message = message & vbCrLf & "Résultats: " & compteur & " affectation(s) trouvée(s)"
    
    MsgBox message, vbInformation, "Extraction Assignments"
    
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
End Sub


' ======================================================================
' MACRO DE VERIFICATION GLOBALE DES TAGS SUR ASSIGNMENTS
' ======================================================================
' Cette macro vérifie que TOUS les assignments ont bien hérité
' des tags de leurs tâches parentes.
' Elle génère un rapport indiquant les assignments correctement tagués
' et ceux manquants (s'il y en a).
' ======================================================================

Sub VerifierTagsAssignments()
    
    Dim pjApp As MSProject.Application
    Dim pjProj As MSProject.Project
    Dim t As Task
    Dim a As Assignment
    
    Set pjApp = MSProject.Application
    Set pjProj = pjApp.ActiveProject
    
    If pjProj Is Nothing Then
        MsgBox "Aucun projet MS Project ouvert.", vbCritical
        Exit Sub
    End If
    
    Dim totalAssignments As Long
    Dim assignmentsTagged As Long
    Dim assignmentsNonTagged As Long
    
    totalAssignments = 0
    assignmentsTagged = 0
    assignmentsNonTagged = 0
    
    Dim rapport As String
    rapport = "===== VERIFICATION DES TAGS SUR ASSIGNMENTS =====" & vbCrLf & vbCrLf
    
    ' Parcours de toutes les tâches
    For Each t In pjProj.Tasks
        If Not t Is Nothing Then
            If Not t.Summary Then
                
                For Each a In t.Assignments
                    If Not a Is Nothing Then
                        totalAssignments = totalAssignments + 1
                        
                        ' Vérifier si au moins un tag est présent
                        Dim hasTag As Boolean
                        hasTag = (Trim(a.Text1) <> "" Or Trim(a.Text2) <> "" Or _
                                  Trim(a.Text3) <> "" Or Trim(a.Text4) <> "" Or Trim(a.Text5) <> "")
                        
                        If hasTag Then
                            assignmentsTagged = assignmentsTagged + 1
                        Else
                            assignmentsNonTagged = assignmentsNonTagged + 1
                            rapport = rapport & "⚠️ Assignment NON TAGUÉ: " & t.Name & " > " & a.ResourceName & vbCrLf
                        End If
                        
                    End If
                Next a
                
            End If
        End If
    Next t
    
    ' Résumé
    rapport = rapport & vbCrLf & "===== RESUME =====" & vbCrLf
    rapport = rapport & "Total assignments: " & totalAssignments & vbCrLf
    rapport = rapport & "Assignments tagués: " & assignmentsTagged & vbCrLf
    rapport = rapport & "Assignments NON tagués: " & assignmentsNonTagged & vbCrLf
    rapport = rapport & vbCrLf
    
    If assignmentsNonTagged = 0 Then
        rapport = rapport & "✅ Tous les assignments sont correctement tagués !" & vbCrLf
    Else
        rapport = rapport & "❌ Certains assignments ne sont pas tagués." & vbCrLf
    End If
    
    ' Affichage dans une boîte de dialogue
    MsgBox rapport, vbInformation, "Vérification Tags"
    
    ' Export vers fichier texte
    Dim fichierRapport As String
    fichierRapport = Environ$("USERPROFILE") & "\Downloads\Verification_Tags_Assignments_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim f As Object
    Set f = fso.CreateTextFile(fichierRapport, True)
    f.Write rapport
    f.Close
    Set f = Nothing
    Set fso = Nothing
    
    MsgBox "Rapport exporté vers:" & vbCrLf & fichierRapport, vbInformation
    
End Sub


' ======================================================================
' MACRO D'EXPORT COMPLET DES ASSIGNMENTS AVEC TAGS
' ======================================================================
' Exporte TOUS les assignments avec leurs tags dans Excel
' (sans filtre)
' ======================================================================

Sub ExporterTousAssignmentsAvecTags()
    
    Dim pjApp As MSProject.Application
    Dim pjProj As MSProject.Project
    Dim t As Task
    Dim a As Assignment
    
    Set pjApp = MSProject.Application
    Set pjProj = pjApp.ActiveProject
    
    If pjProj Is Nothing Then
        MsgBox "Aucun projet MS Project ouvert.", vbCritical
        Exit Sub
    End If
    
    ' Création Excel
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)
    
    ' En-têtes
    xlSheet.Cells(1, 1).Value = "ID Tâche"
    xlSheet.Cells(1, 2).Value = "Tâche"
    xlSheet.Cells(1, 3).Value = "Ressource"
    xlSheet.Cells(1, 4).Value = "Type Ressource"
    xlSheet.Cells(1, 5).Value = "Unités"
    xlSheet.Cells(1, 6).Value = "Work (h)"
    xlSheet.Cells(1, 7).Value = "Tranche (Assignment)"
    xlSheet.Cells(1, 8).Value = "Zone (Assignment)"
    xlSheet.Cells(1, 9).Value = "Sous-Zone (Assignment)"
    xlSheet.Cells(1, 10).Value = "Type/Métier (Assignment)"
    xlSheet.Cells(1, 11).Value = "Entreprise (Assignment)"
    
    With xlSheet.Range("A1:K1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    
    Dim ligneExcel As Long
    ligneExcel = 2
    
    ' Parcours complet
    For Each t In pjProj.Tasks
        If Not t Is Nothing Then
            If Not t.Summary Then
                
                For Each a In t.Assignments
                    If Not a Is Nothing Then
                        
                        xlSheet.Cells(ligneExcel, 1).Value = t.ID
                        xlSheet.Cells(ligneExcel, 2).Value = t.Name
                        xlSheet.Cells(ligneExcel, 3).Value = a.ResourceName
                        
                        ' Type ressource
                        Select Case a.Resource.Type
                            Case pjResourceTypeWork
                                xlSheet.Cells(ligneExcel, 4).Value = "Travail"
                            Case pjResourceTypeMaterial
                                xlSheet.Cells(ligneExcel, 4).Value = "Matériel"
                            Case pjResourceTypeCost
                                xlSheet.Cells(ligneExcel, 4).Value = "Coût"
                            Case Else
                                xlSheet.Cells(ligneExcel, 4).Value = "Inconnu"
                        End Select
                        
                        xlSheet.Cells(ligneExcel, 5).Value = a.Units
                        xlSheet.Cells(ligneExcel, 6).Value = Format(a.Work / 60, "0.00")
                        
                        ' Tags hérités
                        xlSheet.Cells(ligneExcel, 7).Value = a.Text1  ' Tranche
                        xlSheet.Cells(ligneExcel, 8).Value = a.Text2  ' Zone
                        xlSheet.Cells(ligneExcel, 9).Value = a.Text3  ' Sous-Zone
                        xlSheet.Cells(ligneExcel, 10).Value = a.Text4 ' Type
                        xlSheet.Cells(ligneExcel, 11).Value = a.Text5 ' Entreprise
                        
                        ligneExcel = ligneExcel + 1
                        
                    End If
                Next a
                
            End If
        End If
    Next t
    
    ' Ajustement colonnes
    xlSheet.Columns("A:K").AutoFit
    
    MsgBox "Export terminé ! " & (ligneExcel - 2) & " affectation(s) exportée(s).", vbInformation
    
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
End Sub

