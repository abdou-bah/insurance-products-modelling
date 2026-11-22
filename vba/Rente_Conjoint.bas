Attribute VB_Name = "Rente_Conjoint"
' Modèle Rente Conjoint : scénarios de taux
' ---- Tirage des trajectoires de taux (modèle de Vasicek, schéma d'Euler, pas annuel)
Sub TirerTrajectoires(nomFeuille As String)
    Dim wsT As Worksheet: Set wsT = Sheets(nomFeuille)
    Dim nSimu As Long: nSimu = Val(wsT.Range("C3").Value)
    Dim horizon As Integer: horizon = Val(wsT.Range("C4").Value)
    Dim dt As Double: dt = wsT.Range("C5").Value
    Dim r0 As Double: r0 = wsT.Range("C6").Value
    Dim a As Double: a = wsT.Range("C7").Value
    Dim theta As Double: theta = wsT.Range("C8").Value
    Dim sigma As Double: sigma = wsT.Range("C9").Value
    If nSimu <= 0 Or horizon <= 0 Then Exit Sub

    Dim Resultats() As Double: ReDim Resultats(1 To nSimu, 1 To horizon + 1)
    Dim i As Long, j As Integer
    Randomize
    For i = 1 To nSimu
        Resultats(i, 1) = r0
        For j = 2 To horizon + 1
            Resultats(i, j) = Resultats(i, j - 1) + a * (theta - Resultats(i, j - 1)) * dt _
                              + sigma * Sqr(dt) * WorksheetFunction.Norm_S_Inv(Rnd())
        Next j
    Next i

    wsT.Range("B11:ZZ10000").ClearContents
    wsT.Range("E12").Resize(nSimu, horizon + 1).Value = Resultats
    For j = 0 To horizon: wsT.Cells(11, 5 + j).Value = j: Next j
    For i = 1 To nSimu: wsT.Cells(11 + i, 2).Value = "Trajectoire " & i: Next i
End Sub

' ---- Recalcul du BE et de la PVFP pour chaque trajectoire
' Les taux sont injectés sur une ligne de l'onglet Inputs & Hypothèses (taux à 1 an de chaque année) :
' l'année j utilise le taux en début d'année, r(j-1), et l'actualisation se fait année par année.
Sub CalculerBEPVFP(nomFeuille As String, ligneInputs As Long, nbAnnees As Integer, _
                   feuilleBE As String, celluleBE As String, feuillePVFP As String, cellulePVFP As String)
    Dim wsT As Worksheet: Set wsT = Sheets(nomFeuille)
    Dim wsI As Worksheet: Set wsI = Sheets("Inputs & Hypothèses")
    Dim nSimu As Long: nSimu = Val(wsT.Range("C3").Value)
    If nSimu <= 0 Then Exit Sub

    Dim Sauvegarde As Variant
    Sauvegarde = wsI.Cells(ligneInputs, 8).Resize(1, nbAnnees).Value
    Dim Trajectoires As Variant
    Trajectoires = wsT.Range("E12").Resize(nSimu, nbAnnees).Value

    Dim ListeBE() As Double: ReDim ListeBE(1 To nSimu, 1 To 1)
    Dim ListePVFP() As Double: ReDim ListePVFP(1 To nSimu, 1 To 1)
    Dim i As Long, j As Integer
    Application.ScreenUpdating = False
    For i = 1 To nSimu
        For j = 1 To nbAnnees
            wsI.Cells(ligneInputs, 7 + j).Value = Trajectoires(i, j)
        Next j
        Application.Calculate
        ListeBE(i, 1) = Sheets(feuilleBE).Range(celluleBE).Value
        ListePVFP(i, 1) = Sheets(feuillePVFP).Range(cellulePVFP).Value
    Next i
    wsI.Cells(ligneInputs, 8).Resize(1, nbAnnees).Value = Sauvegarde
    Application.Calculate

    wsT.Range("C11").Value = "BE simulé"
    wsT.Range("C12").Resize(nSimu, 1).Value = ListeBE
    wsT.Range("D11").Value = "PVFP simulée"
    wsT.Range("D12").Resize(nSimu, 1).Value = ListePVFP
    Application.ScreenUpdating = True
    MsgBox "BE moyen : " & Format(WorksheetFunction.Average(wsT.Range("C12").Resize(nSimu, 1)), "# ##0.0") & _
           vbCrLf & "PVFP moyenne : " & Format(WorksheetFunction.Average(wsT.Range("D12").Resize(nSimu, 1)), "# ##0.0")
End Sub

Sub Etape1_TirageAleatoire(): TirerTrajectoires "Trajectoires facteurs d'actu": End Sub
Sub P_TirageAleatoire(): TirerTrajectoires "Trajectoires rendement fi": End Sub
Sub Extraction_BE_Et_PVFP(): CalculerBEPVFP "Trajectoires facteurs d'actu", 9, 50, "Best Estimate", "C30", "PVFP", "C30": End Sub
Sub P_Extraction_BE_Et_PVFP(): CalculerBEPVFP "Trajectoires rendement fi", 8, 50, "Best Estimate", "C30", "PVFP", "C30": End Sub
