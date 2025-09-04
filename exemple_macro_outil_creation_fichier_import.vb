Sub Filtre() 
    'Au lancement de la macro, toutes les requêtes des tableaux sont actualisées pour garantir que les données utilisées pour le filtrage sont à jour.
    ' Macro qui permet de filtrer les données en fonction des critères spécifiés dans la feuille "BORNAGE" par l'utilisateur.
    ' Cette macro utilise des filtres automatiques pour afficher uniquement les lignes correspondant aux critères.
    ' Elle vérifie également que les cellules de critères ne sont pas vides avant d'appliquer les filtres.
    ' Si une cellule de critère est vide, un message d'avertissement est affiché.
    ' Après l'application des filtres, la feuille "LISTE" est sélectionnée pour afficher les résultats.
    '

    Worksheets("Conditions fournisseur").ListObjects(1).QueryTable.Refresh False
    Worksheets("Alpha 8 F main").ListObjects(1).QueryTable.Refresh False
    Worksheets("LISTE").ListObjects(1).QueryTable.Refresh False
    Worksheets("Commandes Clients F").ListObjects(1).QueryTable.Refresh False
    Worksheets("Commandes Clients P").ListObjects(1).QueryTable.Refresh False
    Worksheets("Fusionner1 (2)").ListObjects(1).QueryTable.Refresh False
    Worksheets("Fusionner1").ListObjects(1).QueryTable.Refresh False
    
    
    Dim Fournisseur As String
    Dim Article As String
    Dim commande As String
    Dim R_Date_debut As Range
    Dim R_Date_fin As Range
    Dim Date_debut As Date
    Dim Date_fin As Date
    

    Article = Worksheets("BORNAGE").Range("H13").Value
    Fournisseur = Worksheets("BORNAGE").Range("H9").Value
    commande = Worksheets("BORNAGE").Range("H11").Value
    Set R_Date_debut = Worksheets("BORNAGE").Range("H7")
    Set R_Date_fin = Worksheets("BORNAGE").Range("L7")
    
    If Article = "" Then
        MsgBox "La cellule H13 est vide. Veuillez entrer une valeur.", vbExclamation
    ElseIf Article <> "*" Then
        Worksheets("LISTE").Range("L8").AutoFilter Field:=12, Criteria1:=Article
    End If
    
    If Fournisseur = "" Then
        MsgBox "La cellule H9 est vide. Veuillez entrer une valeur.", vbExclamation
    ElseIf Fournisseur <> "*" Then
        Worksheets("LISTE").Range("C8").AutoFilter Field:=3, Criteria1:=Fournisseur
    End If
    
    If commande = "" Then
        MsgBox "La cellule H11 est vide. Veuillez entrer une valeur.", vbExclamation
    ElseIf commande <> "*" Then
        Worksheets("LISTE").Range("E8").AutoFilter Field:=5, Criteria1:=commande
    End If
    
    If (R_Date_debut.Value <> "" And IsDate(R_Date_debut)) And (R_Date_fin.Value <> "" And IsDate(R_Date_fin)) Then
        Date_debut = Format(R_Date_debut.Value, "mm/dd/yyyy")
        Date_fin = Format(R_Date_fin.Value, "mm/dd/yyyy")
        Worksheets("LISTE").Range("Q8").AutoFilter Field:=17, Criteria1:=">=" & Date_debut, Operator:=xlAnd, Criteria2:="<=" & Date_fin
    Else
        MsgBox "La cellule F4 est vide.D8 Veuillez entrer une valeur.", vbExclamation
    End If



    
    Sheets("LISTE").Select
End Sub

Sub Efface_filtre()
    ' Cette macro supprime tous les filtres appliqués sur la feuille "LISTE".

    Worksheets("LISTE").Range("C8").AutoFilter Field:=3
    Worksheets("LISTE").Range("Q8").AutoFilter Field:=17
    Worksheets("LISTE").Range("L8").AutoFilter Field:=12
    Worksheets("LISTE").Range("E8").AutoFilter Field:=5
    Worksheets("LISTE").Range("S8").AutoFilter Field:=19
    
    MsgBox "Les filtres se sont bien effacés"
End Sub

Sub Remplir()
    ' Cette macro copie les valeurs de la colonne "Nouvelle Date Packing" (colonne 19) vers la colonne "Date Packing" (colonne 18)
    ' dans le tableau nommé "Alpha_8_F" situé sur la feuille "LISTE".
    ' Elle parcourt chaque ligne du tableau et effectue la copie de valeur.
    ' Après la copie, elle affiche un message de confirmation.

    Dim tbl As ListObject
    Dim ligne As ListRow
    Dim i As Integer
    
    Set tbl = ThisWorkbook.Sheets("LISTE").ListObjects("Alpha_8_F")
    
    For i = 1 To tbl.ListRows.Count
        tbl.ListColumns(18).DataBodyRange.Cells(i, 1).Value = tbl.ListColumns(19).DataBodyRange.Cells(i, 1).Value
    Next i
        
    Dim Source As Worksheet, Cible As Worksheet
    Dim DernièreLigne As Long
    Dim TableauCible As ListObject
    Dim PlageTableau As Range
    
    Set Source = ThisWorkbook.Sheets("LISTE")
    Set Cible = ThisWorkbook.Sheets("BASE")
    
    Cible.Cells.Clear
    
    Source.Range("A8").CurrentRegion.SpecialCells(xlCellTypeVisible).Copy
    Cible.Range("A7").PasteSpecial xlPasteValues
    
    DernièreLigne = Cible.Cells(Cible.Rows.Count, "A").End(xlUp).row
    Set PlageTableau = Cible.Range("A8:S" & DernièreLigne)
    
    On Error Resume Next
    Set TableauCible = Cible.ListObjects(1)
    If Not TableauCible Is Nothing Then TableauCible.Delete
    On Error GoTo 0

    Set TableauCible = Cible.ListObjects.Add(xlSrcRange, PlageTableau, , xlYes)
    TableauCible.Name = "TableauFiltré"
    TableauCible.TableStyle = "TableStyleLight8"
    
    Application.CutCopyMode = False
    
    MsgBox "Copie terminée !"

End Sub

Sub ConvertTableToFormattedTXT()
    ' Cette macro exporte les données du tableau nommé "TableauFiltré" situé sur la feuille "BASE"
    ' vers un fichier texte formaté (CSV) en respectant une structure spécifique. Voir la partie 3.4. de mon rapport.
    ' Le fichier est enregistré dans un répertoire réseau spécifié afin que l'utilisateur puisse le retrouver facilement
    ' et qu'il n'y ait pas d'erreur sur le nom du fichier. (particularité de l'application S9)


    Dim scriptDir As String, outputFile As String
    Dim lo As ListObject
    Dim currentCommande As String, commande As String
    Dim colPO As Long, colCodeProduit As Long, colLigne As Long, colDatePacking As Long
    Dim fnum As Integer
    Dim ws As Worksheet
    Dim rw As Range
    
  
    Set ws = ThisWorkbook.Sheets("BASE")
    Set lo = ws.ListObjects("TableauFiltré")
    
 
    scriptDir = "\\novafile\datas\Supply_Production\11_Remontée date packing S9\"
    outputFile = scriptDir & "MAJ_DATE_PACKING.csv"
    

    With lo.Sort
        .SortFields.Clear
        .SortFields.Add Key:=lo.ListColumns("N° PO S9").DataBodyRange, _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .Apply
    End With
    
    

    colPO = lo.ListColumns("N° PO S9").Index
    colCodeProduit = lo.ListColumns("Code Produit UP").Index
    colLigne = lo.ListColumns("Ligne").Index
    colDatePacking = lo.ListColumns("Nouvelle Date Packing").Index
    
 
    fnum = FreeFile
    Open outputFile For Output As #fnum
    
    currentCommande = ""
    
  
    For Each rw In lo.DataBodyRange.Rows
        commande = Right("00000000" & CStr(rw.Cells(1, colPO).Value), 8)
        If commande <> currentCommande Then
            Print #fnum, "COMMANDE;" & commande
            currentCommande = commande
        End If
        Print #fnum, "LIGNE;" & rw.Cells(1, colCodeProduit).Value & ";" & _
                     rw.Cells(1, colLigne).Value & ";" & _
                     Format(rw.Cells(1, colDatePacking).Value, "DD/MM/YYYY")
    Next rw
    
    Close #fnum
    
    Worksheets("LISTE").Range("C8").AutoFilter Field:=3
    Worksheets("LISTE").Range("Q8").AutoFilter Field:=17
    Worksheets("LISTE").Range("L8").AutoFilter Field:=12
    Worksheets("LISTE").Range("E8").AutoFilter Field:=5
    Worksheets("LISTE").Range("S8").AutoFilter Field:=19
    
    Worksheets("LISTE").ListObjects("Alpha_8_F").ListColumns("Nouvelle Date Packing").DataBodyRange.ClearContents
    
    MsgBox "Les commandes re-formatées ont été enregistrées dans le fichier:" & vbCrLf & outputFile, vbInformation
End Sub

Sub Copie()

    Dim wsSource As Worksheet, wsCible As Worksheet
    Dim tblSource As ListObject, tblCible As ListObject
    Dim colCleSource As Range, colValeurSource As Range
    Dim colCleCible As Range, colValeurCible As Range
    Dim i As Long, valeurCherchee As Variant, valeurTrouvee As Variant

    Set wsSource = ThisWorkbook.Sheets("LISTE")
    Set wsCible = ThisWorkbook.Sheets("BASE")


    Set tblSource = wsSource.ListObjects("Alpha_8_F")
    Set tblCible = wsCible.ListObjects("TableauFiltré")

    Set colCleSource = tblSource.ListColumns(1).DataBodyRange
    Set colValeurSource = tblSource.ListColumns(18).DataBodyRange
    Set colCleCible = tblCible.ListColumns(1).DataBodyRange
    Set colValeurCible = tblCible.ListColumns(18).DataBodyRange

    For i = 1 To colCleCible.Rows.Count
        valeurCherchee = colCleCible.Cells(i, 1).Value
        

        If Not IsEmpty(valeurCherchee) Then

            On Error Resume Next
            valeurTrouvee = Application.WorksheetFunction.VLookup(valeurCherchee, tblSource.DataBodyRange, 18, False)
            On Error GoTo 0
            
    
            If Not IsError(valeurTrouvee) Then
                colValeurCible.Cells(i, 1).Value = valeurTrouvee
            Else
                colValeurCible.Cells(i, 1).Value = "Non trouvé"
            End If
        Else
            colValeurCible.Cells(i, 1).Value = "Clé vide"
        End If
    Next i

    MsgBox "Modification Manuelle prise en compte", vbInformation
End Sub






