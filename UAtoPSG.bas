Attribute VB_Name = "Module121"
' #####################################################
' #################### VBA_SEQUOIA ####################
' #####################################################
'
' AUTEUR : CHEVEREAU Matthieu
' COORDONNEES : matthieuchevereau@yahoo.fr
'
' #####################################################
Sub UAtoPSG()
'Déclaration des variables
Dim ROW, COLUMN As Variant
Dim Nom_tableau As String
Dim i As Integer
Dim COL As Integer

'##### MISE EN FORME DE UA #####

    'Sélection des cellulles
    Cells.Select
    Selection.Style = "Normal"
    
    Range("O1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    
    ROW = Selection.Rows.Count
    COLUMN = Selection.Columns.Count
    
    'Remplacement des points
    COL = 1
    For i = 1 To COLUMN
        If Left(Cells(1, COL), 5) = "SURF_" Then
            Columns(COL).Select
            Selection.Replace What:=".", Replacement:=".", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
            Selection.NumberFormat = "0.0000"
            COL = COL + 1
        Else
            COL = COL + 1
        End If
    Next
    
    'Mise en forme tableau
    Randomize
    Nom_tableau = "Table991"
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(ROW, COLUMN)), , xlYes).Name = Nom_tableau
    ActiveSheet.ListObjects(Nom_tableau).TableStyle = "TableStyleLight9"
    ActiveSheet.ListObjects(Nom_tableau).ShowTableStyleRowStripes = False
    ActiveSheet.ListObjects(Nom_tableau).ShowTotals = True
    ActiveSheet.Name = "UA"
    
    'version logiciel
    num = Application.Version

    If num = "15.0" Or num = "16.0" Then
        verss = 6
    Else
        verss = xlPivotTableVersion15
    End If

'##### Création de OCS #####
    
    'Création de la feuille de traitement
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "OCS"
        
    'Création du TCD
    TCD = "TCD_OCS"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            "Table991", Version:=verss).CreatePivotTable TableDestination:="OCS!R5C2", _
            TableName:=TCD, DefaultVersion:=verss
    
    'Entrée des valeurs
    With ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables(TCD).AddDataField ActiveSheet. _
        PivotTables(TCD).PivotFields("SURF_COR"), _
        "Somme de SURF_COR", xlSum
    ActiveSheet.PivotTables(TCD).AddDataField ActiveSheet. _
        PivotTables(TCD).PivotFields("SURF_COR"), _
        "Somme de SURF_COR2", xlSum
    With ActiveSheet.PivotTables(TCD).PivotFields("Somme de SURF_COR2")
        .Calculation = xlPercentOfColumn
    End With
    
    'Inscription des légendes
    ActiveSheet.PivotTables(TCD).CompactLayoutRowHeader = "OCCUPATION DU SOL"
    ActiveSheet.PivotTables(TCD).DataPivotField.PivotItems("Somme de SURF_COR"). _
            Caption = "SURFACE (HA)"
    ActiveSheet.PivotTables(TCD).DataPivotField.PivotItems("Somme de SURF_COR2"). _
            Caption = "PROPORTION"
    ActiveSheet.PivotTables(TCD).GrandTotalName = "TOTAL"
    
    Columns("C:C").Select
    Selection.NumberFormat = "0.0000"
    Columns("D:D").Select
    Selection.NumberFormat = "0.0%"
    
    ActiveSheet.PivotTables(TCD).TableStyle2 = _
            "PivotStyleMedium2"
    ActiveSheet.PivotTables(TCD).HasAutoFormat = False
            
    'Inscription du titre
    Range("A1:H1").Select
    Selection.Style = "Titre 1"
    Columns("B:H").Select
    Selection.Columns.AutoFit
    
    Cells(1, 1) = "Répartition des surfaces par occupations du sol"

'##### Création de PC #####
        
    'Création de la feuille de traitement
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "PC"
            
    'Création du TCD
    TCD = "TCD_PC"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            "Table991", Version:=verss).CreatePivotTable TableDestination:="PC!R3C2", _
            TableName:=TCD, DefaultVersion:=verss
        
    'Entrée des valeurs
    With ActiveSheet.PivotTables(TCD).PivotFields("COM_NOM")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(TCD).PivotFields("PARCA")
        .Orientation = xlRowField
        .Position = 2
    End With
        With ActiveSheet.PivotTables(TCD).PivotFields("LIEUDIT")
        .Orientation = xlRowField
        .Position = 3
    End With
    ActiveSheet.PivotTables(TCD).AddDataField ActiveSheet. _
        PivotTables(TCD).PivotFields("SURF_COR"), _
        "Somme de SURF_COR", xlSum
                
    'Mise en plan
    ActiveSheet.PivotTables(TCD).RowAxisLayout xlTabularRow
    ActiveSheet.PivotTables(TCD).RepeatAllLabels xlRepeatLabels
    
    With ActiveSheet.PivotTables(TCD)
    .RowAxisLayout xlTabularRow
    .RepeatAllLabels xlRepeatLabels
    ' defined once per pivottable:
    .ColumnGrand = False
    .RowGrand = False
    ' use RowFields only:
        For Each campos In .RowFields
            ' either this:
            campos.Subtotals(1) = True   ' Automatic on (= all other off)
            campos.Subtotals(1) = False  ' Automatic also off
    
            ' or that (all 12 off):
            'campos.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        Next campos
    End With
    
    'Inscription des légendes
    ActiveSheet.PivotTables(TCD).PivotFields("COM_NOM"). _
            Caption = "COMMUNES"
    ActiveSheet.PivotTables(TCD).PivotFields("PARCA"). _
            Caption = "PARCELLES"
    ActiveSheet.PivotTables(TCD).DataPivotField.PivotItems("Somme de SURF_COR"). _
            Caption = "SURFACE (HA)"
    ActiveSheet.PivotTables(TCD).PivotSelect _
        "SURFACE (HA)", xlDataAndLabel, True
    Selection.NumberFormat = "0.0000"
    ActiveSheet.PivotTables(TCD).GrandTotalName = "TOTAL"
    ActiveSheet.PivotTables(TCD).TableStyle2 = _
            "PivotStyleMedium2"
    ActiveSheet.PivotTables(TCD).HasAutoFormat = False
    With ActiveSheet.PivotTables(TCD).PivotFields("COMMUNES")
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables(TCD).ColumnGrand = True
                
    'Inscription du titre
    Range("A1:J1").Select
    Selection.Style = "Titre 1"
    Columns("B:J").Select
    Selection.Columns.AutoFit
        
    Cells(1, 1) = "Répartition des surfaces par parcelles cadastrales"
    
    ' ***
    
    'Création du TCD2
    TCD = "TCD_PC2"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            "Table991", Version:=verss).CreatePivotTable TableDestination:="PC!R5C7", _
            TableName:=TCD, DefaultVersion:=verss
    
    With ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL").ClearAllFilters
    ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL").CurrentPage = "BOISEE"
    
    With ActiveSheet.PivotTables(TCD).PivotFields("DEP_NOM")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(TCD).PivotFields("COM_NOM")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables(TCD).PivotFields("DEP_NOM")
        .PivotItems("(blank)").Visible = False
    End With
    
    ActiveSheet.PivotTables(TCD).AddDataField ActiveSheet. _
        PivotTables(TCD).PivotFields("SURF_COR"), _
        "Somme de SURF_COR", xlSum
    
    ActiveSheet.PivotTables(TCD).RowAxisLayout xlTabularRow
    ActiveSheet.PivotTables(TCD).RepeatAllLabels xlRepeatLabels
    
    With ActiveSheet.PivotTables(TCD)
    .RowAxisLayout xlTabularRow
    .RepeatAllLabels xlRepeatLabels
    ' defined once per pivottable:
    .ColumnGrand = False
    .RowGrand = False
    ' use RowFields only:
        For Each campos In .RowFields
            ' either this:
            campos.Subtotals(1) = True   ' Automatic on (= all other off)
            campos.Subtotals(1) = False  ' Automatic also off
    
            ' or that (all 12 off):
            'campos.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        Next campos
    End With
    
    ActiveSheet.PivotTables(TCD).ColumnGrand = True
    
    ActiveSheet.PivotTables(TCD).PivotFields("DEP_NOM"). _
            Caption = "DEPARTEMENT"
    ActiveSheet.PivotTables(TCD).PivotFields("COM_NOM"). _
            Caption = "COMMUNES"
    ActiveSheet.PivotTables(TCD).DataPivotField.PivotItems("Somme de SURF_COR"). _
            Caption = "SURFACE (HA)"
    ActiveSheet.PivotTables(TCD).PivotSelect _
        "SURFACE (HA)", xlDataAndLabel, True
    Selection.NumberFormat = "0.0000"
    ActiveSheet.PivotTables(TCD).GrandTotalName = "TOTAL"
    
    ActiveSheet.PivotTables(TCD).TableStyle2 = _
        "PivotStyleMedium9"
    ActiveSheet.PivotTables(TCD).PivotFields("COMMUNES"). _
        AutoSort xlDescending, "SURFACE (HA)"
    
     
'##### Création de PC-PF #####
    
    'Création de la feuille de traitement
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "PCF"
        
    'Création du TCD
    TCD = "TCD_PC-PF"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            "Table991", Version:=verss).CreatePivotTable TableDestination:="PCF!R5C2", _
            TableName:=TCD, DefaultVersion:=verss
    
    'Entrée des valeurs
    With ActiveSheet.PivotTables(TCD).PivotFields("N_PARFOR")
        .Orientation = xlRowField
        .Position = 1
    End With
        With ActiveSheet.PivotTables(TCD).PivotFields("COM_NOM")
        .Orientation = xlRowField
        .Position = 2
    End With
        With ActiveSheet.PivotTables(TCD).PivotFields("COM_NOM")
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables(TCD).PivotFields("PARCA")
        .Orientation = xlRowField
        .Position = 3
    End With
    ActiveSheet.PivotTables(TCD).AddDataField ActiveSheet. _
        PivotTables(TCD).PivotFields("SURF_COR"), _
        "Somme de SURF_COR", xlSum
        
    'Mise en plan
    ActiveSheet.PivotTables(TCD).RowAxisLayout xlTabularRow
    ActiveSheet.PivotTables(TCD).RepeatAllLabels xlRepeatLabels
    
    With ActiveSheet.PivotTables(TCD)
    .RowAxisLayout xlTabularRow
    .RepeatAllLabels xlRepeatLabels
    ' defined once per pivottable:
    .ColumnGrand = False
    .RowGrand = False
    ' use RowFields only:
        For Each campos In .RowFields
            ' either this:
            campos.Subtotals(1) = True   ' Automatic on (= all other off)
            campos.Subtotals(1) = False  ' Automatic also off
    
            ' or that (all 12 off):
            'campos.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        Next campos
    End With
    ActiveSheet.PivotTables(TCD).ColumnGrand = True
    
    'Création du tableur synthèse
    Cells(5, 7) = "Parcelle forestière"
    Cells(5, 8) = "Surface (ha)"
    Cells(5, 9) = "Commune"
    Cells(5, 10) = "Parcelle cadastrale"
    Cells(5, 11) = "Surface (ha)"
    
    ra = 6
    rpf = 6
    
    While Cells(ra, 2) <> ""
        Cells(ra, 7) = Cells(ra, 2)
        Cells(ra, 9) = Cells(ra, 3)
        Cells(ra, 10) = Cells(ra, 4)
        Cells(ra, 11) = Cells(ra, 5)
        
        pf = Application.WorksheetFunction.Sum(pf, Cells(ra, 5))
        If Cells(ra + 1, 2) <> Cells(ra, 2) Then
            Cells(rpf, 8) = pf
            
            Range(Cells(rpf + 1, 7), Cells(ra, 7)).Select
            Selection.ClearContents
            
            Range(Cells(rpf, 7), Cells(ra, 7)).Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.Merge
            
            Range(Cells(rpf, 8), Cells(ra, 8)).Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.Merge
            
            rpf = ra + 1
            pf = 0
        End If
        
        ra = ra + 1
        
    Wend
    
    
    
    'Inscription des légendes
    ActiveSheet.PivotTables(TCD).PivotFields("N_PARFOR"). _
            Caption = "PARCELLES FORESTIERES"
    ActiveSheet.PivotTables(TCD).PivotFields("COM_NOM"). _
            Caption = "COMMUNES"
    ActiveSheet.PivotTables(TCD).PivotFields("PARCA"). _
            Caption = "PARCELLES CADASTRALES"
    ActiveSheet.PivotTables(TCD).DataPivotField.PivotItems("Somme de SURF_COR"). _
            Caption = "SURFACE (HA)"
    ActiveSheet.PivotTables(TCD).PivotSelect _
        "SURFACE (HA)", xlDataAndLabel, True
    Selection.NumberFormat = "0.0000"
    ActiveSheet.PivotTables(TCD).GrandTotalName = "TOTAL"
    ActiveSheet.PivotTables(TCD).TableStyle2 = _
            "PivotStyleMedium2"
    ActiveSheet.PivotTables(TCD).HasAutoFormat = False
    
    Range("G5:K5").Select
    Selection.Style = "Accent1"
    Columns("G:H").Select
    With Selection
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Columns("K:K").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.NumberFormat = "0.0000"
            
    'Inscription du titre
    ActiveSheet.Name = "PC-PF"
    Range("A1:L1").Select
    Selection.Style = "Titre 1"
    Columns("B:L").Select
    Selection.Columns.AutoFit
    
    Cells(1, 1) = "Correspondance parcelles forestières/parcelles cadastrales"

'##### Création de PF #####
    
    'Création de la feuille de traitement
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "PF"
        
    'Création du TCD
    TCD = "TCD_PF"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            "Table991", Version:=verss).CreatePivotTable TableDestination:="PF!R5C2", _
            TableName:=TCD, DefaultVersion:=verss
    
    'Entrée des valeurs
    With ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL").ClearAllFilters
    ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL").CurrentPage = "BOISEE"
    
    With ActiveSheet.PivotTables(TCD).PivotFields("N_PARFOR")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables(TCD).AddDataField ActiveSheet. _
        PivotTables(TCD).PivotFields("SURF_COR"), _
        "Somme de SURF_COR", xlSum
    
    'Inscription des légendes
    ActiveSheet.PivotTables(TCD).CompactLayoutRowHeader = "PARCELLES"
    ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL"). _
            Caption = "OCCUPATION DU SOL"
    ActiveSheet.PivotTables(TCD).PivotFields("N_PARFOR"). _
            Caption = "PARCELLES FORESTIERES"
    ActiveSheet.PivotTables(TCD).DataPivotField.PivotItems("Somme de SURF_COR"). _
            Caption = "SURFACE (HA)"
    ActiveSheet.PivotTables(TCD).PivotSelect _
        "SURFACE (HA)", xlDataAndLabel, True
    Selection.NumberFormat = "0.0000"
    ActiveSheet.PivotTables(TCD).GrandTotalName = "TOTAL"
    ActiveSheet.PivotTables(TCD).TableStyle2 = _
            "PivotStyleMedium2"
    ActiveSheet.PivotTables(TCD).HasAutoFormat = False
            
    'Inscription du titre
    Range("A1:H1").Select
    Selection.Style = "Titre 1"
    Columns("B:H").Select
    Selection.Columns.AutoFit
    
    Cells(1, 1) = "Répartition des surfaces par parcelles forestières"
    
    'Création du résumé
    Cells(5, 2).Select
    Range(Selection, Selection.End(xlDown)).Select
    ROW = Selection.Rows.Count - 2
    
    Cells(5, 5) = "MIN"
    Cells(6, 5) = "MAX"
    Cells(7, 5) = "MOY"
    
    Cells(5, 6) = Application.WorksheetFunction.Min(Range(Cells(6, 3), Cells((5 + ROW), 3)))
    Cells(6, 6) = Application.WorksheetFunction.Max(Range(Cells(6, 3), Cells((5 + ROW), 3)))
    Cells(7, 6) = Application.Average(Range(Cells(6, 3), Cells((5 + ROW), 3)))
    
    Range(Cells(5, 5), Cells(7, 5)).Select
    Selection.Style = "Accent1"
    Range(Cells(5, 6), Cells(7, 6)).Select
    Selection.Style = "Calcul"
    Selection.NumberFormat = "0.0000"
    
    ActiveSheet.PivotTables("TCD_PF").PivotFields("OCCUPATION DU SOL").CurrentPage _
        = "(All)"
    ActiveSheet.PivotTables("TCD_PF").PivotFields("OCCUPATION DU SOL"). _
        EnableMultiplePageItems = True
    With ActiveSheet.PivotTables("TCD_PF").PivotFields("OCCUPATION DU SOL")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    Range("E4").FormulaR1C1 = "valeurs calculées sur les surfaces boisées uniquement:"
    Range("E4").Select
    Selection.Font.Italic = True
    
'##### Création de PLT #####
    
    'Création de la feuille de traitement
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "PLT"
        
    'Création du TCD
    TCD = "TCD_PLT"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            "Table991", Version:=verss).CreatePivotTable TableDestination:="PLT!R5C2", _
            TableName:=TCD, DefaultVersion:=verss
    
    'Entrée des valeurs
    'With ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL")
    '    .Orientation = xlPageField
    '    .Position = 1
    'End With
    'ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL").ClearAllFilters
    'ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL").CurrentPage = "BOISEE"
    
    With ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(TCD).PivotFields("PLT_TYPE")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables(TCD).AddDataField ActiveSheet. _
        PivotTables(TCD).PivotFields("SURF_COR"), _
        "Somme de SURF_COR", xlSum
    ActiveSheet.PivotTables(TCD).AddDataField ActiveSheet.PivotTables( _
        TCD).PivotFields("SURF_COR"), "Somme de SURF_COR2", xlSum
    With ActiveSheet.PivotTables(TCD).PivotFields("Somme de SURF_COR2")
        .Calculation = xlPercentOfColumn
    End With
    
    'Inscription des légendes
    ActiveSheet.PivotTables(TCD).CompactLayoutRowHeader = "PEUPLEMENTS"
    ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL"). _
            Caption = "OCCUPATION DU SOL"
    ActiveSheet.PivotTables(TCD).PivotFields("PLT_TYPE"). _
            Caption = "PEUPLEMENTS"
    ActiveSheet.PivotTables(TCD).DataPivotField.PivotItems("Somme de SURF_COR"). _
            Caption = "SURFACE (HA)"
    ActiveSheet.PivotTables(TCD).DataPivotField.PivotItems( _
        "Somme de SURF_COR2").Caption = "PROPORTION"
    ActiveSheet.PivotTables(TCD).PivotSelect _
        "SURFACE (HA)", xlDataAndLabel, True
    Selection.NumberFormat = "0.0000"
    ActiveSheet.PivotTables(TCD).PivotSelect _
        "SURFACE (HA)", xlDataAndLabel, True
    Selection.NumberFormat = "0.0000"
    ActiveSheet.PivotTables(TCD).GrandTotalName = "TOTAL"
    ActiveSheet.PivotTables(TCD).TableStyle2 = _
            "PivotStyleMedium2"
    ActiveSheet.PivotTables(TCD).HasAutoFormat = False
    Columns("D:D").Select
    Selection.NumberFormat = "0.0%"
            
    'Inscription du titre
    Range("A1:D1").Select
    Selection.Style = "Titre 1"
    Columns("B:D").Select
    Selection.Columns.AutoFit
    
    Cells(1, 1) = "Répartition des surfaces par peuplements"
    
'##### Création de PLT-PF #####
    
    'Création de la feuille de traitement
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "PLTPF"
        
    'Création du TCD
    TCD = "TCD_PLT-PF"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            "Table991", Version:=verss).CreatePivotTable TableDestination:="PLTPF!R5C2", _
            TableName:=TCD, DefaultVersion:=verss
    
    'Entrée des valeurs
    With ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(TCD).PivotFields("PLT_TYPE")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables(TCD).PivotFields("N_PARFOR")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables(TCD).AddDataField ActiveSheet. _
        PivotTables(TCD).PivotFields("SURF_COR"), _
        "Somme de SURF_COR", xlSum
    
    'Inscription des légendes
    ActiveSheet.PivotTables(TCD).CompactLayoutRowHeader = "PEUPLEMENTS"
    ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL"). _
            Caption = "OCCUPATION DU SOL"
    ActiveSheet.PivotTables(TCD).PivotFields("PLT_TYPE"). _
            Caption = "PEUPLEMENTS"
    ActiveSheet.PivotTables(TCD).CompactLayoutColumnHeader = _
        "PARCELLES FORESTIERES"
    ActiveSheet.PivotTables(TCD).DataPivotField.PivotItems("Somme de SURF_COR"). _
            Caption = "SURFACE (HA)"
    ActiveSheet.PivotTables(TCD).PivotSelect _
        "SURFACE (HA)", xlDataAndLabel, True
    Selection.NumberFormat = "0.0000"
    ActiveSheet.PivotTables(TCD).PivotSelect _
        "SURFACE (HA)", xlDataAndLabel, True
    Selection.NumberFormat = "0.0000"
    ActiveSheet.PivotTables(TCD).GrandTotalName = "TOTAL"
    ActiveSheet.PivotTables(TCD).TableStyle2 = _
            "PivotStyleMedium2"
    ActiveSheet.PivotTables(TCD).HasAutoFormat = False
            
    'Inscription du titre
    ActiveSheet.Name = "PLT-PF"
    Range("A1:AZ1").Select
    Selection.Style = "Titre 1"
    Columns("B:AZ").Select
    Selection.Columns.AutoFit
    
    Cells(1, 1) = "Répartition des surfaces par peuplements et parcelles forestières"

'##### Création de PLT-PC #####
    
    'Création de la feuille de traitement
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "PLTPC"
        
    'Création du TCD
    TCD = "TCD_PLT-PC"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            "Table991", Version:=verss).CreatePivotTable TableDestination:="PLTPC!R5C2", _
            TableName:=TCD, DefaultVersion:=verss
    
    'Entrée des valeurs
    With ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL").ClearAllFilters
    ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL").CurrentPage = "BOISEE"
    
    With ActiveSheet.PivotTables(TCD).PivotFields("COM_NOM")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(TCD).PivotFields("PARCA")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables(TCD).PivotFields("PLT_TYPE")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables(TCD).AddDataField ActiveSheet. _
        PivotTables(TCD).PivotFields("SURF_COR"), _
        "Somme de SURF_COR", xlSum
    
    'Inscription des légendes
    ActiveSheet.PivotTables(TCD).CompactLayoutRowHeader = "COMMUNES"
    ActiveSheet.PivotTables(TCD).PivotFields("PARCA"). _
            Caption = "PARCELLE CADASTRALE"
    ActiveSheet.PivotTables(TCD).PivotFields("PLT_TYPE"). _
            Caption = "PEUPLEMENTS"
    ActiveSheet.PivotTables(TCD).DataPivotField.PivotItems("Somme de SURF_COR"). _
            Caption = "SURFACE (HA)"
    ActiveSheet.PivotTables(TCD).PivotSelect _
        "SURFACE (HA)", xlDataAndLabel, True
    Selection.NumberFormat = "0.0000"
    
    ActiveSheet.PivotTables(TCD).GrandTotalName = "TOTAL"
    ActiveSheet.PivotTables(TCD).TableStyle2 = _
            "PivotStyleMedium2"
    ActiveSheet.PivotTables(TCD).HasAutoFormat = False
            
    'Inscription du titre
    ActiveSheet.Name = "PLT-PC"
    Range("A1:AZ1").Select
    Selection.Style = "Titre 1"
    Columns("B:AZ").Select
    Selection.Columns.AutoFit
    
    Cells(1, 1) = "Répartition des surfaces par peuplements et parcelles forestières"

'##### Création de AME #####
    
    'Création de la feuille de traitement
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "AME"
        
    'Création du TCD
    TCD = "TCD_AME"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            "Table991", Version:=verss).CreatePivotTable TableDestination:="AME!R5C2", _
            TableName:=TCD, DefaultVersion:=verss
    
    'Entrée des valeurs
    With ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(TCD).PivotFields("AME_TYPE")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables(TCD).AddDataField ActiveSheet. _
        PivotTables(TCD).PivotFields("SURF_COR"), _
        "Somme de SURF_COR", xlSum
    ActiveSheet.PivotTables(TCD).AddDataField ActiveSheet.PivotTables( _
        TCD).PivotFields("SURF_COR"), "Somme de SURF_COR2", xlSum
    With ActiveSheet.PivotTables(TCD).PivotFields("Somme de SURF_COR2")
        .Calculation = xlPercentOfColumn
    End With
    
    'Inscription des légendes
    ActiveSheet.PivotTables(TCD).CompactLayoutRowHeader = "PEUPLEMENTS"
    ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL"). _
            Caption = "OCCUPATION DU SOL"
    ActiveSheet.PivotTables(TCD).PivotFields("AME_TYPE"). _
            Caption = "AMENAGEMENTS"
    ActiveSheet.PivotTables(TCD).DataPivotField.PivotItems("Somme de SURF_COR"). _
            Caption = "SURFACE (HA)"
    ActiveSheet.PivotTables(TCD).DataPivotField.PivotItems( _
        "Somme de SURF_COR2").Caption = "PROPORTION"
    ActiveSheet.PivotTables(TCD).PivotSelect _
        "SURFACE (HA)", xlDataAndLabel, True
    Selection.NumberFormat = "0.0000"
    ActiveSheet.PivotTables(TCD).PivotSelect _
        "SURFACE (HA)", xlDataAndLabel, True
    Selection.NumberFormat = "0.0000"
    ActiveSheet.PivotTables(TCD).GrandTotalName = "TOTAL"
    ActiveSheet.PivotTables(TCD).TableStyle2 = _
            "PivotStyleMedium2"
    ActiveSheet.PivotTables(TCD).HasAutoFormat = False
    Columns("D:D").Select
    Selection.NumberFormat = "0.0%"
            
    'Inscription du titre
    Range("A1:D1").Select
    Selection.Style = "Titre 1"
    Columns("B:D").Select
    Selection.Columns.AutoFit
    
    Cells(1, 1) = "Répartition des surfaces par peuplements"
    
'##### Création de SSPF #####
    
    'Création de la feuille de traitement
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "SSPF"
        
    'Création du TCD
    TCD = "TCD_SSPF"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            "Table991", Version:=verss).CreatePivotTable TableDestination:="SSPF!R5C2", _
            TableName:=TCD, DefaultVersion:=verss
    
    'Entrée des valeurs
    With ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL").ClearAllFilters
    ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL").CurrentPage = "BOISEE"
    
    With ActiveSheet.PivotTables(TCD).PivotFields("PARFOR")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(TCD).PivotFields("PLT_TYPE")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables(TCD).AddDataField ActiveSheet. _
        PivotTables(TCD).PivotFields("SURF_COR"), _
        "Somme de SURF_COR", xlSum
    
    'Inscription des légendes
    ActiveSheet.PivotTables(TCD).CompactLayoutRowHeader = "PEUPLEMENTS"
    ActiveSheet.PivotTables(TCD).PivotFields("OCCUP_SOL"). _
            Caption = "OCCUPATION DU SOL"
    ActiveSheet.PivotTables(TCD).PivotFields("PLT_TYPE"). _
            Caption = "PEUPLEMENTS"
    ActiveSheet.PivotTables(TCD).CompactLayoutColumnHeader = _
        "PARCELLES FORESTIERES"
    ActiveSheet.PivotTables(TCD).DataPivotField.PivotItems("Somme de SURF_COR"). _
            Caption = "SURFACE (HA)"
    ActiveSheet.PivotTables(TCD).PivotSelect _
        "SURFACE (HA)", xlDataAndLabel, True
    Selection.NumberFormat = "0.0000"
    ActiveSheet.PivotTables(TCD).PivotSelect _
        "SURFACE (HA)", xlDataAndLabel, True
    Selection.NumberFormat = "0.0000"
    ActiveSheet.PivotTables(TCD).GrandTotalName = "TOTAL"
    ActiveSheet.PivotTables(TCD).TableStyle2 = _
            "PivotStyleMedium2"
    ActiveSheet.PivotTables(TCD).HasAutoFormat = False
    ActiveSheet.PivotTables(TCD).RowAxisLayout xlTabularRow
    
    Set pvttbl = ActiveSheet.PivotTables(TCD)
    With pvttbl
        For Each pvtFld In .PivotFields
            pvtFld.Subtotals(1) = False
        Next pvtFld
    End With
            
    'Inscription du titre
    Range("A1:H1").Select
    Selection.Style = "Titre 1"
    Columns("B:H").Select
    Selection.Columns.AutoFit
    
    Cells(1, 1) = "Répartition des surfaces par sous-parcelles forestières"
    
'##### Création de PROG #####
    
    Sheets("SSPF").Select
    Sheets("SSPF").Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "PROG"
    
    With ActiveSheet.PivotTables("TCD_SSPF")
        .ColumnGrand = False
        .RowGrand = False
    End With
    
    Range("B5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ROW = Selection.Rows.Count - 1
    COLUMN = Selection.Columns.Count
    
    Range("E5").FormulaR1C1 = "=RC[-3]"
    Range("E5").Select
    Selection.AutoFill Destination:=Range("E5:G5"), Type:=xlFillDefault
    
    Range(Cells(5, 5), Cells(5, 7)).Select
    Selection.AutoFill Destination:=Range(Cells(5, 5), Cells(ROW + 5, 7)), Type:=xlFillDefault
    
    Range(Cells(5, 5), Cells(ROW + 5, 7)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Columns("E:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:D").Delete Shift:=xlToLeft
    
    Range("D3").FormulaR1C1 = "ANNEE"
    For i = 1 To 15
        Cells(3, i + 4) = i
        Cells(5, i + 4) = Year(Now()) + i
    Next
       
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(5, 2), Cells(ROW + 5, 4 + 15)), , xlYes).Name = _
        "Table992"
    ActiveSheet.ListObjects("Table992").TableStyle = "TableStyleLight9"
    
    Range("D3").Style = "Accent1"
    Range("E3:S3").Style = "20 % - Accent1"
    
    Range(Cells(6, 4), Cells(ROW + 5, 4)).Select
    Selection.NumberFormat = "0.0000"
    
    'Inscription du titre
    Range("A1:T1").Select
    Selection.Style = "Titre 1"
    Columns("B:T").Select
    Selection.Columns.AutoFit
    
    Cells(1, 1) = "Programme des coupes et travaux"

End Sub

