Sub Mensal_Efi()

Application.ScreenUpdating = False
On Error GoTo erro

    With ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Quinzena1")
        .SlicerItems("1ª Q").Selected = True
        .SlicerItems("2ª Q").Selected = True
        .SlicerItems("(vazio)").Selected = False
    End With
    With ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Célula")
        .SlicerItems("1").Selected = True
        .SlicerItems("2").Selected = True
        .SlicerItems("3").Selected = True
        .SlicerItems("4").Selected = True
        .SlicerItems("5").Selected = True
        .SlicerItems("6").Selected = True
        .SlicerItems("7").Selected = True
        .SlicerItems("8").Selected = True
        .SlicerItems("9").Selected = True
        .SlicerItems("10").Selected = True
        .SlicerItems("11").Selected = True
        .SlicerItems("(vazio)").Selected = False
    End With
    Sheets("Tab. Dinâmica").Select
    ActiveWorkbook.ShowPivotTableFieldList = True
    ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Data") _
            .Orientation = xlRowField
    ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Quinzena"). _
        Orientation = xlHidden
    ActiveWorkbook.ShowPivotTableFieldList = False
    Sheets("Dashboard").Select
    ActiveWorkbook.SlicerCaches("NativeTimeline_Data_Eficiência").Slicers("Data Faturamento 2" _
        ).TimelineViewState.Level = xlTimelineLevelMonths

Application.ScreenUpdating = True

Exit Sub
erro: Resume Next
End Sub

Sub Quinzena_Efi()

Application.ScreenUpdating = False
On Error GoTo erro

    With ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Quinzena1")
        .SlicerItems("1ª Q").Selected = True
        .SlicerItems("2ª Q").Selected = True
        .SlicerItems("(vazio)").Selected = False
    End With
    With ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Célula")
        .SlicerItems("1").Selected = True
        .SlicerItems("2").Selected = True
        .SlicerItems("3").Selected = True
        .SlicerItems("4").Selected = True
        .SlicerItems("5").Selected = True
        .SlicerItems("6").Selected = True
        .SlicerItems("7").Selected = True
        .SlicerItems("8").Selected = True
        .SlicerItems("9").Selected = True
        .SlicerItems("10").Selected = True
        .SlicerItems("11").Selected = True
        .SlicerItems("(vazio)").Selected = False
    End With
    Sheets("Tab. Dinâmica").Select
    ActiveWorkbook.ShowPivotTableFieldList = True
    ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Quinzena") _
            .Orientation = xlHidden
    ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Data") _
            .Orientation = xlRowField
    ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Quinzena") _
            .Orientation = xlRowField
    ActiveWorkbook.ShowPivotTableFieldList = False
    Sheets("Dashboard").Select
    ActiveWorkbook.SlicerCaches("NativeTimeline_Data_Eficiência").Slicers("Data Faturamento 2" _
        ).TimelineViewState.Level = xlTimelineLevelMonths

Application.ScreenUpdating = True

Exit Sub
erro: Resume Next
End Sub

Sub Anual_Efi()

Application.ScreenUpdating = False
On Error GoTo erro
    
    With ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Quinzena1")
        .SlicerItems("1ª Q").Selected = True
        .SlicerItems("2ª Q").Selected = True
        .SlicerItems("(vazio)").Selected = False
    End With
    With ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Célula")
        .SlicerItems("1").Selected = True
        .SlicerItems("2").Selected = True
        .SlicerItems("3").Selected = True
        .SlicerItems("4").Selected = True
        .SlicerItems("5").Selected = True
        .SlicerItems("6").Selected = True
        .SlicerItems("7").Selected = True
        .SlicerItems("8").Selected = True
        .SlicerItems("9").Selected = True
        .SlicerItems("10").Selected = True
        .SlicerItems("11").Selected = True
        .SlicerItems("(vazio)").Selected = False
    End With
    Sheets("Tab. Dinâmica").Select
    ActiveWorkbook.ShowPivotTableFieldList = True
    ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Data") _
            .Orientation = xlRowField
    ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Data"). _
        Orientation = xlHidden
    ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Quinzena"). _
        Orientation = xlHidden
    ActiveWorkbook.ShowPivotTableFieldList = False
    Sheets("Dashboard").Select
    ActiveWorkbook.SlicerCaches("NativeTimeline_Data_Eficiência").Slicers("Data Faturamento 2" _
        ).TimelineViewState.Level = xlTimelineLevelYears

Application.ScreenUpdating = True

Exit Sub
erro: Resume Next
End Sub

Sub Por_celulas_efi()

Application.ScreenUpdating = False
On Error GoTo erro

    Sheets("Tab. Dinâmica").Select
    ActiveWorkbook.ShowPivotTableFieldList = True
    ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Célula") _
            .Orientation = xlHidden
    ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Célula") _
            .Orientation = xlRowField
    ActiveWorkbook.ShowPivotTableFieldList = False
    Sheets("Dashboard").Select

Application.ScreenUpdating = True

Exit Sub
erro: Resume Next
End Sub

Sub Total_celulas_efi()

Application.ScreenUpdating = False
On Error GoTo erro
    
    With ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Quinzena1")
        .SlicerItems("1ª Q").Selected = True
        .SlicerItems("2ª Q").Selected = True
        .SlicerItems("(vazio)").Selected = False
    End With
    With ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Célula")
        .SlicerItems("1").Selected = True
        .SlicerItems("2").Selected = True
        .SlicerItems("3").Selected = True
        .SlicerItems("4").Selected = True
        .SlicerItems("5").Selected = True
        .SlicerItems("6").Selected = True
        .SlicerItems("7").Selected = True
        .SlicerItems("8").Selected = True
        .SlicerItems("9").Selected = True
        .SlicerItems("10").Selected = True
        .SlicerItems("11").Selected = True
        .SlicerItems("(vazio)").Selected = False
    End With
    Sheets("Tab. Dinâmica").Select
    ActiveWorkbook.ShowPivotTableFieldList = True
    ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Célula") _
            .Orientation = xlHidden
    ActiveWorkbook.ShowPivotTableFieldList = False
    Sheets("Dashboard").Select

Application.ScreenUpdating = True

Exit Sub
erro: Resume Next
End Sub

Sub Setor_tap2()

Application.ScreenUpdating = False

Sheets("Tab. Dinâmica").Select
ActiveSheet.PivotTables("Tabela dinâmica2").ClearTable
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Data")
        .Orientation = xlRowField
        .AutoGroup
    End With
    Range("M4").Select
    Selection.Group Start:=True, End:=True, Periods:=Array(False, False, False, _
        False, True, False, True)
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Ano")
        .PivotItems("2018").Visible = False
        .PivotItems("2019").Visible = True
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Quinzena")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Setor")
        .PivotItems("Tapeçaria").Visible = True
        .PivotItems("Costura").Visible = False
        .PivotItems("Embalagem").Visible = False
        .PivotItems("Espumação").Visible = False
        .PivotItems("Laminação").Visible = False
        .PivotItems("Montagem").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Setor")
        .Orientation = xlPageField
        .Position = 1
    End With
ActiveSheet.PivotTables("Tabela dinâmica2").CalculatedFields.Add "Efi", _
    "='Minutos Produzidos'/'Minutos totais'", True
ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Efi"). _
    Orientation = xlDataField
Columns("N:N").Select
Selection.NumberFormat = "0.00%"
Sheets("Dashboard").Select
    ActiveWorkbook.SlicerCaches("NativeTimeline_Data").TimelineState. _
        SetFilterDateRange "01/01/2019", "31/12/2019"
    ActiveSheet.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "TAPEÇARIA"
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Fill
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.Brightness = 0.400000006
    End With
    Selection.Format.Line.Visible = msoFalse
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "#.##0,0%"
Range("B36").Select
Application.ScreenUpdating = True
End Sub

Sub Setor_cos2()

Application.ScreenUpdating = False

Sheets("Tab. Dinâmica").Select
ActiveSheet.PivotTables("Tabela dinâmica2").ClearTable
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Data")
        .Orientation = xlRowField
        .AutoGroup
    End With
    Range("M4").Select
    Selection.Group Start:=True, End:=True, Periods:=Array(False, False, False, _
        False, True, False, True)
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Ano")
        .PivotItems("2018").Visible = False
        .PivotItems("2019").Visible = True
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Quinzena")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Setor")
        .PivotItems("Tapeçaria").Visible = False
        .PivotItems("Costura").Visible = True
        .PivotItems("Embalagem").Visible = False
        .PivotItems("Espumação").Visible = False
        .PivotItems("Laminação").Visible = False
        .PivotItems("Montagem").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Setor")
        .Orientation = xlPageField
        .Position = 1
    End With
ActiveSheet.PivotTables("Tabela dinâmica2").CalculatedFields.Add "Efi", _
    "='Minutos Produzidos'/'Minutos totais'", True
ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Efi"). _
    Orientation = xlDataField
Columns("N:N").Select
Selection.NumberFormat = "0.00%"
Sheets("Dashboard").Select
    ActiveWorkbook.SlicerCaches("NativeTimeline_Data").TimelineState. _
        SetFilterDateRange "01/01/2019", "31/12/2019"
    ActiveSheet.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "COSTURA"
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Fill
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.Brightness = 0.400000006
    End With
    Selection.Format.Line.Visible = msoFalse
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "#.##0,0%"
Range("B36").Select
Application.ScreenUpdating = True
End Sub

Sub Setor_emb2()

Application.ScreenUpdating = False

Sheets("Tab. Dinâmica").Select
ActiveSheet.PivotTables("Tabela dinâmica2").ClearTable
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Data")
        .Orientation = xlRowField
        .AutoGroup
    End With
    Range("M4").Select
    Selection.Group Start:=True, End:=True, Periods:=Array(False, False, False, _
        False, True, False, True)
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Ano")
        .PivotItems("2018").Visible = False
        .PivotItems("2019").Visible = True
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Quinzena")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Setor")
        .PivotItems("Tapeçaria").Visible = False
        .PivotItems("Costura").Visible = False
        .PivotItems("Embalagem").Visible = True
        .PivotItems("Espumação").Visible = False
        .PivotItems("Laminação").Visible = False
        .PivotItems("Montagem").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Setor")
        .Orientation = xlPageField
        .Position = 1
    End With
ActiveSheet.PivotTables("Tabela dinâmica2").CalculatedFields.Add "Efi", _
    "='Minutos Produzidos'/'Minutos totais'", True
ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Efi"). _
    Orientation = xlDataField
Columns("N:N").Select
Selection.NumberFormat = "0.00%"
Sheets("Dashboard").Select
    ActiveWorkbook.SlicerCaches("NativeTimeline_Data").TimelineState. _
        SetFilterDateRange "01/01/2019", "31/12/2019"
    ActiveSheet.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "EMBALAGEM"
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Fill
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.Brightness = 0.400000006
    End With
    Selection.Format.Line.Visible = msoFalse
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "#.##0,0%"
Range("B36").Select
Application.ScreenUpdating = True
End Sub

Sub Setor_esp2()

Application.ScreenUpdating = False

Sheets("Tab. Dinâmica").Select
ActiveSheet.PivotTables("Tabela dinâmica2").ClearTable
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Data")
        .Orientation = xlRowField
        .AutoGroup
    End With
    Range("M4").Select
    Selection.Group Start:=True, End:=True, Periods:=Array(False, False, False, _
        False, True, False, True)
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Ano")
        .PivotItems("2018").Visible = False
        .PivotItems("2019").Visible = True
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Quinzena")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Setor")
        .PivotItems("Tapeçaria").Visible = False
        .PivotItems("Costura").Visible = False
        .PivotItems("Embalagem").Visible = False
        .PivotItems("Espumação").Visible = True
        .PivotItems("Laminação").Visible = False
        .PivotItems("Montagem").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Setor")
        .Orientation = xlPageField
        .Position = 1
    End With
ActiveSheet.PivotTables("Tabela dinâmica2").CalculatedFields.Add "Efi", _
    "='Minutos Produzidos'/'Minutos totais'", True
ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Efi"). _
    Orientation = xlDataField
Columns("N:N").Select
Selection.NumberFormat = "0.00%"
Sheets("Dashboard").Select
    ActiveWorkbook.SlicerCaches("NativeTimeline_Data").TimelineState. _
        SetFilterDateRange "01/01/2019", "31/12/2019"
    ActiveSheet.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "ESPUMAÇÃO"
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Fill
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.Brightness = 0.400000006
    End With
    Selection.Format.Line.Visible = msoFalse
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "#.##0,0%"
Range("B36").Select
Application.ScreenUpdating = True
End Sub

Sub Setor_lam2()

Application.ScreenUpdating = False

Sheets("Tab. Dinâmica").Select
ActiveSheet.PivotTables("Tabela dinâmica2").ClearTable
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Data")
        .Orientation = xlRowField
        .AutoGroup
    End With
    Range("M4").Select
    Selection.Group Start:=True, End:=True, Periods:=Array(False, False, False, _
        False, True, False, True)
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Ano")
        .PivotItems("2018").Visible = False
        .PivotItems("2019").Visible = True
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Quinzena")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Setor")
        .PivotItems("Tapeçaria").Visible = False
        .PivotItems("Costura").Visible = False
        .PivotItems("Embalagem").Visible = False
        .PivotItems("Espumação").Visible = False
        .PivotItems("Laminação").Visible = True
        .PivotItems("Montagem").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Setor")
        .Orientation = xlPageField
        .Position = 1
    End With
ActiveSheet.PivotTables("Tabela dinâmica2").CalculatedFields.Add "Efi", _
    "='Minutos Produzidos'/'Minutos totais'", True
ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Efi"). _
    Orientation = xlDataField
Columns("N:N").Select
Selection.NumberFormat = "0.00%"
Sheets("Dashboard").Select
    ActiveWorkbook.SlicerCaches("NativeTimeline_Data").TimelineState. _
        SetFilterDateRange "01/01/2019", "31/12/2019"
    ActiveSheet.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "LAMINAÇÃO"
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Fill
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.Brightness = 0.400000006
    End With
    Selection.Format.Line.Visible = msoFalse
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "#.##0,0%"
Range("B36").Select
Application.ScreenUpdating = True
End Sub

Sub Setor_mon2()

Application.ScreenUpdating = False

Sheets("Tab. Dinâmica").Select
ActiveSheet.PivotTables("Tabela dinâmica2").ClearTable
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Data")
        .Orientation = xlRowField
        .AutoGroup
    End With
    Range("M4").Select
    Selection.Group Start:=True, End:=True, Periods:=Array(False, False, False, _
        False, True, False, True)
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Ano")
        .PivotItems("2018").Visible = False
        .PivotItems("2019").Visible = True
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Quinzena")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Setor")
        .PivotItems("Tapeçaria").Visible = False
        .PivotItems("Costura").Visible = False
        .PivotItems("Embalagem").Visible = False
        .PivotItems("Espumação").Visible = False
        .PivotItems("Laminação").Visible = False
        .PivotItems("Montagem").Visible = True
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Setor")
        .Orientation = xlPageField
        .Position = 1
    End With
ActiveSheet.PivotTables("Tabela dinâmica2").CalculatedFields.Add "Efi", _
    "='Minutos Produzidos'/'Minutos totais'", True
ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("Efi"). _
    Orientation = xlDataField
Columns("N:N").Select
Selection.NumberFormat = "0.00%"
Sheets("Dashboard").Select
    ActiveWorkbook.SlicerCaches("NativeTimeline_Data").TimelineState. _
        SetFilterDateRange "01/01/2019", "31/12/2019"
    ActiveSheet.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "MONTAGEM"
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Fill
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.Brightness = 0.400000006
    End With
    Selection.Format.Line.Visible = msoFalse
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "#.##0,0%"
Range("B36").Select
Application.ScreenUpdating = True
End Sub

Sub Mensal_IVA()

Application.ScreenUpdating = False
On Error GoTo erro
    
    With ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Quinzena")
        .SlicerItems("1ª Q").Selected = True
        .SlicerItems("2ª Q").Selected = True
        .SlicerItems("(vazio)").Selected = False
    End With
    Sheets("Tab. Dinâmica").Select
    ActiveWorkbook.ShowPivotTableFieldList = True
    ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("DATA_FATURAMENTO"). _
        Orientation = xlRowField
    ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Quinzena"). _
        Orientation = xlHidden
    ActiveWorkbook.ShowPivotTableFieldList = False
    Sheets("Dashboard").Select
    ActiveWorkbook.SlicerCaches("NativeTimeline_DATA_FATURAMENTO").Slicers( _
        "Data Faturamento 1").TimelineViewState.Level = xlTimelineLevelMonths

Application.ScreenUpdating = True

Exit Sub
erro: Resume Next
End Sub

Sub Quinzena_IVA()

Application.ScreenUpdating = False
On Error GoTo erro
    
    With ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Quinzena")
        .SlicerItems("1ª Q").Selected = True
        .SlicerItems("2ª Q").Selected = True
        .SlicerItems("(vazio)").Selected = False
    End With
    Sheets("Tab. Dinâmica").Select
    ActiveWorkbook.ShowPivotTableFieldList = True
    ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("DATA_FATURAMENTO"). _
        Orientation = xlRowField
    ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Quinzena") _
            .Orientation = xlRowField
    ActiveWorkbook.ShowPivotTableFieldList = False
    Sheets("Dashboard").Select
    ActiveWorkbook.SlicerCaches("NativeTimeline_DATA_FATURAMENTO").Slicers( _
        "Data Faturamento 1").TimelineViewState.Level = xlTimelineLevelMonths

Application.ScreenUpdating = True

Exit Sub
erro: Resume Next
End Sub

Sub Anual_IVA()

Application.ScreenUpdating = False
On Error GoTo erro

    With ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Quinzena")
        .SlicerItems("1ª Q").Selected = True
        .SlicerItems("2ª Q").Selected = True
        .SlicerItems("(vazio)").Selected = False
    End With
    Sheets("Tab. Dinâmica").Select
    ActiveWorkbook.ShowPivotTableFieldList = True
    ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("DATA_FATURAMENTO"). _
        Orientation = xlHidden
    ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Quinzena"). _
        Orientation = xlHidden
    ActiveWorkbook.ShowPivotTableFieldList = False
    Sheets("Dashboard").Select
    ActiveWorkbook.SlicerCaches("NativeTimeline_DATA_FATURAMENTO").Slicers( _
        "Data Faturamento 1").TimelineViewState.Level = xlTimelineLevelYears

Application.ScreenUpdating = True

Exit Sub
erro: Resume Next
End Sub

Sub Setor_tap()

Application.ScreenUpdating = False

Sheets("Tab. Dinâmica").Select
ActiveSheet.PivotTables("Tabela dinâmica1").ClearTable
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("DATA_FATURAMENTO")
        .Orientation = xlRowField
        .AutoGroup
    End With
    Range("J4").Select
    Selection.Group Start:=True, End:=True, Periods:=Array(False, False, False, _
        False, True, False, True)
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Anos")
        .PivotItems("2017").Visible = False
        .PivotItems("2018").Visible = False
        .PivotItems("2019").Visible = True
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Quinzena")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Filtro 5")
        .PivotItems("1").Visible = True
        .PivotItems("0").Visible = False
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Filtro 5")
        .Orientation = xlPageField
        .Position = 1
    End With
ActiveSheet.PivotTables("Tabela dinâmica1").CalculatedFields.Add "IVA Tap", _
    "=VL_TOTAL_PEDIDOVENDA_ITEM /'Tempo Tapeçaria'", True
ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("IVA Tap"). _
    Orientation = xlDataField
Columns("K:K").Select
Selection.NumberFormat = "0.00"
Sheets("Dashboard").Select
    ActiveWorkbook.SlicerCaches("NativeTimeline_DATA_FATURAMENTO").TimelineState. _
        SetFilterDateRange "01/01/2019", "31/12/2019"
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.ChartTitle.Text = "TAPEÇARIA"
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Fill
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.Brightness = 0.400000006
    End With
    Selection.Format.Line.Visible = msoFalse
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "#.##0,00"
Range("A1").Select
Application.ScreenUpdating = True
End Sub

Sub Setor_cos()

Application.ScreenUpdating = False
    
Sheets("Tab. Dinâmica").Select
ActiveSheet.PivotTables("Tabela dinâmica1").ClearTable
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("DATA_FATURAMENTO")
        .Orientation = xlRowField
        .AutoGroup
    End With
    Range("J4").Select
    Selection.Group Start:=True, End:=True, Periods:=Array(False, False, False, _
        False, True, False, True)
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Anos")
        .PivotItems("2017").Visible = False
        .PivotItems("2018").Visible = False
        .PivotItems("2019").Visible = True
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Quinzena")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Filtro 6")
        .PivotItems("1").Visible = True
        .PivotItems("0").Visible = False
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Filtro 6")
        .Orientation = xlPageField
        .Position = 1
    End With
ActiveSheet.PivotTables("Tabela dinâmica1").CalculatedFields.Add "IVA Cos", _
    "=VL_TOTAL_PEDIDOVENDA_ITEM /'Tempo Costura'", True
ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("IVA Cos"). _
    Orientation = xlDataField
Columns("K:K").Select
Selection.NumberFormat = "0.00"
Sheets("Dashboard").Select
    ActiveWorkbook.SlicerCaches("NativeTimeline_DATA_FATURAMENTO").TimelineState. _
        SetFilterDateRange "01/01/2019", "31/12/2019"
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.ChartTitle.Text = "COSTURA"
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Fill
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.Brightness = 0.400000006
    End With
    Selection.Format.Line.Visible = msoFalse
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "#.##0,00"
Range("A1").Select
Application.ScreenUpdating = True
End Sub

Sub Setor_emb()

Application.ScreenUpdating = False
    
Sheets("Tab. Dinâmica").Select
ActiveSheet.PivotTables("Tabela dinâmica1").ClearTable
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("DATA_FATURAMENTO")
        .Orientation = xlRowField
        .AutoGroup
    End With
    Range("J4").Select
    Selection.Group Start:=True, End:=True, Periods:=Array(False, False, False, _
        False, True, False, True)
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Anos")
        .PivotItems("2017").Visible = False
        .PivotItems("2018").Visible = False
        .PivotItems("2019").Visible = True
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Quinzena")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Filtro 1")
        .PivotItems("1").Visible = True
        .PivotItems("0").Visible = False
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Filtro 1")
        .Orientation = xlPageField
        .Position = 1
    End With
ActiveSheet.PivotTables("Tabela dinâmica1").CalculatedFields.Add "IVA Emb", _
    "=VL_TOTAL_PEDIDOVENDA_ITEM /'Tempo Embalagem'", True
ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("IVA Emb"). _
    Orientation = xlDataField
Columns("K:K").Select
Selection.NumberFormat = "0.00"
Sheets("Dashboard").Select
    ActiveWorkbook.SlicerCaches("NativeTimeline_DATA_FATURAMENTO").TimelineState. _
        SetFilterDateRange "01/01/2019", "31/12/2019"
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.ChartTitle.Text = "EMBALAGEM"
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Fill
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.Brightness = 0.400000006
    End With
    Selection.Format.Line.Visible = msoFalse
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "#.##0,00"

Application.ScreenUpdating = True
Range("A1").Select
End Sub

Sub Setor_esp()

Application.ScreenUpdating = False
    
Sheets("Tab. Dinâmica").Select
ActiveSheet.PivotTables("Tabela dinâmica1").ClearTable
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("DATA_FATURAMENTO")
        .Orientation = xlRowField
        .AutoGroup
    End With
    Range("J4").Select
    Selection.Group Start:=True, End:=True, Periods:=Array(False, False, False, _
        False, True, False, True)
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Anos")
        .PivotItems("2017").Visible = False
        .PivotItems("2018").Visible = False
        .PivotItems("2019").Visible = True
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Quinzena")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Filtro 2")
        .PivotItems("1").Visible = True
        .PivotItems("0").Visible = False
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Filtro 2")
        .Orientation = xlPageField
        .Position = 1
    End With
ActiveSheet.PivotTables("Tabela dinâmica1").CalculatedFields.Add "IVA Esp", _
    "=VL_TOTAL_PEDIDOVENDA_ITEM /'Tempo Espumação'", True
ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("IVA Esp"). _
    Orientation = xlDataField
Columns("K:K").Select
Selection.NumberFormat = "0.00"
Sheets("Dashboard").Select
    ActiveWorkbook.SlicerCaches("NativeTimeline_DATA_FATURAMENTO").TimelineState. _
        SetFilterDateRange "01/01/2019", "31/12/2019"
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.ChartTitle.Text = "ESPUMAÇÃO"
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Fill
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.Brightness = 0.400000006
    End With
    Selection.Format.Line.Visible = msoFalse
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "#.##0,00"

Application.ScreenUpdating = True
Range("A1").Select
End Sub

Sub Setor_lam()

Application.ScreenUpdating = False
    
Sheets("Tab. Dinâmica").Select
ActiveSheet.PivotTables("Tabela dinâmica1").ClearTable
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("DATA_FATURAMENTO")
        .Orientation = xlRowField
        .AutoGroup
    End With
    Range("J4").Select
    Selection.Group Start:=True, End:=True, Periods:=Array(False, False, False, _
        False, True, False, True)
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Anos")
        .PivotItems("2017").Visible = False
        .PivotItems("2018").Visible = False
        .PivotItems("2019").Visible = True
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Quinzena")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Filtro 3")
        .PivotItems("1").Visible = True
        .PivotItems("0").Visible = False
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Filtro 3")
        .Orientation = xlPageField
        .Position = 1
    End With
ActiveSheet.PivotTables("Tabela dinâmica1").CalculatedFields.Add "IVA Lam", _
    "=VL_TOTAL_PEDIDOVENDA_ITEM /'Tempo Laminação'", True
ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("IVA Lam"). _
    Orientation = xlDataField
Columns("K:K").Select
Selection.NumberFormat = "0.00"
Sheets("Dashboard").Select
    ActiveWorkbook.SlicerCaches("NativeTimeline_DATA_FATURAMENTO").TimelineState. _
        SetFilterDateRange "01/01/2019", "31/12/2019"
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.ChartTitle.Text = "LAMINAÇÃO"
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Fill
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.Brightness = 0.400000006
    End With
    Selection.Format.Line.Visible = msoFalse
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "#.##0,00"

Application.ScreenUpdating = True
Range("A1").Select
End Sub

Sub Setor_mon()

Application.ScreenUpdating = False
    
Sheets("Tab. Dinâmica").Select
ActiveSheet.PivotTables("Tabela dinâmica1").ClearTable
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("DATA_FATURAMENTO")
        .Orientation = xlRowField
        .AutoGroup
    End With
    Range("J4").Select
    Selection.Group Start:=True, End:=True, Periods:=Array(False, False, False, _
        False, True, False, True)
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Anos")
        .PivotItems("2017").Visible = False
        .PivotItems("2018").Visible = False
        .PivotItems("2019").Visible = True
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Quinzena")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Filtro 4")
        .PivotItems("1").Visible = True
        .PivotItems("0").Visible = False
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Filtro 4")
        .Orientation = xlPageField
        .Position = 1
    End With
ActiveSheet.PivotTables("Tabela dinâmica1").CalculatedFields.Add "IVA Mon", _
    "=VL_TOTAL_PEDIDOVENDA_ITEM /'Tempo Montagem'", True
ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("IVA Mon"). _
    Orientation = xlDataField
Columns("K:K").Select
Selection.NumberFormat = "0.00"
Sheets("Dashboard").Select
    ActiveWorkbook.SlicerCaches("NativeTimeline_DATA_FATURAMENTO").TimelineState. _
        SetFilterDateRange "01/01/2019", "31/12/2019"
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.ChartTitle.Text = "MONTAGEM"
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Fill
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.Brightness = 0.400000006
    End With
    Selection.Format.Line.Visible = msoFalse
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "#.##0,00"

Application.ScreenUpdating = True
Range("A1").Select
End Sub
