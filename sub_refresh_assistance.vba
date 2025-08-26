Sub Atualizar_1()

Dim x, y, w, z, k As Long

tempoinicial = Timer
    
Application.ScreenUpdating = False
ActiveWorkbook.RefreshAll

    
' Atualiza tabela dinâmica Subgrupos e conta a diferença de linhas.
' Calcula os Subgrupos a partir da difença de linhas
'   e passa formulas pra texto.

    Sheets("Subgrupos").Select
    ActiveSheet.PivotTables("Tabela dinâmica6").PivotCache.Refresh
    x = Range("A3").End(xlDown).Row
    Range("C2:J2").Select
    Selection.Copy
    Range("C4:J" & x).Select
    ActiveSheet.Paste
    Calculate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
' Conta as linhas adicionadas em "Cadastro NC"
'   e as linhas existentes do "Controle NC".
    
    x = Sheets("Controle NC").Range("A" & Rows.Count).End(xlUp).Row
    y = Sheets("Cadastro NC").Range("B" & Rows.Count).End(xlUp).Row - 2

' Passa o conteúdo adicionado em "Cadastro NC"
'   e passa para o "Controle NC" organizando-os.

  For k = 2 To 28
    z = Sheets("Cadastro NC").Cells(1, k).Value
        For w = 1 To y
         Sheets("Controle NC").Cells(x + w, z) = Sheets("Cadastro NC").Cells(w + 2, k)
        Next w
  Next k

' Limpa o "Cadastro NC" para inserção de futuros dados.

    Sheets("Cadastro NC").Select
    w = Sheets("Cadastro NC").Range("B" & Rows.Count).End(xlUp).Row
    Range("B3:AB" & w).Select
    Selection.Clear
    
' Compara a diferença de linhas para copiar e colar
'   fórmulas nas colunas AJ,AK,AL

    Sheets("Controle NC").Select
    x = Range("L2").End(xlDown).Row
    Range("AJ1:AL1").Select
    Selection.Copy
    Range("AJ3:AL" & x).Select
    ActiveSheet.Paste
    Calculate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
' Compara a diferença de linhas para copiar e colar
'   fórmula nas células que faltam na coluna AN de "Controle NC".

    Sheets("Controle NC").Select
    x = Range("AJ2").End(xlDown).Row
    Range("AM3").FormulaLocal = "=SE(L3=$AM$1;AM2;SEERRO(ARRUMAR(SEERRO(ÍNDICE('Base faturamento'!$K:$K;CORRESP('Controle NC'!B3;'Base faturamento'!$AV:$AV;0));ÍNDICE('Base faturamento'!$K:$K;CORRESP('Controle NC'!C3;'Base faturamento'!$AV:$AV;0))));ARRUMAR(PROCV(L3;'Clientes Ajuste'!$A:$B;2;0))))"
    Range("AM3:AM" & x).Select
    Selection.FillDown
    Calculate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
    
'Formata "Controle NC"

    Range("A3:AO" & x).Select
    With Selection
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
    With Selection.Font
        .Name = "Arial"
        .Size = 10
        .Bold = False
    End With
    Selection.HorizontalAlignment = xlCenter
    
' Volta para a "Dashboard".
    
    Sheets("Dashboard").Select
    Application.CutCopyMode = False
    tempofinal = (Timer - tempoinicial) / 60
    MsgBox ("Parte 1 finalizada! Ajuste os nomes dos clientes" & Chr(13) & "Tempo: " & Round(tempofinal, 2) & " minutos")
    
End Sub
' Ajustar clientes na planilha "Clientes Ajustados" antes de executar parte 2.


Sub Atualizar_2()

' Executar Parte 2 somente após ter ajustados clientes
'   na planilha "Clientes Ajustados".

Dim x, y, k As Long, z As String

tempoinicial = Timer

On Error GoTo erro

' Atualização de tela em modo off e pergunta.
    
    Application.ScreenUpdating = False
    x = InputBox("Houve ajuste de clientes?(0-NÃO / 1-SIM)")

' Compara a diferença de linhas para copiar e colar
'   fórmula nas células que faltam na coluna AN de "Controle NC".

If x = 1 Then
    Sheets("Controle NC").Select
    ActiveSheet.ShowAllData
    x = Range("AM1").End(xlDown).Row
    Range("AM3").FormulaLocal = "=SE(L3=$AM$1;AM2;SEERRO(ARRUMAR(SEERRO(ÍNDICE('Base faturamento'!$K:$K;CORRESP('Controle NC'!B3;'Base faturamento'!$AV:$AV;0));ÍNDICE('Base faturamento'!$K:$K;CORRESP('Controle NC'!C3;'Base faturamento'!$AV:$AV;0))));ARRUMAR(PROCV(L3;'Clientes Ajuste'!$A:$B;2;0))))"
    Range("AM3:AM" & x).Select
    Selection.FillDown
    Calculate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
Else
End If
    
' Compara a diferença de linhas para copiar e colar fórmula nas
'   células que faltam nas colunas AU,AV,AW,AX de "Base de faturamento".
    
    Sheets("Base faturamento").Select
    x = Range("A" & Rows.Count).End(xlUp).Row
    Range("AX1:BA1").Select
    Selection.Copy
    Range("AX4:BA" & x).Select
    ActiveSheet.Paste
    Calculate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False

    
' Pôe filtro e copia nome fantasia de clientes com assistência
'   "Base de faturamento" e "Controle NC" e cola em "Planilha de calculo".
'   Após isso, desduplica, enumera clientes com assistência e tira filtro.

    Sheets("Base faturamento").Select
    x = Range("$A" & Rows.Count).End(xlUp).Row
    ActiveSheet.Range("$AX$3:$BD$" & x).AutoFilter Field:=1, Criteria1:="1"
    Range("$K$4:K$" & x).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    Sheets("Planilha de calculo").Select
    Range("$B$17").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    y = Range("$B$17").End(xlDown).Row + 1
    Range("B" & y).Select
    Sheets("Controle NC").Select
    y = Range("$AM" & Rows.Count).End(xlUp).Row
    Range("$AM$3", "$AM" & y).Select
    Selection.Copy
    Sheets("Planilha de calculo").Select
    ActiveSheet.Paste
    ActiveSheet.Range("$B$16:$B$" & Range("$B" & Rows.Count).End(xlUp).Row).RemoveDuplicates Columns:=1, Header:= _
        xlYes
    y = Range("B" & Rows.Count).End(xlUp).Row
    For k = 17 To y
        Cells(k, 3) = k - 16
    Next k
    Sheets("Base faturamento").Range("$AX$3:$BD$" & x).AutoFilter Field:=1
    
' Atualiza tabela dinâmica de produtos com assistência
'   em "Planilha de calculo" e enumera todos.
    
    Sheets("Planilha de calculo").Select
    ActiveSheet.PivotTables("Tabela dinâmica2").PivotCache.Refresh
    x = Range("I" & Rows.Count).End(xlUp).Row - 1
    For k = 62 To x
        Cells(k, 10) = k - 61
    Next k

' Compara linhas para copiar e colar fórmulas nas células
'   vazias em "Controle NC" nas colunas AN e AO.

    Sheets("Controle NC").Select
    x = Range("AM" & Rows.Count).End(xlUp).Row
    Range("AN1:AO1").Select
    Selection.Copy
    Range("AN3:AO" & x).Select
    ActiveSheet.Paste
    Calculate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
    

' Compara linhas para copiar e colar fórmulas nas células
'   vazias em "Base de faturamento" na coluna AZ.

    Sheets("Base faturamento").Select
    x = Range("AZ" & Rows.Count).End(xlUp).Row
    Range("BB1:BD1").Select
    Selection.Copy
    Range("BB4:BD" & x).Select
    ActiveSheet.Paste
    Calculate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
    
    
' Conta ultima célula dos clientes com assistência em "Planilha de calculo"
'   e depois copia e cola fórmulas em todas as células. Após isso,
'   insere a formula para contar e somar as colunas
    
    Sheets("Planilha de calculo").Select
    k = Range("B" & Rows.Count).End(xlUp).Row
    Range("$D$15:$F$15").Copy
    Range("$D$17:$F$17").Select
    ActiveSheet.Paste
    Range("$D$17:$F$" & k).Select
    Selection.FillDown
    Calculate
    Range("D" & k + 1).Select
    ActiveCell.FormulaLocal = "=SOMA(D17:D" & k & ")"
    Range("D" & k + 1 & ":F" & k + 1).Select
    Selection.FillRight
    Calculate
    
' Conta ultima célula dos produtos com assistência em "Planilha de calculo"
'   e depois copia e cola fórmulas em todas as células. Após isso,
'   insere a formula para contar e somar as colunas
    
    Sheets("Planilha de calculo").Select
    k = Range("I" & Rows.Count).End(xlUp).Row - 1
    Range("$K$60:$N$60").Copy
    Range("$K$62:$N62").Select
    ActiveSheet.Paste
    Range("$K$62:$N$" & k).Select
    Selection.FillDown
    Calculate
    Range("K" & k + 1).Select
    ActiveCell.FormulaLocal = "=SOMA(K62:K" & k & ")"
    Range("K" & k + 1 & ":N" & k + 1).Select
    Selection.FillRight
    Calculate
    
'Formata "Planilha de calculo", insere Total geral, põe em negrito e pinta célula
    
    Sheets("Planilha de calculo").Select
    x = Range("B" & Rows.Count).End(xlUp).Row
    y = Range("I" & Rows.Count).End(xlUp).Row
    Range("B17:F" & x).Select
    Range("B" & x + 1) = Range("I" & y)
    Range("B" & x + 1 & ":F" & x + 1).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("J" & y & ":N" & y).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    
' Conta numero de linhas de clientes com assistencia em "Planilha de calculo"
'   e atualiza lista de clientes com assistencia no "Dashboard"
    
    w = Sheets("Planilha de calculo").Range("$B$17").End(xlDown).Row
    Sheets("Dashboard").Select
    Range("C150:G150").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="='Planilha de calculo'!$B$17:$B$" & w
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
'Insere fórmulas no Dashboard dinamicamente e arrasta para baixo
    
    Sheets("Planilha de calculo").Select
    x = Range("B" & Rows.Count).End(xlUp).Row - 1
    y = Range("I" & Rows.Count).End(xlUp).Row - 1
    Sheets("Dashboard").Select
    Range("I4").FormulaLocal = "=(SOMASES('Base faturamento'!$X:$X;'Base faturamento'!$BD:$BD;1;'Base faturamento'!$AX:$AX;1)/1000)-SOMA('Planilha de calculo'!$K$62:$K$" & y & ")"
    Range("L4").FormulaLocal = "=(SOMASES('Base faturamento'!$X:$X;'Base faturamento'!$AX:$AX;1;'Base faturamento'!$BD:$BD;1)/1000)-SOMA('Planilha de calculo'!$D$17:$D$" & x & ")"
    Range("O4").FormulaLocal = "=SOMA('Planilha de calculo'!$F$9:$Q$9)-SOMASES('Controle NC'!$S:$S;'Controle NC'!$AO:$AO;1)"
    Range("R4").FormulaLocal = "=SOMA('Planilha de calculo'!$F$17:$F$" & x & ")-SOMASES('Controle NC'!$S:$S;'Controle NC'!$AO:$AO;1)"
    
    Range("C72").FormulaLocal = "=SEERRO(ÍNDICE('Planilha de calculo'!$B$17:$B$" & x & ";CORRESP(SE(Dashboard!I72=0;"""";Dashboard!I72);'Planilha de calculo'!$D$17:$D$" & x & ";0);1);"""")"
    Range("I72").FormulaLocal = "=MAIOR('Planilha de calculo'!$D$17:$D$" & x & ";'Planilha de calculo'!$H16)"
    Range("K72").FormulaLocal = "=SEERRO(Dashboard!I72/'Planilha de calculo'!K16;"""")"
    Range("M72").FormulaLocal = "=SEERRO(ÍNDICE('Planilha de calculo'!$B$17:$B$" & x & ";CORRESP(SE(Dashboard!S72=0;"""";Dashboard!S72);'Planilha de calculo'!$E$17:$E$" & x & ";0);1);"""")"
    Range("S72").FormulaLocal = "=MAIOR('Planilha de calculo'!$E$17:$E$" & x & ";'Planilha de calculo'!$H16)"
    Range("U72").FormulaLocal = "=SEERRO(Dashboard!S72/'Planilha de calculo'!T16;"""")"
    
    Range("C72:K96,M72:U96").Select
    Selection.FillDown
    
    Range("C104").FormulaLocal = "=ÍNDICE('Planilha de calculo'!$I$62:$I$" & y & ";CORRESP(Dashboard!$I104;'Planilha de calculo'!$K$62:$K$" & y & ";0);1)"
    Range("I104").FormulaLocal = "=MAIOR('Planilha de calculo'!$K$62:$K$" & y & ";'Planilha de calculo'!$R61)"
    Range("K104").FormulaLocal = "=Dashboard!I104/'Planilha de calculo'!S61"
    Range("M104").FormulaLocal = "=SEERRO(ÍNDICE('Planilha de calculo'!$I$62:$I$" & y & ";CORRESP(SE(Dashboard!$S104=0;"""";Dashboard!$S104);'Planilha de calculo'!$L$62:$L$" & y & ";0);1);"""")"
    Range("S104").FormulaLocal = "=MAIOR('Planilha de calculo'!$L$62:$L$" & y & ";'Planilha de calculo'!$R61)"
    Range("U104").FormulaLocal = "=SEERRO(Dashboard!S104/'Planilha de calculo'!Y61;"""")"
    
    Range("C104:K113,M104:U113").Select
    Selection.FillDown
    
    Range("C193").FormulaLocal = "=SE(Dashboard!I193=0;"""";ÍNDICE('Planilha de calculo'!$I$62:$I$" & y & ";CORRESP(Dashboard!I193;'Planilha de calculo'!M$62:M$" & y & ";0);1))"
    Range("I193").FormulaLocal = "=MAIOR('Planilha de calculo'!M$62:M$" & y & ";'Planilha de calculo'!$S87)"
    Range("K193").FormulaLocal = "=SEERRO(Dashboard!I193/'Planilha de calculo'!Z87;"""")"
    Range("M193").FormulaLocal = "=SE(Dashboard!S193=0;"""";ÍNDICE('Planilha de calculo'!$I$62:$I$" & y & ";CORRESP(Dashboard!S193;'Planilha de calculo'!N$62:N$" & y & ";0);1))"
    Range("S193").FormulaLocal = "=MAIOR('Planilha de calculo'!N$62:N$" & y & ";'Planilha de calculo'!$S87)"
    Range("U193").FormulaLocal = "=SEERRO(Dashboard!S193/'Planilha de calculo'!AE87;"""")"
    
    Range("C193:K197,M193:U197").Select
    Selection.FillDown
    
'Formata as bordas das tabelas dashboard
    
    Range("C72:K96,M72:U96,C104:K113,M104:U113,C193:K197,M193:U197").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    Calculate
    tempofinal = (Timer - tempoinicial) / 60
    MsgBox ("Parte 2 finalizada com sucesso!" & Chr(13) & "Tempo: " & Round(tempofinal, 2) & " minutos")
    
Exit Sub
erro: Resume Next

End Sub
