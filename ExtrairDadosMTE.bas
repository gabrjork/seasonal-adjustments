
Attribute VB_Name = "ModuloProcessoCompleto"
Sub ProcessoCompletoTabelas_Final_V5()

    Dim ws As Worksheet, wsSetorial As Worksheet, wsSourceT51 As Worksheet
    Dim ultimaColunaOriginal As Long, coluna As Long, valorCelula As String
    Dim colunasParaManter As Object, dataAtual As Date, dataSetorial As Date
    Dim colunaFinal As Long, novaUltimaColuna As Long, linhaSetorial As Long
    Dim rangeParaCopiar As Range, rangeToCopyT51 As Range
    Dim lastRowSource As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Tabela 6.1")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "A planilha 'Tabela 6.1' não foi encontrada.", vbCritical: Exit Sub

    ultimaColunaOriginal = ws.Cells(6, ws.Columns.Count).End(xlToLeft).Column
    Set colunasParaManter = CreateObject("Scripting.Dictionary")
    For coluna = 1 To ultimaColunaOriginal
        valorCelula = ws.Cells(6, coluna).Value
        If Not IsEmpty(valorCelula) Then
            If InStr(1, LCase(valorCelula), "saldo") > 0 Then
                colunasParaManter.Add coluna, True
            End If
        End If
    Next coluna

    Application.ScreenUpdating = False
    For coluna = ultimaColunaOriginal To 1 Step -1
        If coluna <> 1 And coluna <> 2 And Not colunasParaManter.Exists(coluna) Then
            ws.Columns(coluna).Delete
        End If
    Next coluna

    novaUltimaColuna = ws.Cells(6, ws.Columns.Count).End(xlToLeft).Column
    If novaUltimaColuna > 2 Then
        dataAtual = DateSerial(2020, 1, 1)
        For colunaFinal = 3 To novaUltimaColuna
            ws.Cells(5, colunaFinal).Value = dataAtual
            ws.Cells(5, colunaFinal).NumberFormat = "mm/yyyy"
            dataAtual = DateAdd("m", 1, dataAtual)
        Next colunaFinal
    End If

    Set wsSetorial = Nothing
    On Error Resume Next
    Set wsSetorial = ThisWorkbook.Sheets("Setorial")
    On Error GoTo 0

    If Not wsSetorial Is Nothing Then
        wsSetorial.Cells.ClearContents
    Else
        Set wsSetorial = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsSetorial.Name = "Setorial"
    End If

    If novaUltimaColuna >= 2 Then
        Set rangeParaCopiar = ws.Range(ws.Cells(7, 2), ws.Cells(34, novaUltimaColuna))
        rangeParaCopiar.Copy
        wsSetorial.Range("B2").PasteSpecial Paste:=xlPasteAll, Transpose:=True
        Application.CutCopyMode = False

        dataSetorial = DateSerial(2020, 1, 1)
        For linhaSetorial = 3 To novaUltimaColuna
            wsSetorial.Cells(linhaSetorial, 1).Value = dataSetorial
            wsSetorial.Cells(linhaSetorial, 1).NumberFormat = "mm/yyyy"
            dataSetorial = DateAdd("m", 1, dataSetorial)
        Next linhaSetorial
    End If

    wsSetorial.Columns("B:C").Insert Shift:=xlToRight
    On Error Resume Next
    Set wsSourceT51 = ThisWorkbook.Sheets("Tabela 5.1")
    On Error GoTo 0

    If Not wsSourceT51 Is Nothing Then
        Dim lastRowD As Long, lastRowE As Long
        lastRowD = wsSourceT51.Cells(wsSourceT51.Rows.Count, "D").End(xlUp).Row
        lastRowE = wsSourceT51.Cells(wsSourceT51.Rows.Count, "E").End(xlUp).Row
        lastRowSource = Application.WorksheetFunction.Max(lastRowD, lastRowE)
        If lastRowSource >= 5 Then
            Set rangeToCopyT51 = wsSourceT51.Range("D5:E" & lastRowSource)
            rangeToCopyT51.Copy
            wsSetorial.Range("B2").PasteSpecial Paste:=xlPasteAll
            Application.CutCopyMode = False
        End If
    End If

    Dim wsTemp As Worksheet
    Application.DisplayAlerts = False
    For Each wsTemp In ThisWorkbook.Sheets
        If wsTemp.Name <> "Setorial" Then wsTemp.Delete
    Next wsTemp
    Application.DisplayAlerts = True

    wsSetorial.Rows(1).Delete
    wsSetorial.Range("A1").Value = "Data"

    With wsSetorial.Cells
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = False
        .Font.Color = RGB(0, 0, 0)
        .Interior.ColorIndex = xlNone
        .HorizontalAlignment = xlRight
        .Borders.LineStyle = xlNone
    End With
    wsSetorial.Columns(1).NumberFormat = "dd-mm-yyyy"

    Dim cel As Range
    For Each cel In wsSetorial.UsedRange
        If VarType(cel.Value) = vbString Then
            If InStr(cel.Value, ";") > 0 Then
                cel.Value = Replace(cel.Value, ";", ",")
            End If
        End If
    Next cel

    Dim ultimaLinha As Long
    ultimaLinha = wsSetorial.Cells(wsSetorial.Rows.Count, 1).End(xlUp).Row
    If ultimaLinha >= 2 Then
        wsSetorial.Rows(ultimaLinha).Delete
        wsSetorial.Rows(ultimaLinha - 1).Delete
    End If

    Dim i As Long, j As Long
    Dim ultimaColuna As Long
    ultimaLinha = wsSetorial.Cells(wsSetorial.Rows.Count, 1).End(xlUp).Row
    ultimaColuna = wsSetorial.Cells(1, wsSetorial.Columns.Count).End(xlToLeft).Column
    For i = 2 To ultimaLinha
        For j = 1 To ultimaColuna
            If IsNumeric(wsSetorial.Cells(i, j).Value) Then
                wsSetorial.Cells(i, j).Value = CDbl(wsSetorial.Cells(i, j).Value)
            End If
        Next j
    Next i

    Dim csvContent As String, linhaCSV As String
    Dim csvPath As String
    csvPath = ThisWorkbook.Path & Application.PathSeparator & "2025_Cheio.csv"

    csvContent = ""
    For i = 1 To ultimaLinha
        linhaCSV = ""
        For j = 1 To ultimaColuna
            Dim valorNumerico As Variant
            valorNumerico = wsSetorial.Cells(i, j).Value
            If IsNumeric(valorNumerico) Then
                valorCelula = Replace(Format(CDbl(valorNumerico), "0.00"), ".", ",")
            Else
                valorCelula = wsSetorial.Cells(i, j).Text
            End If
            linhaCSV = linhaCSV & valorCelula
            If j < ultimaColuna Then linhaCSV = linhaCSV & ";"
        Next j
        csvContent = csvContent & linhaCSV & vbCrLf
    Next i

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Charset = "utf-8"
        .Open
        .WriteText csvContent
        .SaveToFile csvPath, 2
        .Close
    End With
    Set stream = Nothing

    Application.ScreenUpdating = True
    MsgBox "Processo finalizado com sucesso.", vbInformation

End Sub
