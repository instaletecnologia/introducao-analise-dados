Sub CalcularDecilesCicloComFiltro()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cicloCol As Long, decileCicloCol As Long, secondCol As Long
    Dim cicloRange As Range, cell As Range
    Dim deciles(1 To 10) As Double
    Dim i As Long
    Dim totalSeconds As Long
    Dim times() As Double

    ' Defina a planilha
    Set ws = ThisWorkbook.Sheets("DADOS")

    ' Identificar a coluna "Ciclo"
    cicloCol = ws.Rows(1).Find(What:="Ciclo", LookIn:=xlValues, LookAt:=xlWhole).Column
    If cicloCol = 0 Then
        MsgBox "Coluna 'Ciclo' não encontrada!", vbExclamation
        Exit Sub
    End If

    ' Verifique se a coluna "Decile Ciclo" já existe
    Dim findResult As Range
    Set findResult = ws.Rows(1).Find(What:="Decile Ciclo", LookIn:=xlValues, LookAt:=xlWhole)

    ' Se não existir, "decileCicloCol" será 0 e você precisará criar a coluna
    If Not findResult Is Nothing Then
        decileCicloCol = findResult.Column
    Else
        decileCicloCol = 0
    End If

    secondCol = cicloCol + 1 ' Coluna para os segundos

    ' Se a coluna "Decile Ciclo" não existir, insira uma nova coluna
    If decileCicloCol = 0 Then
        ws.Cells(1, secondCol).EntireColumn.Insert
        ws.Cells(1, secondCol).Value = "Ciclo (Segundos)"
        ws.Cells(1, secondCol + 1).Value = "Decile Ciclo"
        decileCicloCol = secondCol + 1
    End If

    ' Define o intervalo para "Ciclo"
    lastRow = ws.Cells(ws.Rows.Count, cicloCol).End(xlUp).Row
    Set cicloRange = ws.Range(ws.Cells(2, cicloCol), ws.Cells(lastRow, cicloCol))

    ' Converte os tempos em segundos e armazena numa matriz
    ReDim times(1 To lastRow - 1)
    i = 1
    For Each cell In cicloRange
        totalSeconds = Hour(cell.Value) * 3600 + Minute(cell.Value) * 60 + Second(cell.Value)
        times(i) = totalSeconds
        cell.Offset(0, 1).Value = totalSeconds ' Coloca o valor dos segundos na coluna ao lado
        cell.Offset(0, 1).NumberFormat = "0" ' Formata a célula para exibir números inteiros
        i = i + 1
    Next cell

    ' Calcule os deciles com base nos segundos
    For i = 1 To 10
        deciles(i) = Application.WorksheetFunction.Percentile_Inc(times, i / 10)
    Next i

    ' Classifique os dados da coluna "Ciclo" baseado nos segundos, mas apenas nos dados filtrados
    i = 1
    For Each cell In cicloRange
        If cell.EntireRow.Hidden = False Then ' Verifica se a linha está visível (filtrada)
            For j = 1 To 10
                If times(i) <= deciles(j) Then
                    cell.Offset(0, 2).Value = j ' Coloca o decile na coluna "Decile Ciclo"
                    Exit For
                End If
            Next j
        End If
        i = i + 1
    Next cell

End Sub
