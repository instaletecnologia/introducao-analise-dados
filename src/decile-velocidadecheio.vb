Sub CalcularDecilesVelocidadeMediaCheio()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DADOS") ' Modifique para o nome da sua planilha

    Dim velocidadeMediaCheioCol As Long
    velocidadeMediaCheioCol = ws.Rows(1).Find(What:="Vel. Cheio KM/H", LookIn:=xlValues, LookAt:=xlWhole).Column
    If velocidadeMediaCheioCol = 0 Then
        MsgBox "Coluna 'Vel. Cheio KM/H' não encontrada!", vbExclamation
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, velocidadeMediaCheioCol).End(xlUp).Row

    Dim velocidadeMediaCheioRange As Range
    Set velocidadeMediaCheioRange = ws.Range(ws.Cells(2, velocidadeMediaCheioCol), ws.Cells(lastRow, velocidadeMediaCheioCol))

    ' Adiciona cabeçalho "Decile Velocidade Média" na coluna adjacente à direita
    Dim decileCol As Long
    decileCol = velocidadeMediaCheioCol + 1
    ws.Cells(1, decileCol).Value = "Decile Velocidade Média Cheio"

    ' Calcula os valores dos deciles
    Dim deciles(1 To 10) As Double
    For i = 1 To 10
        deciles(i) = Application.WorksheetFunction.Percentile_Inc(velocidadeMediaCheioRange, i / 10)
    Next i

    ' Classificar cada entrada de velocidade média em seu respectivo decile
    Dim cell As Range
    For Each cell In velocidadeMediaCheioRange
        If IsNumeric(cell.Value) Then
            Dim valor As Double
            valor = CDbl(cell.Value)
            For i = 1 To 10
                If valor <= deciles(i) Then
                    ws.Cells(cell.Row, decileCol).Value = i
                    Exit For
                End If
            Next i
        End If
    Next cell

MsgBox "Deciles para Velocidade Média Cheio calculados com sucesso!"

End Sub
