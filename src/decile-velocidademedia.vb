Sub CalcularDecilesVelocidadeMedia()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DADOS") ' Modifique para o nome da sua planilha

    Dim velocidadeMediaCol As Long
    velocidadeMediaCol = ws.Rows(1).Find(What:="Vel. Média KM/H", LookIn:=xlValues, LookAt:=xlWhole).Column
    If velocidadeMediaCol = 0 Then
        MsgBox "Coluna 'Vel. Média KM/H' não encontrada!", vbExclamation
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, velocidadeMediaCol).End(xlUp).Row

    Dim velocidadeMediaRange As Range
    Set velocidadeMediaRange = ws.Range(ws.Cells(2, velocidadeMediaCol), ws.Cells(lastRow, velocidadeMediaCol))

    ' Adiciona cabeçalho "Decile Velocidade Média" na coluna adjacente à direita
    Dim decileCol As Long
    decileCol = velocidadeMediaCol + 1
    ws.Cells(1, decileCol).Value = "Decile Velocidade Média"

    ' Calcula os valores dos deciles
    Dim deciles(1 To 10) As Double
    For i = 1 To 10
        deciles(i) = Application.WorksheetFunction.Percentile_Inc(velocidadeMediaRange, i / 10)
    Next i

    ' Classificar cada entrada de velocidade média em seu respectivo decile
    Dim cell As Range
    For Each cell In velocidadeMediaRange
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

    MsgBox "Deciles para Velocidade Média calculados com sucesso!"

End Sub
