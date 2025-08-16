Attribute VB_Name = "Modulo 5"
Sub MoverparaRFQ()

    ' Desativa a atualização da tela para a macro rodar mais rápido
    Application.ScreenUpdating = False

    ' Declaração das variáveis das planilhas
    Dim wsCriacao As Worksheet
    Dim wsRFQ As Worksheet
    Dim wsPortal As Worksheet

    ' Declaração de variáveis de controle
    Dim ultLinhaCriacao As Long
    Dim ultLinhaPortal As Long
    Dim linhaDestinoInicio As Long
    Dim ultLinhaFinal As Long
    Dim i As Long
    Dim rangeFinal As Range

    ' Define as planilhas para facilitar a referência
    Set wsCriacao = ThisWorkbook.Sheets("Criação")
    Set wsRFQ = ThisWorkbook.Sheets("RFQ")
    Set wsPortal = ThisWorkbook.Sheets("Planilha Portal")

    ' *** NOVO: Encontra a primeira linha em branco na aba RFQ para iniciar a inserção ***
    linhaDestinoInicio = wsRFQ.Cells(wsRFQ.rows.Count, "A").End(xlUp).Row + 1
    ' Garante que, se a planilha estiver vazia, comece na linha 2 (caso tenha cabeçalho na linha 1)
    If linhaDestinoInicio < 2 Then linhaDestinoInicio = 2

    ' Encontra a última linha com dados na coluna A da aba "Criação"
    ultLinhaCriacao = wsCriacao.Cells(wsCriacao.rows.Count, "A").End(xlUp).Row
    
    ' Sai da macro se não houver dados novos para adicionar na aba "Criação"
    If ultLinhaCriacao < 2 Then
        MsgBox "Não há dados novos na aba 'Criação' para serem processados.", vbInformation
        Exit Sub
    End If

    ' Copia o intervalo de A2 até K da aba "Criação" e cola apenas os VALORES na próxima linha livre da "RFQ"
    wsCriacao.Range("A2:K" & ultLinhaCriacao).Copy
    wsRFQ.Range("A" & linhaDestinoInicio).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False ' Limpa a área de transferência

    ' Encontra a última linha com dados na coluna R da aba "Planilha Portal"
    ultLinhaPortal = wsPortal.Cells(wsPortal.rows.Count, "R").End(xlUp).Row

    ' Copia os dados da coluna R da "Planilha Portal" e cola apenas os VALORES na coluna I da "RFQ"
    wsPortal.Range("R2:R" & ultLinhaPortal).Copy
    wsRFQ.Range("I" & linhaDestinoInicio).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False ' Limpa a área de transferência

    ' Define a última linha do bloco recém-adicionado
    ultLinhaFinal = wsRFQ.Cells(wsRFQ.rows.Count, "A").End(xlUp).Row

    ' Percorre APENAS O BLOCO NOVO de baixo para cima para remover as linhas com coluna I vazia
    For i = ultLinhaFinal To linhaDestinoInicio Step -1
        If IsEmpty(wsRFQ.Cells(i, "I").Value) Or wsRFQ.Cells(i, "I").Value = "" Then
            wsRFQ.rows(i).Delete
        End If
    Next i

    ' Encontra a última linha novamente após a possível exclusão de linhas
    ultLinhaFinal = wsRFQ.Cells(wsRFQ.rows.Count, "A").End(xlUp).Row

    ' Aplica formatação e data SE algum dado novo foi efetivamente adicionado
    If ultLinhaFinal >= linhaDestinoInicio Then
        
        ' Define o intervalo final que contém apenas os novos dados válidos
        Set rangeFinal = wsRFQ.Range("A" & linhaDestinoInicio & ":L" & ultLinhaFinal)

        ' Preenche a data atual na coluna L para todas as linhas novas
        wsRFQ.Range("L" & linhaDestinoInicio & ":L" & ultLinhaFinal).Value = Date

        ' *** NOVO: Aplica formatação (centralização e BORDAS BRANCAS) ***
        With rangeFinal
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            With .Borders
                .LineStyle = xlContinuous
                .Color = vbWhite ' Define a cor da borda como branca
                .Weight = xlThin
            End With
        End With
    End If

    ' Reativa a atualização da tela
    Application.ScreenUpdating = True

    ' Exibe uma mensagem de conclusão
    MsgBox "Finalizado.", vbInformation

    Call LimparBase

End Sub
