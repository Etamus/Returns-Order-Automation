Attribute VB_Name = "M�dulo6"
Sub MoverparaRFQ()

    ' Desativa a atualiza��o da tela para a macro rodar mais r�pido
    Application.ScreenUpdating = False

    ' Declara��o das vari�veis das planilhas
    Dim wsCriacao As Worksheet
    Dim wsRFQ As Worksheet
    Dim wsPortal As Worksheet

    ' Declara��o de vari�veis de controle
    Dim ultLinhaCriacao As Long
    Dim ultLinhaPortal As Long
    Dim linhaDestinoInicio As Long
    Dim ultLinhaFinal As Long
    Dim i As Long
    Dim rangeFinal As Range

    ' Define as planilhas para facilitar a refer�ncia
    Set wsCriacao = ThisWorkbook.Sheets("Cria��o")
    Set wsRFQ = ThisWorkbook.Sheets("RFQ")
    Set wsPortal = ThisWorkbook.Sheets("Planilha Portal")

    ' *** NOVO: Encontra a primeira linha em branco na aba RFQ para iniciar a inser��o ***
    linhaDestinoInicio = wsRFQ.Cells(wsRFQ.rows.Count, "A").End(xlUp).Row + 1
    ' Garante que, se a planilha estiver vazia, comece na linha 2 (caso tenha cabe�alho na linha 1)
    If linhaDestinoInicio < 2 Then linhaDestinoInicio = 2

    ' Encontra a �ltima linha com dados na coluna A da aba "Cria��o"
    ultLinhaCriacao = wsCriacao.Cells(wsCriacao.rows.Count, "A").End(xlUp).Row
    
    ' Sai da macro se n�o houver dados novos para adicionar na aba "Cria��o"
    If ultLinhaCriacao < 2 Then
        MsgBox "N�o h� dados novos na aba 'Cria��o' para serem processados.", vbInformation
        Exit Sub
    End If

    ' Copia o intervalo de A2 at� K da aba "Cria��o" e cola apenas os VALORES na pr�xima linha livre da "RFQ"
    wsCriacao.Range("A2:K" & ultLinhaCriacao).Copy
    wsRFQ.Range("A" & linhaDestinoInicio).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False ' Limpa a �rea de transfer�ncia

    ' Encontra a �ltima linha com dados na coluna R da aba "Planilha Portal"
    ultLinhaPortal = wsPortal.Cells(wsPortal.rows.Count, "R").End(xlUp).Row

    ' Copia os dados da coluna R da "Planilha Portal" e cola apenas os VALORES na coluna I da "RFQ"
    wsPortal.Range("R2:R" & ultLinhaPortal).Copy
    wsRFQ.Range("I" & linhaDestinoInicio).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False ' Limpa a �rea de transfer�ncia

    ' Define a �ltima linha do bloco rec�m-adicionado
    ultLinhaFinal = wsRFQ.Cells(wsRFQ.rows.Count, "A").End(xlUp).Row

    ' Percorre APENAS O BLOCO NOVO de baixo para cima para remover as linhas com coluna I vazia
    For i = ultLinhaFinal To linhaDestinoInicio Step -1
        If IsEmpty(wsRFQ.Cells(i, "I").Value) Or wsRFQ.Cells(i, "I").Value = "" Then
            wsRFQ.rows(i).Delete
        End If
    Next i

    ' Encontra a �ltima linha novamente ap�s a poss�vel exclus�o de linhas
    ultLinhaFinal = wsRFQ.Cells(wsRFQ.rows.Count, "A").End(xlUp).Row

    ' Aplica formata��o e data SE algum dado novo foi efetivamente adicionado
    If ultLinhaFinal >= linhaDestinoInicio Then
        
        ' Define o intervalo final que cont�m apenas os novos dados v�lidos
        Set rangeFinal = wsRFQ.Range("A" & linhaDestinoInicio & ":L" & ultLinhaFinal)

        ' Preenche a data atual na coluna L para todas as linhas novas
        wsRFQ.Range("L" & linhaDestinoInicio & ":L" & ultLinhaFinal).Value = Date

        ' *** NOVO: Aplica formata��o (centraliza��o e BORDAS BRANCAS) ***
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

    ' Reativa a atualiza��o da tela
    Application.ScreenUpdating = True

    ' Exibe uma mensagem de conclus�o
    MsgBox "Finalizado.", vbInformation

    Call LimparBase

End Sub
