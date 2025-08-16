Attribute VB_Name = "Módulo4"
Sub AjustarEnvio()
    ' Desativa a atualização de tela para melhorar o desempenho
    Application.ScreenUpdating = False

    ' Declaração de variáveis
    Dim wsCriacao As Worksheet
    Dim wsPortal As Worksheet
    Dim lastRowCriacao As Long
    Dim i As Long
    
    Dim textoColunaI As String
    Dim textoBusca As String
    Dim posInicioNum As Long
    Dim numeroExtraido As String

    ' Define as abas "Criação" e "Planilha Portal"
    On Error Resume Next
    Set wsCriacao = ThisWorkbook.Sheets("Criação")
    Set wsPortal = ThisWorkbook.Sheets("Planilha Portal")
    On Error GoTo 0

    ' Verifica se as planilhas existem para evitar erros
    If wsCriacao Is Nothing Or wsPortal Is Nothing Then
        MsgBox "Uma das planilhas necessárias ('Criação' ou 'Planilha Portal') não foi encontrada. Verifique os nomes e tente novamente.", vbCritical
        Exit Sub
    End If
    
    ' Encontra a última linha com dados na coluna I da aba "Criação"
    lastRowCriacao = wsCriacao.Cells(wsCriacao.rows.Count, "I").End(xlUp).Row
    
    ' Limpa os dados antigos a partir da SEGUNDA linha para preservar o cabeçalho
    If wsPortal.rows.Count > 1 Then
        wsPortal.Range("R2:U" & wsPortal.rows.Count).ClearContents
    End If

    ' Define o texto a ser procurado para extrair o número
    textoBusca = "Dev. NF Cliente Parc "

    ' =========================================================================
    ' PARTE 1 E 2: Processamento integrado linha por linha
    ' O loop agora vai até a última linha da "Criação" e usa o mesmo índice 'i'
    ' para ler da "Criação" e escrever na "Planilha Portal", garantindo a ordem.
    ' =========================================================================
    For i = 2 To lastRowCriacao
        ' Pega o valor da célula na coluna I da "Criação"
        textoColunaI = Trim(wsCriacao.Cells(i, "I").Value)
        
        ' Se a célula de origem estiver vazia, a linha de destino permanecerá vazia.
        ' O código continua para a próxima iteração.
        If textoColunaI <> "" Then
            ' Procura a posição do texto "Dev. NF Cliente Parc "
            posInicioNum = InStr(1, textoColunaI, textoBusca, vbTextCompare)

            ' Se encontrou o texto específico...
            If posInicioNum > 0 Then
                numeroExtraido = Mid(textoColunaI, posInicioNum + Len(textoBusca), 8)
                
                ' Verifica se os caracteres extraídos são realmente um número
                If IsNumeric(numeroExtraido) Then
                    ' --- PREENCHE COLUNA R ---
                    ' Se for numérico, cola na coluna R da "Planilha Portal" na MESMA linha
                    wsPortal.Cells(i, "R").Value = CLng(numeroExtraido)
                    
                    ' --- LÓGICA DE T e U ---
                    ' Como R foi preenchido, aplicamos a regra para T e U
                    If InStr(1, wsPortal.Cells(i, "L").Value, "509") > 0 Then
                        ' Se contiver 509, coloca um "X" na coluna U
                        wsPortal.Cells(i, "U").Value = "X"
                    Else
                        ' Se não contiver 509, preenche a coluna T com a data
                        wsPortal.Cells(i, "T").Value = Application.WorksheetFunction.WorkDay(Date, 5)
                        wsPortal.Cells(i, "T").NumberFormat = "dd/mm/yyyy" ' Formata a data
                    End If
                Else
                    ' --- PREENCHE COLUNA S ---
                    ' Se o que foi extraído não for número, trata como texto normal e cola na coluna S
                    wsPortal.Cells(i, "S").Value = textoColunaI
                End If
            Else
                ' --- PREENCHE COLUNA S ---
                ' Se não encontrou o padrão, cola o texto inteiro na coluna S
                wsPortal.Cells(i, "S").Value = textoColunaI
            End If
        End If
        ' Se textoColunaI estava vazio, nada acontece nesta iteração,
        ' e a linha 'i' na Planilha Portal permanece em branco (R, S, T, U),
        ' que é o comportamento desejado.
    Next i

    ' Reativa a atualização de tela
    Application.ScreenUpdating = True

   Call SalvarAbaComoArquivo

End Sub
