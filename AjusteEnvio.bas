Attribute VB_Name = "M�dulo4"
Sub AjustarEnvio()
    ' Desativa a atualiza��o de tela para melhorar o desempenho
    Application.ScreenUpdating = False

    ' Declara��o de vari�veis
    Dim wsCriacao As Worksheet
    Dim wsPortal As Worksheet
    Dim lastRowCriacao As Long
    Dim i As Long
    
    Dim textoColunaI As String
    Dim textoBusca As String
    Dim posInicioNum As Long
    Dim numeroExtraido As String

    ' Define as abas "Cria��o" e "Planilha Portal"
    On Error Resume Next
    Set wsCriacao = ThisWorkbook.Sheets("Cria��o")
    Set wsPortal = ThisWorkbook.Sheets("Planilha Portal")
    On Error GoTo 0

    ' Verifica se as planilhas existem para evitar erros
    If wsCriacao Is Nothing Or wsPortal Is Nothing Then
        MsgBox "Uma das planilhas necess�rias ('Cria��o' ou 'Planilha Portal') n�o foi encontrada. Verifique os nomes e tente novamente.", vbCritical
        Exit Sub
    End If
    
    ' Encontra a �ltima linha com dados na coluna I da aba "Cria��o"
    lastRowCriacao = wsCriacao.Cells(wsCriacao.rows.Count, "I").End(xlUp).Row
    
    ' Limpa os dados antigos a partir da SEGUNDA linha para preservar o cabe�alho
    If wsPortal.rows.Count > 1 Then
        wsPortal.Range("R2:U" & wsPortal.rows.Count).ClearContents
    End If

    ' Define o texto a ser procurado para extrair o n�mero
    textoBusca = "Dev. NF Cliente Parc "

    ' =========================================================================
    ' PARTE 1 E 2: Processamento integrado linha por linha
    ' O loop agora vai at� a �ltima linha da "Cria��o" e usa o mesmo �ndice 'i'
    ' para ler da "Cria��o" e escrever na "Planilha Portal", garantindo a ordem.
    ' =========================================================================
    For i = 2 To lastRowCriacao
        ' Pega o valor da c�lula na coluna I da "Cria��o"
        textoColunaI = Trim(wsCriacao.Cells(i, "I").Value)
        
        ' Se a c�lula de origem estiver vazia, a linha de destino permanecer� vazia.
        ' O c�digo continua para a pr�xima itera��o.
        If textoColunaI <> "" Then
            ' Procura a posi��o do texto "Dev. NF Cliente Parc "
            posInicioNum = InStr(1, textoColunaI, textoBusca, vbTextCompare)

            ' Se encontrou o texto espec�fico...
            If posInicioNum > 0 Then
                numeroExtraido = Mid(textoColunaI, posInicioNum + Len(textoBusca), 8)
                
                ' Verifica se os caracteres extra�dos s�o realmente um n�mero
                If IsNumeric(numeroExtraido) Then
                    ' --- PREENCHE COLUNA R ---
                    ' Se for num�rico, cola na coluna R da "Planilha Portal" na MESMA linha
                    wsPortal.Cells(i, "R").Value = CLng(numeroExtraido)
                    
                    ' --- L�GICA DE T e U ---
                    ' Como R foi preenchido, aplicamos a regra para T e U
                    If InStr(1, wsPortal.Cells(i, "L").Value, "509") > 0 Then
                        ' Se contiver 509, coloca um "X" na coluna U
                        wsPortal.Cells(i, "U").Value = "X"
                    Else
                        ' Se n�o contiver 509, preenche a coluna T com a data
                        wsPortal.Cells(i, "T").Value = Application.WorksheetFunction.WorkDay(Date, 5)
                        wsPortal.Cells(i, "T").NumberFormat = "dd/mm/yyyy" ' Formata a data
                    End If
                Else
                    ' --- PREENCHE COLUNA S ---
                    ' Se o que foi extra�do n�o for n�mero, trata como texto normal e cola na coluna S
                    wsPortal.Cells(i, "S").Value = textoColunaI
                End If
            Else
                ' --- PREENCHE COLUNA S ---
                ' Se n�o encontrou o padr�o, cola o texto inteiro na coluna S
                wsPortal.Cells(i, "S").Value = textoColunaI
            End If
        End If
        ' Se textoColunaI estava vazio, nada acontece nesta itera��o,
        ' e a linha 'i' na Planilha Portal permanece em branco (R, S, T, U),
        ' que � o comportamento desejado.
    Next i

    ' Reativa a atualiza��o de tela
    Application.ScreenUpdating = True

   Call SalvarAbaComoArquivo

End Sub
