Attribute VB_Name = "Modulo 2"
Sub AjustarPlanilha()

    '======================================================================================================================
    ' PARTE 0: CONFIGURAÇÕES INICIAIS
    ' Declaração de todas as variáveis que serão utilizadas na macro.
    '======================================================================================================================
    Dim wsPortal As Worksheet, wsCriacao As Worksheet, wsCliente As Worksheet
    Dim lastRowPortal As Long, lastRowCriacao As Long, lastRowCliente As Long
    Dim i As Long
    
    ' Dicionários são objetos que armazenam pares de chave-valor, muito eficientes para buscas.
    Dim dictClientes As Object
    Set dictClientes = CreateObject("Scripting.Dictionary")
    
    Dim dictDuplicados As Object
    Set dictDuplicados = CreateObject("Scripting.Dictionary")
    
    ' Variáveis para a lógica de duplicados
    Dim chaveConcatenada As Variant
    Dim grupoDuplicado As Object
    Dim chaveSubgrupo As Variant
    Dim dictSubgrupos As Object
    Dim linha As Variant
    Dim somaH As Double
    Dim primeiraLinha As Boolean
    
    ' Cor verde para o preenchimento e bordas
    Dim corVerde As Long
    corVerde = RGB(169, 208, 142) 'Um tom de verde claro

    ' Define as planilhas para evitar erros e facilitar a leitura do código
    On Error GoTo TratamentoErro
    Set wsPortal = ThisWorkbook.Worksheets("Planilha Portal")
    Set wsCriacao = ThisWorkbook.Worksheets("Criação")
    Set wsCliente = ThisWorkbook.Worksheets("Cliente")
    
    ' Desativa a atualização de tela para a macro rodar muito mais rápido
    Application.ScreenUpdating = False
    
    '======================================================================================================================
    ' PARTE 1: TRANSFERÊNCIA E FORMATAÇÃO DOS DADOS
    '======================================================================================================================
    
    lastRowCriacao = wsCriacao.Cells(wsCriacao.rows.Count, "A").End(xlUp).Row
    If lastRowCriacao > 1 Then
        wsCriacao.Range("A2:K" & lastRowCriacao).ClearContents
        wsCriacao.Range("A2:K" & lastRowCriacao).Interior.Color = xlNone
        wsCriacao.Range("A2:K" & lastRowCriacao).Borders.LineStyle = xlNone
    End If

    lastRowPortal = wsPortal.Cells(wsPortal.rows.Count, "D").End(xlUp).Row
    
    If lastRowPortal > 1 Then
        For i = 2 To lastRowPortal
            With wsCriacao.Cells(i, "A")
                .Value = wsPortal.Cells(i, "D").Value
                .NumberFormat = "0"
            End With
            wsCriacao.Cells(i, "B").Value = wsPortal.Cells(i, "A").Value
            wsCriacao.Cells(i, "C").Value = wsPortal.Cells(i, "C").Value
            wsCriacao.Cells(i, "D").Value = wsPortal.Cells(i, "I").Value
            wsCriacao.Cells(i, "E").Value = wsPortal.Cells(i, "J").Value
            
            With wsCriacao.Cells(i, "F")
                .Value = wsPortal.Cells(i, "G").Value
                .NumberFormat = "000000000-1"
            End With

            wsCriacao.Cells(i, "G").Value = wsPortal.Cells(i, "L").Value
            wsCriacao.Cells(i, "H").Value = wsPortal.Cells(i, "E").Value
            
            wsCriacao.Range("A" & i & ":H" & i).HorizontalAlignment = xlCenter
        Next i
    End If
    
    '======================================================================================================================
    ' PARTE 2: BUSCA DE CLIENTES (COM NOVA REGRA)
    ' Procura o código do cliente ou aplica a regra especial do "509".
    '======================================================================================================================
    
    lastRowCliente = wsCliente.Cells(wsCliente.rows.Count, "A").End(xlUp).Row
    
    If lastRowCliente > 1 Then
        For i = 2 To lastRowCliente
            If Not dictClientes.Exists(Trim(wsCliente.Cells(i, "A").Value)) Then
                dictClientes.Add Trim(wsCliente.Cells(i, "A").Value), wsCliente.Cells(i, "C").Value
            End If
        Next i
    End If
    
    lastRowCriacao = wsCriacao.Cells(wsCriacao.rows.Count, "A").End(xlUp).Row
    
    If lastRowCriacao > 1 Then
        For i = 2 To lastRowCriacao
            ' ---> INÍCIO DA ALTERAÇÃO <---
            ' Verifica se a coluna G contém "509". A função InStr procura um texto dentro de outro.
            If InStr(wsCriacao.Cells(i, "G").Value, "509") > 0 Then
                ' Se encontrou, aplica o código especial
                wsCriacao.Cells(i, "J").Value = "5002359"
            Else
                ' Se não encontrou, executa a busca normal
                Dim chaveCliente As String
                chaveCliente = Trim(wsCriacao.Cells(i, "B").Value)
                
                If dictClientes.Exists(chaveCliente) Then
                    wsCriacao.Cells(i, "J").Value = dictClientes(chaveCliente)
                Else
                    wsCriacao.Cells(i, "J").Value = "Sem Cadastro"
                End If
            End If
            ' ---> FIM DA ALTERAÇÃO <---
            
            wsCriacao.Cells(i, "J").HorizontalAlignment = xlCenter
        Next i
    End If

    '======================================================================================================================
    ' PARTE 3 E 4: VERIFICAÇÃO E PROCESSAMENTO DE DUPLICADOS (LÓGICA OTIMIZADA)
    '======================================================================================================================
    
    lastRowCriacao = wsCriacao.Cells(wsCriacao.rows.Count, "A").End(xlUp).Row
    
    If lastRowCriacao > 1 Then
        ' 1. Cria uma "chave" para cada linha e agrupa os números das linhas duplicadas
        For i = 2 To lastRowCriacao
            Dim strChave As String ' Variável temporária para construir a chave
            strChave = wsCriacao.Cells(i, "A").Value & "|" & wsCriacao.Cells(i, "B").Value & "|" & wsCriacao.Cells(i, "C").Value & "|" & wsCriacao.Cells(i, "F").Value & "|" & wsCriacao.Cells(i, "G").Value & "|" & wsCriacao.Cells(i, "J").Value
            
            If Not dictDuplicados.Exists(strChave) Then
                Set grupoDuplicado = New Collection
                grupoDuplicado.Add i
                dictDuplicados.Add strChave, grupoDuplicado
            Else
                dictDuplicados(strChave).Add i
            End If
        Next i
        
        ' 2. Agora, percorre apenas os grupos que têm duplicados
        For Each chaveConcatenada In dictDuplicados.Keys
            Set grupoDuplicado = dictDuplicados(chaveConcatenada)
            
            If grupoDuplicado.Count > 1 Then ' Se for maior que 1, é um grupo de duplicados
                
                ' PARTE 3: Pinta e marca TODAS as linhas do grupo como "Duplicado"
                For Each linha In grupoDuplicado
                    With wsCriacao.Range("A" & linha & ":K" & linha)
                        .Interior.Color = corVerde
                        .Borders.Color = corVerde
                        .Borders.LineStyle = xlContinuous
                        .Borders.Weight = xlThin
                    End With
                    wsCriacao.Cells(linha, "K").Value = "Duplicado"
                    wsCriacao.Cells(linha, "K").HorizontalAlignment = xlCenter
                Next linha
                
                ' PARTE 4: Cria subgrupos pela Coluna D e processa a soma
                Set dictSubgrupos = CreateObject("Scripting.Dictionary")
                For Each linha In grupoDuplicado
                    chaveSubgrupo = wsCriacao.Cells(linha, "D").Value
                    If Not dictSubgrupos.Exists(chaveSubgrupo) Then
                        Set dictSubgrupos(chaveSubgrupo) = New Collection
                    End If
                    dictSubgrupos(chaveSubgrupo).Add linha
                Next linha
                
                ' Processa os subgrupos que também são duplicados (baseado na Coluna D)
                For Each chaveSubgrupo In dictSubgrupos.Keys
                    If dictSubgrupos(chaveSubgrupo).Count > 1 Then
                        somaH = 0
                        For Each linha In dictSubgrupos(chaveSubgrupo)
                            somaH = somaH + wsCriacao.Cells(linha, "H").Value
                        Next linha
                        
                        primeiraLinha = True
                        For Each linha In dictSubgrupos(chaveSubgrupo)
                            If primeiraLinha Then
                                wsCriacao.Cells(linha, "H").Value = somaH
                                primeiraLinha = False
                            Else
                                wsCriacao.Cells(linha, "H").Value = 0
                                wsCriacao.Cells(linha, "I").Value = "X"
                                wsCriacao.Cells(linha, "I").HorizontalAlignment = xlCenter
                            End If
                        Next linha
                    End If
                Next chaveSubgrupo
            End If
        Next chaveConcatenada
    End If

Conclusao:
    ' Libera a memória dos objetos
    Set wsPortal = Nothing
    Set wsCriacao = Nothing
    Set wsCliente = Nothing
    Set dictClientes = Nothing
    Set dictDuplicados = Nothing
    Set dictSubgrupos = Nothing
    
    Application.ScreenUpdating = True
    MsgBox "Finalizado.", vbInformation, "Finalizado"
    Exit Sub

TratamentoErro:
    Application.ScreenUpdating = True
    MsgBox "Ocorreu um erro: " & vbCrLf & Err.Description, vbCritical, "Erro na Macro"
    
End Sub
