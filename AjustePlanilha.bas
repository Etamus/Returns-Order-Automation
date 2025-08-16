Attribute VB_Name = "M�dulo1"
Sub AjustarPlanilha()

    '======================================================================================================================
    ' PARTE 0: CONFIGURA��ES INICIAIS
    ' Declara��o de todas as vari�veis que ser�o utilizadas na macro.
    '======================================================================================================================
    Dim wsPortal As Worksheet, wsCriacao As Worksheet, wsCliente As Worksheet
    Dim lastRowPortal As Long, lastRowCriacao As Long, lastRowCliente As Long
    Dim i As Long
    
    ' Dicion�rios s�o objetos que armazenam pares de chave-valor, muito eficientes para buscas.
    Dim dictClientes As Object
    Set dictClientes = CreateObject("Scripting.Dictionary")
    
    Dim dictDuplicados As Object
    Set dictDuplicados = CreateObject("Scripting.Dictionary")
    
    ' Vari�veis para a l�gica de duplicados
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

    ' Define as planilhas para evitar erros e facilitar a leitura do c�digo
    On Error GoTo TratamentoErro
    Set wsPortal = ThisWorkbook.Worksheets("Planilha Portal")
    Set wsCriacao = ThisWorkbook.Worksheets("Cria��o")
    Set wsCliente = ThisWorkbook.Worksheets("Cliente")
    
    ' Desativa a atualiza��o de tela para a macro rodar muito mais r�pido
    Application.ScreenUpdating = False
    
    '======================================================================================================================
    ' PARTE 1: TRANSFER�NCIA E FORMATA��O DOS DADOS
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
    ' Procura o c�digo do cliente ou aplica a regra especial do "509".
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
            ' ---> IN�CIO DA ALTERA��O <---
            ' Verifica se a coluna G cont�m "509". A fun��o InStr procura um texto dentro de outro.
            If InStr(wsCriacao.Cells(i, "G").Value, "509") > 0 Then
                ' Se encontrou, aplica o c�digo especial
                wsCriacao.Cells(i, "J").Value = "5002359"
            Else
                ' Se n�o encontrou, executa a busca normal
                Dim chaveCliente As String
                chaveCliente = Trim(wsCriacao.Cells(i, "B").Value)
                
                If dictClientes.Exists(chaveCliente) Then
                    wsCriacao.Cells(i, "J").Value = dictClientes(chaveCliente)
                Else
                    wsCriacao.Cells(i, "J").Value = "Sem Cadastro"
                End If
            End If
            ' ---> FIM DA ALTERA��O <---
            
            wsCriacao.Cells(i, "J").HorizontalAlignment = xlCenter
        Next i
    End If

    '======================================================================================================================
    ' PARTE 3 E 4: VERIFICA��O E PROCESSAMENTO DE DUPLICADOS (L�GICA OTIMIZADA)
    '======================================================================================================================
    
    lastRowCriacao = wsCriacao.Cells(wsCriacao.rows.Count, "A").End(xlUp).Row
    
    If lastRowCriacao > 1 Then
        ' 1. Cria uma "chave" para cada linha e agrupa os n�meros das linhas duplicadas
        For i = 2 To lastRowCriacao
            Dim strChave As String ' Vari�vel tempor�ria para construir a chave
            strChave = wsCriacao.Cells(i, "A").Value & "|" & wsCriacao.Cells(i, "B").Value & "|" & wsCriacao.Cells(i, "C").Value & "|" & wsCriacao.Cells(i, "F").Value & "|" & wsCriacao.Cells(i, "G").Value & "|" & wsCriacao.Cells(i, "J").Value
            
            If Not dictDuplicados.Exists(strChave) Then
                Set grupoDuplicado = New Collection
                grupoDuplicado.Add i
                dictDuplicados.Add strChave, grupoDuplicado
            Else
                dictDuplicados(strChave).Add i
            End If
        Next i
        
        ' 2. Agora, percorre apenas os grupos que t�m duplicados
        For Each chaveConcatenada In dictDuplicados.Keys
            Set grupoDuplicado = dictDuplicados(chaveConcatenada)
            
            If grupoDuplicado.Count > 1 Then ' Se for maior que 1, � um grupo de duplicados
                
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
                
                ' Processa os subgrupos que tamb�m s�o duplicados (baseado na Coluna D)
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
    ' Libera a mem�ria dos objetos
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
