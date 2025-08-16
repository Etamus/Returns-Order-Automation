Attribute VB_Name = "Modulo 6"
Sub SalvarAbaComoArquivo()
    ' Desativa atualizações de tela para performance.
    Application.ScreenUpdating = False

    Dim abaOrigem As Worksheet
    Dim novoWorkbook As Workbook
    Dim novaAba As Worksheet
    Dim caminhoDesktop As String
    Dim nomePasta As String
    Dim caminhoCompleto As String
    Dim nomeArquivo As String
    Dim ultimaLinha As Long
    Dim dadosRange As Range
    Dim cabecalhoRange As Range

    ' --- CONFIGURAÇÃO INICIAL ---
    ' Define a aba de onde os dados serão copiados.
    On Error GoTo ErroPlanilhaInexistente
    Set abaOrigem = ThisWorkbook.Sheets("Planilha Portal")
    On Error GoTo 0 ' Restaura o tratamento de erro padrão

    ' Encontra a última linha com dados na coluna A.
    ultimaLinha = abaOrigem.Cells(abaOrigem.rows.Count, "A").End(xlUp).Row
    
    ' Se não houver dados (além do cabeçalho), encerra a macro.
    If ultimaLinha <= 1 Then
        MsgBox "Não há dados para exportar na aba 'Planilha Portal'.", vbExclamation, "Aviso"
        Exit Sub
    End If

    ' Define o intervalo de dados e o intervalo do cabeçalho.
    Set dadosRange = abaOrigem.Range("A1:U" & ultimaLinha)
    Set cabecalhoRange = abaOrigem.Range("A1:U1")

    ' --- CRIAÇÃO DO NOVO ARQUIVO ---
    ' Cria uma nova pasta de trabalho.
    Set novoWorkbook = Workbooks.Add
    Set novaAba = novoWorkbook.Sheets(1)

    ' --- CÓPIA E COLAGEM DOS DADOS ---
    ' 1. Copia os dados da origem.
    dadosRange.Copy
    ' 2. Cola como valores (sem formatação).
    novaAba.Range("A1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    ' --- APLICAÇÃO DAS FORMATAÇÕES ---
    ' 1. Copia o cabeçalho da origem.
    cabecalhoRange.Copy
    ' 2. Cola apenas a formatação no novo cabeçalho.
    novaAba.Range("A1").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    ' 3. APLICA A FORMATAÇÃO DE DATA NA COLUNA T
    ' Formata a coluna inteira, exceto o cabeçalho, como data.
    novaAba.Range("T2:T" & ultimaLinha).NumberFormat = "dd/mm/yyyy"

    ' --- SALVANDO O ARQUIVO ---
    ' Define o caminho completo diretamente.
    caminhoCompleto = "C:\Users\lopesm21\Downloads\Macros\Novas Macros\Histórico de Criação"

    ' Cria a pasta se ela não existir.
    If Dir(caminhoCompleto, vbDirectory) = "" Then
       MkDir caminhoCompleto
    End If

    ' Define o nome do arquivo com a data atual.
    nomeArquivo = "Retorno " & Format(Date, "dd.mm.yyyy") & ".xlsx"

    ' Salva e fecha o novo arquivo.
    On Error GoTo ErroAoSalvar
    With novoWorkbook
        ' Desativa alertas (como "sobrescrever arquivo?") para salvar automaticamente.
        Application.DisplayAlerts = False
        .SaveAs Filename:=caminhoCompleto & "\" & nomeArquivo, FileFormat:=xlOpenXMLWorkbook
        Application.DisplayAlerts = True
        .Close SaveChanges:=False
    End With
    On Error GoTo 0

    ' Reativa as atualizações de tela.
    Application.ScreenUpdating = True

    ' --- MENSAGEM DE CONCLUSÃO ---
    MsgBox "Arquivo '" & nomeArquivo & "' foi criado.", vbInformation, "Exportação Concluída"
    Exit Sub

' --- SEÇÕES DE TRATAMENTO DE ERRO ---
ErroPlanilhaInexistente:
    MsgBox "Erro: A aba 'Planilha Portal' não foi encontrada." & vbCrLf & _
           "Por favor, verifique o nome da aba e tente novamente.", _
           vbCritical, "Erro na Execução"
    Application.ScreenUpdating = True
    Exit Sub

ErroAoSalvar:
    MsgBox "Erro ao salvar o arquivo." & vbCrLf & _
           "Verifique se o arquivo '" & nomeArquivo & "' já está aberto.", _
           vbCritical, "Erro ao Salvar"
    ' Garante que os alertas e a atualização de tela sejam reativados.
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

End Sub

