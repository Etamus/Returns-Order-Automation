Attribute VB_Name = "M�dulo5"
Sub SalvarAbaComoArquivo()
    ' Desativa atualiza��es de tela para performance.
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

    ' --- CONFIGURA��O INICIAL ---
    ' Define a aba de onde os dados ser�o copiados.
    On Error GoTo ErroPlanilhaInexistente
    Set abaOrigem = ThisWorkbook.Sheets("Planilha Portal")
    On Error GoTo 0 ' Restaura o tratamento de erro padr�o

    ' Encontra a �ltima linha com dados na coluna A.
    ultimaLinha = abaOrigem.Cells(abaOrigem.rows.Count, "A").End(xlUp).Row
    
    ' Se n�o houver dados (al�m do cabe�alho), encerra a macro.
    If ultimaLinha <= 1 Then
        MsgBox "N�o h� dados para exportar na aba 'Planilha Portal'.", vbExclamation, "Aviso"
        Exit Sub
    End If

    ' Define o intervalo de dados e o intervalo do cabe�alho.
    Set dadosRange = abaOrigem.Range("A1:U" & ultimaLinha)
    Set cabecalhoRange = abaOrigem.Range("A1:U1")

    ' --- CRIA��O DO NOVO ARQUIVO ---
    ' Cria uma nova pasta de trabalho.
    Set novoWorkbook = Workbooks.Add
    Set novaAba = novoWorkbook.Sheets(1)

    ' --- C�PIA E COLAGEM DOS DADOS ---
    ' 1. Copia os dados da origem.
    dadosRange.Copy
    ' 2. Cola como valores (sem formata��o).
    novaAba.Range("A1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    ' --- APLICA��O DAS FORMATA��ES ---
    ' 1. Copia o cabe�alho da origem.
    cabecalhoRange.Copy
    ' 2. Cola apenas a formata��o no novo cabe�alho.
    novaAba.Range("A1").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    ' 3. APLICA A FORMATA��O DE DATA NA COLUNA T
    ' Formata a coluna inteira, exceto o cabe�alho, como data.
    novaAba.Range("T2:T" & ultimaLinha).NumberFormat = "dd/mm/yyyy"

    ' --- SALVANDO O ARQUIVO ---
    ' Define o caminho completo diretamente.
    caminhoCompleto = "C:\Users\lopesm21\Downloads\Macros\Novas Macros\Hist�rico de Cria��o"

    ' Cria a pasta se ela n�o existir.
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

    ' Reativa as atualiza��es de tela.
    Application.ScreenUpdating = True

    ' --- MENSAGEM DE CONCLUS�O ---
    MsgBox "Arquivo '" & nomeArquivo & "' foi criado.", vbInformation, "Exporta��o Conclu�da"
    Exit Sub

' --- SE��ES DE TRATAMENTO DE ERRO ---
ErroPlanilhaInexistente:
    MsgBox "Erro: A aba 'Planilha Portal' n�o foi encontrada." & vbCrLf & _
           "Por favor, verifique o nome da aba e tente novamente.", _
           vbCritical, "Erro na Execu��o"
    Application.ScreenUpdating = True
    Exit Sub

ErroAoSalvar:
    MsgBox "Erro ao salvar o arquivo." & vbCrLf & _
           "Verifique se o arquivo '" & nomeArquivo & "' j� est� aberto.", _
           vbCritical, "Erro ao Salvar"
    ' Garante que os alertas e a atualiza��o de tela sejam reativados.
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

End Sub

