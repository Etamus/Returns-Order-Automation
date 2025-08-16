Attribute VB_Name = "Modulo 4"
Sub LimparBase()
    ' Desativa a atualização de tela para a macro rodar mais rápido
    Application.ScreenUpdating = False

    ' Chama a rotina de limpeza para cada aba especificada
    Call LimparAba("Planilha Portal")
    Call LimparAba("Criação")

    ' Reativa a atualização de tela
    Application.ScreenUpdating = True
    
    ' Exibe uma mensagem informando que o processo foi concluído
    MsgBox "Finalizado.", vbInformation, "Processo Finalizado"
End Sub

Private Sub LimparAba(nomeDaAba As String)
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim intervalo As Range

    ' Tenta acessar a aba pelo nome fornecido
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nomeDaAba)
    On Error GoTo 0

    ' Verifica se a aba existe antes de continuar
    If Not ws Is Nothing Then
        ' Encontra a última linha com dados na coluna A
        ultimaLinha = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
        
        ' Define o intervalo a ser limpo e formatado, a partir da linha 2
        ' Se houver dados, o intervalo vai da linha 2 até a última linha com conteúdo.
        ' Se a aba estiver vazia a partir da linha 2, nada acontece.
        If ultimaLinha >= 2 Then
            Set intervalo = ws.Range("A2:" & ws.Cells(ultimaLinha, ws.Columns.Count).Address)
            
            ' 1. Limpa todo o conteúdo do intervalo
            intervalo.ClearContents
            
            ' 2. Pinta o interior de todas as células de branco
            With intervalo.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = vbWhite
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            
            ' 3. Aplica bordas brancas em todas as células do intervalo
            With intervalo.Borders
                .LineStyle = xlContinuous
                .Color = vbWhite
                .Weight = xlThin
            End With
        End If
    Else
        ' Alerta o usuário se uma das abas não for encontrada
        MsgBox "A aba '" & nomeDaAba & "' não foi encontrada. Verifique o nome e tente novamente.", vbExclamation, "Aba Não Encontrada"
    End If
End Sub
