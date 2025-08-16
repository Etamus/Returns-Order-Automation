Attribute VB_Name = "M�dulo2"
Sub LimparBase()
    ' Desativa a atualiza��o de tela para a macro rodar mais r�pido
    Application.ScreenUpdating = False

    ' Chama a rotina de limpeza para cada aba especificada
    Call LimparAba("Planilha Portal")
    Call LimparAba("Cria��o")

    ' Reativa a atualiza��o de tela
    Application.ScreenUpdating = True
    
    ' Exibe uma mensagem informando que o processo foi conclu�do
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
        ' Encontra a �ltima linha com dados na coluna A
        ultimaLinha = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
        
        ' Define o intervalo a ser limpo e formatado, a partir da linha 2
        ' Se houver dados, o intervalo vai da linha 2 at� a �ltima linha com conte�do.
        ' Se a aba estiver vazia a partir da linha 2, nada acontece.
        If ultimaLinha >= 2 Then
            Set intervalo = ws.Range("A2:" & ws.Cells(ultimaLinha, ws.Columns.Count).Address)
            
            ' 1. Limpa todo o conte�do do intervalo
            intervalo.ClearContents
            
            ' 2. Pinta o interior de todas as c�lulas de branco
            With intervalo.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = vbWhite
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            
            ' 3. Aplica bordas brancas em todas as c�lulas do intervalo
            With intervalo.Borders
                .LineStyle = xlContinuous
                .Color = vbWhite
                .Weight = xlThin
            End With
        End If
    Else
        ' Alerta o usu�rio se uma das abas n�o for encontrada
        MsgBox "A aba '" & nomeDaAba & "' n�o foi encontrada. Verifique o nome e tente novamente.", vbExclamation, "Aba N�o Encontrada"
    End If
End Sub
