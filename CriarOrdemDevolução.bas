Sub CriarOrdemDevolucao()

' Declaração de todas as variáveis utilizadas na macro.
Dim SapGuiAuto As Object
Dim app As Object
Dim Connection As Object
Dim session As Object

Dim dictInvoices As Object ' Dicionário para agrupar linhas por Nota Fiscal
Dim invoiceKey As Variant
Dim invoiceRows As Collection
Dim itemRow As Variant
Dim resultRow As Variant

Dim lastRow As Long
Dim i As Long
Dim firstRow As Long
Dim isFirstItem As Boolean

Dim strTexto As String
Dim tipo As String
Dim erroitem As String
Dim iteninex As String
Dim ois As String
Dim DOCNUM As String
Dim teste As String
Dim teste55 As String
Dim tipozdvp As String
Dim parceiro As String
Dim Msg36 As String
Dim mensagemerro As String
Dim ois1 As String
Dim Itemq As Double ' Alterado para Double
Dim Itemloop As String
Dim Item2 As String
Dim notaFiscalOriginal As String
Dim notaFiscalFormatada As String

' Variáveis para a nova lógica de busca e verificação
Dim motivoOriginalCriacao As String
Dim codigoParaBusca As String
Dim motivoFinalParaSAP As String
Dim wsCodigo As Worksheet
Dim lastRowCodigo As Long
Dim j As Long
Dim posAbreParenteses As Integer
Dim posFechaParenteses As Integer
Dim skipGroup As Boolean

' Ativa a aba "Criação" e define o cabeçalho para o nome do solicitante.
Sheets("Criação").Activate
Range("L1").Value = "NOME"

' Validação para garantir que o nome do solicitante foi preenchido.
If Range("L2").Value = "" Then
    MsgBox "É obrigatório preencher o NOME na célula L2!", vbCritical, "CONTROLE DE DADOS"
    Range("L2").Select
    Exit Sub
End If

Application.DisplayAlerts = False

' --- CONEXÃO COM O SAP (EXECUTADA APENAS UMA VEZ) ---
On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
If SapGuiAuto Is Nothing Then
    MsgBox "Não foi possível encontrar o SAP GUI. Verifique se ele está em execução.", vbCritical, "Erro de Conexão"
    Exit Sub
End If

Set app = SapGuiAuto.GetScriptingEngine
If app Is Nothing Then
    MsgBox "Não foi possível obter o Scripting Engine do SAP. Verifique as configurações de script do SAP.", vbCritical, "Erro de Conexão"
    Exit Sub
End If

Set Connection = app.Children(0)
If Connection Is Nothing Then
    MsgBox "Nenhuma conexão SAP encontrada. Verifique se você está logado em um sistema SAP.", vbCritical, "Erro de Conexão"
    Exit Sub
End If

Set session = Connection.Children(0)
If session Is Nothing Then
    MsgBox "Nenhuma sessão SAP encontrada. Verifique se há uma janela de sessão aberta.", vbCritical, "Erro de Conexão"
    Exit Sub
End If
On Error GoTo 0 ' Restaura o tratamento de erros padrão

If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject app, "on"
End If
' --- FIM DA CONEXÃO SAP ---


' Cria um objeto Dicionário para agrupar as ordens.
Set dictInvoices = CreateObject("Scripting.Dictionary")
lastRow = Sheets("Criação").Cells(Sheets("Criação").rows.Count, "A").End(xlUp).Row

' --- Agrupamento de Linhas por Nota Fiscal e Cliente ---
For i = 2 To lastRow
    Dim key As String
    Dim nf As String
    Dim cliente As String
    
    nf = Trim(Sheets("Criação").Cells(i, "F").Value)
    cliente = Trim(Sheets("Criação").Cells(i, "B").Value)

    If nf <> "" And cliente <> "" Then
        ' Agrupa apenas se a coluna K for "duplicado". Caso contrário, cria uma chave única.
        If UCase(Trim(Sheets("Criação").Cells(i, "K").Value)) = "DUPLICADO" Then
            key = nf & "|" & cliente ' Chave compartilhada para itens duplicados
        Else
            key = nf & "|" & cliente & "|" & i ' Chave única para itens individuais
        End If

        If Not dictInvoices.Exists(key) Then
            Set dictInvoices(key) = New Collection
        End If
        dictInvoices(key).Add i
    End If
Next i

' --- Loop Principal de Processamento ---
For Each invoiceKey In dictInvoices.Keys
    
    Set invoiceRows = dictInvoices(invoiceKey)
    firstRow = invoiceRows(1)

    ' Verifica se a primeira linha do grupo já foi processada
    If Trim(Sheets("Criação").Cells(firstRow, "I").Value) <> "" Then
        GoTo ProximoGrupo ' Pula para o próximo grupo de NF/Cliente
    End If

    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nVA01"
    session.findById("wnd[0]").sendVKey 0

    If InStr(1, Sheets("Criação").Cells(firstRow, "G").Value, "668") > 0 Then
        tipo = "ROB"
    Else
        tipo = "REB"
    End If

Tipo1:
    session.findById("wnd[0]/usr/ctxtVBAK-AUART").Text = tipo
    session.findById("wnd[0]/usr/ctxtVBAK-VKORG").Text = ""
    session.findById("wnd[0]/usr/ctxtVBAK-VTWEG").Text = ""
    session.findById("wnd[0]/usr/ctxtVBAK-SPART").Text = "90"
    session.findById("wnd[0]/usr/ctxtVBAK-VKBUR").Text = ""
    session.findById("wnd[0]/usr/ctxtVBAK-VKGRP").Text = ""
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[1]").sendVKey 4
    
    notaFiscalOriginal = Trim(Sheets("Criação").Cells(firstRow, "F").Value)
    notaFiscalFormatada = Format(notaFiscalOriginal, "000000000") & "-1"
    
    session.findById("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB005/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").Text = notaFiscalFormatada
    session.findById("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB005/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[4,24]").Text = Trim(Sheets("Criação").Cells(firstRow, "B").Value)
    session.findById("wnd[2]/tbar[0]/btn[0]").press

    On Error Resume Next
    strTexto = session.findById("wnd[0]/sbar").Text

    If strTexto = "Nenhum valor para esta seleção" Then
        session.findById("wnd[2]/tbar[0]/btn[12]").press
        session.findById("wnd[1]/tbar[0]/btn[12]").press
        
        For Each resultRow In invoiceRows
            Sheets("Criação").Cells(resultRow, "I").Value = "NF inexistente para o código de cliente informado"
        Next resultRow
        
        strTexto = ""
        GoTo ProximoGrupo
    End If
    On Error GoTo 0

    session.findById("wnd[0]").sendVKey 0
inicio10:
    On Error Resume Next
    session.findById("wnd[0]").sendVKey 5
    teste55 = Left(session.findById("wnd[2]/usr/txtMESSTXT1").Text, 11)
    If teste55 = "O doc.venda" Then
        Application.Wait Now + TimeValue("00:00:05")
        session.findById("wnd[0]").sendVKey 0
        teste55 = ""
        GoTo inicio10
    End If

    tipozdvp = session.findById("wnd[2]/usr/txtMESSTXT2").Text
    parceiro = session.findById("wnd[1]").Text

    If parceiro = "Seleção de parceiro" Then
        parceiro = ""
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End If

    If tipozdvp = "ZVPC para REB" Then
        tipo = "ZDVP"
        session.findById("wnd[2]/tbar[0]/btn[0]").press
        session.findById("wnd[1]/tbar[0]/btn[12]").press
        tipozdvp = ""
        GoTo Tipo1
    End If

    If tipozdvp = "ZLD2 para REB" Or tipozdvp = "ZDOA para REB" Then
        If tipozdvp = "ZLD2 para REB" Then
            mensagemerro = "Peça Livre de Débito"
        Else
            mensagemerro = "Tipo e NF não gera devolução - Doação"
        End If
        
        For Each resultRow In invoiceRows
            Sheets("Criação").Cells(resultRow, "I").Value = mensagemerro
        Next resultRow
        
        tipozdvp = ""
        GoTo ProximoGrupo
    End If

volta3:
    teste = session.findById("wnd[1]/usr/txtMESSTXT1").Text
        If Left(teste, 2) = "Já" Then
        teste = ""
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        GoTo volta3
    End If
    teste = ""
    
    Msg36 = session.findById("wnd[1]").Text
    If Left(Msg36, 5) = "Texto" Then
        Msg36 = ""
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End If
    On Error GoTo 0

    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBKD-BSTDK").SetFocus
    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBKD-BSTDK").caretPosition = 0
    session.findById("wnd[0]").sendVKey 4
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-LIFSK").key = " "
    
motivoOriginalCriacao = Trim(CStr(Sheets("Criação").Cells(firstRow, "G").Value))
posAbreParenteses = InStr(motivoOriginalCriacao, "(")
posFechaParenteses = InStr(motivoOriginalCriacao, ")")

If posAbreParenteses > 0 And posFechaParenteses > posAbreParenteses Then
    codigoParaBusca = Trim(Mid(motivoOriginalCriacao, posAbreParenteses + 1, posFechaParenteses - posAbreParenteses - 1))
Else
    codigoParaBusca = ""
End If

' --- Ajuste para tratar 90 como 090 e 92 como 092 ---
If codigoParaBusca = "90" Then
    codigoParaBusca = "090"
ElseIf codigoParaBusca = "92" Then
    codigoParaBusca = "092"
End If
' ----------------------------------------------------

motivoFinalParaSAP = ""
If codigoParaBusca <> "" Then
    Set wsCodigo = ThisWorkbook.Sheets("Código")
    lastRowCodigo = wsCodigo.Cells(wsCodigo.rows.Count, "A").End(xlUp).Row

    For j = 1 To lastRowCodigo
        If Trim(CStr(Left(wsCodigo.Cells(j, "A").Value, Len(codigoParaBusca)))) = codigoParaBusca Then
            motivoFinalParaSAP = Trim(CStr(Left(wsCodigo.Cells(j, "A").Value, 3)))
            Exit For
        End If
    Next j
End If


    If motivoFinalParaSAP <> "" Then
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-AUGRU").key = motivoFinalParaSAP
    Else
        Sheets("Criação").Cells(firstRow, "I").Value = "ERRO: Código '" & codigoParaBusca & "' não encontrado na aba 'Código'."
        GoTo ProximoGrupo
    End If
    
    If InStr(1, Sheets("Criação").Cells(firstRow, "G").Value, "668") > 0 Then
        GoTo OIROB
    End If
    
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_MKAL").press
    
    For Each itemRow In invoiceRows
        If Sheets("Criação").Cells(itemRow, "H").Value > 0 Then
            
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POPO").press
            
            Itemloop = Sheets("Criação").Cells(itemRow, "D").Value
            session.findById("wnd[1]/usr/ctxtRV45A-PO_MATNR").Text = Itemloop
            session.findById("wnd[1]/tbar[0]/btn[0]").press

            On Error Resume Next
            erroitem = session.findById("wnd[2]/usr/txtMESSTXT1").Text
            If erroitem = "Item  inexistente" Then
                iteninex = Sheets("Criação").Cells(itemRow, "D").Value
                erroitem = ""
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                
                For Each resultRow In invoiceRows
                    Sheets("Criação").Cells(resultRow, "I").Value = "Item " & iteninex & " inexistente para NF"
                Next resultRow
                
                iteninex = ""
                GoTo ProximoGrupo
            End If
            On Error GoTo 0
            erroitem = ""

            session.findById("wnd[0]").sendVKey 9

            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,0]").SetFocus
            Itemq = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,0]").Text

            If Itemq < Sheets("Criação").Cells(itemRow, "H").Value Then
                For Each resultRow In invoiceRows
                    Sheets("Criação").Cells(resultRow, "I").Value = "QUANTIDADE SOLICITADA (" & Sheets("Criação").Cells(itemRow, "H").Text & "pç) É MAIOR QUE FATURADA(" & Itemq & "pç) para o item " & Itemloop & ""
                Next resultRow
                GoTo ProximoGrupo
            End If

            If session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,0]").Text >= Sheets("Criação").Cells(itemRow, "H").Value Then
                session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,0]").Text = Sheets("Criação").Cells(itemRow, "H").Value
            End If
            
        End If
    Next itemRow

    '************************************************************************************************
    '*** INÍCIO DA ALTERAÇÃO - Reintroduzindo a verificação de segurança do código antigo ***
    '************************************************************************************************
    On Error Resume Next
    Item2 = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,1]").Text
    On Error GoTo 0
    
    If Item2 = "" Then
        GoTo pularDelete ' Se não houver um segundo item, pule a exclusão
    End If
    
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POLO").press
    On Error Resume Next
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    On Error GoTo 0
    
pularDelete:
    '************************************************************************************************
    '*** FIM DA ALTERAÇÃO ***
    '************************************************************************************************

OIROB:
    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04").Select

    If InStr(1, Sheets("Criação").Cells(firstRow, "G").Value, "668") = 0 Then
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").Text = "e1-1"
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").SetFocus
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").caretPosition = 3
    End If

    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07").Select
    
    Dim contsp As Integer
    contsp = 1
trs:
    If contsp > 7 Then
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,8]").key = "SP"
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,8]").Text = Sheets("Criação").Cells(firstRow, "J").Value
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,8]").SetFocus
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,8]").caretPosition = 7
        session.findById("wnd[0]").sendVKey 0
    Else
        If session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & contsp & "]").Text = "SP Transportador" Then
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1," & contsp & "]").Text = Sheets("Criação").Cells(firstRow, "J").Value
        Else
            contsp = contsp + 1
            GoTo trs
        End If
    End If

    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08").Select
    
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text = Sheets("Criação").Cells(firstRow, "A").Text & " - " & Sheets("Criação").Range("L2").Text
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes 11, 11

    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09").Select
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/ctxtVBKD-BSARK").Text = "ZLR1"

    If InStr(1, Sheets("Criação").Cells(firstRow, "G").Value, "509") > 0 Then
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBAK-SUBMI").Text = "01"
    End If

    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBKD-BSTKD").Text = Sheets("Criação").Cells(firstRow, "A").Text
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBKD-BSTKD_E").Text = Sheets("Criação").Cells(firstRow, "A").Text
    session.findById("wnd[0]/tbar[0]/btn[11]").press

    ois = ""
    ois1 = ""
    DOCNUM = ""

    If tipo = "ROB" Or tipo = "ZDVP" Then
        ois1 = session.findById("wnd[0]/sbar").Text
        If Left(ois1, 2) = "Já" Then
            GoTo voltaOI
        End If

        On Error Resume Next
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nvl01n"
        session.findById("wnd[0]").sendVKey 0

Inicio_VL01N:
        session.findById("wnd[0]/usr/ctxtLIKP-VSTEL").Text = "1350"
        ois = session.findById("wnd[0]/usr/ctxtLV50C-VBELN").Text
        session.findById("wnd[0]/usr/ctxtLV50C-ABPOS").Text = ""
        session.findById("wnd[0]/usr/ctxtLV50C-BIPOS").Text = ""
        session.findById("wnd[0]/usr/ctxtLIKP-LFART").Text = ""
        Application.Wait Now + TimeValue("00:00:01")
        session.findById("wnd[0]").sendVKey 0

        On Error Resume Next
        If Left(session.findById("wnd[0]/sbar").Text, 6) = "Ordem " Then
            GoTo Inicio_VL01N
        End If
        On Error GoTo 0

        session.findById("wnd[0]/tbar[0]/btn[11]").press
        DOCNUM = Left(Right(session.findById("wnd[0]/sbar").Text, 20), 9) * 1
        On Error GoTo 0

        If DOCNUM = "" Then
            GoTo ProximoGrupo
        End If
        
        For Each resultRow In invoiceRows
            Sheets("Criação").Cells(resultRow, "I").Value = "Dev. NF Cliente Parc " & ois & " foi gravado (foi gerado o fornecimento " & DOCNUM & ")"
        Next resultRow

    Else
voltaOI:
        For Each resultRow In invoiceRows
            Sheets("Criação").Cells(resultRow, "I").Value = session.findById("wnd[0]/sbar").Text
        Next resultRow
    End If
    
    ActiveWorkbook.Save
    session.findById("wnd[0]/tbar[0]/btn[3]").press

ProximoGrupo:
    On Error Resume Next
    On Error GoTo 0

Next invoiceKey

' --- NOVO: PÓS-PROCESSAMENTO PARA CORRIGIR LINHAS COM "X" ---
Dim nfToFind As String
Dim clientToFind As String
Dim masterResult As String

For i = 2 To lastRow
    ' Verifica se a célula na coluna I contém "X"
    If UCase(Trim(Sheets("Criação").Cells(i, "I").Value)) = "X" Then
        ' Armazena a NF e o Cliente da linha "X" para procurar a linha mestre
        nfToFind = Trim(Sheets("Criação").Cells(i, "F").Value)
        clientToFind = Trim(Sheets("Criação").Cells(i, "B").Value)
        masterResult = "" ' Reseta a variável de resultado

        ' Procura pela linha mestre (mesma NF/Cliente, com "Duplicado" e um resultado válido)
        For j = 2 To lastRow
            If Trim(Sheets("Criação").Cells(j, "F").Value) = nfToFind And _
               Trim(Sheets("Criação").Cells(j, "B").Value) = clientToFind Then
                
                ' Verifica se é a linha mestre (marcada como "Duplicado" e com um resultado que não seja "X" ou vazio)
                If UCase(Trim(Sheets("Criação").Cells(j, "K").Value)) = "DUPLICADO" And _
                   Trim(Sheets("Criação").Cells(j, "I").Value) <> "" And _
                   UCase(Trim(Sheets("Criação").Cells(j, "I").Value)) <> "X" Then
                    
                    masterResult = Sheets("Criação").Cells(j, "I").Value
                    Exit For ' Encontrou a linha mestre, pode sair do loop de busca
                End If
            End If
        Next j

        ' Se um resultado mestre foi encontrado, atualiza a linha "X"
        If masterResult <> "" Then
            Sheets("Criação").Cells(i, "I").Value = masterResult
        End If
    End If
Next i


MsgBox "Finalizado.", vbInformation

End Sub