Attribute VB_Name = "M�dulo3"
Sub CriarOrdemDevolucao()

' Declara��o de todas as vari�veis utilizadas na macro.
Dim SapGuiAuto As Object
Dim app As Object
Dim Connection As Object
Dim session As Object

Dim dictInvoices As Object ' Dicion�rio para agrupar linhas por Nota Fiscal
Dim invoiceKey As Variant
Dim invoiceRows As Collection
Dim itemRow As Variant
Dim resultRow As Variant

Dim lastRow As Long
Dim i As Long
Dim firstRow As Long

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
Dim Itemq As Double
Dim Itemloop As String
Dim Item2 As String
Dim notaFiscalOriginal As String
Dim notaFiscalFormatada As String

Dim motivoOriginalCriacao As String
Dim codigoParaBusca As String
Dim motivoFinalParaSAP As String
Dim wsCodigo As Worksheet
Dim lastRowCodigo As Long
Dim j As Long
Dim posAbreParenteses As Integer
Dim posFechaParenteses As Integer

' *** Vari�veis do c�digo antigo que ser�o restauradas ***
Dim cont As Integer
Dim conterr As Integer
Dim contmaior As Integer

' Ativa a aba "Cria��o" e define o cabe�alho para o nome do solicitante.
Sheets("Cria��o").Activate
Range("L1").Value = "NOME"

' Valida��o para garantir que o nome do solicitante foi preenchido.
If Range("L2").Value = "" Then
    MsgBox "� obrigat�rio preencher o NOME na c�lula L2!", vbCritical, "CONTROLE DE DADOS"
    Range("L2").Select
    Exit Sub
End If

Application.DisplayAlerts = False

' --- CONEX�O COM O SAP (EXECUTADA APENAS UMA VEZ) ---
On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
If SapGuiAuto Is Nothing Then
    MsgBox "N�o foi poss�vel encontrar o SAP GUI. Verifique se ele est� em execu��o.", vbCritical, "Erro de Conex�o"
    Exit Sub
End If

Set app = SapGuiAuto.GetScriptingEngine
If app Is Nothing Then
    MsgBox "N�o foi poss�vel obter o Scripting Engine do SAP. Verifique as configura��es de script do SAP.", vbCritical, "Erro de Conex�o"
    Exit Sub
End If

Set Connection = app.Children(0)
If Connection Is Nothing Then
    MsgBox "Nenhuma conex�o SAP encontrada. Verifique se voc� est� logado em um sistema SAP.", vbCritical, "Erro de Conex�o"
    Exit Sub
End If

Set session = Connection.Children(0)
If session Is Nothing Then
    MsgBox "Nenhuma sess�o SAP encontrada. Verifique se h� uma janela de sess�o aberta.", vbCritical, "Erro de Conex�o"
    Exit Sub
End If
On Error GoTo 0 ' Restaura o tratamento de erros padr�o

If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject app, "on"
End If
' --- FIM DA CONEX�O SAP ---


' Cria um objeto Dicion�rio para agrupar as ordens.
Set dictInvoices = CreateObject("Scripting.Dictionary")
lastRow = Sheets("Cria��o").Cells(Sheets("Cria��o").rows.Count, "A").End(xlUp).Row

' --- Agrupamento de Linhas por Nota Fiscal e Cliente ---
For i = 2 To lastRow
    Dim key As String
    Dim nf As String
    Dim cliente As String
    
    nf = Trim(Sheets("Cria��o").Cells(i, "F").Value)
    cliente = Trim(Sheets("Cria��o").Cells(i, "B").Value)

    If nf <> "" And cliente <> "" Then
        If UCase(Trim(Sheets("Cria��o").Cells(i, "K").Value)) = "DUPLICADO" Then
            key = nf & "|" & cliente
        Else
            key = nf & "|" & cliente & "|" & i
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

    If Trim(Sheets("Cria��o").Cells(firstRow, "I").Value) <> "" Then
        GoTo ProximoGrupo
    End If

    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nVA01"
    session.findById("wnd[0]").sendVKey 0

    If InStr(1, Sheets("Cria��o").Cells(firstRow, "G").Value, "668") > 0 Then
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
    
    notaFiscalOriginal = Trim(Sheets("Cria��o").Cells(firstRow, "F").Value)
    notaFiscalFormatada = Format(notaFiscalOriginal, "000000000") & "-1"
    
    session.findById("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB005/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").Text = notaFiscalFormatada
    session.findById("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB005/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[4,24]").Text = Trim(Sheets("Cria��o").Cells(firstRow, "B").Value)
    session.findById("wnd[2]/tbar[0]/btn[0]").press

    On Error Resume Next
    strTexto = session.findById("wnd[0]/sbar").Text

    If strTexto = "Nenhum valor para esta sele��o" Then
        session.findById("wnd[2]/tbar[0]/btn[12]").press
        session.findById("wnd[1]/tbar[0]/btn[12]").press
        
        For Each resultRow In invoiceRows
            Sheets("Cria��o").Cells(resultRow, "I").Value = "NF inexistente para o c�digo de cliente informado"
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

    If parceiro = "Sele��o de parceiro" Then
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
            mensagemerro = "Pe�a Livre de D�bito"
        Else
            mensagemerro = "Tipo e NF n�o gera devolu��o - Doa��o"
        End If
        
        For Each resultRow In invoiceRows
            Sheets("Cria��o").Cells(resultRow, "I").Value = mensagemerro
        Next resultRow
        
        tipozdvp = ""
        GoTo ProximoGrupo
    End If

volta3:
    teste = session.findById("wnd[1]/usr/txtMESSTXT1").Text
        If Left(teste, 2) = "J�" Then
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
    
    '--- In�cio da L�gica de Motivo (do c�digo novo)
    motivoOriginalCriacao = Trim(CStr(Sheets("Cria��o").Cells(firstRow, "G").Value))
    posAbreParenteses = InStr(motivoOriginalCriacao, "(")
    posFechaParenteses = InStr(motivoOriginalCriacao, ")")

    If posAbreParenteses > 0 And posFechaParenteses > posAbreParenteses Then
        codigoParaBusca = Trim(Mid(motivoOriginalCriacao, posAbreParenteses + 1, posFechaParenteses - posAbreParenteses - 1))
    Else
        codigoParaBusca = ""
    End If

    If codigoParaBusca = "90" Then
        codigoParaBusca = "090"
    ElseIf codigoParaBusca = "92" Then
        codigoParaBusca = "092"
    End If

    motivoFinalParaSAP = ""
    If codigoParaBusca <> "" Then
        Set wsCodigo = ThisWorkbook.Sheets("C�digo")
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
        Sheets("Cria��o").Cells(firstRow, "I").Value = "ERRO: C�digo '" & codigoParaBusca & "' n�o encontrado na aba 'C�digo'."
        GoTo ProximoGrupo
    End If
    '--- Fim da L�gica de Motivo
    
    If InStr(1, Sheets("Cria��o").Cells(firstRow, "G").Value, "668") > 0 Then
        GoTo OIROB
    End If
    
    '************************************************************************************************
    '*** IN�CIO DA L�GICA DE ITENS 100% RESTAURADA DO SCRIPT ANTIGO ***
    '************************************************************************************************
    
    cont = 1
    Do While cont <= invoiceRows.Count
        itemRow = invoiceRows(cont)

        If Sheets("Cria��o").Cells(itemRow, "H").Value <= 0 Then
            ' Pula para o pr�ximo item se a quantidade for zero
            cont = cont + 1
        Else
            If cont > 1 Then GoTo mais
            
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_MKAL").press 'SELECIONA TD
mais:
            If cont > 1 Then
                session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_MKLO").press
                session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_MKAL").press
            End If
            
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POPO").press 'PROCURA
            
            Itemloop = Sheets("Cria��o").Cells(itemRow, "D").Value
            session.findById("wnd[1]/usr/ctxtRV45A-PO_MATNR").Text = Itemloop
            session.findById("wnd[1]/tbar[0]/btn[0]").press

            On Error Resume Next
            erroitem = session.findById("wnd[2]/usr/txtMESSTXT1").Text
            If erroitem = "Item  inexistente" Then
                iteninex = Sheets("Cria��o").Cells(itemRow, "D").Value
                erroitem = ""
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                
                For Each resultRow In invoiceRows
                    Sheets("Cria��o").Cells(resultRow, "I").Value = "Item " & iteninex & " inexistente para NF"
                Next resultRow
                
                iteninex = ""
                GoTo ProximoGrupo
            End If
            On Error GoTo 0
            
            session.findById("wnd[0]").sendVKey 9 'Desmarca a linha

            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,0]").SetFocus
            Itemq = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,0]").Text

            If Itemq < Sheets("Cria��o").Cells(itemRow, "H").Value Then
                For Each resultRow In invoiceRows
                    Sheets("Cria��o").Cells(resultRow, "I").Value = "QUANTIDADE SOLICITADA (" & Sheets("Cria��o").Cells(itemRow, "H").Text & "p�) � MAIOR QUE FATURADA(" & Itemq & "p�) para o item " & Itemloop & ""
                Next resultRow
                GoTo ProximoGrupo
            End If

            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,0]").Text = Sheets("Cria��o").Cells(itemRow, "H").Value
            
            cont = cont + 1
        End If
    Loop
    
    On Error Resume Next
    Item2 = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,1]").Text
    If Err.Number <> 0 Then Item2 = ""
    On Error GoTo 0
    
    If Item2 = "" Then
        GoTo pularDelete
    End If
    
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POLO").press
    On Error Resume Next
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    On Error GoTo 0
    
pularDelete:
    '************************************************************************************************
    '*** FIM DA L�GICA RESTAURADA ***
    '************************************************************************************************

OIROB:
    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
    ' ... O resto do c�digo continua igual ...
    
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04").Select

    If InStr(1, Sheets("Cria��o").Cells(firstRow, "G").Value, "668") = 0 Then
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
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,8]").Text = Sheets("Cria��o").Cells(firstRow, "J").Value
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,8]").SetFocus
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,8]").caretPosition = 7
        session.findById("wnd[0]").sendVKey 0
    Else
        If session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & contsp & "]").Text = "SP Transportador" Then
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1," & contsp & "]").Text = Sheets("Cria��o").Cells(firstRow, "J").Value
        Else
            contsp = contsp + 1
            GoTo trs
        End If
    End If

    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08").Select
    
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text = Sheets("Cria��o").Cells(firstRow, "A").Text & " - " & Sheets("Cria��o").Range("L2").Text
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes 11, 11

    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09").Select
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/ctxtVBKD-BSARK").Text = "ZLR1"

    If InStr(1, Sheets("Cria��o").Cells(firstRow, "G").Value, "509") > 0 Then
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBAK-SUBMI").Text = "01"
    End If

    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBKD-BSTKD").Text = Sheets("Cria��o").Cells(firstRow, "A").Text
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBKD-BSTKD_E").Text = Sheets("Cria��o").Cells(firstRow, "A").Text
    session.findById("wnd[0]/tbar[0]/btn[11]").press

    ois = ""
    ois1 = ""
    DOCNUM = ""

    If tipo = "ROB" Or tipo = "ZDVP" Then
        ois1 = session.findById("wnd[0]/sbar").Text
        If Left(ois1, 2) = "J�" Then
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
            Sheets("Cria��o").Cells(resultRow, "I").Value = "Dev. NF Cliente Parc " & ois & " foi gravado (foi gerado o fornecimento " & DOCNUM & ")"
        Next resultRow

    Else
voltaOI:
        For Each resultRow In invoiceRows
            Sheets("Cria��o").Cells(resultRow, "I").Value = session.findById("wnd[0]/sbar").Text
        Next resultRow
    End If
    
    ActiveWorkbook.Save
    session.findById("wnd[0]/tbar[0]/btn[3]").press

ProximoGrupo:
    On Error Resume Next
    On Error GoTo 0

Next invoiceKey

' --- P�S-PROCESSAMENTO PARA CORRIGIR LINHAS COM "X" ---
Dim nfToFind As String
Dim clientToFind As String
Dim masterResult As String

For i = 2 To lastRow
    If UCase(Trim(Sheets("Cria��o").Cells(i, "I").Value)) = "X" Then
        nfToFind = Trim(Sheets("Cria��o").Cells(i, "F").Value)
        clientToFind = Trim(Sheets("Cria��o").Cells(i, "B").Value)
        masterResult = ""

        For j = 2 To lastRow
            If Trim(Sheets("Cria��o").Cells(j, "F").Value) = nfToFind And _
               Trim(Sheets("Cria��o").Cells(j, "B").Value) = clientToFind Then
                
                If UCase(Trim(Sheets("Cria��o").Cells(j, "K").Value)) = "DUPLICADO" And _
                   Trim(Sheets("Cria��o").Cells(j, "I").Value) <> "" And _
                   UCase(Trim(Sheets("Cria��o").Cells(j, "I").Value)) <> "X" Then
                    
                    masterResult = Sheets("Cria��o").Cells(j, "I").Value
                    Exit For
                End If
            End If
        Next j

        If masterResult <> "" Then
            Sheets("Cria��o").Cells(i, "I").Value = masterResult
        End If
    End If
Next i


MsgBox "Finalizado.", vbInformation

End Sub

