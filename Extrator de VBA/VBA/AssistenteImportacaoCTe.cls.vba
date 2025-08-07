Attribute VB_Name = "AssistenteImportacaoCTe"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private EnumContribuicoes As New clsEnumeracoesSPEDContribuicoes
Private EnumFiscal As New clsEnumeracoesSPEDFiscal
Private GerenciadorSPED As New clsRegistrosSPED
Private assImportacao As AssistenteImportacao
Private ExpReg As ExportadorRegistros

Public Sub ImportarDTe(ByVal tpImportacao As String)
    
    Call ProcessarDocumentos.CarregarXMLS(tpImportacao)
    If tpImportacao = "Arquivo" Then DocsFiscais.arrCTe.addRange DocsFiscais.arrTodos
    
    CNPJBase = VBA.Left(CNPJContribuinte, 8)
    If DocsFiscais.arrCTe.Count = 0 Then Exit Sub
    
    Call InicializarObjetos
    Call Util.AtualizarBarraStatus("Iniciando importação dos XMLs...")
    
    Call ProcessarXMLS
    
    Call ExpReg.ExportarRegistros("0000", "0000_Contr", "0001", "0005", "0100", "0110", _
        "0140", "0150", "D001", "D010", "D100", "D101", "D101_Contr", "D105", "D110", "D190")
        
    Call Util.MsgInformativa("Registros gerados com sucesso!", "Importação CTe/NCFe", Inicio)
    
    Call LimparObjetos
    
End Sub

Private Function ProcessarXMLS()

Dim b As Long
Dim XML As Variant
Dim CTe As DOMDocument60
    
    b = 0
    Comeco = Timer
    DocsSemValidade = 0
    For Each XML In DocsFiscais.arrCTe
        
        Call Util.AntiTravamento(b, 100, "Importando XML " & b + 1 & " de " & DocsFiscais.arrCTe.Count, DocsFiscais.arrCTe.Count, Comeco)
        
        Set CTe = assImportacao.ExtrairDadosXML(XML)
        If Not CTe Is Nothing Then Call GerarRegistrosSPED(CTe)
        
    Next XML
    
End Function

Private Sub GerarRegistrosSPED(ByRef CTe As DOMDocument60)
    
    With assImportacao
        
        If Not dtoRegSPED.r0000.Exists(DadosXML.ARQUIVO) Then Call .CriarRegistro0000(CTe)
        If Not dtoRegSPED.r0000_Contr.Exists(DadosXML.ARQUIVO) Then Call .CriarRegistro0000_Contr(CTe)
        If Not dtoRegSPED.r0001.Exists(DadosXML.ARQUIVO) Then Call .CriarRegistro0001
        If Not dtoRegSPED.r0005.Exists(DadosXML.ARQUIVO) Then Call .CriarRegistro0005(CTe)
        If Not dtoRegSPED.r0100.Exists(DadosXML.ARQUIVO) Then Call .CriarRegistro0100
        If Not dtoRegSPED.r0110.Exists(DadosXML.ARQUIVO) Then Call .CriarRegistro0110
        If Not dtoRegSPED.r0140.Exists(DadosXML.ARQUIVO & DadosXML.CNPJ_ESTABELECIMENTO) Then Call .CriarRegistro0140(CTe)
        If Not dtoRegSPED.rD001.Exists(DadosXML.ARQUIVO) Then Call .CriarRegistroD001
        If Not dtoRegSPED.rD010.Exists(DadosXML.ARQUIVO & DadosXML.CNPJ_ESTABELECIMENTO) Then Call .CriarRegistroD010
        
        Call ProcessarXML(CTe)
        
    End With
    
End Sub

Private Sub ProcessarXML(ByRef CTe As DOMDocument60)

Dim Chave As String
    
    If dtoRegSPED.rD100 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroD100("IND_OPER", "IND_EMIT", "CHV_CTe")
    
    With CamposD100
        
        .IND_OPER = assImportacao.IdentificarTipoOperacao()
        .IND_EMIT = assImportacao.IdentificarTipoEmissao(CTe)
        .CHV_CTE = VBA.Right(ValidarTag(CTe, "//@Id"), 44)
        
        Chave = Util.UnirCampos(.IND_OPER, .IND_EMIT, .CHV_CTE)
        If dtoRegSPED.rD100.Exists(Chave) Then Call CarregarCamposD100(Chave) Else Call CriarRegistroD100(CTe)
        
        If .COD_PART <> "" Then Call assImportacao.CriarRegistro0150(CTe, .COD_PART)
        Call CriarRegistroD101(CTe)
        Call CriarRegistroD105(CTe)
        
    End With
    
    'Incluir rotinas de processamento dos sub registros
    
End Sub

Private Sub CarregarCamposD100(ByVal Chave As String)

Dim Campos As Variant
Dim i As Long
    
    Campos = dtoRegSPED.rD100(Chave)
    If IsEmpty(Campos) Then Exit Sub
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    With CamposD100
        
        .REG = Campos(dtoTitSPED.tD100("REG"))
        .ARQUIVO = Campos(dtoTitSPED.tD100("ARQUIVO"))
        .CHV_REG = Campos(dtoTitSPED.tD100("CHV_REG"))
        .CHV_PAI_FISCAL = Campos(dtoTitSPED.tD100("CHV_PAI_FISCAL"))
        .CHV_PAI_CONTRIBUICOES = Campos(dtoTitSPED.tD100("CHV_PAI_CONTRIBUICOES"))
        .IND_OPER = Campos(dtoTitSPED.tD100("IND_OPER"))
        .IND_EMIT = Campos(dtoTitSPED.tD100("IND_EMIT"))
        .COD_PART = Campos(dtoTitSPED.tD100("COD_PART"))
        .COD_MOD = Campos(dtoTitSPED.tD100("COD_MOD"))
        .COD_SIT = Campos(dtoTitSPED.tD100("COD_SIT"))
        .SER = Campos(dtoTitSPED.tD100("SER"))
        .NUM_DOC = Campos(dtoTitSPED.tD100("NUM_DOC"))
        .CHV_CTE = Campos(dtoTitSPED.tD100("CHV_CTE"))
        .DT_DOC = Campos(dtoTitSPED.tD100("DT_DOC"))
        .DT_A_P = Campos(dtoTitSPED.tD100("DT_A_P"))
        .TP_CTe = Campos(dtoTitSPED.tD100("TP_CT_E"))
        .CHV_CTE_REF = Campos(dtoTitSPED.tD100("CHV_CTE_REF"))
        .VL_DOC = Campos(dtoTitSPED.tD100("VL_DOC"))
        .VL_DESC = Campos(dtoTitSPED.tD100("VL_DESC"))
        .IND_FRT = Campos(dtoTitSPED.tD100("IND_FRT"))
        .VL_SERV = Campos(dtoTitSPED.tD100("VL_SERV"))
        .VL_BC_ICMS = Campos(dtoTitSPED.tD100("VL_BC_ICMS"))
        .VL_ICMS = Campos(dtoTitSPED.tD100("VL_ICMS"))
        .VL_NT = Campos(dtoTitSPED.tD100("VL_NT"))
        .COD_INF = Campos(dtoTitSPED.tD100("COD_INF"))
        .COD_CTA = Campos(dtoTitSPED.tD100("COD_CTA"))
        .COD_MUN_ORIG = Campos(dtoTitSPED.tD100("COD_MUN_ORIG"))
        .COD_MUN_DEST = Campos(dtoTitSPED.tD100("COD_MUN_DEST"))
        
    End With
    
End Sub

Public Sub CriarRegistroD100(ByRef CTe As DOMDocument60)

Dim Campos As Variant
Dim Chave  As String
    
    If dtoRegSPED.rD100 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroD100("IND_OPER", "IND_EMIT", "CHV_CTE")
    
    With CamposD100
        
        .REG = "D100"
        .ARQUIVO = DadosXML.ARQUIVO
        .COD_MOD = fnXML.ValidarTag(CTe, "//mod")
        .COD_PART = assImportacao.ExtrairCodigoParticipante()
        .COD_SIT = EnumContribuicoes.ValidarEnumeracao_COD_SIT(fnSPED.GerarCodigoSituacao(fnXML.ValidarSituacao(ValidarTag(CTe, "//cStat"))))
        .SER = VBA.Format(fnXML.ValidarTag(CTe, "//serie"), "000")
        .SUB = fnXML.ValidarTag(CTe, "//subserie")
        .NUM_DOC = fnXML.ValidarTag(CTe, "//nNF")
        .DT_DOC = fnXML.ExtrairDataDocumento(CTe)
        .DT_A_P = .DT_DOC
        .TP_CTe = EnumContribuicoes.ValidarEnumeracao_TP_CT_E(ValidarTag(CTe, "//tpCTe"))
        .CHV_CTE_REF = ""
        .VL_DOC = fnXML.ValidarValores(CTe, "//vRec")
        .VL_DESC = 0
        .IND_FRT = EnumContribuicoes.ValidarEnumeracao_IND_FRT(ValidarTag(CTe, "//modFrete"))
        .VL_SERV = fnXML.ValidarValores(CTe, "//vPrest/vRec")
        .VL_BC_ICMS = fnXML.ValidarValores(CTe, "//imp/ICMS//vBC")
        .VL_ICMS = fnXML.ValidarValores(CTe, "//imp/ICMS//vICMS")
        .VL_NT = 0
        .COD_INF = ""
        .COD_CTA = ""
        .COD_MUN_ORIG = ValidarTag(CTe, "//cMunIni")
        .COD_MUN_DEST = ValidarTag(CTe, "//cMunFim")
        .CHV_PAI_FISCAL = assImportacao.ExtrairChaveRegD001()
        .CHV_PAI_CONTRIBUICOES = assImportacao.ExtrairChaveRegD010()
        
        If .COD_PART <> "" Then .COD_PART = Util.FormatarTexto(.COD_PART)
        
        'Verifica se a nota está cancelada
        If DocsFiscais.arrChavesCanceladas.contains(VBA.Replace(.CHV_CTE, "'", "")) Then _
            .COD_SIT = EnumContribuicoes.ValidarEnumeracao_COD_SIT(fnSPED.GerarCodigoSituacao(fnXML.ValidarSituacao("101")))
            
        'Elimina as informações das notas canceladas
        If VBA.Left(.COD_SIT, 2) = "02" Or VBA.Left(.COD_SIT, 2) = "03" Then
            
            .COD_PART = ""
            .VL_DOC = 0
            .VL_DESC = 0
            .IND_FRT = ""
            .VL_SERV = ""
            .VL_BC_ICMS = 0
            .VL_ICMS = 0
            .VL_NT = 0
            .COD_MUN_ORIG = ""
            .COD_MUN_DEST = ""
            
        End If
        
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .IND_OPER, .IND_EMIT, .COD_PART, .COD_MOD, .SER, .SUB, .NUM_DOC, .CHV_CTE)
        Chave = Util.UnirCampos(.IND_OPER, .IND_EMIT, .CHV_CTE)
        
        Campos = Array(.REG, ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, .IND_OPER, _
            .IND_EMIT, Util.FormatarTexto(.COD_PART), .COD_MOD, .COD_SIT, .SER, .SUB, .NUM_DOC, .CHV_CTE, _
            .DT_DOC, .DT_A_P, .TP_CTe, .CHV_CTE_REF, CDbl(.VL_DOC), CDbl(.VL_DESC), .IND_FRT, CDbl(.VL_SERV), _
            CDbl(.VL_BC_ICMS), CDbl(.VL_ICMS), CDbl(.VL_NT), .COD_INF, .COD_CTA, .COD_MUN_ORIG, .COD_MUN_DEST)
            
        Chave = Util.UnirCampos(.IND_OPER, .IND_EMIT, .CHV_CTE)
        dtoRegSPED.rD100(Chave) = Campos
        
    End With
    
End Sub

Public Sub CriarRegistroD101(ByRef CTe As DOMDocument60)

Dim Campos As Variant
Dim total As Double
    
    If dtoRegSPED.rD101 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroD101("CHV_PAI_FISCAL")
    
    With CamposD101
        
        If Not dtoRegSPED.rD101.Exists(CamposD100.CHV_REG) Then
            
            .REG = "D101"
            .ARQUIVO = DadosXML.ARQUIVO
            .VL_FCP_UF_DEST = fnXML.ValidarValores(CTe, "//ICMSTot/vFCPUFDest")
            .VL_ICMS_UF_DEST = fnXML.ValidarValores(CTe, "//ICMSTot/vICMSUFDest")
            .VL_ICMS_UF_REM = fnXML.ValidarValores(CTe, "//ICMSTot/vICMSUFRemet")
            .CHV_PAI_FISCAL = CamposD100.CHV_REG
            .CHV_PAI_CONTRIBUICOES = ""
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, "D101")
            
            Campos = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, _
                CDbl(.VL_FCP_UF_DEST), CDbl(.VL_ICMS_UF_DEST), CDbl(.VL_ICMS_UF_REM))
            
            total = CDbl(.VL_FCP_UF_DEST) + CDbl(.VL_ICMS_UF_DEST) + CDbl(.VL_ICMS_UF_REM)
            If total > 0 Then dtoRegSPED.rD101(CamposD100.CHV_REG) = Campos
            
        End If
        
    End With
    
End Sub

'Private Function CriarRegistroD101(ByRef Campos As Variant, ByVal VL_DIFAL As Double, ByVal VL_FCP As Double)
'
'Dim CHV_PAI As String, CHV_CTe$, CHV_CTE$, COD_MUN_ORIG$, COD_MUN_DEST$
'
'    CHV_CTe = Campos(dicTitulosApuracao("CHV_CTe"))
'    If dicCorrelacoesCTeCTe.Exists(CHV_CTe) Then
'
'        CHV_CTE = dicCorrelacoesCTeCTe(CHV_CTe)(dicTitulosCTeCTe("CHV_CTE"))
'        If dicDadosD100.Exists(CHV_CTE) Then
'
'            CHV_PAI = dicDadosD100(CHV_CTE)(dicTitulosD100("CHV_REG"))
'            COD_MUN_ORIG = Util.ApenasNumeros(dicDadosD100(CHV_CTE)(dicTitulosD100("COD_MUN_ORIG")))
'            COD_MUN_DEST = Util.ApenasNumeros(dicDadosD100(CHV_CTE)(dicTitulosD100("COD_MUN_DEST")))
'
'            If VBA.Left(COD_MUN_ORIG, 2) <> VBA.Left(COD_MUN_DEST, 2) Then
'
'                With CamposD101
'
'                    .REG = "D101"
'                    .ARQUIVO = Campos(dicTitulosApuracao("ARQUIVO"))
'                    .CHV_PAI = CHV_PAI
'                    .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, "D101")
'                    .VL_FCP_UF_DEST = VL_FCP
'                    .VL_ICMS_UF_DEST = VL_DIFAL
'                    .VL_ICMS_UF_REM = 0
'
'                    Call AtualizarRegistroE310(CHV_E310, .VL_ICMS_UF_DEST, .VL_FCP_UF_DEST, 0, 0, 0, 0)
'                    If dicDadosD101.Exists(.CHV_REG) Then Call AtualizarRegistroD101(.CHV_REG)
'
'                    dicDadosD101(.CHV_REG) = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI, "", _
'                        CDbl(.VL_FCP_UF_DEST), CDbl(.VL_ICMS_UF_DEST), CDbl(.VL_ICMS_UF_REM))
'
'                End With
'
'            End If
'
'        End If
'
'     End If
'
'End Function

Public Sub CriarRegistroD101_Contr(ByVal Campos As Variant)

Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    With CamposD101_Contr
        
        .REG = "D101"
        .ARQUIVO = DadosXML.ARQUIVO
        .IND_NAT_FRT = ""
        .VL_ITEM = Campos(dtoTitSPED.tD100("VL_DOC") - i)
        .CST_PIS = assImportacao.ExtrairCST_PIS_COFINS_AquisicaoFrete(.ARQUIVO)
        .NAT_BC_CRED = ""
        .VL_BC_PIS = .VL_ITEM
        .ALIQ_PIS = assImportacao.ExtrairALIQ_PIS_AquisicaoFrete(.ARQUIVO)
        .VL_PIS = VBA.Round(.VL_BC_PIS * .ALIQ_PIS, 2)
        .COD_CTA = fnExcel.FormatarTexto(Campos(dtoTitSPED.tD100("COD_CTA") - i))
        .CHV_PAI_CONTRIBUICOES = Campos(dtoTitSPED.tD100("CHV_REG") - i)
        
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI_CONTRIBUICOES, .CST_PIS, .ALIQ_PIS, .COD_CTA)
        
        Campos = Array(.REG, .ARQUIVO, .CHV_REG, "", .CHV_PAI_CONTRIBUICOES, .IND_NAT_FRT, CDbl(.VL_ITEM), _
            .CST_PIS, .NAT_BC_CRED, CDbl(.VL_BC_PIS), CDbl(.ALIQ_PIS), CDbl(.VL_PIS), .COD_CTA)
            
        dtoRegSPED.rD101_Contr(.CHV_REG) = Campos
        
    End With
    
End Sub

Public Sub CriarRegistroD105(ByVal Campos As Variant)

Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    With CamposD105
        
        .REG = "D105"
        .ARQUIVO = Campos(dtoTitSPED.tD100("ARQUIVO") - i)
        .IND_NAT_FRT = ""
        .VL_ITEM = Campos(dtoTitSPED.tD100("VL_DOC") - i)
        .CST_COFINS = assImportacao.ExtrairCST_PIS_COFINS_AquisicaoFrete(.ARQUIVO)
        .NAT_BC_CRED = ""
        .VL_BC_COFINS = .VL_ITEM
        .ALIQ_COFINS = assImportacao.ExtrairALIQ_COFINS_AquisicaoFrete(.ARQUIVO)
        .VL_COFINS = VBA.Round(.VL_BC_COFINS * .ALIQ_COFINS, 2)
        .COD_CTA = fnExcel.FormatarTexto(Campos(dtoTitSPED.tD100("COD_CTA") - i))
        .CHV_PAI_CONTRIBUICOES = Campos(dtoTitSPED.tD100("CHV_REG") - i)
        
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI_CONTRIBUICOES, .CST_COFINS, .ALIQ_COFINS, .COD_CTA)
        
        Campos = Array(.REG, .ARQUIVO, .CHV_REG, "", .CHV_PAI_CONTRIBUICOES, .IND_NAT_FRT, CDbl(.VL_ITEM), _
            .CST_COFINS, .NAT_BC_CRED, CDbl(.VL_BC_COFINS), CDbl(.ALIQ_COFINS), CDbl(.VL_COFINS), .COD_CTA)
            
        dtoRegSPED.rD105(.CHV_REG) = Campos
        
    End With
    
End Sub

Public Sub CriarRegistroD200(ByRef CTe As IXMLDOMNode, ByRef dicDadosD200 As Dictionary, ByRef dicTitulosD200 As Dictionary, _
    ByRef dicDadosD201 As Dictionary, ByRef dicDadosD205 As Dictionary, ByVal ARQUIVO As String, ByVal CHV_PAI As String)
    
Dim Chave As String, NUM_DOC_INI$, NUM_DOC_FIN$
Dim Campos As Variant, CamposDic
Dim VL_DOC As Double, vICMS#
Dim i As Byte

    With CamposD200
        
        .REG = "D200"
        .COD_MOD = ValidarTag(CTe, "//mod")
        .COD_SIT = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_COD_SIT(fnSPED.GerarCodigoSituacao(fnXML.ValidarSituacao(ValidarTag(CTe, "//cStat"))))
        .SER = VBA.Format(ValidarTag(CTe, "//serie"), "000")
        .SUB = ValidarTag(CTe, "//subserie")
        .NUM_DOC_INI = ValidarTag(CTe, "//nCT")
        .NUM_DOC_FIN = ValidarTag(CTe, "//nCT")
        .CFOP = ValidarTag(CTe, "//CFOP")
        .DT_REF = VBA.Format(VBA.Left(ValidarTag(CTe, "//dhEmi"), 10), "yyyy-mm-dd")
        .VL_DOC = fnXML.ValidarValores(CTe, "//vRec")
        .VL_DESC = 0
        .CHV_PAI = CHV_PAI
        
        Chave = fnSPED.GerarChaveRegistro(.COD_MOD, .SER, .COD_SIT, .CFOP, .DT_REF)
        If dicDadosD200.Exists(Chave) Then
            
            'carrega campos do dicionário
            CamposDic = dicDadosD200(Chave)
            If LBound(CamposDic) = 0 Then i = 1 Else i = 0
            
            'carrega informações a serem atualizadas no registro
            NUM_DOC_INI = Util.ApenasNumeros(CamposDic(dicTitulosD200("NUM_DOC_INI") - i))
            NUM_DOC_FIN = Util.ApenasNumeros(CamposDic(dicTitulosD200("NUM_DOC_FIN") - i))
            VL_DOC = CamposDic(dicTitulosD200("VL_DOC") - i)
            
            .VL_DOC = VL_DOC + CDbl(.VL_DOC)
            If NUM_DOC_INI < .NUM_DOC_INI Then .NUM_DOC_INI = NUM_DOC_INI
            If NUM_DOC_FIN > .NUM_DOC_FIN Then .NUM_DOC_FIN = NUM_DOC_FIN
            
        End If
        
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .COD_MOD, .COD_SIT, .SER, .SUB, .NUM_DOC_INI, .NUM_DOC_FIN, .CFOP, .DT_REF)
        Campos = Array(.REG, ARQUIVO, .CHV_REG, "", .CHV_PAI, .COD_MOD, .COD_SIT, "'" & .SER, .SUB, _
            "'" & .NUM_DOC_INI, "'" & .NUM_DOC_FIN, .CFOP, .DT_REF, CDbl(.VL_DOC), CDbl(.VL_DESC))
        
        dicDadosD200(Chave) = Campos
        
        vICMS = fnXML.ValidarValores(CTe, "//imp/ICMS//vICMS")
        Call IncuirRegistroD201(CDbl(.VL_DOC), vICMS, dicDadosD201, ARQUIVO)
        Call IncuirRegistroD205(CDbl(.VL_DOC), vICMS, dicDadosD205, ARQUIVO)
        
    End With
    
End Sub

Public Sub IncuirRegistroD201(ByRef vRec As Double, ByRef vICMS As Double, ByRef dicDadosD201 As Dictionary, ByVal ARQUIVO As String)

Dim Campos As Variant
    
    With CamposD201
        
        .REG = "D201"
        .CST_PIS = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS("01")
        .VL_ITEM = vRec
        .VL_BC_PIS = vRec - vICMS
        .ALIQ_PIS = 0.0065
        .VL_PIS = VBA.Round(.VL_BC_PIS * .ALIQ_PIS, 2)
        .COD_CTA = ""
        .CHV_PAI = CamposD200.CHV_REG
        
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .CST_PIS, .ALIQ_PIS, .COD_CTA)
        Campos = Array(.REG, ARQUIVO, .CHV_REG, "", .CHV_PAI, "'" & .CST_PIS, CDbl(.VL_ITEM), CDbl(.VL_BC_PIS), CDbl(.ALIQ_PIS), CDbl(.VL_PIS), .COD_CTA)
        
        dicDadosD201(.CHV_REG) = Campos
        
    End With
    
End Sub

Public Sub IncuirRegistroD205(ByRef vRec As Double, ByRef vICMS As Double, ByRef dicDadosD205 As Dictionary, ByVal ARQUIVO As String)

Dim Campos As Variant
    
    With CamposD205
        
        .REG = "D205"
        .CST_COFINS = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS("01")
        .VL_ITEM = vRec
        .VL_BC_COFINS = vRec - vICMS
        .ALIQ_COFINS = 0.03
        .VL_COFINS = VBA.Round(.VL_BC_COFINS * .ALIQ_COFINS, 2)
        .COD_CTA = ""
        .CHV_PAI = CamposD200.CHV_REG
        
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .CST_COFINS, .ALIQ_COFINS, .COD_CTA)
        Campos = Array(.REG, ARQUIVO, .CHV_REG, "", .CHV_PAI, "'" & .CST_COFINS, CDbl(.VL_ITEM), CDbl(.VL_BC_COFINS), CDbl(.ALIQ_COFINS), CDbl(.VL_COFINS), .COD_CTA)
        
        dicDadosD205(.CHV_REG) = Campos
        
    End With
    
End Sub

Private Sub InicializarObjetos()
    
    Call Util.DesabilitarControles
    
    Set ExpReg = New ExportadorRegistros
    Set assImportacao = New AssistenteImportacao
    
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    Call CarregarRegistrosSPED
    
End Sub

Private Sub CarregarRegistrosSPED()
    
    With GerenciadorSPED
        
        Call .CarregarDadosRegistro0000("ARQUIVO")
        Call .CarregarDadosRegistro0000_Contr("ARQUIVO")
        Call .CarregarDadosRegistro0001("ARQUIVO")
        Call .CarregarDadosRegistro0005("ARQUIVO")
        Call .CarregarDadosRegistro0100("ARQUIVO")
        Call .CarregarDadosRegistro0110("ARQUIVO")
        Call .CarregarDadosRegistro0140("ARQUIVO", "CNPJ")
        Call .CarregarDadosRegistro0150("CHV_PAI_FISCAL", "COD_PART")
        Call .CarregarDadosRegistro0190("CHV_PAI_FISCAL", "UNID")
        Call .CarregarDadosRegistro0200("CHV_PAI_FISCAL", "COD_ITEM")
        Call .CarregarDadosRegistroD001("ARQUIVO")
        Call .CarregarDadosRegistroD010("ARQUIVO", "CNPJ")
        Call .CarregarDadosRegistroD100("IND_OPER", "IND_EMIT", "CHV_NFE")
        Call .CarregarDadosRegistroD101("CHV_PAI_FISCAL")
        Call .CarregarDadosRegistroD101_Contr("CHV_PAI_CONTRIBUICOES", "CST_PIS", "ALIQ_PIS", "COD_CTA")
        Call .CarregarDadosRegistroD105("CHV_PAI_CONTRIBUICOES", "CST_COFINS", "ALIQ_COFINS", "COD_CTA")
        Call .CarregarDadosRegistroD110("CHV_PAI_FISCAL", "NUM_ITEM")
        Call .CarregarDadosRegistroD190("CHV_PAI_FISCAL", "CST_ICMS", "CFOP", "ALIQ_ICMS")
        
    End With
    
End Sub

Private Sub LimparObjetos()
    
    Set ExpReg = Nothing
    Set assImportacao = Nothing
    
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    
    Call Util.AtualizarBarraStatus(False)
    Call Util.HabilitarControles
    
End Sub
