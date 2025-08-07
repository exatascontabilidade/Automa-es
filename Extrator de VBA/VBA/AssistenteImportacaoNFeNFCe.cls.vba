Attribute VB_Name = "AssistenteImportacaoNFeNFCe"
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
Private CNPJEstabelecimento As String
Private CNPJEmit As String
Private CNPJDest As String
Private Periodo As String
Private ARQUIVO As String
Private tpNF As String

Public Sub ImportarNFeNFCe(ByVal tpImportacao As String)
    
    Call ProcessarDocumentos.CarregarXMLS(tpImportacao)
    If tpImportacao = "Arquivo" Then DocsFiscais.arrNFeNFCe.addRange DocsFiscais.arrTodos
    
    CNPJBase = VBA.Left(CNPJContribuinte, 8)
    If DocsFiscais.arrNFeNFCe.Count = 0 Then Exit Sub
    
    Call InicializarObjetos
    Call Util.AtualizarBarraStatus("Iniciando importação dos XMLs...")
    
    Call ProcessarXMLS
    
    Call ExpReg.ExportarRegistros("0000", "0000_Contr", "0001", "0005", "0100", "0110", "0140", "0150", _
        "0190", "0200", "C001", "C010", "C100", "C101", "C110", "C113", "C120", "C140", "C141", "C170")
        
    Call Util.MsgInformativa("Registros gerados com sucesso!", "Importação NFe/NCFe", Inicio)
    
    Call LimparObjetos
    
End Sub

Private Function ProcessarXMLS()

Dim b As Long
Dim XML As Variant
Dim NFe As DOMDocument60
    
    b = 0
    Comeco = Timer
    DocsSemValidade = 0
    For Each XML In DocsFiscais.arrNFeNFCe
        
        Call Util.AntiTravamento(b, 100, "Importando XML " & b + 1 & " de " & DocsFiscais.arrNFeNFCe.Count, DocsFiscais.arrNFeNFCe.Count, Comeco)
        
        Set NFe = assImportacao.ExtrairDadosXML(XML)
        If Not NFe Is Nothing Then Call GerarRegistrosSPED(NFe)
        
    Next XML
    
End Function

Private Sub GerarRegistrosSPED(ByRef NFe As DOMDocument60)
    
    With assImportacao
        
        If Not dtoRegSPED.r0000.Exists(DadosXML.ARQUIVO) Then Call .CriarRegistro0000(NFe)
        If Not dtoRegSPED.r0000_Contr.Exists(DadosXML.ARQUIVO) Then Call .CriarRegistro0000_Contr(NFe)
        If Not dtoRegSPED.r0001.Exists(DadosXML.ARQUIVO) Then Call .CriarRegistro0001
        If Not dtoRegSPED.r0005.Exists(DadosXML.ARQUIVO) Then Call .CriarRegistro0005(NFe)
        If Not dtoRegSPED.r0100.Exists(DadosXML.ARQUIVO) Then Call .CriarRegistro0100
        If Not dtoRegSPED.r0110.Exists(DadosXML.ARQUIVO) Then Call .CriarRegistro0110
        If Not dtoRegSPED.r0140.Exists(DadosXML.ARQUIVO & DadosXML.CNPJ_ESTABELECIMENTO) Then Call .CriarRegistro0140(NFe)
        If Not dtoRegSPED.rC001.Exists(DadosXML.ARQUIVO) Then Call .CriarRegistroC001
        If Not dtoRegSPED.rC010.Exists(DadosXML.ARQUIVO & DadosXML.CNPJ_ESTABELECIMENTO) Then Call .CriarRegistroC010
        
        Call ProcessarXML(NFe)
        
    End With
    
End Sub

Private Sub InicializarObjetos()
    
    Call Util.DesabilitarControles
    
    Set ExpReg = New ExportadorRegistros
    Set assImportacao = New AssistenteImportacao
    
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    Call DTO_Correlacoes.CarregarCorrelacionamentos
    Call CarregarRegistrosSPED
    
End Sub

Private Sub LimparObjetos()
    
    Set ExpReg = Nothing
    Set assImportacao = Nothing
    
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    Call DTO_Correlacoes.ResetarDadosCorrelacionamento
    
    Call Util.AtualizarBarraStatus(False)
    Call Util.HabilitarControles
    
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
        Call .CarregarDadosRegistro0220("ARQUIVO", "UNID_COM")
        Call .CarregarDadosRegistroC001("ARQUIVO")
        Call .CarregarDadosRegistroC010("ARQUIVO", "CNPJ")
        Call .CarregarDadosRegistroC100("IND_OPER", "IND_EMIT", "CHV_NFE")
        Call .CarregarDadosRegistroC101("CHV_PAI_FISCAL")
        Call .CarregarDadosRegistroC110("CHV_PAI_FISCAL", "COD_INF")
        Call .CarregarDadosRegistroC113("CHV_PAI_FISCAL", "CHV_DOCE")
        Call .CarregarDadosRegistroC120("CHV_PAI_FISCAL", "NUM_DOC_IMP", "NUM_ACDRAW")
        Call .CarregarDadosRegistroC170("CHV_PAI_FISCAL", "NUM_ITEM")
        
    End With
    
End Sub

Private Sub ProcessarXML(ByRef NFe As DOMDocument60)

Dim Chave As String
    
    If dtoRegSPED.rC100 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC100("IND_OPER", "IND_EMIT", "CHV_NFE")
    
    With CamposC100
        
        .IND_OPER = assImportacao.IdentificarTipoOperacao()
        .IND_EMIT = assImportacao.IdentificarTipoEmissao(NFe)
        .CHV_NFE = VBA.Right(ValidarTag(NFe, "//@Id"), 44)
        
        Chave = Util.UnirCampos(.IND_OPER, .IND_EMIT, .CHV_NFE)
        If dtoRegSPED.rC100.Exists(Chave) Then Call CarregarCamposC100(NFe, Chave) Else Call CriarRegistroC100(NFe)
        
        If .COD_PART <> "" Then Call assImportacao.CriarRegistro0150(NFe, .COD_PART)
        Call CriarRegistroC101(NFe)
        'Call CriarRegistroC110(NFe)
        Call CriarRegistroC140(NFe)
        Call ProcessarProdutosNFe(NFe)
        
    End With
    
    'Incluir rotinas de processamento dos sub registros
    
End Sub

Private Sub CarregarCamposC100(ByRef NFe As IXMLDOMNode, ByVal Chave As String)

Dim Campos As Variant
Dim i As Long
    
    Campos = dtoRegSPED.rC100(Chave)
    If IsEmpty(Campos) Then Exit Sub
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    With CamposC100
        
        .REG = Campos(dtoTitSPED.tC100("REG"))
        .ARQUIVO = Campos(dtoTitSPED.tC100("ARQUIVO"))
        .CHV_REG = Campos(dtoTitSPED.tC100("CHV_REG"))
        .CHV_PAI_FISCAL = Campos(dtoTitSPED.tC100("CHV_PAI_FISCAL"))
        .CHV_PAI_CONTRIBUICOES = Campos(dtoTitSPED.tC100("CHV_PAI_CONTRIBUICOES"))
        .IND_OPER = Campos(dtoTitSPED.tC100("IND_OPER"))
        .IND_EMIT = Campos(dtoTitSPED.tC100("IND_EMIT"))
        .COD_PART = Campos(dtoTitSPED.tC100("COD_PART"))
        .COD_MOD = Campos(dtoTitSPED.tC100("COD_MOD"))
        .COD_SIT = Campos(dtoTitSPED.tC100("COD_SIT"))
        .SER = Campos(dtoTitSPED.tC100("SER"))
        .NUM_DOC = Campos(dtoTitSPED.tC100("NUM_DOC"))
        .CHV_NFE = Campos(dtoTitSPED.tC100("CHV_NFE"))
        .DT_DOC = Campos(dtoTitSPED.tC100("DT_DOC"))
        .DT_E_S = Campos(dtoTitSPED.tC100("DT_E_S"))
        .VL_DOC = Campos(dtoTitSPED.tC100("VL_DOC"))
        .IND_PGTO = Campos(dtoTitSPED.tC100("IND_PGTO"))
        .VL_DESC = Campos(dtoTitSPED.tC100("VL_DESC"))
        .VL_ABAT_NT = Campos(dtoTitSPED.tC100("VL_ABAT_NT"))
        .VL_MERC = Campos(dtoTitSPED.tC100("VL_MERC"))
        .IND_FRT = Campos(dtoTitSPED.tC100("IND_FRT"))
        .VL_FRT = Campos(dtoTitSPED.tC100("VL_FRT"))
        .VL_SEG = Campos(dtoTitSPED.tC100("VL_SEG"))
        .VL_OUT_DA = Campos(dtoTitSPED.tC100("VL_OUT_DA"))
        .VL_BC_ICMS = Campos(dtoTitSPED.tC100("VL_BC_ICMS"))
        .VL_ICMS = Campos(dtoTitSPED.tC100("VL_ICMS"))
        .VL_BC_ICMS_ST = Campos(dtoTitSPED.tC100("VL_BC_ICMS_ST"))
        .VL_ICMS_ST = Campos(dtoTitSPED.tC100("VL_ICMS_ST"))
        .VL_IPI = Campos(dtoTitSPED.tC100("VL_IPI"))
        .VL_PIS = Campos(dtoTitSPED.tC100("VL_PIS"))
        .VL_COFINS = Campos(dtoTitSPED.tC100("VL_COFINS"))
        .VL_PIS_ST = Campos(dtoTitSPED.tC100("VL_PIS_ST"))
        .VL_COFINS_ST = Campos(dtoTitSPED.tC100("VL_COFINS_ST"))
        
    End With
    
End Sub

Public Sub CriarRegistroC100(ByRef NFe As DOMDocument60)

Dim Campos As Variant
Dim Chave  As String
    
    If dtoRegSPED.rC100 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC100("IND_OPER", "IND_EMIT", "CHV_NFE")
    
    With CamposC100
        
        .REG = "C100"
        .ARQUIVO = DadosXML.ARQUIVO
        .COD_MOD = fnXML.ValidarTag(NFe, "//mod")
        .COD_PART = assImportacao.ExtrairCodigoParticipante()
        .COD_SIT = EnumContribuicoes.ValidarEnumeracao_COD_SIT(fnSPED.GerarCodigoSituacao(fnXML.ValidarSituacao(ValidarTag(NFe, "//cStat"))))
        .SER = VBA.Format(fnXML.ValidarTag(NFe, "//serie"), "000")
        .NUM_DOC = fnXML.ValidarTag(NFe, "//nNF")
        .DT_DOC = fnXML.ExtrairDataDocumento(NFe)
        .DT_E_S = fnXML.ExtrairDataEntradaSaida(NFe)
        .VL_DOC = fnXML.ValidarValores(NFe, "//ICMSTot/vNF")
        .IND_PGTO = fnXML.ExtrairTipoPagamento(NFe, .DT_DOC)
        .VL_DESC = fnXML.ValidarValores(NFe, "//ICMSTot/vDesc")
        .VL_ABAT_NT = fnXML.ValidarValores(NFe, "//ICMSTot/vICMSDeson")
        .VL_MERC = fnXML.ValidarValores(NFe, "//ICMSTot/vProd")
        .IND_FRT = EnumContribuicoes.ValidarEnumeracao_IND_FRT(ValidarTag(NFe, "//modFrete"))
        .VL_FRT = fnXML.ValidarValores(NFe, "//ICMSTot/vFrete")
        .VL_SEG = fnXML.ValidarValores(NFe, "//ICMSTot/vSeg")
        .VL_OUT_DA = fnXML.ValidarValores(NFe, "//ICMSTot/vOutro")
        .VL_BC_ICMS = fnXML.ExtrairBaseICMSTotal(NFe)
        .VL_ICMS = fnXML.ExtrairICMSTotal(NFe)
        .VL_BC_ICMS_ST = fnXML.ValidarValores(NFe, "//ICMSTot/vBCST")
        .VL_FCP_ST = fnXML.ValidarValores(NFe, "//ICMSTot/vFCPST")
        .VL_ICMS_ST = fnXML.ValidarValores(NFe, "//ICMSTot/vST") + CDbl(.VL_FCP_ST)
        .VL_IPI = fnXML.ValidarValores(NFe, "//ICMSTot/vIPI")
        .VL_PIS = fnXML.ValidarValores(NFe, "//ICMSTot/vPIS")
        .VL_COFINS = fnXML.ValidarValores(NFe, "//ICMSTot/vCOFINS")
        .VL_PIS_ST = 0
        .VL_COFINS_ST = 0
        .CHV_PAI_CONTRIBUICOES = assImportacao.ExtrairChaveRegC010()
        .CHV_PAI_FISCAL = assImportacao.ExtrairChaveRegC001()
        
        If .COD_MOD = "65" Then .COD_PART = ""
        If .IND_PGTO = "" Then .IND_PGTO = "0 - Á Vista"
        If .DT_E_S = "" Then .DT_E_S = .DT_DOC
        If .COD_PART <> "" Then .COD_PART = Util.FormatarTexto(.COD_PART)
        
        'Verifica se a nota está cancelada
        If DocsFiscais.arrChavesCanceladas.contains(VBA.Replace(.CHV_NFE, "'", "")) Then _
            .COD_SIT = EnumContribuicoes.ValidarEnumeracao_COD_SIT(fnSPED.GerarCodigoSituacao(fnXML.ValidarSituacao("101")))
            
        'Elimina as informações das notas canceladas
        If VBA.Left(.COD_SIT, 2) = "02" Or VBA.Left(.COD_SIT, 2) = "03" Then
            
            .COD_PART = ""
            .VL_DOC = 0
            .VL_PIS = 0
            .VL_COFINS = 0
            .VL_DESC = 0
            .VL_MERC = 0
            .VL_OUT_DA = 0
            .VL_BC_ICMS = 0
            .VL_ICMS = 0
            .VL_PIS_ST = 0
            .VL_COFINS_ST = 0
            
        End If
        
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .IND_OPER, .IND_EMIT, .COD_PART, .COD_MOD, .SER, .NUM_DOC, .CHV_NFE)
        Chave = Util.UnirCampos(.IND_OPER, .IND_EMIT, .CHV_NFE)
        
        Campos = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, .IND_OPER, .IND_EMIT, _
            Util.FormatarTexto(.COD_PART), .COD_MOD, .COD_SIT, "'" & .SER, "'" & .NUM_DOC, "'" & .CHV_NFE, .DT_DOC, _
            .DT_E_S, CDbl(.VL_DOC), .IND_PGTO, CDbl(.VL_DESC), CDbl(.VL_ABAT_NT), CDbl(.VL_MERC), .IND_FRT, _
            CDbl(.VL_FRT), CDbl(.VL_SEG), CDbl(.VL_OUT_DA), CDbl(.VL_BC_ICMS), CDbl(.VL_ICMS), CDbl(.VL_BC_ICMS_ST), _
            CDbl(.VL_ICMS_ST), CDbl(.VL_IPI), CDbl(.VL_PIS), CDbl(.VL_COFINS), CDbl(.VL_PIS_ST), CDbl(.VL_COFINS_ST))
        
        Chave = Util.UnirCampos(.IND_OPER, .IND_EMIT, .CHV_NFE)
        dtoRegSPED.rC100(Chave) = Campos
        
    End With
    
End Sub

Private Sub ProcessarProdutosNFe(ByRef NFe As IXMLDOMNode)

Dim Produtos As IXMLDOMNodeList
Dim Produto As IXMLDOMNode
    
    Set Produtos = NFe.SelectNodes("//det")
    
    For Each Produto In Produtos
        
        Call CriarRegistroC170(Produto)
        
    Next Produto
    
End Sub

Public Sub CriarRegistroC170(ByVal Produto As IXMLDOMNode)
    
Dim Chave As String

    If CamposC100.COD_SIT Like "02*" Then Exit Sub
    
    With CamposC170
        
        .NUM_ITEM = CInt(fnXML.ValidarnItem(Produto))
        .CHV_PAI_FISCAL = CamposC100.CHV_REG
        Chave = Util.UnirCampos(.CHV_PAI_FISCAL, .NUM_ITEM)
        If Not dtoRegSPED.rC170.Exists(Chave) Then
            
            .REG = "C170"
            .ARQUIVO = DadosXML.ARQUIVO
            .COD_ITEM = fnXML.ValidarTag(Produto, "prod/cProd")
            .DESCR_COMPL = Util.RemoverPipes(ValidarTag(Produto, "prod/xProd"))
            .QTD = fnXML.ValidarValores(Produto, "prod/qCom")
            .UNID = VBA.UCase(fnXML.ValidarTag(Produto, "prod/uCom"))
            .VL_ITEM = fnXML.ValidarValores(Produto, "prod/vProd")
            .VL_DESC = fnXML.ValidarValores(Produto, "prod/vDesc")
            .IND_MOV = "0 - SIM"
            .CST_ICMS = fnXML.ExtrairCST_CSOSN_ICMS(Produto)
            .CFOP = fnXML.ValidarValores(Produto, "prod/CFOP")
            .COD_NAT = ""
            .VL_BC_ICMS = fnXML.ExtrairBaseICMS(Produto)
            .ALIQ_ICMS = fnXML.ExtrairAliquotaICMS(Produto)
            .VL_ICMS = fnXML.ExtrairValorICMS(Produto)
            .VL_BC_ICMS_ST = fnXML.ValidarValores(Produto, "imposto/ICMS//vBCST")
            .ALIQ_ST = fnXML.ExtrairAliquotaICMSST(Produto)
            .VL_ICMS_ST = fnXML.ExtrairValorICMSST(Produto)
            .IND_APUR = ""
            .CST_IPI = EnumContribuicoes.ValidarEnumeracao_CST_IPI(ValidarTag(Produto, "imposto/IPI//CST"))
            .COD_ENQ = ""
            .VL_BC_IPI = fnXML.ValidarValores(Produto, "imposto/IPI//vBC")
            .ALIQ_IPI = fnXML.ValidarPercentual(Produto, "imposto/IPI//pIPI")
            .VL_IPI = fnXML.ValidarValores(Produto, "imposto/IPI//vIPI")
            .CST_PIS = EnumContribuicoes.ValidarEnumeracao_CST_PIS_COFINS(ValidarTag(Produto, "imposto/PIS//CST"))
            .VL_BC_PIS = fnXML.ValidarValores(Produto, "imposto/PIS//vBC")
            .ALIQ_PIS = fnXML.ValidarPercentual(Produto, "imposto/PIS//pPIS")
            .QUANT_BC_PIS = fnXML.ValidarValores(Produto, "imposto/PIS//qBCProd")
            .ALIQ_PIS_QUANT = fnXML.ValidarValores(Produto, "imposto/PIS//vAliqProd")
            .VL_PIS = fnXML.ValidarValores(Produto, "imposto/PIS//vPIS")
            .CST_COFINS = EnumContribuicoes.ValidarEnumeracao_CST_PIS_COFINS(ValidarTag(Produto, "imposto/PIS//CST"))
            .VL_BC_COFINS = fnXML.ValidarValores(Produto, "imposto/COFINS//vBC")
            .ALIQ_COFINS = fnXML.ValidarPercentual(Produto, "imposto/COFINS//pCOFINS")
            .QUANT_BC_COFINS = fnXML.ValidarValores(Produto, "imposto/COFINS//qBCProd")
            .ALIQ_COFINS_QUANT = fnXML.ValidarValores(Produto, "imposto/COFINS//vAliqProd")
            .VL_COFINS = fnXML.ValidarValores(Produto, "imposto/COFINS//vCOFINS")
            .COD_CTA = ""
            .VL_ABAT_NT = fnXML.ValidarValores(Produto, "imposto/ICMS//vICMSDeson")
            
            If .VL_ICMS = 0 Then .VL_BC_ICMS = 0: .ALIQ_ICMS = 0
            If .QUANT_BC_PIS = 0 Then .QUANT_BC_PIS = ""
            If .ALIQ_PIS_QUANT = 0 Then .ALIQ_PIS_QUANT = ""
            If .QUANT_BC_COFINS = 0 Then .QUANT_BC_COFINS = ""
            If .ALIQ_COFINS_QUANT = 0 Then .ALIQ_COFINS_QUANT = ""
            
            .CHV_PAI_CONTRIBUICOES = CamposC100.CHV_REG
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI_FISCAL, .NUM_ITEM)
            
            Call IncluirRegistrosCadastrais(Produto)
            
            Campos = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, CInt(.NUM_ITEM), _
                "'" & .COD_ITEM, .DESCR_COMPL, CDbl(.QTD), "'" & .UNID, CDbl(.VL_ITEM), CDbl(.VL_DESC), _
                .IND_MOV, "'" & .CST_ICMS, CInt(.CFOP), .COD_NAT, CDbl(.VL_BC_ICMS), CDbl(.ALIQ_ICMS), _
                CDbl(.VL_ICMS), CDbl(.VL_BC_ICMS_ST), CDbl(.ALIQ_ST), CDbl(.VL_ICMS_ST), .IND_APUR, _
                "'" & .CST_IPI, "'" & .COD_ENQ, CDbl(.VL_BC_IPI), CDbl(.ALIQ_IPI), CDbl(.VL_IPI), _
                "'" & .CST_PIS, CDbl(.VL_BC_PIS), CDbl(.ALIQ_PIS), .QUANT_BC_PIS, .ALIQ_PIS_QUANT, _
                CDbl(.VL_PIS), "'" & .CST_COFINS, CDbl(.VL_BC_COFINS), CDbl(.ALIQ_COFINS), _
                .QUANT_BC_COFINS, .ALIQ_COFINS_QUANT, CDbl(.VL_COFINS), .COD_CTA, CDbl(.VL_ABAT_NT), "")
                
            dtoRegSPED.rC170(.CHV_REG) = Campos
            
        End If
        
    End With
    
End Sub

Private Sub IncluirRegistrosCadastrais(ByRef Produto As IXMLDOMNode)
    
Dim Chave As String
Dim Campos As Variant
Dim FatConv As Double
Dim i As Long
    
    With CamposC170
        
        If CamposC100.IND_EMIT Like "1*" Then
                
                .CFOP = fnXML.AjustarCFOPEntrada(fnXML.ValidarCFOPEntrada(.CFOP))
                .CST_IPI = assImportacao.AjustarCST_IPI(.CST_IPI)
                
                Chave = Util.RemoverAspaSimples(Util.UnirCampos(Util.FormatarCNPJ(DadosXML.CNPJ_EMITENTE), .COD_ITEM, .UNID))
                If CadItensFornecProprios Then
                    
                    Call assImportacao.CadastrarItensTerceirosComoProprios(Produto)
                    
                ElseIf Correlacionamento.dicCorrelacoes.Exists(Chave) Then
                    
                    Call TratarCorrelacionamentos(Produto, Chave)
                    
                Else
                    
                    .COD_ITEM = "SEM CORRELAÇÃO"
                    
                End If
                
            Else
                
                Call assImportacao.CriarRegistro0190(.UNID)
                
                Campos0200.COD_ITEM = .COD_ITEM
                Campos0200.DESCR_ITEM = .DESCR_COMPL
                Campos0200.UNID_INV = .UNID
                
                Call assImportacao.CriarRegistro0200(Produto)
                
        End If
        
    End With
    
End Sub

Private Function TratarCorrelacionamentos(ByVal Produto As IXMLDOMNode, ByVal Chave As String)

Dim i As Long
Dim FatConv As Double
Dim UNID_COM As String
    
    With Correlacionamento
        
        Campos = .dicCorrelacoes(Chave)
        If LBound(Campos) = 0 Then i = 1 Else i = 0
        
        FatConv = Util.ValidarValores(Campos(.dicTitulosCorrelacoes("FAT_CONV") - i))
        UNID_COM = Campos(.dicTitulosCorrelacoes("UND_FORNEC") - i)
        
        Campos0200.COD_ITEM = Campos(.dicTitulosCorrelacoes("COD_ITEM") - i)
        Campos0200.DESCR_ITEM = VBA.UCase(Campos(.dicTitulosCorrelacoes("DESCR_ITEM") - i))
        Campos0200.UNID_INV = VBA.UCase(Campos(.dicTitulosCorrelacoes("UND_INV") - i))
        
        CamposC170.COD_ITEM = Campos0200.COD_ITEM
        CamposC170.DESCR_COMPL = Campos0200.DESCR_ITEM
        
        Call assImportacao.CriarRegistro0190(UNID_COM)
        Call assImportacao.CriarRegistro0190(Campos0200.UNID_INV)
        Call assImportacao.CriarRegistro0200(Produto)
        Call assImportacao.CriarRegistro0220(UNID_COM, FatConv)
        
    End With
    
End Function

Public Sub CriarRegistroC101(ByRef NFe As DOMDocument60)

Dim Campos As Variant
Dim total As Double
    
    If dtoRegSPED.rC101 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC101("CHV_PAI_FISCAL")
        
    With CamposC101
        
        If Not dtoRegSPED.rC101.Exists(CamposC100.CHV_REG) Then
            
            .REG = "C101"
            .ARQUIVO = DadosXML.ARQUIVO
            .VL_FCP_UF_DEST = fnXML.ValidarValores(NFe, "//ICMSTot/vFCPUFDest")
            .VL_ICMS_UF_DEST = fnXML.ValidarValores(NFe, "//ICMSTot/vICMSUFDest")
            .VL_ICMS_UF_REM = fnXML.ValidarValores(NFe, "//ICMSTot/vICMSUFRemet")
            .CHV_PAI_FISCAL = CamposC100.CHV_REG
            .CHV_PAI_CONTRIBUICOES = ""
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, "C101")
            
            Campos = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, _
                CDbl(.VL_FCP_UF_DEST), CDbl(.VL_ICMS_UF_DEST), CDbl(.VL_ICMS_UF_REM))
            
            total = CDbl(.VL_FCP_UF_DEST) + CDbl(.VL_ICMS_UF_DEST) + CDbl(.VL_ICMS_UF_REM)
            If total > 0 Then dtoRegSPED.rC101(CamposC100.CHV_REG) = Campos
            
        End If
        
    End With
    
End Sub

Public Sub CriarRegistroC110(ByVal NFe As DOMDocument60)

Dim Chave As String
Dim Campos As Variant
Dim NotaRef As IXMLDOMNode
Dim NotasRef As IXMLDOMNodeList
    
    If dtoRegSPED.rC110 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC110("CHV_PAI_FISCAL", "COD_INF")
    
    With CamposC110
        
        .COD_INF = "INF_NF"
        .CHV_PAI_FISCAL = CamposC100.CHV_REG
        Chave = Util.UnirCampos(.CHV_PAI_FISCAL, .COD_INF)
        
        If Not dtoRegSPED.rC110.Exists(Chave) Then
            
            .REG = "C110"
            .ARQUIVO = DadosXML.ARQUIVO
            .TXT_COMPL = fnXML.ValidarTag(NFe, "//infAdFisco")
            .CHV_PAI_CONTRIBUICOES = CamposC100.CHV_REG
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI_FISCAL, .COD_INF)
            
            Campos = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, .COD_INF, .TXT_COMPL)
            If .COD_INF <> "" Then dtoRegSPED.rC110(.CHV_REG) = Campos
            
            Set NotasRef = NFe.SelectNodes("//NFref")
            For Each NotaRef In NotasRef
                
                Call CriarRegistroC113(NotaRef)
                
            Next NotaRef
            
        End If
        
    End With
    
End Sub

Public Sub CriarRegistroC113(ByRef NotaRef As IXMLDOMNode)

Dim Chave As String, CNPJ_EMIT$
Dim Campos As Variant
    
    If dtoRegSPED.rC113 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC113("CHV_PAI_FISCAL", "CHV_DOCE")
    
     With CamposC113
         
        .CHV_DOCE = fnXML.ValidarTag(NotaRef, "refNFe")
        CNPJ_EMIT = VBA.Mid(.CHV_DOCE, 7, 14)
        Chave = Util.UnirCampos(.CHV_DOCE)
        
        If Not dtoRegSPED.rC113.Exists(Chave) Then
            
            .REG = "C113"
            .ARQUIVO = DadosXML.ARQUIVO
            .IND_EMIT = EnumFiscal.ValidarEnumeracao_IND_EMIT(assImportacao.IdentificarTipoEmissao_C113(VBA.Mid(.CHV_DOCE, 35, 1), CNPJ_EMIT))
            .IND_OPER = EnumFiscal.ValidarEnumeracao_IND_OPER(assImportacao.IdentificarTipoOperacao_C113(VBA.Left(.IND_EMIT, 1), CNPJ_EMIT))
            .COD_PART = CNPJ_EMIT
            .COD_MOD = VBA.Mid(.CHV_DOCE, 21, 2)
            .SER = VBA.Mid(.CHV_DOCE, 23, 3)
            .SUB = ""
            .NUM_DOC = VBA.Mid(.CHV_DOCE, 26, 9)
            .DT_DOC = "20" & VBA.Mid(.CHV_DOCE, 3, 2) & "-" & VBA.Mid(.CHV_DOCE, 5, 2) & "-01"
            .CHV_PAI_FISCAL = ""
            .CHV_PAI_CONTRIBUICOES = CamposC100.CHV_REG
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI_FISCAL, .IND_EMIT, .COD_PART, .COD_MOD, .SER, .NUM_DOC)
            
            Campos = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, .IND_OPER, _
                .IND_EMIT, .COD_PART, .COD_MOD, .SER, .SUB, .NUM_DOC, .DT_DOC, "'" & .CHV_DOCE)
                
            dtoRegSPED.rC113(Chave) = Campos
            
        End If
        
    End With
    
End Sub

Public Sub CriarRegistroC120(ByRef Produto As IXMLDOMNode)
    
Dim Chave As String
Dim Campos As Variant
    
    If dtoRegSPED.rC120 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC120("CHV_PAI_FISCAL", "NUM_DOC_IMP", "NUM_ACDRAW")
    
    With CamposC120
        
        .CHV_PAI_FISCAL = CamposC100.CHV_REG
        .NUM_DOC_IMP = ValidarTag(Produto, "prod/DI/nDI")
        .NUM_ACDRAW = Util.FormatarTexto(ValidarTag(Produto, "prod//nDraw"))
        Chave = Util.UnirCampos(.CHV_PAI_FISCAL, .NUM_DOC_IMP, .NUM_ACDRAW)
        
        .REG = "C120"
        .ARQUIVO = DadosXML.ARQUIVO
        .COD_DOC_IMP = ""
        .PIS_IMP = fnXML.ValidarValores(Produto, "imposto/PIS//vPIS")
        .COFINS_IMP = fnXML.ValidarValores(Produto, "imposto/COFINS//vCOFINS")
        .CHV_PAI_CONTRIBUICOES = CamposC100.CHV_REG
        
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI_FISCAL, .NUM_DOC_IMP, .NUM_ACDRAW)
        
        Campos = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, _
            .COD_DOC_IMP, "'" & .NUM_DOC_IMP, CDbl(.PIS_IMP), CDbl(.COFINS_IMP), .NUM_ACDRAW)
            
        dtoRegSPED.rC120(Chave) = Campos
        
    End With
    
Exit Sub
Tratar:

End Sub

Public Sub CriarRegistroC140(ByRef NFe As DOMDocument60)

Dim Campos As Variant
Dim Produto As IXMLDOMNode
Dim Duplicatas As IXMLDOMNodeList
    
    If dtoRegSPED.rC140 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC140("CHV_PAI_FISCAL")
    
    With CamposC140
        
        .CHV_REG = fnSPED.GerarChaveRegistro(CStr(CamposC100.CHV_REG), "C140")
        If Not dtoRegSPED.rC140.Exists(.CHV_REG) Then
            
            Set Duplicatas = NFe.SelectNodes("//dup")
            
            .REG = "C140"
            .ARQUIVO = DadosXML.ARQUIVO
            .IND_EMIT = EnumFiscal.ValidarEnumeracao_IND_EMIT(VBA.Left(CamposC100.IND_EMIT, 1))
            .IND_TIT = EnumFiscal.ValidarEnumeracao_IND_TIT(fnXML.ConverterTipoPagamento(fnXML.ValidarTag(NFe, "//pag/detPag/tPag")))
            .DESC_TIT = fnXML.ValidarDescricaoTitulo(.IND_TIT)
            .NUM_TIT = fnXML.ValidarTag(NFe, "//fat/nFat")
            .QTD_PARC = ""
            .VL_TIT = fnXML.ValidarValores(NFe, "//cobr/fat/vLiq")
            .CHV_PAI_FISCAL = CamposC100.CHV_REG
            .CHV_PAI_CONTRIBUICOES = ""
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI_FISCAL, .REG)
            
            If .VL_TIT = 0 Then .VL_TIT = fnXML.ValidarValores(NFe, "//pag/detPag/vPag")
            If Duplicatas.Length > 0 Then
                
                .QTD_PARC = Duplicatas.Length
                Call CriarRegistroC141(Duplicatas)
                
            End If
            
            Campos = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, _
                .IND_EMIT, .IND_TIT, .DESC_TIT, .NUM_TIT, .QTD_PARC, CDbl(.VL_TIT))
                
            If .VL_TIT > 0 Then dtoRegSPED.rC140(.CHV_REG) = Campos
            
        End If
        
    End With
    
End Sub

Private Sub CriarRegistroC141(ByVal Duplicatas As IXMLDOMNodeList)

Dim Duplicata As IXMLDOMNode
Dim Campos As Variant
Dim Chave As String
Dim i As Byte
    
    If dtoRegSPED.rC141 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC141("CHV_PAI_FISCAL", "NUM_PARC")
    
    i = 0
    For Each Duplicata In Duplicatas
        
        i = i + 1
        With CamposC141
            
            .REG = "C141"
            .ARQUIVO = DadosXML.ARQUIVO
            .NUM_PARC = VBA.Format(i, "00")
            .DT_VCTO = Util.FormatarData(fnXML.ValidarTag(Duplicata, "dVenc"))
            .VL_PARC = fnXML.ValidarValores(Duplicata, "vDup")
            .CHV_PAI_FISCAL = CamposC140.CHV_REG
            .CHV_PAI_CONTRIBUICOES = ""
            
            .CHV_REG = fnSPED.GerarChaveRegistro(CStr(CamposC140.CHV_REG), CStr(.NUM_PARC))
            
            Campos = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, .NUM_PARC, .DT_VCTO, CDbl(.VL_PARC))
            
            dtoRegSPED.rC141(.CHV_REG) = Campos
            
        End With
    
    Next Duplicata
    
End Sub
