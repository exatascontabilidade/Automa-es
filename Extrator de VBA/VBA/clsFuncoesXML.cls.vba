Attribute VB_Name = "clsFuncoesXML"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private EnumFiscal As New clsEnumeracoesSPEDFiscal
Private EnumContrib As New clsEnumeracoesSPEDContribuicoes
Private ValidacoesGerais As New clsRegrasFiscaisGerais

Public Sub CriarRegistro0140(ByVal NFe As IXMLDOMNode, ByRef dicDados As Dictionary, ByRef dicTitulos As Dictionary, _
    ByVal ARQUIVO As String, ByVal CHV_PAI As String, ByRef tpCont As String)
    
    With Campos0140
        
        .CNPJ = ValidarTag(NFe, "//" & tpCont & "/CNPJ")
        
        .CHV_REG = fnSPED.GerarChaveRegistro(CHV_PAI, .CNPJ)
        If Not dicDados.Exists(.CHV_REG) And .CHV_REG <> "" Then
            
            .REG = "'0140"
            .COD_EST = ""
            .NOME = ValidarTag(NFe, "//" & tpCont & "/xNome")
            .UF = ValidarTag(NFe, "//" & tpCont & "//UF")
            .IE = ValidarTag(NFe, "//" & tpCont & "/IE")
            .COD_MUN = ValidarTag(NFe, "//" & tpCont & "//cMun")
            .IM = ""
            .SUFRAMA = ValidarTag(NFe, "//" & tpCont & "/SUFRAMA")
            
            dicDados(.CNPJ) = Array(.REG, ARQUIVO, .CHV_REG, "", CHV_PAI, .COD_EST, .NOME, "'" & .CNPJ, .UF, "'" & .IE, .COD_MUN, .IM, .SUFRAMA)
            
        End If
        
    End With
    
End Sub

Public Sub CriarRegistroC010(ByVal NFe As IXMLDOMNode, ByRef dicDados As Dictionary, ByRef dicTitulos As Dictionary, _
    ByVal ARQUIVO As String, ByVal CHV_PAI As String, ByRef tpCont As String)

    With CamposC010
        
        .CNPJ = ValidarTag(NFe, "//" & tpCont & "/CNPJ")
        
        .CHV_REG = fnSPED.GerarChaveRegistro(CHV_PAI, .CNPJ)
        If Not dicDados.Exists(.CHV_REG) And .CHV_REG <> "" Then
            
            .REG = "'C010"
            .IND_ESCRI = 2
            
            dicDados(.CNPJ) = Array(.REG, ARQUIVO, .CHV_REG, "", CHV_PAI, "'" & .CNPJ, .IND_ESCRI)
            
        End If
        
    End With

End Sub

Public Sub CriarRegistroD010(ByVal NFe As IXMLDOMNode, ByRef dicDados As Dictionary, ByRef dicTitulos As Dictionary, _
    ByVal ARQUIVO As String, ByVal CHV_PAI As String, ByRef tpCont As String)

    With CamposD010
        
        .CNPJ = ValidarTag(NFe, "//" & tpCont & "/CNPJ")
        
        .CHV_REG = fnSPED.GerarChaveRegistro(CHV_PAI, .CNPJ)
        If Not dicDados.Exists(.CHV_REG) And .CHV_REG <> "" Then
            
            .REG = "'D010"
            dicDados(.CNPJ) = Array(.REG, ARQUIVO, .CHV_REG, "", CHV_PAI, "'" & .CNPJ)
            
        End If
        
    End With

End Sub

Public Sub CriarRegistro0150(ByVal NFe As IXMLDOMNode, ByRef dicDados As Dictionary, ByRef dicTitulos As Dictionary, _
    ByVal CHV_REG As String, ByVal ARQUIVO As String, ByVal CHV_PAI As String, ByRef tpPart As String)
    
    If Not dicDados.Exists(CHV_REG) And CHV_REG <> "" Then
        
        With Campos0150
            
            .REG = "'0150"
            .COD_PART = CHV_REG
            .NOME = VBA.Trim(VBA.Left(ValidarTag(NFe, "//" & tpPart & "/xNome"), 100))
            .COD_PAIS = fnExcel.FormatarTexto(ValidarTag(NFe, "//" & tpPart & "//cPais"))
            .CNPJ = fnExcel.FormatarTexto(ValidarTag(NFe, "//" & tpPart & "/CNPJ"))
            .CPF = fnExcel.FormatarTexto(ValidarTag(NFe, "//" & tpPart & "/CPF"))
            .IE = fnExcel.FormatarTexto(ValidarTag(NFe, "//" & tpPart & "/IE"))
            .COD_MUN = fnExcel.FormatarTexto(ValidarTag(NFe, "//" & tpPart & "//cMun"))
            .SUFRAMA = fnExcel.FormatarTexto(ValidarTag(NFe, "//" & tpPart & "/SUFRAMA"))
            .END = VBA.Trim(VBA.Left(ValidarTag(NFe, "//" & tpPart & "//xLgr"), 60))
            .NUM = VBA.Trim(VBA.Left(fnExcel.FormatarTexto(ValidarTag(NFe, "//" & tpPart & "//nro")), 10))
            .COMPL = VBA.Trim(VBA.Left(ValidarTag(NFe, "//" & tpPart & "//xCpl"), 60))
            .BAIRRO = VBA.Trim(VBA.Left(ValidarTag(NFe, "//" & tpPart & "//xBairro"), 60))
            If .COD_PAIS = "'" Then .COD_PAIS = "1058"
            
            .CHV_REG = fnSPED.GerarChaveRegistro(CHV_PAI, .COD_PART)
            dicDados(.CHV_REG) = Array(.REG, ARQUIVO, .CHV_REG, CHV_PAI, "", Util.FormatarTexto(.COD_PART), .NOME, _
                                    .COD_PAIS, .CNPJ, .CPF, .IE, .COD_MUN, .SUFRAMA, .END, .NUM, .COMPL, .BAIRRO)
            
        End With
        
    End If

End Sub

Public Sub CriarRegistro0190(ByVal Produtos As IXMLDOMNodeList, ByRef dicDados As Dictionary, _
    ByRef dicTitulos As Dictionary, ByVal ARQUIVO As String, ByVal CHV_PAI As String)
    
Dim CHV_REG As String
Dim Produto As IXMLDOMNode
    
    With Campos0190
        
        For Each Produto In Produtos
            
            .UNID = VBA.UCase(ValidarTag(Produto, "prod/uCom"))
            CHV_REG = fnSPED.GerarChaveRegistro(CHV_PAI, .UNID)
            If Not dicDados.Exists(CHV_REG) Then
                
                .REG = "'0190"
                .DESCR = VBA.UCase(.UNID)
                
                dicDados(CHV_REG) = Array(.REG, ARQUIVO, CHV_REG, CHV_PAI, "", .UNID, .DESCR)
                
            End If
            
        Next Produto
        
    End With

End Sub

Public Sub IncluirRegistro0190(ByRef dicDados As Dictionary, ByVal UNID As String, ByVal ARQUIVO As String, ByVal CHV_PAI As String)
    
    With Campos0190
        
        .CHV_REG = fnSPED.GerarChaveRegistro(CHV_PAI, UNID)
        If Not dicDados.Exists(.CHV_REG) Then
            
            .REG = "'0190"
            .DESCR = VBA.UCase(UNID)
            
            dicDados(.CHV_REG) = Array(.REG, ARQUIVO, .CHV_REG, CHV_PAI, "", UNID, .DESCR)
            
        End If

    End With

End Sub

Public Sub CriarRegistro0200(ByVal Produtos As IXMLDOMNodeList, ByRef dicDados As Dictionary, _
    ByRef dicTitulos As Dictionary, ByVal ARQUIVO As String, ByVal CHV_PAI As String)

Dim Produto As IXMLDOMNode
Dim Campos As Variant
    
    With Campos0200
        
        For Each Produto In Produtos
            
            .COD_ITEM = ValidarTag(Produto, "prod/cProd")
            .CHV_REG = fnSPED.GerarChaveRegistro(CHV_PAI, .COD_ITEM)
            If Not dicDados.Exists(.CHV_REG) Then
                
                .REG = "'0200"
                .DESCR_ITEM = Util.RemoverPipes(ValidarTag(Produto, "prod/xProd"))
                .COD_BARRA = fnExcel.FormatarTexto(Util.FormatarTexto(fnXML.ExtrairCodigoBarrasProduto(Produto)))
                .COD_ANT_ITEM = ""
                .UNID_INV = VBA.UCase(ValidarTag(Produto, "prod/uCom"))
                .TIPO_ITEM = ""
                .COD_NCM = fnExcel.FormatarTexto(VBA.Format(ValidarTag(Produto, "prod/NCM"), String(8, "0")))
                .EX_IPI = fnExcel.FormatarTexto(ValidarTag(Produto, "prod/EXTIPI"))
                .COD_GEN = fnExcel.FormatarTexto(VBA.Left(Util.ApenasNumeros(.COD_NCM), 2))
                .COD_LST = ""
                .ALIQ_ICMS = ValidarPercentual(Produto, "imposto/ICMS//pICMS")
                .CEST = ExtrairCEST(Produto)
                
                If .CEST = "0" Then .CEST = ""
                
                dicDados(.CHV_REG) = Array(.REG, ARQUIVO, .CHV_REG, CHV_PAI, "", "'" & .COD_ITEM, .DESCR_ITEM, .COD_BARRA, .COD_ANT_ITEM, _
                    .UNID_INV, .TIPO_ITEM, .COD_NCM, .EX_IPI, .COD_GEN, .COD_LST, CDbl(.ALIQ_ICMS), fnExcel.FormatarTexto(.CEST))
                
            End If
            
        Next Produto
        
    End With
    
End Sub

Public Sub IncluirRegistro0200(ByVal Produto As IXMLDOMNode, ByRef dicDados As Dictionary, _
    ByVal COD_ITEM As String, ByVal DESCR_ITEM As String, ByVal UNID_INV As String, ByVal ARQUIVO As String, ByVal CHV_PAI As String)

Dim Campos As Variant
    
    With Campos0200
        
        .COD_ITEM = ValidarTag(Produto, "prod/cProd")
        If COD_ITEM <> "" Then .COD_ITEM = COD_ITEM
        .CHV_REG = fnSPED.GerarChaveRegistro(CHV_PAI, .COD_ITEM)
        If Not dicDados.Exists(.CHV_REG) Then
            
            .REG = "'0200"
            .DESCR_ITEM = Util.RemoverPipes(ValidarTag(Produto, "prod/xProd"))
            .COD_BARRA = fnExcel.FormatarTexto(Util.FormatarTexto(fnXML.ExtrairCodigoBarrasProduto(Produto)))
            .COD_ANT_ITEM = ""
            .UNID_INV = ValidarTag(Produto, "prod/uCom")
            .TIPO_ITEM = ""
            .COD_NCM = fnExcel.FormatarTexto(VBA.Format(ValidarTag(Produto, "prod/NCM"), String(8, "0")))
            .EX_IPI = fnExcel.FormatarTexto(ValidarTag(Produto, "prod/EXTIPI"))
            .COD_GEN = fnExcel.FormatarTexto(VBA.Left(Util.ApenasNumeros(.COD_NCM), 2))
            .COD_LST = ""
            .ALIQ_ICMS = ValidarPercentual(Produto, "imposto/ICMS//pICMS")
            .CEST = ExtrairCEST(Produto)
            
            If .CEST = "0" Then .CEST = ""
            If CamposC100.COD_MOD = "65" Then .TIPO_ITEM = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_TIPO_ITEM("00")
            
            If DESCR_ITEM <> "" Then .DESCR_ITEM = DESCR_ITEM
            If UNID_INV <> "" Then .UNID_INV = UNID_INV
            If .COD_BARRA = "'SEM GTIN" Then .COD_BARRA = ""
            
            dicDados(.CHV_REG) = Array(.REG, ARQUIVO, .CHV_REG, CHV_PAI, "", "'" & .COD_ITEM, .DESCR_ITEM, .COD_BARRA, .COD_ANT_ITEM, _
                .UNID_INV, .TIPO_ITEM, .COD_NCM, .EX_IPI, .COD_GEN, .COD_LST, CDbl(.ALIQ_ICMS), fnExcel.FormatarTexto(.CEST))
            
        End If
        
    End With
        
End Sub

Public Sub CriarRegistro0450(ByRef Produtos As IXMLDOMNodeList, ByRef dicDados As Dictionary, ByVal ARQUIVO As String)

Dim Produto As IXMLDOMNode
    
    On Error GoTo Tratar:
    
    With Campos0450
        
        .REG = "0450"
        .COD_INF = "OBC110"
        .TXT = "Observação do registro C110"
        .CHV_PAI = Campos0001.CHV_REG
        
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, "0450", dicDados.Count)
        
        If .COD_INF <> "" Then dicDados(.CHV_REG) = Array(.REG, ARQUIVO, .CHV_REG, .CHV_PAI, "", .COD_INF, .TXT)
        
    End With
    
Exit Sub
Tratar:

End Sub

Public Sub CriarRegistroC100(ByRef NFe As IXMLDOMNode, ByRef dicDadosC100 As Dictionary, _
    ByRef arrCanceladas As ArrayList, ByVal CHV_PAI As String, Optional ByVal SPEDContrib As Boolean)
    
Dim Campos As Variant
    
    With CamposC100
        
        .REG = "C100"
        .IND_OPER = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_OPER(fnXML.IdentificarTipoOperacao(NFe, ValidarTag(NFe, "//tpNF"), SPEDContrib))
        .IND_EMIT = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_EMIT(fnXML.IdentificarTipoEmissao(NFe, SPEDContrib))
        .COD_MOD = ValidarTag(NFe, "//mod")
        .COD_PART = fnXML.IdentificarParticipante(NFe, VBA.Left(.IND_OPER, 1), VBA.Left(.IND_EMIT, 1))
        .COD_SIT = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_COD_SIT(fnSPED.GerarCodigoSituacao(fnXML.ValidarSituacao(ValidarTag(NFe, "//cStat"))))
        .SER = VBA.Format(ValidarTag(NFe, "//serie"), "000")
        .NUM_DOC = ValidarTag(NFe, "//nNF")
        .CHV_NFE = VBA.Right(ValidarTag(NFe, "//@Id"), 44)
        .DT_DOC = ExtrairDataDocumento(NFe)
        .DT_E_S = ExtrairDataEntradaSaida(NFe)
        .VL_DOC = ValidarValores(NFe, "//ICMSTot/vNF")
        .IND_PGTO = ExtrairTipoPagamento(NFe, .DT_DOC) 'ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_PGTO(ValidarTag(NFe, "//indPag"))
        .VL_DESC = ValidarValores(NFe, "//ICMSTot/vDesc")
        .VL_ABAT_NT = ValidarValores(NFe, "//ICMSTot/vICMSDeson")
        .VL_MERC = ValidarValores(NFe, "//ICMSTot/vProd")
        .IND_FRT = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_FRT(ValidarTag(NFe, "//modFrete"))
        .VL_FRT = ValidarValores(NFe, "//ICMSTot/vFrete")
        .VL_SEG = ValidarValores(NFe, "//ICMSTot/vSeg")
        .VL_OUT_DA = ValidarValores(NFe, "//ICMSTot/vOutro")
        .VL_BC_ICMS = ExtrairBaseICMSTotal(NFe)
        .VL_ICMS = ExtrairICMSTotal(NFe)
        .VL_BC_ICMS_ST = ValidarValores(NFe, "//ICMSTot/vBCST")
        .VL_FCP_ST = ValidarValores(NFe, "//ICMSTot/vFCPST")
        .VL_ICMS_ST = ValidarValores(NFe, "//ICMSTot/vST") + CDbl(.VL_FCP_ST)
        .VL_IPI = ValidarValores(NFe, "//ICMSTot/vIPI")
        .VL_PIS = ValidarValores(NFe, "//ICMSTot/vPIS")
        .VL_COFINS = ValidarValores(NFe, "//ICMSTot/vCOFINS")
        .VL_PIS_ST = 0
        .VL_COFINS_ST = 0
        .CHV_PAI = CHV_PAI
        
        If .COD_MOD = "65" Then .COD_PART = ""
        If .IND_PGTO = "" Then .IND_PGTO = "0 - Á Vista"
        If .DT_E_S = "" Then .DT_E_S = .DT_DOC
        
        'Verifica se a nota está cancelada
        If arrCanceladas.contains(VBA.Replace(.CHV_NFE, "'", "")) Then _
            .COD_SIT = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_COD_SIT(fnSPED.GerarCodigoSituacao(fnXML.ValidarSituacao("101")))
            
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
        
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .IND_OPER, .IND_EMIT, .COD_PART, .COD_MOD, .COD_SIT, .SER, .NUM_DOC, .CHV_NFE)
        Campos = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI, "", .IND_OPER, .IND_EMIT, Util.FormatarTexto(.COD_PART), .COD_MOD, _
                       .COD_SIT, "'" & .SER, "'" & .NUM_DOC, "'" & .CHV_NFE, .DT_DOC, .DT_E_S, CDbl(.VL_DOC), _
                       .IND_PGTO, CDbl(.VL_DESC), CDbl(.VL_ABAT_NT), CDbl(.VL_MERC), .IND_FRT, CDbl(.VL_FRT), _
                       CDbl(.VL_SEG), CDbl(.VL_OUT_DA), CDbl(.VL_BC_ICMS), CDbl(.VL_ICMS), CDbl(.VL_BC_ICMS_ST), _
                       CDbl(.VL_ICMS_ST), CDbl(.VL_IPI), CDbl(.VL_PIS), CDbl(.VL_COFINS), CDbl(.VL_PIS_ST), CDbl(.VL_COFINS_ST))
        
        dicDadosC100(.CHV_REG) = Campos
        
    End With
    
End Sub

Public Function ExtrairICMSTotal(ByRef NFe As IXMLDOMNode) As Double

Dim VL_ICMS As Double, VL_FCP#, VL_CRED_SN#
Dim Produtos As IXMLDOMNodeList
Dim Produto As IXMLDOMNode
    
    VL_ICMS = ValidarValores(NFe, "//ICMSTot/vICMS")
    VL_FCP = ValidarValores(NFe, "//ICMSTot/vFCP")
    
    If VL_ICMS + VL_FCP = 0 Then
        
        Set Produtos = NFe.SelectNodes("//det")
        For Each Produto In Produtos
            
            If ValidarTag(Produto, "imposto/ICMS//CSOSN") <> "" Then _
                VL_CRED_SN = VL_CRED_SN + ExtrairValorICMS(Produto)
            
        Next Produto
        
    End If
    
    ExtrairICMSTotal = VL_ICMS + VL_FCP + VL_CRED_SN
    
End Function

Public Function ExtrairBaseICMSTotal(ByRef NFe As IXMLDOMNode) As Double

Dim VL_BC_ICMS As Double, VL_FCP#, VL_CRED_SN#
Dim Produtos As IXMLDOMNodeList
Dim Produto As IXMLDOMNode
    
    VL_BC_ICMS = ValidarValores(NFe, "//ICMSTot/vBC")
    
    If VL_BC_ICMS = 0 Then
        
        Set Produtos = NFe.SelectNodes("//det")
        For Each Produto In Produtos
            
            If ValidarTag(Produto, "imposto/ICMS//CSOSN") <> "" Then _
                VL_BC_ICMS = VL_BC_ICMS + ExtrairBaseICMS(Produto)
            
        Next Produto
        
    End If
    
    ExtrairBaseICMSTotal = VL_BC_ICMS
    
End Function

Public Sub CriarRegistroC101(ByRef NFe As IXMLDOMNode, ByRef dicDadosC101 As Dictionary, ByVal ARQUIVO As String)

Dim Produto As IXMLDOMNode
Dim chNFe As String
Dim total As Double
    
    On Error GoTo Tratar:
    
    If Not dicDadosC101.Exists(CamposC100.CHV_REG) Then
        
        With CamposC101
            
            .REG = "C101"
            .VL_FCP_UF_DEST = fnXML.ValidarValores(NFe, "//ICMSTot/vFCPUFDest")
            .VL_ICMS_UF_DEST = fnXML.ValidarValores(NFe, "//ICMSTot/vICMSUFDest")
            .VL_ICMS_UF_REM = fnXML.ValidarValores(NFe, "//ICMSTot/vICMSUFRemet")
            .CHV_PAI = CamposC100.CHV_REG
            
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, "C101")
            total = CDbl(.VL_FCP_UF_DEST) + CDbl(.VL_ICMS_UF_DEST) + CDbl(.VL_ICMS_UF_REM)
            If total > 0 Then dicDadosC101(.CHV_REG) = Array(.REG, ARQUIVO, .CHV_REG, _
                .CHV_PAI, "", CDbl(.VL_FCP_UF_DEST), CDbl(.VL_ICMS_UF_DEST), CDbl(.VL_ICMS_UF_REM))
            
        End With
        
    End If
    
Exit Sub
Tratar:
        
End Sub

Public Sub CriarRegistroC110(ByRef Produtos As IXMLDOMNodeList, ByRef dicDados As Dictionary, ByRef dicDados0450 As Dictionary, ByVal ARQUIVO As String)

Dim Produto As IXMLDOMNode
    
    On Error GoTo Tratar:
    
    With CamposC110
        
        For Each Produto In Produtos
            
            .REG = "C110"
            .COD_INF = ""
            .TXT_COMPL = ""
            .CHV_PAI = CamposC100.CHV_REG
            
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, "C110", dicDados.Count)
            
            If .COD_INF <> "" Then dicDados(.CHV_REG) = Array(.REG, ARQUIVO, .CHV_REG, .CHV_PAI, "", .COD_INF, .TXT_COMPL)
            
        Next Produto
        
    End With
    
Exit Sub
Tratar:

End Sub

Public Sub CriarRegistroC113(ByRef Notas As IXMLDOMNodeList, ByRef dicDados As Dictionary, ByVal ARQUIVO As String)

Dim Nota As IXMLDOMNode

    On Error GoTo Tratar:
        
    With CamposC113
        
        For Each Nota In Notas
            
            .REG = "C113"
            .IND_OPER = ""
            .IND_EMIT = ""
            .COD_PART = ""
            .COD_MOD = ""
            .SER = ""
            .SUB = ""
            .NUM_DOC = ""
            .DT_DOC = ""
            .CHV_DOCE = ""
            .CHV_PAI = CamposC100.CHV_REG
            
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, "C113", dicDados.Count)
            
            If .CHV_DOCE <> "" Then dicDados(.CHV_REG) = Array(.REG, ARQUIVO, .CHV_REG, .CHV_PAI, "", _
                .IND_OPER, .IND_EMIT, .COD_PART, .COD_MOD, .SER, .SUB, .NUM_DOC, .DT_DOC, "'" & .CHV_DOCE)
            
        Next Nota
        
    End With
    
Exit Sub
Tratar:
        
End Sub

Public Sub CriarRegistroC120(ByRef Produtos As IXMLDOMNodeList, ByRef dicDadosC120 As Dictionary, ByVal ARQUIVO As String)

Dim Produto As IXMLDOMNode

    On Error GoTo Tratar:
        
    With CamposC120
        
        For Each Produto In Produtos
        
            .REG = "C120"
            .COD_DOC_IMP = ""
            .NUM_DOC_IMP = ValidarTag(Produto, "prod/DI/nDI")
            .PIS_IMP = ValidarValores(Produto, "imposto/PIS//vPIS")
            .COFINS_IMP = ValidarValores(Produto, "imposto/COFINS//vCOFINS")
            .NUM_ACDRAW = Util.FormatarTexto(ValidarTag(Produto, "prod//nDraw"))
            .CHV_PAI = CamposC100.CHV_REG
            
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, "C120", dicDadosC120.Count)
            
            If .NUM_DOC_IMP <> "" Then dicDadosC120(.CHV_REG) = Array(.REG, ARQUIVO, .CHV_REG, _
                .CHV_PAI, "", .COD_DOC_IMP, "'" & .NUM_DOC_IMP, CDbl(.PIS_IMP), CDbl(.COFINS_IMP), .NUM_ACDRAW)
        
        Next Produto
        
    End With
    
Exit Sub
Tratar:
        
End Sub

Public Sub CriarRegistroC140(ByRef NFe As IXMLDOMNode, ByRef dicDadosC140 As Dictionary, ByRef dicDadosC141 As Dictionary, ByVal ARQUIVO As String)

Dim Produto As IXMLDOMNode
Dim chNFe As String
Dim total As Double
Dim Duplicatas As IXMLDOMNodeList
    
    On Error GoTo Tratar:
    
    With CamposC140
                
        .CHV_REG = fnSPED.GerarChaveRegistro(CStr(CamposC100.CHV_REG), "C140")
        If Not dicDadosC140.Exists(.CHV_REG) Then
            
            Set Duplicatas = NFe.SelectNodes("//dup")
            
            .REG = "C140"
            .CHV_PAI = CamposC100.CHV_REG
            .IND_EMIT = CamposC100.IND_EMIT
            .IND_TIT = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_TIT(fnXML.ValidarTag(NFe, "//pag/detPag/tPag"))
            .DESC_TIT = ValidarDescricaoTitulo(.IND_TIT)
            .NUM_TIT = fnXML.ValidarTag(NFe, "//fat/nFat")
            .QTD_PARC = ""
            .VL_TIT = fnXML.ValidarValores(NFe, "//cobr/fat/vLiq")
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .REG)
            
            If .VL_TIT = 0 Then .VL_TIT = fnXML.ValidarValores(NFe, "//pag/detPag/vPag")
            If Duplicatas.Length > 0 Then
                
                .QTD_PARC = Duplicatas.Length
                Call CriarRegistroC141(Duplicatas, .CHV_REG, dicDadosC141, ARQUIVO)
                
            End If
            
            .IND_TIT = ConverterTipoPagamento(.IND_TIT)
            If .VL_TIT > 0 Then dicDadosC140(.CHV_REG) = Array(.REG, ARQUIVO, .CHV_REG, _
                .CHV_PAI, "", .IND_EMIT, .IND_TIT, .DESC_TIT, .NUM_TIT, .QTD_PARC, CDbl(.VL_TIT))
        End If
        
    End With
    
Exit Sub
Tratar:

Resume
        
End Sub

Public Sub CriarRegistroC141(ByVal Duplicatas As IXMLDOMNodeList, ByVal CHV_PAI As String, ByRef dicDadosC141 As Dictionary, ARQUIVO As String)

Dim Duplicata As IXMLDOMNode
Dim Chave As String, CNPJEmit$
Dim i As Byte

    On Error GoTo Tratar:
        
        i = 0
        For Each Duplicata In Duplicatas
            
            i = i + 1
            With CamposC141
                
                .REG = "C141"
                .CHV_PAI = CamposC140.CHV_REG
                .NUM_PARC = VBA.Format(i, "00")
                .DT_VCTO = Util.FormatarData(fnXML.ValidarTag(Duplicata, "dVenc"))
                .VL_PARC = fnXML.ValidarValores(Duplicata, "vDup")
                
                .CHV_REG = fnSPED.GerarChaveRegistro(CStr(CamposC140.CHV_REG), CStr(.NUM_PARC))
                dicDadosC141(.CHV_REG) = Array(.REG, ARQUIVO, .CHV_REG, .CHV_PAI, "", .NUM_PARC, .DT_VCTO, CDbl(.VL_PARC))
                
            End With
        
        Next Duplicata
        
Exit Sub
Tratar:

Resume
        
End Sub

Public Sub CriarRegistroC170(ByVal Produtos As IXMLDOMNodeList, ByRef VL_PIS_ST As String, ByRef VL_COFINS_ST As String, _
    ByVal CNPJEmit As String, ByVal CNPJDest As String, ByRef dicDados As Dictionary, ARQUIVO As String, _
    ByRef dicCorrelacoes As Dictionary, ByRef dicTitulosCorrelacoes As Dictionary, ByRef dicDados0000 As Dictionary, _
    ByRef dicTitulos0000 As Dictionary, ByRef dicDados0190 As Dictionary, ByRef dicDados0200 As Dictionary, _
    ByRef dicDados0220 As Dictionary, ByRef dicTitulos0220 As Dictionary, Optional ByVal SPEDContrib As Boolean)
    
Dim Chave As String, uCom$, COD_ITEM$, UND_INV$, DESCR_ITEM$, CHV_0220 As String, CHV_0000$, CHV_0001$, CHV_0200$, CHV_PAI$
Dim Produto As IXMLDOMNode
Dim FatConv As Double
Dim i As Long
    
    If VBA.Left(CamposC100.COD_SIT, 2) <> "02" Then
        
        With CamposC170
            
            For Each Produto In Produtos
                
                .NUM_ITEM = CInt(ValidarnItem(Produto))
                .CHV_PAI = CamposC100.CHV_REG
                .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .NUM_ITEM)
                If Not dicDados.Exists(.CHV_REG) Then
                    
                    .REG = "C170"
                    .COD_ITEM = ValidarTag(Produto, "prod/cProd")
                    .DESCR_COMPL = Util.RemoverPipes(ValidarTag(Produto, "prod/xProd"))
                    .QTD = ValidarValores(Produto, "prod/qCom")
                    .UNID = VBA.UCase(ValidarTag(Produto, "prod/uCom"))
                    .VL_ITEM = ValidarValores(Produto, "prod/vProd")
                    .VL_DESC = ValidarValores(Produto, "prod/vDesc")
                    .IND_MOV = "0 - SIM"
                    .CST_ICMS = ExtrairCST_CSOSN_ICMS(Produto)
                    .CFOP = ValidarValores(Produto, "prod/CFOP")
                    .COD_NAT = ""
                    .VL_BC_ICMS = ExtrairBaseICMS(Produto)
                    .ALIQ_ICMS = ExtrairAliquotaICMS(Produto)
                    .VL_ICMS = ExtrairValorICMS(Produto)
                    .VL_BC_ICMS_ST = ValidarValores(Produto, "imposto/ICMS//vBCST")
                    .ALIQ_ST = ExtrairAliquotaICMSST(Produto)
                    .VL_ICMS_ST = ExtrairValorICMSST(Produto)
                    .IND_APUR = ""
                    .CST_IPI = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_IPI(ValidarTag(Produto, "imposto/IPI//CST"))
                    .COD_ENQ = "" 'imposto/IPI//cEnq"
                    .VL_BC_IPI = ValidarValores(Produto, "imposto/IPI//vBC")
                    .ALIQ_IPI = ValidarPercentual(Produto, "imposto/IPI//pIPI")
                    .VL_IPI = ValidarValores(Produto, "imposto/IPI//vIPI")
                    .CST_PIS = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(ValidarTag(Produto, "imposto/PIS//CST"))
                    .VL_BC_PIS = ValidarValores(Produto, "imposto/PIS//vBC")
                    .ALIQ_PIS = ValidarPercentual(Produto, "imposto/PIS//pPIS")
                    .QUANT_BC_PIS = ValidarValores(Produto, "imposto/PIS//qBCProd")
                    .ALIQ_PIS_QUANT = ValidarValores(Produto, "imposto/PIS//vAliqProd")
                    .VL_PIS = ValidarValores(Produto, "imposto/PIS//vPIS")
                    .CST_COFINS = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(ValidarTag(Produto, "imposto/PIS//CST"))
                    .VL_BC_COFINS = ValidarValores(Produto, "imposto/COFINS//vBC")
                    .ALIQ_COFINS = ValidarPercentual(Produto, "imposto/COFINS//pCOFINS")
                    .QUANT_BC_COFINS = ValidarValores(Produto, "imposto/COFINS//qBCProd")
                    .ALIQ_COFINS_QUANT = ValidarValores(Produto, "imposto/COFINS//vAliqProd")
                    .VL_COFINS = ValidarValores(Produto, "imposto/COFINS//vCOFINS")
                    .COD_CTA = ""
                    .VL_ABAT_NT = ValidarValores(Produto, "imposto/ICMS//vICMSDeson")
                    
                    If .VL_ICMS = 0 Then .VL_BC_ICMS = 0: .ALIQ_ICMS = 0
                    
                    VL_PIS_ST = ValidarValores(Produto, "imposto/PIS/PISST/vPIS") + Util.FormatarValores(VL_PIS_ST)
                    VL_COFINS_ST = ValidarValores(Produto, "imposto/COFINS/COFINSST/vCOFINS") + Util.FormatarValores(VL_COFINS_ST)
                    
                    If .QUANT_BC_PIS = 0 Then .QUANT_BC_PIS = ""
                    If .ALIQ_PIS_QUANT = 0 Then .ALIQ_PIS_QUANT = ""
                    If .QUANT_BC_COFINS = 0 Then .QUANT_BC_COFINS = ""
                    If .ALIQ_COFINS_QUANT = 0 Then .ALIQ_COFINS_QUANT = ""
                    
                    'Verifica se a nota fiscal é de emissão de terceiros
                    If CamposC100.IND_EMIT Like "1*" Then
                        
                        'Cria chave do produto para buscar as informações de correlação
                        Chave = Util.FormatarCNPJ(CNPJEmit) & .COD_ITEM & .UNID
                        Chave = VBA.Replace(Chave, "'", "")
                        
                        'Verifica se o usuário selecionou a opção de cadastrar itens do fornecedor como próprios
                        If CadItensFornecProprios Then
                            
                            .COD_ITEM = CNPJEmit & " - " & .COD_ITEM
                            
                            If dicDados0000.Exists(ARQUIVO) Then
                                If LBound(dicDados0000(ARQUIVO)) = 0 Then i = 1 Else i = 0
                                CHV_0000 = Util.RemoverAspaSimples(dicDados0000(ARQUIVO)(dicTitulos0000("CHV_REG") - i))
                                CHV_0001 = fnSPED.GerarChaveRegistro(CHV_0000, "0001")
                                CHV_PAI = Util.SelecionarChaveSPED(CHV_0001, CNPJBase, CNPJEmit, CNPJDest, SPEDContrib)
                            End If
                            
                            Call IncluirRegistro0190(dicDados0190, .UNID, ARQUIVO, CHV_PAI)
                            Call IncluirRegistro0200(Produto, dicDados0200, .COD_ITEM, .DESCR_COMPL, .UNID, ARQUIVO, CHV_PAI)
                            
                        'Verifica se existe uma correlação cadastrada para o produto atual
                        ElseIf dicCorrelacoes.Exists(Chave) Then
                                                    
                            If LBound(dicCorrelacoes(Chave)) = 0 Then i = 1 Else i = 0
                            .COD_ITEM = dicCorrelacoes(Chave)(dicTitulosCorrelacoes("COD_ITEM") - i)
                            UND_INV = VBA.UCase(dicCorrelacoes(Chave)(dicTitulosCorrelacoes("UND_INV") - i))
                            DESCR_ITEM = VBA.UCase(dicCorrelacoes(Chave)(dicTitulosCorrelacoes("DESCR_ITEM") - i))
                            COD_ITEM = VBA.UCase(dicCorrelacoes(Chave)(dicTitulosCorrelacoes("COD_ITEM") - i))
                            
                            If dicDados0000.Exists(ARQUIVO) Then
                                If LBound(dicDados0000(ARQUIVO)) = 0 Then i = 1 Else i = 0
                                CHV_0000 = Util.RemoverAspaSimples(dicDados0000(ARQUIVO)(dicTitulos0000("CHV_REG") - i))
                                CHV_0001 = fnSPED.GerarChaveRegistro(CHV_0000, "0001")
                                CHV_PAI = Util.SelecionarChaveSPED(CHV_0001, CNPJBase, CNPJEmit, CNPJDest, SPEDContrib)
                                Call IncluirRegistro0190(dicDados0190, .UNID, ARQUIVO, CHV_0001)
                                Call IncluirRegistro0200(Produto, dicDados0200, COD_ITEM, DESCR_ITEM, UND_INV, ARQUIVO, CHV_PAI)
                                
                                CHV_0200 = fnSPED.GerarChaveRegistro(CHV_0001, .COD_ITEM)
                            End If
                            
                            If LBound(dicCorrelacoes(Chave)) = 0 Then i = 1 Else i = 0
                            FatConv = Util.ValidarValores(dicCorrelacoes(Chave)(dicTitulosCorrelacoes("FAT_CONV") - i))
                            uCom = dicCorrelacoes(Chave)(dicTitulosCorrelacoes("UND_FORNEC") - i)
                            
                            CHV_0220 = fnSPED.GerarChaveRegistro(CHV_0200, uCom)
                            If Not dicDados0220.Exists(CHV_0220) And FatConv > 0 Then
                                dicDados0220(CHV_0220) = Array("'0220", ARQUIVO, CHV_0220, CHV_0200, uCom, FatConv, "")
                                FatConv = 0
                            End If
                            
                        Else
                            
                            .COD_ITEM = "SEM CORRELAÇÃO"
                            
                        End If
                        
                        .CFOP = AjustarCFOPEntrada(ValidarCFOPEntrada(.CFOP))
                        
                        Select Case True
                            
                            Case Util.ApenasNumeros(.CST_IPI) = "99"
                                .CST_IPI = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_IPI(49)
                                
                            Case .CST_IPI Like "5*"
                                .CST_IPI = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_IPI("0" & VBA.Right(Util.ApenasNumeros(.CST_IPI), 1))
                                
                        End Select
                        
                    Else
                    
                        If dicDados0000.Exists(ARQUIVO) Then
                            If LBound(dicDados0000(ARQUIVO)) = 0 Then i = 1 Else i = 0
                            CHV_0000 = Util.RemoverAspaSimples(dicDados0000(ARQUIVO)(dicTitulos0000("CHV_REG") - i))
                            CHV_0001 = fnSPED.GerarChaveRegistro(CHV_0000, "0001")
                            CHV_PAI = Util.SelecionarChaveSPED(CHV_0001, CNPJBase, CNPJEmit, CNPJDest, SPEDContrib)
                            
                            Call IncluirRegistro0190(dicDados0190, .UNID, ARQUIVO, CHV_PAI)
                            Call IncluirRegistro0190(dicDados0190, UND_INV, ARQUIVO, CHV_PAI)
                            Call IncluirRegistro0200(Produto, dicDados0200, COD_ITEM, DESCR_ITEM, UND_INV, ARQUIVO, CHV_PAI)
                        End If
                    
                    End If
                                        
                    .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .NUM_ITEM)
                    dicDados(.CHV_REG) = Array(.REG, ARQUIVO, .CHV_REG, .CHV_PAI, .CHV_PAI, "'" & .NUM_ITEM, "'" & .COD_ITEM, .DESCR_COMPL, _
                        CDbl(.QTD), "'" & .UNID, CDbl(.VL_ITEM), CDbl(.VL_DESC), .IND_MOV, "'" & .CST_ICMS, CInt(.CFOP), .COD_NAT, _
                        CDbl(.VL_BC_ICMS), CDbl(.ALIQ_ICMS), CDbl(.VL_ICMS), CDbl(.VL_BC_ICMS_ST), CDbl(.ALIQ_ST), CDbl(.VL_ICMS_ST), _
                        .IND_APUR, "'" & .CST_IPI, "'" & .COD_ENQ, CDbl(.VL_BC_IPI), CDbl(.ALIQ_IPI), CDbl(.VL_IPI), "'" & .CST_PIS, _
                        CDbl(.VL_BC_PIS), CDbl(.ALIQ_PIS), .QUANT_BC_PIS, .ALIQ_PIS_QUANT, CDbl(.VL_PIS), "'" & .CST_COFINS, _
                        CDbl(.VL_BC_COFINS), CDbl(.ALIQ_COFINS), .QUANT_BC_COFINS, .ALIQ_COFINS_QUANT, CDbl(.VL_COFINS), .COD_CTA, _
                        CDbl(.VL_ABAT_NT), "")
                                                   
                End If
                
            Next Produto
        
        End With
    
    End If
        
End Sub

Public Function ExtrairAliquotaICMS(ByRef Produto As IXMLDOMNode) As Double

Dim pCredSN As Double, pICMS#, pFCP#
    
    pCredSN = ValidarPercentual(Produto, "imposto/ICMS//pCredSN")
    pICMS = ValidarPercentual(Produto, "imposto/ICMS//pICMS")
    pFCP = ValidarPercentual(Produto, "imposto/ICMS//pFCP")
    
    ExtrairAliquotaICMS = pCredSN + pICMS + pFCP
    
End Function

Public Function ExtrairTotaisICMS(ByRef NFe As IXMLDOMNode) As Double

Dim vICMS As Double, vFCP#
    
    vICMS = ValidarValores(NFe, "//ICMSTot/vICMS")
    vFCP = ValidarValores(NFe, "//ICMSTot/vFCP")
    
    ExtrairTotaisICMS = vICMS + vFCP
    
End Function

Public Function ExtrairTotaisICMSST(ByRef NFe As IXMLDOMNode) As Double

Dim vICMS As Double, vFCP#
    
    vICMS = ValidarValores(NFe, "//ICMSTot/vST")
    vFCP = ValidarValores(NFe, "//ICMSTot/vFCPST")
    
    ExtrairTotaisICMSST = vICMS + vFCP
    
End Function

Public Function ExtrairValorICMS(ByRef Produto As IXMLDOMNode) As Double

Dim vCredICMSSN As Double, vICMS#, vFCP#
    
    vCredICMSSN = ValidarValores(Produto, "imposto/ICMS//vCredICMSSN")
    vICMS = ValidarValores(Produto, "imposto/ICMS//vICMS")
    vFCP = ValidarValores(Produto, "imposto/ICMS//vFCP")
    
    ExtrairValorICMS = vCredICMSSN + vICMS + vFCP
    
End Function

Public Function ExtrairValorICMSST(ByRef Produto As IXMLDOMNode) As Double

Dim vICMSST As Double, vFCPST#
    
    vICMSST = ValidarValores(Produto, "imposto/ICMS//vICMSST")
    vFCPST = ValidarValores(Produto, "imposto/ICMS//vFCPST")
    
    ExtrairValorICMSST = vICMSST + vFCPST
    
End Function

Public Function ExtrairAliquotaICMSST(ByRef Produto As IXMLDOMNode) As Double

Dim pICMSST As Double, pFCPST#
    
    pICMSST = ValidarPercentual(Produto, "imposto/ICMS//pICMSST")
    pFCPST = ValidarPercentual(Produto, "imposto/ICMS//pPFCPST")
    
    ExtrairAliquotaICMSST = pICMSST + pFCPST
    
End Function

Public Function ExtrairBaseICMS(ByRef Produto As IXMLDOMNode) As Double

Dim CSOSN As String
Dim VL_BC_ICMS As Double, VL_PROD#, VL_FRT#, VL_SEG#, VL_OUTRO#, vCredICMSSN#, vICMS#
    
    VL_BC_ICMS = ValidarValores(Produto, "imposto/ICMS//vBC")
    CSOSN = ValidarTag(Produto, "imposto/ICMS//CSOSN")
    vCredICMSSN = ValidarValores(Produto, "imposto/ICMS//vCredICMSSN")
    vICMS = ValidarValores(Produto, "imposto/ICMS//vICMS")
    
    vICMS = vCredICMSSN + vICMS
    
    If VL_BC_ICMS = 0 And vICMS > 0 Then
        
        VL_PROD = ValidarValores(Produto, "prod/vProd")
        VL_FRT = ValidarValores(Produto, "prod/vFrete")
        VL_SEG = ValidarValores(Produto, "prod/vSeg")
        VL_OUTRO = ValidarValores(Produto, "prod/vOutro")
        
    End If
    
    ExtrairBaseICMS = VL_BC_ICMS + VL_PROD + VL_FRT + VL_SEG + VL_OUTRO
    
End Function

Public Sub CriarRegistroC175Contr(ByVal Produtos As IXMLDOMNodeList, _
    ByRef dicTitulos As Dictionary, ByRef dicDados As Dictionary, ByVal CHV_PAI As String)

Dim Produto As IXMLDOMNode
Dim Campos As Variant
    
    For Each Produto In Produtos
        
         With CamposC175Contrib
             
             If VBA.Left(CamposC800.COD_SIT, 2) <> "02" Then
                
                .REG = "C175"
                .ARQUIVO = CamposC100.ARQUIVO
                .CFOP = ValidarValores(Produto, "prod/CFOP")
                .VL_OPER = ValidarValores(Produto, "prod/vItem")
                .VL_DESC = ValidarValores(Produto, "imposto/ICMS//vDesc")
                .CST_PIS = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(VBA.Format(ValidarValores(Produto, "imposto/PIS//CST"), "00"))
                .VL_BC_PIS = ValidarValores(Produto, "imposto/PIS//vBC")
                .ALIQ_PIS = ValidarPercentual(Produto, "imposto/PIS//pPIS")
                .QUANT_BC_PIS = ValidarValores(Produto, "imposto/PIS/PISQtde/qBCProd")
                .ALIQ_PIS_QUANT = ValidarValores(Produto, "imposto/PIS/PISQtde/vAliqProd")
                .VL_PIS = ValidarPercentual(Produto, "imposto/PIS//vPIS")
                .CST_COFINS = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(VBA.Format(ValidarValores(Produto, "imposto/COFINS//CST"), "00"))
                .VL_BC_COFINS = ValidarValores(Produto, "imposto/COFINS//vBC")
                .ALIQ_COFINS = ValidarPercentual(Produto, "imposto/COFINS//pCOFINS")
                .QUANT_BC_COFINS = ValidarValores(Produto, "imposto/COFINS/COFINSQtde/qBCProd")
                .ALIQ_COFINS_QUANT = ValidarValores(Produto, "imposto/COFINS/COFINSQtde/vAliqProd")
                .VL_COFINS = ValidarPercentual(Produto, "imposto/COFINS//vCOFINS")
                .COD_CTA = ""
                .INFO_COMPL = ""
                .CHV_PAI = CHV_PAI
                .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .CFOP, .CST_PIS, .CST_COFINS, .ALIQ_PIS, .ALIQ_COFINS)
                
                If .QUANT_BC_PIS = 0 Then .QUANT_BC_PIS = ""
                If .ALIQ_PIS_QUANT = 0 Then .ALIQ_PIS_QUANT = ""
                If .QUANT_BC_COFINS = 0 Then .QUANT_BC_COFINS = ""
                If .ALIQ_COFINS_QUANT = 0 Then .ALIQ_COFINS_QUANT = ""
                
                If dicDados.Exists(.CHV_REG) Then
                    
                    'Soma valores valores do C175Contr para registros com a mesma chave
                    .VL_OPER = dicDados(.CHV_REG)(dicTitulos("VL_OPER") - 1) + .VL_OPER
                    .VL_DESC = dicDados(.CHV_REG)(dicTitulos("VL_DESC") - 1) + .VL_DESC
                    .VL_BC_PIS = dicDados(.CHV_REG)(dicTitulos("VL_BC_PIS") - 1) + .VL_BC_PIS
                    .VL_PIS = dicDados(.CHV_REG)(dicTitulos("VL_PIS") - 1) + .VL_PIS
                    .VL_BC_COFINS = dicDados(.CHV_REG)(dicTitulos("VL_BC_COFINS") - 1) + .VL_BC_COFINS
                    .VL_COFINS = dicDados(.CHV_REG)(dicTitulos("VL_COFINS") - 1) + .VL_COFINS
                    
                End If
                
                Campos = Array(.REG, ARQUIVO, .CHV_REG, "", .CHV_PAI, CInt(.CFOP), CDbl(.VL_OPER), CDbl(.VL_DESC), .CST_PIS, CDbl(.VL_BC_PIS), _
                           CDbl(.ALIQ_PIS), .QUANT_BC_PIS, .ALIQ_PIS_QUANT, CDbl(.VL_PIS), .CST_COFINS, CDbl(.VL_BC_COFINS), _
                           CDbl(.ALIQ_COFINS), .QUANT_BC_COFINS, .ALIQ_COFINS_QUANT, CDbl(.VL_COFINS), .COD_CTA, .INFO_COMPL)
                           
                dicDados(.CHV_REG) = Campos
                
            End If
            
        End With
        
    Next Produto
    
End Sub

Public Sub CriarRegistroC190(ByVal Produtos As IXMLDOMNodeList, ByRef dicTitulos As Dictionary, ByRef dicDados As Dictionary, _
    ByRef dicTitulosC191 As Dictionary, ByRef dicDadosC191 As Dictionary)

Dim Produto As IXMLDOMNode
Dim Campos As Variant, Chave
Dim VL_FCP_OP As Double, ALIQ_FCP#, VL_FCP_ST#, VL_PROD#, VL_FRETE#, VL_SEG#, VL_OUTRAS#, VL_DESC#
    
    If VBA.Left(CamposC100.COD_SIT, 2) <> "02" Then
        
        For Each Produto In Produtos
            
             With CamposC190
                 
                VL_FCP_OP = ValidarValores(Produto, "imposto/ICMS//vFCP")
                ALIQ_FCP = ValidarPercentual(Produto, "imposto/ICMS//pFCP")
                VL_FCP_ST = ValidarValores(Produto, "imposto/ICMS//vFCPST")
                VL_PROD = ValidarValores(Produto, "prod/vProd")
                VL_FRETE = ValidarValores(Produto, "prod/vFrete")
                VL_SEG = ValidarValores(Produto, "prod/vSeg")
                VL_OUTRAS = ValidarValores(Produto, "prod/vOutro")
                VL_DESC = ValidarValores(Produto, "prod/vDesc")
                
                .REG = "C190"
                .ARQUIVO = CamposC100.ARQUIVO
                .CST_ICMS = fnExcel.FormatarTexto(ExtrairCST_CSOSN_ICMS(Produto))
                .CFOP = ValidarValores(Produto, "prod/CFOP")
                .ALIQ_ICMS = ValidarPercentual(Produto, "imposto/ICMS//pICMS") + ALIQ_FCP
                .VL_OPR = VL_PROD + VL_FRETE + VL_SEG + VL_OUTRAS - VL_DESC
                .VL_BC_ICMS = ValidarValores(Produto, "imposto/ICMS//vBC")
                .VL_ICMS = ValidarValores(Produto, "imposto/ICMS//vICMS") + VL_FCP_OP
                .VL_BC_ICMS_ST = ValidarValores(Produto, "imposto/ICMS//vBCST")
                .VL_ICMS_ST = ValidarValores(Produto, "imposto/ICMS//vICMSST") + VL_FCP_ST
                .VL_RED_BC = 0
                .VL_IPI = ValidarValores(Produto, "imposto/IPI//vIPI")
                .COD_OBS = ""
                .CHV_PAI = CamposC100.CHV_REG
                .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .CFOP, .CST_ICMS, .ALIQ_ICMS)
                
                'Calcula redução de base do ICMS caso exista
                If .CST_ICMS Like "*20" Or .CST_ICMS Like "*70" Then .VL_RED_BC = .VL_OPR - .VL_BC_ICMS - .VL_ICMS_ST - .VL_IPI
                
                'Verifica se a nota fiscal é de emissão de terceiros para ajustar o CFOP
                If CamposC100.IND_EMIT Like "1*" Then .CFOP = AjustarCFOPEntrada(ValidarCFOPEntrada(.CFOP))
                    
                If dicDados.Exists(.CHV_REG) Then
                    
                    'Soma valores valores do C190 para registros com a mesma chave
                    .VL_OPR = dicDados(.CHV_REG)(dicTitulos("VL_OPR") - 1) + .VL_OPR
                    .VL_BC_ICMS = dicDados(.CHV_REG)(dicTitulos("VL_BC_ICMS") - 1) + .VL_BC_ICMS
                    .VL_ICMS = dicDados(.CHV_REG)(dicTitulos("VL_ICMS") - 1) + .VL_ICMS
                    .VL_BC_ICMS_ST = dicDados(.CHV_REG)(dicTitulos("VL_BC_ICMS_ST") - 1) + .VL_BC_ICMS_ST
                    .VL_ICMS_ST = dicDados(.CHV_REG)(dicTitulos("VL_ICMS_ST") - 1) + .VL_ICMS_ST
                    .VL_RED_BC = dicDados(.CHV_REG)(dicTitulos("VL_RED_BC") - 1) + .VL_RED_BC
                    .VL_IPI = dicDados(.CHV_REG)(dicTitulos("VL_IPI") - 1) + .VL_IPI
                    
                End If
                
                dicDados(.CHV_REG) = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI, "", .CST_ICMS, .CFOP, CDbl(.ALIQ_ICMS), CDbl(.VL_OPR), _
                    CDbl(.VL_BC_ICMS), CDbl(.VL_ICMS), CDbl(.VL_BC_ICMS_ST), CDbl(.VL_ICMS_ST), CDbl(.VL_RED_BC), CDbl(.VL_IPI), .COD_OBS)
                
                If (VL_FCP_OP + VL_FCP_ST) > 0 Then Call CriarRegistroC191(Produto, dicTitulosC191, dicDadosC191)
                
            End With
                        
        Next Produto
    
    End If
    
    'TODO: Investigar causa das divergências entre C100 e C190 (Hipótese: Soma dos valores de IPI antes de atualizar os registros do SPED Fiscal [Assistente de Apuração do ICMS])
End Sub

Public Sub CriarRegistroC191(ByVal Produto As IXMLDOMNode, ByRef dicTitulos As Dictionary, ByRef dicDados As Dictionary)
    
    With CamposC191
       
       .REG = "C191"
       .ARQUIVO = CamposC100.ARQUIVO
       .VL_FCP_OP = ValidarValores(Produto, "imposto/ICMS//vFCP")
       .VL_FCP_ST = ValidarValores(Produto, "imposto/ICMS//vFCPST")
       .VL_FCP_RET = ValidarPercentual(Produto, "imposto/ICMS//vICMSSTRet")
       
       .CHV_PAI = CamposC190.CHV_REG
       .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, "C191")
       
       If dicDados.Exists(.CHV_REG) Then
           
           'Soma valores valores do C191 para registros com a mesma chave
           .VL_FCP_OP = dicDados(.CHV_REG)(dicTitulos("VL_FCP_OP") - 1) + .VL_FCP_OP
           .VL_FCP_ST = dicDados(.CHV_REG)(dicTitulos("VL_FCP_ST") - 1) + .VL_FCP_ST
           .VL_FCP_RET = dicDados(.CHV_REG)(dicTitulos("VL_FCP_RET") - 1) + .VL_FCP_RET
           
       End If
       
       If (CDbl(.VL_FCP_OP) + CDbl(.VL_FCP_ST) + CDbl(.VL_FCP_RET)) > 0 Then _
           dicDados(.CHV_REG) = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI, "", CDbl(.VL_FCP_OP), CDbl(.VL_FCP_ST), CDbl(.VL_FCP_RET))
       
    End With
    
End Sub

Public Sub CriarRegistroC800(ByRef CFe As IXMLDOMNode, ByRef dicDadosC800 As Dictionary, ByRef arrCanceladas As ArrayList)

Dim Campos As Variant
    
    With CamposC800
        
        .REG = "C800"
        .COD_MOD = ValidarTag(CFe, "//mod")
        .NUM_CFE = ValidarTag(CFe, "//nCFe")
        .DT_DOC = VBA.Format(ValidarTag(CFe, "//dEmi"), "0000-00-00")
        .VL_CFE = ValidarValores(CFe, "//vCFe")
        .VL_PIS = ValidarValores(CFe, "//ICMSTot/vPIS")
        .VL_COFINS = ValidarValores(CFe, "//ICMSTot/vCOFINS")
        .CNPJ_CPF = ""
        .NR_SAT = Util.FormatarTexto(ValidarTag(CFe, "//nserieSAT"))
        .CHV_CFE = Util.FormatarTexto(VBA.Right(ValidarTag(CFe, "//@Id"), 44))
        .VL_DESC = ValidarValores(CFe, "//ICMSTot/vDesc")
        .VL_MERC = ValidarValores(CFe, "//ICMSTot/vProd")
        .VL_OUT_DA = ValidarValores(CFe, "//ICMSTot/vOutro")
        .VL_ICMS = ValidarValores(CFe, "//ICMSTot/vICMS")
        .VL_PIS_ST = ValidarValores(CFe, "//ICMSTot/vPISST")
        .VL_COFINS_ST = ValidarValores(CFe, "//ICMSTot/vCOFINSST")
        
        If arrCanceladas.contains(VBA.Replace(.CHV_CFE, "'", "")) Then .COD_SIT = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_COD_SIT(fnSPED.GerarCodigoSituacao(fnXML.ValidarSituacao("101")))
        If VBA.Left(.COD_SIT, 2) = "02" Or VBA.Left(.COD_SIT, 2) = "03" Then
            .VL_CFE = 0
            .VL_PIS = 0
            .VL_COFINS = 0
            .VL_DESC = 0
            .VL_MERC = 0
            .VL_OUT_DA = 0
            .VL_ICMS = 0
            .VL_PIS_ST = 0
            .VL_COFINS_ST = 0
        End If
        
        .CHV_PAI = fnSPED.GerarChaveRegistro(.ARQUIVO, "C001")
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, "'" & .COD_SIT, .NUM_CFE, .NR_SAT, .DT_DOC)
        Campos = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI, "", .COD_MOD, .COD_SIT, .NUM_CFE, _
                       .DT_DOC, CDbl(.VL_CFE), CDbl(.VL_PIS), CDbl(.VL_COFINS), .CNPJ_CPF, _
                       .NR_SAT, .CHV_CFE, CDbl(.VL_DESC), CDbl(.VL_MERC), CDbl(.VL_OUT_DA), _
                       CDbl(.VL_ICMS), CDbl(.VL_PIS_ST), CDbl(.VL_COFINS_ST))
                                   
        dicDadosC800(.CHV_REG) = Campos
        
    End With
        
End Sub

Public Sub CriarRegistroC810(ByVal Produtos As IXMLDOMNodeList, ByRef dicDados As Dictionary)

Dim Produto As IXMLDOMNode
Dim Chave As String, CNPJEmit$
    
    For Each Produto In Produtos
        
         With CamposC810
             
            .NUM_ITEM = CInt(ValidarnItem(Produto))
            .CHV_PAI = CamposC800.CHV_REG
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .NUM_ITEM)
            If Not dicDados.Exists(.CHV_REG) And VBA.Left(CamposC800.COD_SIT, 2) <> "02" Then
            
                .REG = "C810"
                .ARQUIVO = CamposC800.ARQUIVO
                .COD_ITEM = ValidarTag(Produto, "prod/cProd")
                .QTD = ValidarValores(Produto, "prod/qCom")
                .UNID = VBA.UCase(ValidarTag(Produto, "prod/uCom"))
                .VL_ITEM = ValidarValores(Produto, "prod/vProd")
                .CST_ICMS = ExtrairCST_CSOSN_ICMS(Produto)
                .CFOP = ValidarValores(Produto, "prod/CFOP")

                dicDados(.CHV_REG) = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI, "", .NUM_ITEM, _
                    .COD_ITEM, CDbl(.QTD), .UNID, CDbl(.VL_ITEM), "'" & .CST_ICMS, .CFOP)
                
            End If
        
        End With
    
    Next Produto
        
End Sub

Public Sub CriarRegistroC850(ByVal Produtos As IXMLDOMNodeList, ByRef dicTitulos As Dictionary, ByRef dicDados As Dictionary)

Dim Produto As IXMLDOMNode
Dim Campos As Variant, Chave
        
    For Each Produto In Produtos
            
         With CamposC850
             
             If VBA.Left(CamposC800.COD_SIT, 2) <> "02" Then
                
                .REG = "C850"
                .ARQUIVO = CamposC800.ARQUIVO
                .CST_ICMS = ExtrairCST_CSOSN_ICMS(Produto)
                .CFOP = ValidarValores(Produto, "prod/CFOP")
                .ALIQ_ICMS = ValidarPercentual(Produto, "imposto/ICMS//pICMS")
                .VL_OPR = ValidarValores(Produto, "prod/vItem")
                .VL_BC_ICMS = ValidarValores(Produto, "prod/vItem")
                .VL_ICMS = ValidarValores(Produto, "imposto/ICMS//vICMS")
                .COD_OBS = ""
                .CHV_PAI = CamposC800.CHV_REG
                .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .CFOP, .CST_ICMS, .ALIQ_ICMS)
            
                If .VL_ICMS = 0 Then .VL_BC_ICMS = 0
                If .VL_ICMS = "" Then .VL_ICMS = 0
                If .VL_BC_ICMS = "" Then .VL_BC_ICMS = 0

                If dicDados.Exists(.CHV_REG) Then
                    
                    'Soma valores valores do C850 para registros com a mesma chave
                    .VL_OPR = dicDados(.CHV_REG)(dicTitulos("VL_OPR") - 1) + .VL_OPR
                    .VL_BC_ICMS = dicDados(.CHV_REG)(dicTitulos("VL_BC_ICMS") - 1) + .VL_BC_ICMS
                    .VL_ICMS = dicDados(.CHV_REG)(dicTitulos("VL_ICMS") - 1) + .VL_ICMS
                    
                End If
                
                dicDados(.CHV_REG) = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI, "", "'" & .CST_ICMS, _
                    .CFOP, CDbl(.ALIQ_ICMS), CDbl(.VL_OPR), CDbl(.VL_BC_ICMS), CDbl(.VL_ICMS), .COD_OBS)
                
            End If
        
        End With
                               
    Next Produto
        
End Sub

Public Sub CriarRegistroD100(ByRef CTe As IXMLDOMNode, ByRef dicDadosD100 As Dictionary, ByRef dicDadosD101 As Dictionary, ByRef dicDadosD105 As Dictionary, _
    ByRef dicDadosD190 As Dictionary, ByVal ARQUIVO As String, ByVal CHV_PAI As String, Optional SPEDContr As Boolean)

Dim Campos As Variant
    
    With CamposD100
        
        .REG = "D100"
        .IND_OPER = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_D100_IND_OPER(fnXML.IdentificarTipoOperacaoCTe(CTe, ValidarTag(CTe, "//tpCTe")))
        .IND_EMIT = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_EMIT(fnXML.IdentificarTipoEmissaoCTe(CTe))
        .COD_PART = fnExcel.FormatarTexto(fnXML.IdentificarParticipanteCTe(CTe))
        .COD_MOD = ValidarTag(CTe, "//mod")
        .COD_SIT = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_COD_SIT(fnSPED.GerarCodigoSituacao(fnXML.ValidarSituacao(ValidarTag(CTe, "//cStat"))))
        .SER = VBA.Format(ValidarTag(CTe, "//serie"), "000")
        .SUB = ValidarTag(CTe, "//subserie")
        .NUM_DOC = ValidarTag(CTe, "//nCT")
        .CHV_CTE = fnExcel.FormatarTexto(VBA.Right(ValidarTag(CTe, "//@Id"), 44))
        .DT_DOC = VBA.Format(VBA.Left(ValidarTag(CTe, "//dhEmi"), 10), "yyyy-mm-dd")
        .DT_A_P = .DT_DOC
        .TP_CTe = ValidarEnumeracao_TP_CT_E(ValidarTag(CTe, "//tpCTe"))
        .CHV_CTE_REF = ""
        .VL_DOC = ValidarValores(CTe, "//vRec")
        .VL_DESC = 0
        .IND_FRT = fnXML.ValidarEnumeracao_IND_FRTCTe(CTe)
        .VL_SERV = ValidarValores(CTe, "//vPrest/vRec")
        .VL_BC_ICMS = ValidarValores(CTe, "//imp/ICMS//vBC")
        .VL_ICMS = ValidarValores(CTe, "//imp/ICMS//vICMS")
        .VL_NT = 0
        .COD_INF = ""
        .COD_CTA = ""
        .COD_MUN_ORIG = ValidarTag(CTe, "//cMunIni")
        .COD_MUN_DEST = ValidarTag(CTe, "//cMunFim")
        .CHV_PAI = CHV_PAI
        
        If .VL_SERV = "" Then .VL_SERV = 0
        
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .IND_EMIT, .NUM_DOC, .COD_MOD, .SER, .SUB, .COD_PART, .CHV_CTE)
        Campos = Array(.REG, ARQUIVO, .CHV_REG, .CHV_PAI, "", .IND_OPER, .IND_EMIT, .COD_PART, .COD_MOD, _
                       .COD_SIT, .SER, .SUB, .NUM_DOC, .CHV_CTE, .DT_DOC, .DT_A_P, .TP_CTe, .CHV_CTE_REF, _
                       CDbl(.VL_DOC), CDbl(.VL_DESC), .IND_FRT, CDbl(.VL_SERV), CDbl(.VL_BC_ICMS), _
                       CDbl(.VL_ICMS), CDbl(.VL_NT), .COD_INF, .COD_CTA, .COD_MUN_ORIG, .COD_MUN_DEST)
                    
        dicDadosD100(.CHV_REG) = Campos
        
        'Gera registros filhos
        Call IncuirRegistroD101Contr(CDbl(.VL_SERV), dicDadosD101, ARQUIVO)
        Call IncuirRegistroD105(CDbl(.VL_SERV), dicDadosD105, ARQUIVO)
        Call IncluirRegistroD190(CTe, dicDadosD190, VBA.Val(.IND_OPER), .CHV_REG, ARQUIVO)
        
    End With
    
End Sub

Public Sub IncuirRegistroD101Contr(ByRef vRec As Double, ByRef dicDadosD101 As Dictionary, ByVal ARQUIVO As String)

Dim Campos As Variant
    
    With CamposD101_Contr
        
        .REG = "D101"
        .IND_NAT_FRT = ""
        .VL_ITEM = vRec
        .CST_PIS = ""
        .NAT_BC_CRED = ""
        .VL_BC_PIS = 0
        .ALIQ_PIS = 0
        .VL_PIS = 0
        .COD_CTA = ""
        .CHV_PAI = CamposD100.CHV_REG
        
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .IND_NAT_FRT, .CST_PIS, .NAT_BC_CRED, .ALIQ_PIS, .COD_CTA)
        Campos = Array(.REG, ARQUIVO, .CHV_REG, "", .CHV_PAI, .IND_NAT_FRT, CDbl(.VL_ITEM), _
            .CST_PIS, .NAT_BC_CRED, CDbl(.VL_BC_PIS), CDbl(.ALIQ_PIS), CDbl(.VL_PIS), .COD_CTA)
                    
        dicDadosD101(.CHV_REG) = Campos
        
    End With
    
End Sub

Public Sub IncuirRegistroD105(ByRef vRec As Double, ByRef dicDadosD105 As Dictionary, ByVal ARQUIVO As String)

Dim Campos As Variant
    
    With CamposD105
        
        .REG = "D105"
        .IND_NAT_FRT = ""
        .VL_ITEM = vRec
        .CST_COFINS = ""
        .NAT_BC_CRED = ""
        .VL_BC_COFINS = 0
        .ALIQ_COFINS = 0
        .VL_COFINS = 0
        .COD_CTA = ""
        .CHV_PAI = CamposD100.CHV_REG
        
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .IND_NAT_FRT, .CST_COFINS, .NAT_BC_CRED, .ALIQ_COFINS, .COD_CTA)
        Campos = Array(.REG, ARQUIVO, .CHV_REG, "", .CHV_PAI, .IND_NAT_FRT, CDbl(.VL_ITEM), _
            .CST_COFINS, .NAT_BC_CRED, CDbl(.VL_BC_COFINS), CDbl(.ALIQ_COFINS), CDbl(.VL_COFINS), .COD_CTA)
                    
        dicDadosD105(.CHV_REG) = Campos
        
    End With
    
End Sub

Public Sub IncluirRegistroD190(ByRef CTe As IXMLDOMNode, ByRef dicDadosD190 As Dictionary, ByVal IND_OPER As String, ByVal CHV_PAI As String, ByVal ARQUIVO As String)

Dim UFIni As String
Dim UFFim As String

Dim Campos As Variant
    
    On Error GoTo Tratar:
        
        With CamposD190
            
            .REG = "D190"
            .CST_ICMS = VBA.Format(ValidarTag(CTe, "//CST"), "000")
            .CFOP = ValidarTag(CTe, "//CFOP")
            .ALIQ_ICMS = ValidarPercentual(CTe, "//pICMS")
            .VL_OPR = ValidarValores(CTe, "//vRec")
            .VL_BC_ICMS = ValidarValores(CTe, "//vBC")
            .VL_ICMS = ValidarValores(CTe, "//vICMS")
            .VL_RED_BC = 0
            .COD_OBS = ""
            .CHV_PAI = CHV_PAI
                        
            If IND_OPER = "0" Then .CFOP = AjustarCFOPEntrada(ValidarCFOPEntrada(.CFOP))
        
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .CFOP, .CST_ICMS, .ALIQ_ICMS)
            Campos = Array(.REG, ARQUIVO, .CHV_REG, .CHV_PAI, "", "'" & .CST_ICMS, .CFOP, CDbl(.ALIQ_ICMS), _
                           CDbl(.VL_OPR), CDbl(.VL_BC_ICMS), CDbl(.VL_ICMS), CDbl(.VL_RED_BC), .COD_OBS)
                        
            dicDadosD190(.CHV_REG) = Campos
            
        End With
        
Exit Sub
Tratar:

Resume

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
        .VL_DOC = ValidarValores(CTe, "//vRec")
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
        
        vICMS = ValidarValores(CTe, "//imp/ICMS//vICMS")
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

Public Function ImportarDocumentosEletronicos(ByVal arrXMLs As ArrayList, ByRef arrChaves As ArrayList, ByRef dicEntradasNFe As Dictionary, _
                                              ByRef dicSaidasNFe As Dictionary, ByRef dicEntradasCTes As Dictionary, ByRef dicSaidasCTes As Dictionary, _
                                              ByRef dicSaidasNFCe As Dictionary, ByRef dicSaidasCFe As Dictionary, ByRef arrCanceladas As ArrayList) As Boolean

Dim VerificarStatusXML As Boolean
Dim Arq As Variant, Registro
Dim Doce As New DOMDocument60
        
    a = 0
    Comeco = Timer
    Application.StatusBar = "Importando informações dos documentos eletrônicos, por favor aguarde..."
    For Each Arq In arrXMLs
        
        'Rotina para evitar travamentos
        Call Util.AntiTravamento(a, 200, "", arrXMLs.Count, Comeco)
        Set Doce = fnXML.RemoverNamespaces(Arq)
        
        'Rotinas de validação dos documentos eletrônicos
        If ValidarXMLNFe(Doce) Or ValidarXMLCTe(Doce) Or ValidarXMLCFe(Doce) Then
            
            With DadosDoce
                
                .chNFe = fnXML.ValidarTag(Doce, "//@Id")
                If .chNFe = "" Then GoTo Pular:
                
                .chNFe = VBA.Right(.chNFe, 44)
                .Modelo = Doce.SelectSingleNode("//mod").text
                If Not arrChaves.contains(Replace(.chNFe, "'", "")) And (.Modelo = "55" Or .Modelo = "65" Or .Modelo = "57" Or .Modelo = "59" Or .Modelo = "67") Then
                    
                    .CNPJEmit = VBA.Mid(.chNFe, 7, 14)
                    If .Modelo = 59 Then .dtEmi = VBA.Format(Doce.SelectSingleNode("//dEmi").text, "0000-00-00") Else .dtEmi = VBA.Left(Doce.SelectSingleNode("//dhEmi").text, 10)
                    
                    If arrCanceladas.contains(.chNFe) Then
                        
                        .Status = "Cancelada"
                        
                    ElseIf .Modelo = "59" Then
                        
                        .Status = ValidarSituacao("100")
                        
                    Else
                                                
                        .Status = ValidarSituacao(fnXML.ValidarTag(Doce, "//cStat"))
                        '.Status = ValidarSituacao(Doce.SelectSingleNode("//cStat").text)
                    
                    End If
                    
                    If .Modelo = "55" Or .Modelo = "65" Then
                        
                        .CNPJDest = ExtrairCNPJDestinatario(Doce)
                        .nNF = Doce.SelectSingleNode("//nNF").text
                        .vNF = Util.FormatarValores(Doce.SelectSingleNode("//vNF").text)
                        .tpNF = ValidarOperacao(Doce.SelectSingleNode("//tpNF").text)
                        .UF = ValidarUFNFeNFCe(Doce, .CNPJEmit, .CNPJDest)
                        
                        If .Modelo = "65" Then
                            .UFEmit = ValidarTag(Doce, "//emit/enderEmit/UF")
                            .UF = .UFEmit
                        End If
                        
                    ElseIf .Modelo = "59" Then
                        
                        .CNPJDest = ExtrairCNPJDestinatario(Doce)
                        .nNF = Doce.SelectSingleNode("//nCFe").text
                        .vNF = Util.FormatarValores(Doce.SelectSingleNode("//vCFe").text)
                        .tpNF = ValidarOperacao("1")
                        If Util.ValidarUF(Util.ConverterIBGE_UF(VBA.Left(.chNFe, 2))) Then .UF = Util.ConverterIBGE_UF(VBA.Left(.chNFe, 2))
                        
                    ElseIf .Modelo = "57" Or .Modelo = "67" Then
                        
                        .CNPJTomador = ExtrairCNPJTomador(Doce)
                        .nNF = Doce.SelectSingleNode("//nCT").text
                        .vNF = Util.FormatarValores(Doce.SelectSingleNode("//vRec").text)
                        .tpNF = ValidarEnumeracao_D100_IND_OPER(Doce.SelectSingleNode("//tpCTe").text)
                        .CNPJEmit = ValidarTag(Doce, "//emit/CNPJ")
                        .CNPJRem = ValidarTag(Doce, "//rem/CNPJ")
                        .UF = ValidarUFCTe(Doce, .CNPJEmit, .CNPJTomador)
                        .Tomador = ValidarTomador(ValidarTag(Doce, "//toma"))
                        
                    End If
                    
                    If (CNPJContribuinte <> .CNPJTomador) And (CNPJContribuinte <> .CNPJEmit) And (CNPJContribuinte <> .CNPJDest) Then GoTo Pular:
                    .CNPJPart = ExtrairEmitenteDestinatario(Doce, .CNPJEmit, .tpNF)
                    .RazaoPart = ValidarRazaoEmitenteDestinatario(Doce, .CNPJEmit, .tpNF)
                    .StatusSPED = ""
                    .DivergNF = ""
                    .OBSERVACOES = ""
                    
                    .UFDest = ValidarTag(Doce, "//dest/enderDest/UF")
                    Call Util.GerarObservacoes(.Status, .CNPJEmit, .UFDest, .tpNF)
                    
                    Registro = Array(.nNF, "'" & .CNPJPart, .RazaoPart, .dtEmi, CDbl(.vNF), "'" & .chNFe, .UF, .Status, .tpNF, .StatusSPED, .DivergNF, .OBSERVACOES)
                    Call Util.ClassificarNotaFiscal(.CNPJEmit, .tpNF, .Modelo, Registro, arrChaves, dicEntradasNFe, dicSaidasNFe, dicSaidasCTes, dicEntradasCTes, dicSaidasNFCe, dicSaidasCFe)
                    
                End If
                
            End With
            
        End If
        
Pular:
        DadosDoce.CNPJDest = "": DadosDoce.CNPJEmit = "": DadosDoce.CNPJTomador = ""
        
    Next Arq
    
    ImportarDocumentosEletronicos = True
    Application.StatusBar = False
    
End Function

Public Function ColetarFornecedoresSN(ByRef dicFornecSN As Dictionary, ByVal Arqs As ArrayList)

Dim Produtos As IXMLDOMNodeList
Dim Produto As IXMLDOMNode
Dim NFe As New DOMDocument60
Dim chNFe As String, CRT$, CFOP$, NITEM$, Chave$
Dim pCredSN As Double
Dim Arq As Variant
    
    For Each Arq In Arqs
        
        Set NFe = fnXML.RemoverNamespaces(Arq)
        If ValidarNFe(NFe) Then
        
            chNFe = VBA.Right(NFe.SelectSingleNode("//@Id").text, 44)
            
            CRT = ValidarTag(NFe, "//CRT")
            If CRT = "1" Then
            
                Set Produtos = NFe.SelectNodes("//det")
                For Each Produto In Produtos
                    
                    NITEM = ValidarTag(Produto, "@nItem")
                    CFOP = ValidarValores(Produto, "prod/CFOP")
                    pCredSN = ValidarPercentual(Produto, "imposto/ICMS//pCredSN")
                    
                    Chave = chNFe & NITEM
                    dicFornecSN(Chave) = Array(CFOP, pCredSN)
                
                Next Produto
            
            End If
        
        End If
        
    Next Arq
    
End Function

Public Sub ExtrairSaidasProdutos(ByVal Arqs As Variant, ByRef dicSaidasProdutos As Dictionary, _
                                 ByRef dicEntradasProdutos As Dictionary, ByRef dicReferencias As Dictionary)

Dim Produtos As IXMLDOMNodeList
Dim Produto As IXMLDOMNode
Dim NFe As New DOMDocument60
Dim Arq, Itens

On Error GoTo Tratar:
    
    If VarType(Arqs) = 8204 Then
        
        For Each Arq In Arqs
            
            Set NFe = fnXML.RemoverNamespaces(Arq)
            
            DadosDoce.chNFe = VBA.Right(ValidarTag(NFe, "//@Id"), 44)
            If ValidarXMLNFe(NFe) Then
                
                With DadosDoce
                                        
                    Set Produtos = NFe.SelectNodes("//det")
                    For Each Produto In Produtos
                            
                        .xProd = ValidarTag(Produto, "prod/xProd")
                            
                            .NITEM = ValidarTag(Produto, "@nItem")
                            .cProd = ValidarTag(Produto, "prod/cProd")
                            .cBarra = ValidarTag(Produto, "prod/cEAN")
                            .CFOP = ValidarTag(Produto, "prod/CFOP")
                            .qCom = ValidarValores(Produto, "prod/qCom")
                            .uCom = ValidarTag(Produto, "prod/uCom")
                            .vUnit = ValidarValores(Produto, "prod/vUnCom")
                            .vProd = ValidarValores(Produto, "prod/vProd")
                            
                            .Chave = .cProd & .CFOP & .uCom
                            
                            If dicReferencias.Exists(.cProd) Then .Referencia = dicReferencias(.cProd)(4)
                            
                            If dicSaidasProdutos.Exists(.cProd) Then
                                .qCom = dicSaidasProdutos(.cProd)(8) + CDbl(.qCom)
                                .vProd = dicSaidasProdutos(.cProd)(11) + CDbl(.vProd)
                                .vUnit = CDbl(.vProd) / CDbl(.qCom)
                            End If
                            
                            If dicEntradasProdutos.Exists(.cProd) Then
                                .qCom = dicEntradasProdutos(.cProd)(8) + CDbl(.qCom)
                                .vProd = dicEntradasProdutos(.cProd)(11) + CDbl(.vProd)
                                .vUnit = CDbl(.vProd) / CDbl(.qCom)
                            End If
                            
                            .Chave = .chNFe & .NITEM
                            
                            Select Case VBA.Left(.CFOP, 1)
                            
                                Case "1", "2", "3"
                                    dicEntradasProdutos(.Chave) = Array("'" & .chNFe, .NITEM, "'" & .cProd, .xProd, .Referencia, .cBarra, .CFOP, CDbl(.qCom), .uCom, CDbl(.vUnit), CDbl(.vProd))
                                    
                                    
                                Case "5", "6", "7"
                                    dicSaidasProdutos(.Chave) = Array("'" & .chNFe, .NITEM, "'" & .cProd, .xProd, .Referencia, .cBarra, .CFOP, CDbl(.qCom), .uCom, CDbl(.vUnit), CDbl(.vProd))
                                        
                            End Select
                            
                            .Referencia = ""
                            
                    Next Produto
                    
                End With
                
            End If
            
        Next Arq
        
    End If
    
Exit Sub
Tratar:

Resume

End Sub

Public Sub ImportarCTe(ByVal Arqs As Variant, Dicionario As Dictionary)

Dim CTe As New DOMDocument60
Dim Arq As Variant

    If VarType(Arqs) <> 11 Then
        
        For Each Arq In Arqs
            
            CTe.Load (Arq)
            If ValidarXMLCTe(CTe) Then
                
                With DadosCTe
                    
                    .nCTe = CTe.SelectSingleNode("//nCT").text
                    .CNPJEmit = CTe.SelectSingleNode("//emit/CNPJ").text
                    .RazaoEmit = CTe.SelectSingleNode("//emit/xNome").text
                    .dhEmi = VBA.Left(CTe.SelectSingleNode("//dhEmi").text, 10)
                    .vCTe = CTe.SelectSingleNode("//vRec").text
                    .chCTe = VBA.Right(CTe.SelectSingleNode("//@Id").text, 44)
                    .UFOrig = CTe.SelectSingleNode("//UFIni").text
                    .Stituacao = ValidarSituacao(CTe.SelectSingleNode("//cStat").text)
                    .tpOperacao = ValidarOperacao(CTe.SelectSingleNode("//tpCTe").text)
                    .dtLancamento = ""
                    .DivCTe = ""
                    .OBSERVACOES = ""
                    
                    If Not Dicionario.Exists(.chCTe) Then
                        Dicionario(.chCTe) = Array(.nCTe, "'" & .CNPJEmit, .RazaoEmit, .dhEmi, .vCTe, "'" & .chCTe, _
                                                   .UFOrig, .Stituacao, .tpOperacao, .dtLancamento, .DivCTe, .OBSERVACOES)
                    End If
                    
                End With
                
            End If
            
        Next Arq
        
    End If
    
End Sub

Public Sub ImportarXMLSParaAnalise(ByVal Arqs As Variant, ByRef dicDivergencias As Dictionary)

Dim arrCanceladas As New ArrayList
Dim NFe As New DOMDocument60
Dim Arq As Variant, Campos
    
    If VarType(Arqs) <> 11 Then
        
        Call fnXML.CarregarProtocolosCancelamento(Arqs, arrCanceladas)
        
        a = 0
        Comeco = Timer
        For Each Arq In Arqs
            
            Call Util.AntiTravamento(a, 100, "Importando dados dos XMLS selecionados, por favor aguarde...", Arqs.Count, Comeco)
            With RelDiverg
                
                Set NFe = fnXML.RemoverNamespaces(Arq)
                
                .CHV_NFE = VBA.Right(ValidarTag(NFe, "//@Id"), 44)
                If ValidarXMLNFe(NFe) And dicDivergencias.Exists(.CHV_NFE) Then
                    
                    If dicDivergencias(.CHV_NFE)(27) <> "DIVERGÊNCIA" Then
                         
                        .DOC_CONTRIB = CadContrib.Range("CNPJContribuinte").value
                        .Operacao = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_OPER(fnXML.IdentificarTipoOperacao(NFe, ValidarTag(NFe, "//tpNF")))
                        .TP_EMISSAO = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_EMIT(fnXML.IdentificarTipoEmissao(NFe))
                        .DOC_PART = "'" & fnXML.IdentificarParticipante(NFe, VBA.Left(.Operacao, 1), VBA.Left(.TP_EMISSAO, 1))
                        .Modelo = ValidarTag(NFe, "//mod")
                        .Situacao = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_COD_SIT(fnSPED.GerarCodigoSituacao(fnXML.ValidarSituacao(ValidarTag(NFe, "//cStat"))))
                        .SERIE = VBA.Format(ValidarTag(NFe, "//serie"), "000")
                        .NUM_DOC = VBA.Format(ValidarTag(NFe, "//nNF"), String(9, "0"))
                        .DT_DOC = VBA.Format(VBA.Left(ValidarTag(NFe, "//dhEmi"), 10), "yyyy-mm-dd")
                        .TP_PAGAMENTO = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_PGTO(ValidarTag(NFe, "//indPag"))
                        .TP_FRETE = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_FRT(ValidarTag(NFe, "//modFrete"))
                        .VL_DOC = ValidarValores(NFe, "//ICMSTot/vNF")
                        .VL_DESC = ValidarValores(NFe, "//ICMSTot/vDesc")
                        .VL_ABATIMENTO = ValidarValores(NFe, "//ICMSTot/vICMSDeson")
                        .VL_PROD = ValidarValores(NFe, "//ICMSTot/vProd")
                        .VL_FRETE = ValidarValores(NFe, "//ICMSTot/vFrete")
                        .VL_SEG = ValidarValores(NFe, "//ICMSTot/vSeg")
                        .VL_OUTRO = ValidarValores(NFe, "//ICMSTot/vOutro")
                        .VL_BC_ICMS = ValidarValores(NFe, "//ICMSTot/vBC")
                        .VL_ICMS = CDbl(ValidarValores(NFe, "//ICMSTot/vFCP")) + CDbl(ValidarValores(NFe, "//ICMSTot/vICMS"))
                        .VL_BC_ICMS_ST = ValidarValores(NFe, "//ICMSTot/vBCST")
                        .VL_ICMS_ST = CDbl(ValidarValores(NFe, "//ICMSTot/vST")) + CDbl(ValidarValores(NFe, "//ICMSTot/vFCPST"))
                        .VL_IPI = ValidarValores(NFe, "//ICMSTot/vIPI")
                        .VL_PIS = ValidarValores(NFe, "//ICMSTot/vPIS")
                        .VL_COFINS = ValidarValores(NFe, "//ICMSTot/vCOFINS")
                        
                        If .Modelo = "65" Then .DOC_PART = ""
                        
                        If DesconsiderarPISCOFINS = True Then
                            .VL_PIS = 0
                            .VL_COFINS = 0
                        End If
                        
                        If DesconsiderarAbatimento = True Then
                            .VL_ABATIMENTO = 0
                        End If
                        
                        If SomarIPIProdutos = True Then
                            .VL_PROD = CDbl(.VL_PROD) + CDbl(.VL_IPI)
                            .VL_IPI = 0
                        End If
                        
                        If SomarICMSSTProdutos = True Then
                            .VL_PROD = CDbl(.VL_PROD) + CDbl(.VL_ICMS_ST)
                            .VL_BC_ICMS_ST = 0
                            .VL_ICMS_ST = 0
                        End If
                        
                        If arrCanceladas.contains(.CHV_NFE) Then
                            
                            .Situacao = "02 - Documento Cancelado"
                            .VL_DOC = 0
                            .VL_DESC = 0
                            .VL_ABATIMENTO = 0
                            .VL_PROD = 0
                            .VL_FRETE = 0
                            .VL_SEG = 0
                            .VL_OUTRO = 0
                            .VL_BC_ICMS = 0
                            .VL_ICMS = 0
                            .VL_BC_ICMS_ST = 0
                            .VL_ICMS_ST = 0
                            .VL_IPI = 0
                            .VL_PIS = 0
                            .VL_COFINS = 0
                            
                        End If
                        
                        Campos = Array(.DOC_CONTRIB, .DOC_PART, .Modelo, .Operacao, .TP_EMISSAO, .Situacao, .SERIE, _
                                       .NUM_DOC, "'" & .CHV_NFE, .DT_DOC, .TP_PAGAMENTO, .TP_FRETE, CDbl(.VL_DOC), _
                                       CDbl(.VL_PROD), CDbl(.VL_FRETE), CDbl(.VL_SEG), CDbl(.VL_OUTRO), CDbl(.VL_DESC), _
                                       CDbl(.VL_ABATIMENTO), CDbl(.VL_BC_ICMS), CDbl(.VL_ICMS), CDbl(.VL_BC_ICMS_ST), _
                                       CDbl(.VL_ICMS_ST), CDbl(.VL_IPI), CDbl(.VL_PIS), CDbl(.VL_COFINS), "XML IMPORTADO", "")
                                        
                        Call AnalisarDadosSPEDXML(dicDivergencias, Campos, .CHV_NFE)
                        
                    End If
                    
                End If
                
            End With
Prx:
        Next Arq
        
    End If
    
End Sub

Public Function AnalisarDadosSPEDXML(ByRef dicDivergencias As Dictionary, ByRef Campos As Variant, ByVal CHV_NFE As String)

Dim dicCampos As Variant
Dim i As Long, c As Long
Dim t As Double
    
    dicCampos = dicDivergencias(CHV_NFE)
        
        c = 0: t = 0
        For i = LBound(dicCampos) To UBound(Campos)
            
            If dicCampos(27) <> "DIVERGÊNCIA" Then
                
                Select Case i
                    
                    Case Is <= 2
                        If FomatarCPFouCNPJ(dicCampos(i)) = FomatarCPFouCNPJ(Campos(i - 1)) Then dicCampos(i) = "OK" Else dicCampos(i) = "DIVERGENTE"
                        
                    Case Is <= 12
                        If i <> 1 And i <> 2 And i <> 9 Then If dicCampos(i) = Campos(i - 1) Then dicCampos(i) = "OK" Else dicCampos(i) = "DIVERGENTE"
                        If dicCampos(i) = "OK" Then c = c + 1
                        
                    Case Is <= 26
                        dicCampos(i) = VBA.Round(CDbl(dicCampos(i)) - CDbl(Campos(i - 1)), 2)
                        t = t + VBA.Abs(CDbl(dicCampos(i)))
                        
                End Select
                
            End If
            
        Next i
        
        If c < 11 Or t > 0 Then dicCampos(27) = "DIVERGÊNCIA": dicDivergencias(CHV_NFE) = dicCampos
        If c = 11 And t = 0 Then Call dicDivergencias.Remove(CHV_NFE)

End Function

Private Function FomatarCPFouCNPJ(ByVal Campo As String)

    Campo = VBA.Replace(Campo, "'", "")
    Campo = VBA.Format(Campo, VBA.String(14, "0"))
    FomatarCPFouCNPJ = Campo

End Function

Public Sub CriarRegistroTotaisNFe(ByVal Arqs As Variant, ByRef Dicionario As Dictionary, ByRef Chaves As Dictionary)

Dim NFe As New DOMDocument60
Dim Arq As Variant

    For Each Arq In Arqs
        
        With DadosDoce
        
            Set NFe = fnXML.RemoverNamespaces(Arq)
            
            .chNFe = VBA.Right(ValidarTag(NFe, "//@Id"), 44)
            If Not Chaves.Exists(.chNFe) And ValidarXMLNFe(NFe) Then
                
                .CNPJEmit = ExtrairCNPJEmitente(NFe)
                .CNPJDest = ExtrairCNPJDestinatario(NFe)
                .RazaoEmit = ValidarTag(NFe, "//emit/xNome")
                .nNF = ValidarTag(NFe, "//nNF")
                .dtEmi = VBA.Left(ValidarTag(NFe, "//dhEmi"), 10)
                .vNF = -ValidarValores(NFe, "//ICMSTot/vNF")
                .vProd = -ValidarValores(NFe, "//ICMSTot/vProd")
                .vFrete = -ValidarValores(NFe, "//ICMSTot/vFrete")
                .vSeg = -ValidarValores(NFe, "//ICMSTot/vSeg")
                .vOutro = -ValidarValores(NFe, "//ICMSTot/vOutro")
                .vDesc = -ValidarValores(NFe, "//ICMSTot/vDesc")
                .vICMSDeson = -ValidarValores(NFe, "//ICMSTot/vICMSDeson")
                .vBCICMS = -ValidarValores(NFe, "//ICMSTot/vBC")
                .vICMS = -ValidarValores(NFe, "//ICMSTot/vICMS")
                .vFCP = -ValidarValores(NFe, "//ICMSTot/vFCP")
                .vBCST = -ValidarValores(NFe, "//ICMSTot/vBCST")
                .vST = -ValidarValores(NFe, "//ICMSTot/vST")
                .vFCPST = -ValidarValores(NFe, "//ICMSTot/vFCPST")
                .vIPI = -ValidarValores(NFe, "//ICMSTot/vIPI")
                
                .vICMS = CDbl(.vICMS) + CDbl(.vFCP)
                .vICMSST = CDbl(.vST) + CDbl(.vFCPST)
                
                If Dicionario.Count Mod 50 = 0 Then DoEvents
                
                Dicionario(.chNFe) = Array(.CNPJDest, .CNPJEmit, .RazaoEmit, .nNF, .dtEmi, "'" & .chNFe, .tpOperacao, CDbl(.vNF), _
                                           CDbl(.vProd), CDbl(.vFrete), CDbl(.vSeg), CDbl(.vOutro), CDbl(.vDesc), CDbl(.vICMSDeson), _
                                           CDbl(.vBCICMS), CDbl(.vICMS), CDbl(.vBCST), CDbl(.vST), CDbl(.vIPI), "SIM", "")
                                    
            End If
        
        End With
Prx:
    Next Arq

End Sub

Private Function CalcularICMSTotal(ByVal prod As IXMLDOMNode) As Double
   
Dim Tag
       
    For Each Tag In Split("vICMS,vICMSST,vFCP,vFCPST", ",")
    
        If Not prod.SelectSingleNode("imposto/ICMS//" & Tag) Is Nothing Then
            CalcularICMSTotal = CalcularICMSTotal + Replace(prod.SelectSingleNode("imposto/ICMS//" & Tag).text, ".", ",")
        End If
                    
    Next Tag
    
End Function

Private Function ValidarResponsavel(ByVal NFe As IXMLDOMNode) As String

    If Not NFe.SelectSingleNode("imposto/ICMS//vICMSST") Is Nothing Then
        ValidarResponsavel = 1
    ElseIf Not NFe.SelectSingleNode("imposto/ICMS//vICMSSTRet") Is Nothing Then
        ValidarResponsavel = 2
    Else
        ValidarResponsavel = 3
    End If

End Function

Public Function ValidarXMLNFe(ByVal NFe As IXMLDOMNode) As Boolean
    If Not NFe.SelectSingleNode("nfeProc") Is Nothing Or _
       Not NFe.SelectSingleNode("retConsNFeLog") Is Nothing Then ValidarXMLNFe = True
End Function

Public Function ValidarParticipante(ByVal NFe As IXMLDOMNode, Optional ByRef SPEDContrib As Boolean) As Boolean

Dim CNPJEmit As String
Dim CNPJDest As String
Dim CNPJBase As String

    If Not NFe.SelectSingleNode("//emit/CNPJ") Is Nothing Then CNPJEmit = NFe.SelectSingleNode("//emit/CNPJ").text
    If Not NFe.SelectSingleNode("//dest/CNPJ") Is Nothing Then CNPJDest = NFe.SelectSingleNode("//dest/CNPJ").text
        
    If (CNPJEmit = CNPJContribuinte) Or (CNPJDest = CNPJContribuinte) Then ValidarParticipante = True
    If SPEDContrib Then
        CNPJBase = VBA.Left(CNPJContribuinte, 8)
        If (CNPJEmit Like CNPJBase & "*") Or (CNPJDest Like CNPJBase & "*") Then ValidarParticipante = True
    End If
    
End Function

Public Function ValidarDestinatario(ByVal CNPJDest As String) As Boolean
    If CNPJDest = CNPJContribuinte Then ValidarDestinatario = True
End Function

Public Function ValidarParticipanteCTe(ByVal CTe As IXMLDOMNode) As Boolean

Dim CNPJEmit As String
Dim CNPJToma As String

    CNPJEmit = ValidarTag(CTe, "//emit/CNPJ")
    CNPJToma = ExtrairCNPJTomador(CTe)
    
    If (CNPJEmit = CNPJContribuinte) Or (CNPJToma = CNPJContribuinte) Then ValidarParticipanteCTe = True
    
End Function

Public Function ValidarProtocolo(ByVal NFe As IXMLDOMNode) As Boolean

    Select Case True
        
        Case (Not NFe.SelectSingleNode("procEventoNFe") Is Nothing) Or (Not NFe.SelectSingleNode("procEventoCTe") Is Nothing)
            ValidarProtocolo = True
            
    End Select
    
End Function

Public Function ExtrairChaveAcesso(ByVal NFe As IXMLDOMNode) As String

Dim Tag As String

    Select Case True
        
        Case Not NFe.SelectSingleNode("procEventoNFe") Is Nothing
            Tag = "//chNFe"
            
        Case Not NFe.SelectSingleNode("procEventoCTe") Is Nothing
            Tag = "//chCTe"
            
    End Select
    
    If Tag <> "" Then ExtrairChaveAcesso = VBA.Right(NFe.SelectSingleNode(Tag).text, 44)
    
End Function

Public Function ExtrairChaveAcessoNFe(ByVal NFe As IXMLDOMNode) As String
    ExtrairChaveAcessoNFe = VBA.Right(ValidarTag(NFe, "//@Id"), 44)
End Function

Public Function ExtrairNumeroDocumentoNFe(ByVal Produto As IXMLDOMNode) As String
    ExtrairNumeroDocumentoNFe = ValidarTag(Produto, "//nNF")
End Function

Public Function ExtrairSerieDocumentoNFe(ByVal NFe As IXMLDOMNode) As String
    ExtrairSerieDocumentoNFe = ValidarTag(NFe, "//serie")
End Function

Public Function ExtrairModeloDocumentoNFe(ByVal NFe As IXMLDOMNode) As String
    ExtrairModeloDocumentoNFe = ValidarTag(NFe, "//mod")
End Function

Public Function ExtrairNumeroItemProduto(ByVal Produto As IXMLDOMNode) As Integer

Dim Tags As Variant, Tag
    
    Tags = Array("@nItem", "nItem")
    
    For Each Tag In Tags
        
        If Not Util.VerificarStringVazia(ValidarTag(Produto, Tag)) Then
            
            ExtrairNumeroItemProduto = ValidarValores(Produto, Tag)
            Exit For
            
        End If
        
    Next Tag
    
End Function

Public Function ExtrairProdutosNFe(ByVal NFe As IXMLDOMNode) As IXMLDOMNodeList
    Set ExtrairProdutosNFe = NFe.SelectNodes("//det")
End Function

Public Function ValidarXMLCTe(ByVal CTe As IXMLDOMNode) As Boolean
    If Not CTe.SelectSingleNode("cteProc") Is Nothing Then ValidarXMLCTe = True
End Function

Public Function ValidarXMLCFe(ByVal CFe As IXMLDOMNode) As Boolean
    If Not CFe.SelectSingleNode("CFe") Is Nothing Then ValidarXMLCFe = True
End Function

Public Function ValidarTag(ByVal NFe As IXMLDOMNode, ByVal Tag As String) As String
    If Not NFe.SelectSingleNode(Tag) Is Nothing Then ValidarTag = NFe.SelectSingleNode(Tag).text
End Function

Public Function SetarTag(ByVal NFe As IXMLDOMNode, Tag As String) As IXMLDOMNode
    If Not NFe.SelectSingleNode(Tag) Is Nothing Then Set SetarTag = NFe.SelectSingleNode(Tag)
End Function

Public Function ValidarnItem(ByVal NFe As IXMLDOMNode) As String
    
    If Not NFe.SelectSingleNode("@nItem") Is Nothing Then
        ValidarnItem = NFe.SelectSingleNode("@nItem").text
    ElseIf Not NFe.SelectSingleNode("nItem") Is Nothing Then
         ValidarnItem = NFe.SelectSingleNode("nItem").text
    End If
    
End Function

Public Function ValidarValores(ByVal NFe As IXMLDOMNode, ByVal Tag As String) As Double
    If Not NFe.SelectSingleNode(Tag) Is Nothing Then ValidarValores = Replace(NFe.SelectSingleNode(Tag).text, ".", ",")
End Function

Public Function ExtrairCNPJEmitente(ByVal NFe As IXMLDOMNode)

    If Not NFe.SelectSingleNode("//emit/CNPJ") Is Nothing Then
        ExtrairCNPJEmitente = VBA.Format(NFe.SelectSingleNode("//emit/CNPJ").text, String(14, "0"))
    ElseIf Not NFe.SelectSingleNode("//emit/idEstrangeiro") Is Nothing Then
        ExtrairCNPJEmitente = NFe.SelectSingleNode("//emit/idEstrangeiro").text
    ElseIf Not NFe.SelectSingleNode("//emit/CPF") Is Nothing Then
        ExtrairCNPJEmitente = NFe.SelectSingleNode("//emit/CPF").text
    End If

End Function

Public Function ExtrairCNPJContribuinte(ByVal NFe As IXMLDOMNode, ByVal tpCont As String) As String

    If Not NFe.SelectSingleNode("//" & tpCont & "/CNPJ") Is Nothing Then
        ExtrairCNPJContribuinte = VBA.Format(NFe.SelectSingleNode("//" & tpCont & "/CNPJ").text, String(14, "0"))
    End If

End Function

Public Function ExtrairCNPJDestinatario(ByVal NFe As IXMLDOMNode) As String

    If Not NFe.SelectSingleNode("//dest/CNPJ") Is Nothing Then
        ExtrairCNPJDestinatario = VBA.Format(NFe.SelectSingleNode("//dest/CNPJ").text, String(14, "0"))
    ElseIf Not NFe.SelectSingleNode("//dest/idEstrangeiro") Is Nothing Then
        ExtrairCNPJDestinatario = NFe.SelectSingleNode("//dest/idEstrangeiro").text
    ElseIf Not NFe.SelectSingleNode("//dest/CPF") Is Nothing Then
        ExtrairCNPJDestinatario = NFe.SelectSingleNode("//dest/CPF").text
    End If

End Function

Public Function ExtrairCST_CSOSN_ICMS(ByVal Produto As IXMLDOMNode) As String

Dim orig As String
Dim CST As String
Dim CSOSN As String
    
    orig = ExtrairDigitoOrigemCST_CSOSN_ICMS(Produto)
    CST = ExtrairTabelaBdoCST_ICMS(Produto)
    If Util.VerificarStringVazia(CST) Then CSOSN = ExtrairTabelaBdoCSOSN(Produto)
    
    ExtrairCST_CSOSN_ICMS = orig & CST & CSOSN
    
End Function

Public Function ExtrairDigitoOrigemCST_CSOSN_ICMS(ByVal Produto As IXMLDOMNode) As String

Dim orig As String
Dim Tags As Variant, Tag
    
    Tags = Array("orig", "Orig")
    
    For Each Tag In Tags
        
        Tag = "imposto/ICMS//" & Tag
        orig = ValidarTag(Produto, Tag)
        If Not Util.VerificarStringVazia(orig) Then
            
            ExtrairDigitoOrigemCST_CSOSN_ICMS = ValidarTag(Produto, Tag)
            Exit For
            
        End If
        
    Next Tag
    
End Function

Public Function ExtrairTabelaBdoCST_ICMS(ByVal Produto As IXMLDOMNode) As String

Dim Tag As String
    
    Tag = "imposto/ICMS//CST"
    If Not IsEmpty(ValidarTag(Produto, Tag)) Then ExtrairTabelaBdoCST_ICMS = ValidarTag(Produto, Tag)
    
End Function

Public Function ExtrairTabelaBdoCSOSN(ByVal Produto As IXMLDOMNode) As String

Dim Tag As String
    
    Tag = "imposto/ICMS//CSOSN"
    If Not IsEmpty(ValidarTag(Produto, Tag)) Then ExtrairTabelaBdoCSOSN = ValidarTag(Produto, Tag)
    
End Function

Public Function ExtrairTabelaBdoCST_CSOSN_ICMS(ByVal Produto As IXMLDOMNode)

Dim CST As String
Dim Tags As Variant, Tag
    
    Tags = Array("CST", "CSOSN")
    
    For Each Tag In Tags
        
        Tag = "imposto/ICMS//" & Tag
        If Not IsEmpty(ValidarTag(Produto, Tag)) Then
            
            ExtrairTabelaBdoCST_CSOSN_ICMS = ValidarTag(Produto, Tag)
            Exit For
            
        End If
        
    Next Tag
    
End Function

Public Function ExtrairCodigoBarrasProduto(ByRef Produto As IXMLDOMNode)

Dim Tags As Variant, Tag
Dim codBarras As String
    
    Tags = Array("cEANTrib", "cBarra")
    
    For Each Tag In Tags
        
        If Not IsEmpty(ValidarTag(Produto, "prod/" & Tag)) Then
            
            codBarras = Util.ApenasNumeros(ValidarTag(Produto, "prod/" & Tag))
            If codBarras <> "" Then
                
                codBarras = codBarras * 1
                If RegrasFiscais.Geral.CodigoBarras.ValidarCodigoBarras(codBarras) Then
                    
                    ExtrairCodigoBarrasProduto = "'" & codBarras
                    Exit For
                    
                End If
            
            End If
            
        End If
        
    Next Tag
    
End Function

Public Sub GerarDicionarioProdutos(ByVal Arqs As Variant, ByVal Planilha As Worksheet, ByRef Dicionario As Dictionary)

Dim dicProdutos As New Dictionary
Dim Produtos As IXMLDOMNodeList
Dim Produto As IXMLDOMNode
Dim NFe As New DOMDocument60
Dim Arq As Variant
    
    If VarType(Arqs) <> 11 Then
    
        For Each Arq In Arqs
            
            Set NFe = fnXML.RemoverNamespaces(Arq)
            If ValidarXMLNFe(NFe) Then
            
                With DadosDoce
                    
                    .CNPJEmit = ExtrairCNPJEmitente(NFe)
                    .RazaoEmit = ValidarTag(NFe, "//emit/xNome")
                    
                    Set Produtos = NFe.SelectNodes("//det")
                    For Each Produto In Produtos
                        .cProd = Produto.SelectSingleNode("prod/cProd").text
                        .xProd = Produto.SelectSingleNode("prod/xProd").text
                        .cBarra = fnXML.ExtrairCodigoBarrasProduto(Produto)
                        .uCom = Produto.SelectSingleNode("prod/uCom").text
                        .Chave = Cripto.MD5(.CNPJEmit & .cProd)
                        
                        If Not Dicionario.Exists(.Chave) Then dicProdutos(.Chave) = Array(.Chave, .CNPJEmit, .RazaoEmit, "'" & .cProd, .xProd, "'" & .cBarra, .uCom)
                    
                    Next Produto
                    
                End With
                
            End If
            
        Next Arq
        
        Call Util.AdicionarDadosDicionario(Planilha, dicProdutos)
        
    End If
    
End Sub

Public Function ValidarSituacao(ByVal CodigoSituacao As String)
    
    Select Case CodigoSituacao
    
        Case "100", "150"
            ValidarSituacao = "Autorizada"
        
        Case "101", "151"
            ValidarSituacao = "Cancelada"
            
        Case "102"
            ValidarSituacao = "Inutilizada"
            
        Case "110", "302"
            ValidarSituacao = "Denegada"
        
        Case ""
            ValidarSituacao = "Sem Autorização"
        
    End Select
    
End Function

Public Function ValidarOperacao(ByVal tpOperacao As String)
    
    Select Case tpOperacao
        
        Case "0"
            ValidarOperacao = "Entrada"
            
        Case "1"
            ValidarOperacao = "Saida"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_D100_IND_OPER(ByVal tpOperacao As String)
    
    Select Case tpOperacao
        
        Case "2"
            ValidarEnumeracao_D100_IND_OPER = "Entrada"
            
        Case "0", "1", "2"
            ValidarEnumeracao_D100_IND_OPER = "Saida"
            
    End Select
    
End Function

Private Function ExtrairEmitenteDestinatario(ByRef NFe As DOMDocument60, ByVal CNPJEmit As String, ByVal tpNF As String)

Dim tpPart As String
    
    tpPart = "emit"
    If CNPJContribuinte = DadosDoce.CNPJEmit Then tpPart = "dest" 'And tpNF = "Saida"
        
    If Not NFe.SelectSingleNode("//" & tpPart & "/CNPJ") Is Nothing Then
        ExtrairEmitenteDestinatario = NFe.SelectSingleNode("//" & tpPart & "/CNPJ").text
    ElseIf Not NFe.SelectSingleNode("//" & tpPart & "/CPF") Is Nothing Then
        ExtrairEmitenteDestinatario = NFe.SelectSingleNode("//" & tpPart & "/CPF").text
    ElseIf Not NFe.SelectSingleNode("//" & tpPart & "/idEstrangeiro") Is Nothing Then
        ExtrairEmitenteDestinatario = NFe.SelectSingleNode("//" & tpPart & "/idEstrangeiro").text
    End If
    
End Function

Public Function ValidarRazaoEmitenteDestinatario(ByRef NFe As DOMDocument60, ByVal CNPJEmit As String, ByVal tpNF As String)

Dim tpPart As String
    
    tpPart = "emit"
    If CNPJContribuinte = DadosDoce.CNPJEmit Then tpPart = "dest"  'And tpNF = "Saida"
    If Not NFe.SelectSingleNode("//" & tpPart & "/xNome") Is Nothing Then ValidarRazaoEmitenteDestinatario = NFe.SelectSingleNode("//" & tpPart & "/xNome").text
    
End Function

Public Function ExtrairDataDocumento(ByRef NFe As DOMDocument60)
    
    Select Case True
        
        Case Not NFe.SelectSingleNode("//dhEmi") Is Nothing
            ExtrairDataDocumento = VBA.Left(ValidarTag(NFe, "//dhEmi"), 10)
            
        Case Not NFe.SelectSingleNode("//dEmi") Is Nothing
            ExtrairDataDocumento = VBA.Left(ValidarTag(NFe, "//dEmi"), 10)
            
    End Select
    
End Function

Public Function ExtrairTipoPagamento(ByRef NFe As DOMDocument60, ByVal DT_DOC As String)

Dim tPag As String
    
    tPag = ValidarTag(NFe, "//tPag")
    tPag = EnumerarTipoPagamento(tPag)
    tPag = VerificarDataPagamento(DT_DOC, tPag)
    tPag = VerificarDadosPagamento(NFe, tPag)
    
    ExtrairTipoPagamento = EnumFiscal.ValidarEnumeracao_IND_PGTO(tPag)
    
End Function

Private Function VerificarDadosPagamento(ByRef NFe As DOMDocument60, ByVal tPag As String) As String

Dim Dup As IXMLDOMNodeList
Dim detPag As IXMLDOMNodeList
    
    Set Dup = NFe.SelectNodes("//dup")
    Set detPag = NFe.SelectNodes("//detPag")
    
    If Dup.Length > 1 Or detPag.Length > 1 Then tPag = "1"
    VerificarDadosPagamento = tPag
    
End Function

Private Function VerificarDataPagamento(ByVal DT_DOC As String, ByVal tPag As String) As String
    
    If DT_DOC <> "" Then If CDate(DT_DOC) < CDate("2012-07-01") And tPag = "2" Then tPag = "9"
    VerificarDataPagamento = tPag
    
End Function

Public Function ExtrairDataEntradaSaida(ByRef NFe As IXMLDOMNode)
    
    ExtrairDataEntradaSaida = VBA.Left(ValidarTag(NFe, "//dhSaiEnt"), 10)
    
End Function

Public Function CalcularValorTotalItemImposto(NFe As IXMLDOMNode) As Double

Dim Tag As Variant
    
    For Each Tag In Split("vProd,vFrete,vSeg,vOutro,vDesc,vIPI,vICMSST,vFCPST", ",")
        
        Select Case Tag
            
            Case "vDesc"
                If Not NFe.SelectSingleNode("prod/" & Tag) Is Nothing Then
                    CalcularValorTotalItemImposto = CalcularValorTotalItemImposto - Replace(NFe.SelectSingleNode("prod/" & Tag).text, ".", ",")
                End If
                
            Case "vIPI"
                If Not NFe.SelectSingleNode("imposto/IPI//" & Tag) Is Nothing Then
                    CalcularValorTotalItemImposto = CalcularValorTotalItemImposto + Replace(NFe.SelectSingleNode("imposto/IPI//" & Tag).text, ".", ",")
                End If
                
            Case "vICMSST", "vFCPST"
                If Not NFe.SelectSingleNode("imposto/ICMS//" & Tag) Is Nothing Then
                    CalcularValorTotalItemImposto = CalcularValorTotalItemImposto + Replace(NFe.SelectSingleNode("imposto/ICMS//" & Tag).text, ".", ",")
                End If
                
            Case Else
                If Not NFe.SelectSingleNode("prod/" & Tag) Is Nothing Then
                    CalcularValorTotalItemImposto = CalcularValorTotalItemImposto + Replace(NFe.SelectSingleNode("prod/" & Tag).text, ".", ",")
                End If
                
        End Select
        
    Next Tag
    
End Function

Public Function CalcularValorTotalItem(NFe As IXMLDOMNode) As Double

Dim Tag As Variant
    
    For Each Tag In Split("vProd,vFrete,vSeg,vOutro,vDesc", ",")
        
        Select Case Tag
            
            Case "vDesc"
                If Not NFe.SelectSingleNode("prod/" & Tag) Is Nothing Then
                    CalcularValorTotalItem = CalcularValorTotalItem - Replace(NFe.SelectSingleNode("prod/" & Tag).text, ".", ",")
                End If
                
            Case Else
                If Not NFe.SelectSingleNode("prod/" & Tag) Is Nothing Then
                    CalcularValorTotalItem = CalcularValorTotalItem + Replace(NFe.SelectSingleNode("prod/" & Tag).text, ".", ",")
                End If
                
        End Select
        
    Next Tag
    
    CalcularValorTotalItem = Round(CalcularValorTotalItem, 2)
    
End Function

Public Function ListarDadosDicionario(ByVal Planilha As Worksheet, ByVal Coluna As String, ByVal LinhaFinal As Long) As Dictionary

Dim Dicionario As New Dictionary
Dim Chaves, Chave

    Chaves = Planilha.Range(Coluna & "2:" & Coluna & LinhaFinal)

    For Each Chave In Chaves
        Dicionario(VBA.Right(Chave, 44)) = ""
    Next Chave
    
    Set ListarDadosDicionario = Dicionario
    
End Function

Public Function ValidarUF(ByRef Doc As IXMLDOMNode, ByVal tpXML As String, ByVal Modelo As String) As String
    
    Select Case tpXML
        
        Case "Entrada"
            If Not Doc.SelectSingleNode("//enderEmit/UF") Is Nothing Then ValidarUF = Doc.SelectSingleNode("//enderEmit/UF").text
            
        Case "Saida"
            If Not Doc.SelectSingleNode("//enderDest/UF") Is Nothing Then ValidarUF = Doc.SelectSingleNode("//enderDest/UF").text
            
    End Select
    
    If ValidarUF = "" And Modelo = "65" Then ValidarUF = Doc.SelectSingleNode("//enderEmit/UF").text
    
End Function

Public Function ValidarUFNFeNFCe(ByRef Doc As IXMLDOMNode, ByVal CNPJEmit As String, ByVal CNPJDest As String) As String
    If CNPJEmit = CNPJContribuinte Then If Not Doc.SelectSingleNode("//enderDest/UF") Is Nothing Then ValidarUFNFeNFCe = Doc.SelectSingleNode("//enderDest/UF").text
    If CNPJDest = CNPJContribuinte Then If Not Doc.SelectSingleNode("//enderEmit/UF") Is Nothing Then ValidarUFNFeNFCe = Doc.SelectSingleNode("//enderEmit/UF").text
End Function

Public Function ValidarUFCTe(ByRef Doc As IXMLDOMNode, ByVal CNPJEmit As String, ByVal CNPJTomador As String) As String
    If CNPJEmit = CNPJContribuinte Then If Not Doc.SelectSingleNode("//UFFim") Is Nothing Then ValidarUFCTe = Doc.SelectSingleNode("//UFFim").text
    If CNPJTomador = CNPJContribuinte Then If Not Doc.SelectSingleNode("//UFIni") Is Nothing Then ValidarUFCTe = Doc.SelectSingleNode("//UFIni").text
End Function

Public Function ValidarTomador(ByVal CodigoSituacao As String)
    
    Select Case CodigoSituacao
    
        Case "0"
            ValidarTomador = "Remetente"
        
        Case "1"
            ValidarTomador = "Expedidor"
            
        Case "2"
            ValidarTomador = "Recebedor"
            
        Case "3"
            ValidarTomador = "Destinatário"
            
        Case "4"
            ValidarTomador = "Outros"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_TP_CT_E(ByVal tpCTe As String)
    
    Select Case tpCTe
        
        Case "0"
            ValidarEnumeracao_TP_CT_E = "0 - CT-e ou BP-e Normal"
        
        Case "1"
            ValidarEnumeracao_TP_CT_E = "1 - CT-e de Complemento de Valores"
        
        Case "2"
            ValidarEnumeracao_TP_CT_E = "2 - CT-e emitido em hipótese de anulação de débito"
        
        Case "3"
            ValidarEnumeracao_TP_CT_E = "3 - CTE substituto do CT-e anulado ou BP-e substituição"
    
    End Select
    
End Function

Public Function ValidarEnumeracao_IND_FRTCTe(ByRef CTe As IXMLDOMNode)
    
Dim Tomador As String
                
    If Not CTe.SelectSingleNode("//toma") Is Nothing Then Tomador = CTe.SelectSingleNode("//toma").text
    
    Select Case Tomador
        
        Case "0"
            ValidarEnumeracao_IND_FRTCTe = "1 - Por conta do destinatário/remetente"
        
        Case "1"
            ValidarEnumeracao_IND_FRTCTe = "2 - Por conta de terceiros"
        
        Case "2"
            ValidarEnumeracao_IND_FRTCTe = "2 - Por conta de terceiros"
        
        Case "3"
            ValidarEnumeracao_IND_FRTCTe = "1 - Por conta do destinatário/remetente"
        
        Case "4"
            ValidarEnumeracao_IND_FRTCTe = "2 - Por conta de terceiros"
            
    End Select
        
End Function

Public Function ExtrairCNPJTomador(ByRef CTe As IXMLDOMNode) As String
    
Dim Tomador As String
                
    If Not CTe.SelectSingleNode("//toma") Is Nothing Then
        
        Tomador = CTe.SelectSingleNode("//toma").text
    
        Select Case Tomador
        
            Case "0"
                Tomador = "rem"
            
            Case "1"
                Tomador = "exped"
                
            Case "2"
                Tomador = "receb"
                
            Case "3"
                Tomador = "dest"
                
            Case "4"
                Tomador = "toma4"
                
        End Select
    
        If Not CTe.SelectSingleNode("//" & Tomador & "/CNPJ") Is Nothing Then ExtrairCNPJTomador = CTe.SelectSingleNode("//" & Tomador & "/CNPJ").text
    
    End If
    
End Function

Public Function AlterarCNPJTomador(ByRef CTe As IXMLDOMNode, ByVal CNPJ As String, ByVal CNPJRef As String, ByVal Razao As String, ByVal FANTASIA As String) As Boolean
    
Dim Tomador As String
                
    If Not CTe.SelectSingleNode("//toma") Is Nothing Then
        
        Tomador = CTe.SelectSingleNode("//toma").text
    
        Select Case Tomador
        
            Case "0"
                Tomador = "rem"
            
            Case "1"
                Tomador = "exped"
                
            Case "2"
                Tomador = "receb"
                
            Case "3"
                Tomador = "dest"
                
            Case "4"
                Tomador = "toma4"
                
        End Select
    
        If Not CTe.SelectSingleNode("//" & Tomador & "/CNPJ") Is Nothing Then
            
            If CNPJRef = CTe.SelectSingleNode("//" & Tomador & "/CNPJ").text Then
                CTe.SelectSingleNode("//" & Tomador & "/CNPJ").text = CNPJ
                If Not CTe.SelectSingleNode("//" & Tomador & "/xNome") Is Nothing Then CTe.SelectSingleNode("//" & Tomador & "/xNome").text = Razao
                If Not CTe.SelectSingleNode("//" & Tomador & "/xFant") Is Nothing Then CTe.SelectSingleNode("//" & Tomador & "/xFant").text = FANTASIA
                
                AlterarCNPJTomador = True
                
            End If
        
        End If
        
    End If
    
End Function

Public Function ValidarNFe(ByRef NFe As DOMDocument60) As Boolean
    
    Select Case Util.RemoverNamespaces(NFe.DocumentElement.nodeName)
        
        Case "nfeProc", "procNfe", "NFe", "retConsNFeLog"
            ValidarNFe = ValidarAmbiente(NFe)
            
        Case "proc"
            If Not NFe.SelectSingleNode("proc/nfeProc") Is Nothing Then ValidarNFe = True
            
        Case "NFeDFe"
            If Not NFe.SelectSingleNode("NFeDFe/nfeProc") Is Nothing Then ValidarNFe = ValidarAmbiente(NFe)
            
    End Select
    
End Function

Private Function ValidarAmbiente(ByRef Doce As DOMDocument60) As Boolean

Dim Ambiente As String
    
    If Not Doce.SelectSingleNode("//tpAmb") Is Nothing Then
        
        'Verifica o ambiente em que a nota foi emitida
        Ambiente = ValidarTag(Doce, "//tpAmb")
        If Ambiente = "1" Or chNFSemValidade Then
            ValidarAmbiente = True
        Else
            DocsSemValidade = DocsSemValidade + 1
        End If
        
    End If
    
End Function

Public Function ValidarCFe(ByRef CFe As DOMDocument60) As Boolean

    Select Case Util.RemoverNamespaces(CFe.DocumentElement.nodeName)
                
        Case "CFeProc", "CFe"
            ValidarCFe = ValidarAmbiente(CFe)
            
    End Select
    
End Function

Public Function ValidarCTe(ByRef CTe As DOMDocument60) As Boolean
    
    Select Case Util.RemoverNamespaces(CTe.DocumentElement.nodeName)
                
        Case "cteProc", "CTe"
            ValidarCTe = ValidarAmbiente(CTe)
            
    End Select
    
End Function

Public Function ValidarCTe_Old(ByRef NFe As IXMLDOMNode) As Boolean
    If Not NFe.SelectSingleNode("cteProc") Is Nothing Then ValidarCTe_Old = True
End Function

Public Function ValidarNFSe(ByRef NFSe As DOMDocument60) As Boolean
    
    Select Case Util.RemoverNamespaces(NFSe.DocumentElement.nodeName)
        
        Case "ListaNfse", "CompNfse", "ConsultarNfseResposta", "ConsultarNfseFaixaResposta"
            ValidarNFSe = True
            
    End Select

End Function

Public Function ValidarPercentual(ByRef NFe As IXMLDOMNode, ByVal Tag As String) As Double
    If Not NFe.SelectSingleNode(Tag) Is Nothing Then ValidarPercentual = Replace(NFe.SelectSingleNode(Tag).text, ".", ",") / 100
End Function

Public Function IdentificarParticipante(ByRef NFe As IXMLDOMNode, ByVal tpNF As String, ByVal tpEmit As String)

Dim CNPJEmit As String
Dim CNPJDest As String
    
    CNPJEmit = ExtrairCNPJEmitente(NFe)
    CNPJDest = ExtrairCNPJDestinatario(NFe)
    
    Select Case True
        
        Case (tpNF = "1" And tpEmit = "0") Or (tpNF = "1" And tpEmit = "1") Or (tpNF = "0" And tpEmit = "0")
            IdentificarParticipante = CNPJDest
            
        Case tpNF = "0" And tpEmit = "1"
            IdentificarParticipante = CNPJEmit
            
    End Select
    
End Function

Public Function IdentificarParticipanteCTe(ByRef CTe As IXMLDOMNode)

Dim CNPJEmit As String
Dim CNPJToma As String

    CNPJEmit = ExtrairCNPJEmitente(CTe)
    CNPJToma = ExtrairCNPJTomador(CTe)
    
    Select Case True
            
        Case (CNPJEmit = CNPJContribuinte)
            IdentificarParticipanteCTe = CNPJToma
            
        Case (CNPJToma = CNPJContribuinte)
            IdentificarParticipanteCTe = CNPJEmit
        
        Case Else
            IdentificarParticipanteCTe = ""
            
    End Select
    
End Function

Public Function IdentificarTipoEmissao(ByRef NFe As IXMLDOMNode, Optional ByVal SPEDContrib As Boolean) As String

Dim CNPJEmit As String
Dim CNPJDest As String

    CNPJEmit = ExtrairCNPJEmitente(NFe)
    CNPJDest = ExtrairCNPJDestinatario(NFe)
    
    If SPEDContrib Then
        
        Select Case True
            
            Case (CNPJEmit Like CNPJBase & "*")
                IdentificarTipoEmissao = "0"
                
            Case Else
                IdentificarTipoEmissao = "1"
                
        End Select
    
    Else
        
        Select Case True
            
            Case (CNPJEmit = CNPJContribuinte)
                IdentificarTipoEmissao = "0"
                
            Case Else
                IdentificarTipoEmissao = "1"
                
        End Select
        
    End If
    
End Function

Public Function ExtrairTipoEmissao(ByRef NFe As IXMLDOMNode, Optional ByVal SPEDContrib As Boolean) As String

Dim CNPJEmit As String
Dim CNPJDest As String

    CNPJEmit = ExtrairCNPJEmitente(NFe)
    CNPJDest = ExtrairCNPJDestinatario(NFe)
    
    If SPEDContrib Then
        
        Select Case True
            
            Case (CNPJEmit Like CNPJBase & "*")
                ExtrairTipoEmissao = EnumFiscal.ValidarEnumeracao_IND_EMIT("0")
                
            Case Else
                ExtrairTipoEmissao = EnumFiscal.ValidarEnumeracao_IND_EMIT("1")
                
        End Select
        
    Else
        
        Select Case True
            
            Case (CNPJEmit = CNPJContribuinte)
                ExtrairTipoEmissao = EnumFiscal.ValidarEnumeracao_IND_EMIT("0")
                
            Case Else
                ExtrairTipoEmissao = EnumFiscal.ValidarEnumeracao_IND_EMIT("1")
                
        End Select
        
    End If
    
End Function

Public Function ExtrairParticipante(ByRef NFe As IXMLDOMNode) As String

Dim CNPJEmit As String
Dim CNPJDest As String
Dim CNPJPart As String
Dim tpPart As String
    
    CNPJEmit = ExtrairCNPJEmitente(NFe)
    CNPJDest = ExtrairCNPJDestinatario(NFe)
    
    Select Case True
        
        Case CNPJContribuinte = CNPJDest
            ExtrairParticipante = CNPJEmit
            CNPJPart = CNPJEmit
            tpPart = "emit"
            
        Case CNPJContribuinte = CNPJEmit
            ExtrairParticipante = CNPJDest
            CNPJPart = CNPJDest
            tpPart = "dest"
            
        Case Else
            Exit Function
            
    End Select
    
    Call IncluirRegistro0150(NFe, CNPJPart, tpPart)
    
End Function

Public Function ExtrairInscricaoParticipante(ByRef NFe As IXMLDOMNode) As String

Dim IEEmit As String
Dim IEDest As String
Dim tpPart As String
Dim CNPJEmit As String
Dim CNPJDest As String
        
    CNPJEmit = ExtrairCNPJEmitente(NFe)
    CNPJDest = ExtrairCNPJDestinatario(NFe)
    
    Select Case True
        
        Case CNPJContribuinte = CNPJDest
            tpPart = "emit"
            
        Case CNPJContribuinte = CNPJEmit
            tpPart = "dest"
            
    End Select
    
    If tpPart = "" Then Exit Function
    
    Call IncluirRegistro0150(NFe, CNPJDest, tpPart)
    ExtrairInscricaoParticipante = "'" & ValidarTag(NFe, "//" & tpPart & "/IE")
    
End Function

Public Function ExtrairTipoOperacao(ByRef NFe As IXMLDOMNode, Optional ByVal SPEDContrib As Boolean) As String

Dim CNPJEmit As String
Dim CNPJDest As String
Dim tpNF As String
    
    CNPJEmit = ExtrairCNPJEmitente(NFe)
    CNPJDest = ExtrairCNPJDestinatario(NFe)
    
    tpNF = ValidarTag(NFe, "//tpNF")
    
    If SPEDContrib Then
        
        Select Case True
            
            Case (tpNF = "1" And CNPJEmit = CNPJDest And CNPJEmit Like CNPJBase & "*")
                ExtrairTipoOperacao = EnumFiscal.ValidarEnumeracao_IND_OPER("1")
                
            Case (tpNF = "1" And CNPJEmit Like CNPJBase & "*")
                ExtrairTipoOperacao = EnumFiscal.ValidarEnumeracao_IND_OPER("1")
                
            Case (tpNF = "0" And CNPJEmit = CNPJDest And CNPJEmit Like CNPJBase & "*")
                ExtrairTipoOperacao = EnumFiscal.ValidarEnumeracao_IND_OPER("0")
                
            Case Else
                ExtrairTipoOperacao = EnumFiscal.ValidarEnumeracao_IND_OPER("0")
                
        End Select
        
    Else
        
        Select Case True
            
            Case (tpNF = "1" And CNPJEmit = CNPJContribuinte And CNPJDest = CNPJContribuinte)
                ExtrairTipoOperacao = EnumFiscal.ValidarEnumeracao_IND_OPER("1")
                
            Case (CNPJEmit = CNPJContribuinte And tpNF = "1")
                ExtrairTipoOperacao = EnumFiscal.ValidarEnumeracao_IND_OPER("1")
                
            Case (tpNF = "0" And CNPJEmit = CNPJContribuinte And CNPJDest = CNPJContribuinte)
                ExtrairTipoOperacao = EnumFiscal.ValidarEnumeracao_IND_OPER("0")
                
            Case Else
                ExtrairTipoOperacao = EnumFiscal.ValidarEnumeracao_IND_OPER("0")
                
        End Select
        
    End If
    
End Function

Public Function ExtrairNomeRazaoParticipante(ByRef NFe As IXMLDOMNode, ByVal COD_MOD As String, ByVal COD_SIT As String, Optional ByVal SPEDContrib As Boolean)

Dim tpNF As String
Dim tpPart As String
Dim CNPJEmit As String
        
    tpPart = "emit"
    tpNF = ValidarTag(NFe, "//tpNF")
    CNPJEmit = ExtrairCNPJEmitente(NFe)
    
    If CNPJContribuinte = CNPJEmit Then tpPart = "dest"
    If Not NFe.SelectSingleNode("//" & tpPart & "/xNome") Is Nothing Then _
        ExtrairNomeRazaoParticipante = VBA.Left(ValidarTag(NFe, "//" & tpPart & "/xNome"), 100)
        
End Function

Public Function IdentificarTipoEmissaoCTe(ByRef CTe As IXMLDOMNode, Optional ByVal SPEDContrib As Boolean) As String

Dim CNPJEmit As String
Dim CNPJToma As String

    CNPJEmit = ExtrairCNPJEmitente(CTe)
    CNPJToma = ExtrairCNPJTomador(CTe)
    
    If SPEDContrib Then
        
        Select Case True
            
            Case (CNPJEmit Like CNPJBase & "*")
                IdentificarTipoEmissaoCTe = "0"
                
            Case Else
                IdentificarTipoEmissaoCTe = "1"
                
        End Select
    
    Else
        
        Select Case True
            
            Case (CNPJEmit = CNPJContribuinte)
                IdentificarTipoEmissaoCTe = "0"
                
            Case Else
                IdentificarTipoEmissaoCTe = "1"
                
        End Select
        
    End If
    
End Function

Public Function IdentificarTipoOperacao(ByRef NFe As IXMLDOMNode, ByVal tpNF As String, Optional ByVal SPEDContrib As Boolean) As String

Dim CNPJEmit As String
Dim CNPJDest As String
    
    CNPJEmit = ExtrairCNPJEmitente(NFe)
    CNPJDest = ExtrairCNPJDestinatario(NFe)
    
    If SPEDContrib Then
        
        Select Case True
            
            Case (tpNF = "1" And CNPJEmit = CNPJDest And CNPJEmit Like CNPJBase & "*")
                IdentificarTipoOperacao = "1"
                
            Case (tpNF = "1" And CNPJEmit Like CNPJBase & "*")
                IdentificarTipoOperacao = "1"
                
            Case (tpNF = "0" And CNPJEmit = CNPJDest And CNPJEmit Like CNPJBase & "*")
                IdentificarTipoOperacao = "0"
                
            Case Else
                IdentificarTipoOperacao = "0"
                
        End Select
        
    Else
        
        Select Case True
            
            Case (tpNF = "1" And CNPJEmit = CNPJContribuinte And CNPJDest = CNPJContribuinte)
                IdentificarTipoOperacao = "1"
                
            Case (CNPJEmit = CNPJContribuinte And tpNF = "1")
                IdentificarTipoOperacao = "1"
                
            Case (tpNF = "0" And CNPJEmit = CNPJContribuinte And CNPJDest = CNPJContribuinte)
                IdentificarTipoOperacao = "0"
                
            Case Else
                IdentificarTipoOperacao = "0"
                
        End Select
        
    End If
    
End Function

Public Function IdentificarTipoOperacaoCTe(ByRef CTe As IXMLDOMNode, ByVal tpCT As String, Optional ByVal SPEDContrib As Boolean) As String

Dim CNPJEmit As String
Dim CNPJToma As String
    
    CNPJEmit = ExtrairCNPJEmitente(CTe)
    CNPJToma = ExtrairCNPJTomador(CTe)
    
    If SPEDContrib Then
        
        Select Case True
            
            Case (tpCT = "1" And CNPJEmit = CNPJToma And CNPJEmit Like CNPJBase & "*")
                IdentificarTipoOperacaoCTe = "1"
                
            Case (tpCT = "1" And CNPJEmit Like CNPJBase & "*")
                IdentificarTipoOperacaoCTe = "1"
                
            Case (tpCT = "0" And CNPJEmit = CNPJToma And CNPJEmit Like CNPJBase & "*")
                IdentificarTipoOperacaoCTe = "0"
                
            Case Else
                IdentificarTipoOperacaoCTe = "0"
                
        End Select
        
    Else
        
        Select Case True
            
            Case tpCT <> "2" And CNPJEmit = CNPJContribuinte
                IdentificarTipoOperacaoCTe = "1"
                
            Case tpCT = "2" And CNPJEmit = CNPJContribuinte
                IdentificarTipoOperacaoCTe = "0"
                
            Case tpCT <> "2" And CNPJToma = CNPJContribuinte
                IdentificarTipoOperacaoCTe = "0"
                
            Case Else
                IdentificarTipoOperacaoCTe = "0"
                
        End Select
        
    End If
    
End Function

Public Function ValidarCFOPEntrada(ByRef CFOP As String) As String
    
    Select Case True
        
        Case CFOP Like "5###"
            ValidarCFOPEntrada = "1" & VBA.Right(CFOP, 3)
            
        Case CFOP Like "6###"
            ValidarCFOPEntrada = "2" & VBA.Right(CFOP, 3)
            
        Case Else
            ValidarCFOPEntrada = CFOP
            
    End Select
    
End Function

Public Function AjustarCFOPEntrada(ByRef CFOP As String) As String

    Select Case True
        
        Case CFOP Like "#404" Or CFOP Like "#405"
            AjustarCFOPEntrada = VBA.Left(CFOP, 1) & "403"

        Case Else
            AjustarCFOPEntrada = CFOP
            
    End Select
    
End Function

Public Function CorrelacionarProdutoFornecedor(ByVal Chave As String, _
    ByRef dicCorrelacoes As Dictionary, ByRef cProd As String, ByRef xProd As String)

    If dicCorrelacoes.Exists(Chave) Then
        cProd = dicCorrelacoes(Chave)(5)
        xProd = dicCorrelacoes(Chave)(6)
    Else
        cProd = "Código não encontrado"
        xProd = "Nenhuma correlação foi feita para este produto"
    End If
    
End Function

Public Sub CriarRegistroProdutoFornecedor(ByRef dicProdFornec As Dictionary, ByVal Arqs As ArrayList)

Dim cProd As String, xProd$, CNPJForn$, Razao$, uCom$, Chave$, CNPJDest$, CNPJBase$
Dim Produtos As IXMLDOMNodeList
Dim NFe As New DOMDocument60
Dim Produto As IXMLDOMNode
Dim Arq
    
    For Each Arq In Arqs
        
        Set NFe = fnXML.RemoverNamespaces(Arq)
        
        If ValidarXMLNFe(NFe) And ValidarParticipante(NFe, True) Then
            
            CNPJForn = fnXML.ExtrairCNPJEmitente(NFe)
            CNPJDest = fnXML.ExtrairCNPJDestinatario(NFe)
            CNPJBase = VBA.Left(CNPJDest, 8)
            
            If (CNPJDest = CNPJContribuinte And CNPJForn <> CNPJContribuinte) Or _
               (CNPJContribuinte Like CNPJBase & "*" And Not CNPJForn Like CNPJBase & "*") Then
                
                Razao = NFe.SelectSingleNode("//emit/xNome").text
                
                Set Produtos = NFe.SelectNodes("//det")
                For Each Produto In Produtos
                    
                    cProd = Produto.SelectSingleNode("prod/cProd").text

                    xProd = Produto.SelectSingleNode("prod/xProd").text
                    uCom = VBA.UCase(Produto.SelectSingleNode("prod/uCom").text)
                    
                    Chave = CNPJForn & cProd & uCom
                    If Not dicProdFornec.Exists(Chave) Then
                        dicProdFornec(Chave) = Array("'" & CNPJForn, Razao, "'" & cProd, xProd, "'" & uCom, "", "", "", "", "")
                    End If
                    
                Next Produto
                
            End If
            
        End If
        
    Next Arq
    
End Sub

Public Function ConverterTipoPagamento(ByRef tPag As String)
    
    Select Case tPag
        
        Case "02"
            ConverterTipoPagamento = "01"
            
        Case Else
            ConverterTipoPagamento = "99"
            
    End Select
    
End Function

Private Function EnumerarTipoPagamento(ByVal tPag As String) As String
    
    Select Case tPag
        
        Case "01", "04", "10", "11", "12", "13", "15", "16", "17", "18", "19"
            EnumerarTipoPagamento = "0"
            
        Case "02", "03", "05"
            EnumerarTipoPagamento = "1"
            
        Case "90", "99"
            EnumerarTipoPagamento = "2"
            
    End Select
    
End Function

Public Function ValidarDescricaoTitulo(ByVal tPag As String)
    
    Select Case tPag
        
        Case "01"
            ValidarDescricaoTitulo = "Dinheiro"
        
        Case "02"
            ValidarDescricaoTitulo = "Cheque"
                
        Case "03"
            ValidarDescricaoTitulo = "Cartão de Crédito"
                
        Case "04"
            ValidarDescricaoTitulo = "Cartão de Débito"
                
        Case "05"
            ValidarDescricaoTitulo = "Crédito Loja"
                
        Case "10"
            ValidarDescricaoTitulo = "Vale Alimentação"
                
        Case "11"
            ValidarDescricaoTitulo = "Vale Refeição"
                
        Case "12"
            ValidarDescricaoTitulo = "Vale Presente"
                
        Case "13"
            ValidarDescricaoTitulo = "Vale Combustível"
        
        Case "14"
            ValidarDescricaoTitulo = "Duplicata Mercantil"
            
        Case "15"
            ValidarDescricaoTitulo = "Boleto Bancário"
                
        Case "16"
            ValidarDescricaoTitulo = "Depósito Bancário"
                
        Case "17"
            ValidarDescricaoTitulo = "Pagamento Instantâneo(PIX)"
                
        Case "18"
            ValidarDescricaoTitulo = "Transferência bancária, Carteira Digital"
                
        Case "19"
            ValidarDescricaoTitulo = "Programa de fidelidade, Cashback, Crédito Virtual"
                
        Case "90"
            ValidarDescricaoTitulo = "Sem Pagamento"
                
        Case "99"
            ValidarDescricaoTitulo = "Outros"
    
    End Select
    
End Function

Public Function DefinirParticipanteNFe(ByVal NFe As IXMLDOMNode, Optional ByVal SPEDContrib As Boolean) As String

Dim CNPJEmit As String
Dim CNPJDest As String
Dim CNPJBase As String

    If Not NFe.SelectSingleNode("//emit/CNPJ") Is Nothing Then CNPJEmit = NFe.SelectSingleNode("//emit/CNPJ").text
    If Not NFe.SelectSingleNode("//dest/CNPJ") Is Nothing Then CNPJDest = NFe.SelectSingleNode("//dest/CNPJ").text
        
    If (CNPJEmit = CNPJContribuinte) Then DefinirParticipanteNFe = "dest"
    If (CNPJDest = CNPJContribuinte) Then DefinirParticipanteNFe = "emit"
    
    If SPEDContrib Then
        CNPJBase = VBA.Left(CNPJContribuinte, 8)
        If CNPJEmit Like CNPJBase & "*" Then DefinirParticipanteNFe = "dest"
        If CNPJDest Like CNPJBase & "*" Then DefinirParticipanteNFe = "emit"
    End If
    
End Function

Public Function DefinirContribuinteNFe(ByVal NFe As IXMLDOMNode, Optional ByVal SPEDContrib As Boolean) As String

Dim CNPJEmit As String
Dim CNPJDest As String
    
    If Not NFe.SelectSingleNode("//emit/CNPJ") Is Nothing Then CNPJEmit = NFe.SelectSingleNode("//emit/CNPJ").text
    If Not NFe.SelectSingleNode("//dest/CNPJ") Is Nothing Then CNPJDest = NFe.SelectSingleNode("//dest/CNPJ").text
    
    If (CNPJEmit = CNPJContribuinte) Then DefinirContribuinteNFe = "emit"
    If (CNPJDest = CNPJContribuinte) Then DefinirContribuinteNFe = "dest"
    
    If SPEDContrib Then
        If CNPJEmit Like CNPJBase & "*" Then DefinirContribuinteNFe = "emit"
        If CNPJDest Like CNPJBase & "*" Then DefinirContribuinteNFe = "dest"
    End If
    
End Function

Public Function DefinirContribuinteCTe(ByVal CTe As IXMLDOMNode, Optional ByVal SPEDContrib As Boolean) As String

Dim CNPJEmit As String
Dim CNPJToma As String
Dim toma As String
    
    toma = ColetarTomadorCTe(CTe)
    
    If Not CTe.SelectSingleNode("//emit/CNPJ") Is Nothing Then CNPJEmit = CTe.SelectSingleNode("//emit/CNPJ").text
    If Not CTe.SelectSingleNode("//" & toma & "/CNPJ") Is Nothing Then CNPJToma = CTe.SelectSingleNode("//" & toma & "/CNPJ").text
    
    If (CNPJEmit = CNPJContribuinte) Then DefinirContribuinteCTe = "emit"
    If (CNPJToma = CNPJContribuinte) Then DefinirContribuinteCTe = "dest"
    
    If SPEDContrib Then
        If CNPJEmit Like CNPJBase & "*" Then DefinirContribuinteCTe = "emit"
        If CNPJToma Like CNPJBase & "*" Then DefinirContribuinteCTe = toma
    End If
    
End Function

Public Function DefinirParticipanteCTe(ByVal CTe As IXMLDOMNode) As String

Dim CNPJEmit As String
Dim CNPJToma As String
    
    CNPJEmit = ValidarTag(CTe, "//emit/CNPJ")
    CNPJToma = ExtrairCNPJTomador(CTe)
    
    If (CNPJEmit = CNPJContribuinte) Then DefinirParticipanteCTe = ColetarTomadorCTe(CTe)
    If (CNPJToma = CNPJContribuinte) Then DefinirParticipanteCTe = "emit"
    
End Function

Public Function ColetarTomadorCTe(ByRef CTe As IXMLDOMNode) As String
    
Dim Tomador As String
                
    If Not CTe.SelectSingleNode("//toma") Is Nothing Then Tomador = CTe.SelectSingleNode("//toma").text
    
    Select Case Tomador
    
        Case "0"
            Tomador = "rem"
        
        Case "1"
            Tomador = "exped"
            
        Case "2"
            Tomador = "receb"
            
        Case "3"
            Tomador = "dest"
            
        Case "4"
            Tomador = "toma4"
            
    End Select
    
    If Tomador <> "" Then ColetarTomadorCTe = Tomador
    
End Function

Public Function ValidarProtocoloCancelamento(ByRef No As IXMLDOMNode) As Boolean

Dim Tags As Variant, Tag

    Tags = Array("//procEventoNFe", "//procEventoCTe", "//CFeCanc", "//evento")
    For Each Tag In Tags
        
        If Not No.SelectSingleNode(Tag) Is Nothing Then
            
            If VBA.LCase(Tag) Like "*evento*" Then
                
                Select Case ValidarTag(No, "//tpEvento")
                    
                    Case "110111", "110112"
                        
                        ValidarProtocoloCancelamento = True
                        Exit Function
                        
                End Select
                
            Else
                
                ValidarProtocoloCancelamento = True
                Exit Function
                
            End If
            
        End If
        
    Next Tag
    
End Function

Public Sub CarregarProtocolosCancelamento(ByRef XMLS As Variant, ByRef arrCanceladas As ArrayList, Optional ByVal Msg As String)

Dim Doce As New MSXML2.DOMDocument60
Dim Comeco As Double
Dim XML As Variant
Dim b As Long
    
    b = 0
    Comeco = Timer()
    For Each XML In XMLS
        
        Call Util.AntiTravamento(b, 10, Msg & "Identificando protocolos de cancelamento.", XMLS.Count, Comeco)
        
        Set Doce = fnXML.RemoverNamespaces(XML)
        
        If ValidarProtocoloCancelamento(Doce) Then
            
            If Not fnXML.ValidarXML(Doce) Then GoTo Prx:
            Select Case True
            
                Case Not Doce.SelectSingleNode("procEventoNFe") Is Nothing Or Not Doce.SelectSingleNode("evento") Is Nothing
                    arrCanceladas.Add Doce.SelectSingleNode("//chNFe").text
                
                Case Not Doce.SelectSingleNode("procEventoCTe") Is Nothing
                    arrCanceladas.Add Doce.SelectSingleNode("//chCTe").text
                 
                Case Not Doce.SelectSingleNode("CFeCanc") Is Nothing
                    arrCanceladas.Add VBA.Right(Doce.SelectSingleNode("//@chCanc").text, 44)
                    
            End Select
            
        End If
Prx:
    Next XML
    
    Application.StatusBar = False
    
End Sub

Public Sub CarregarChavesReferenciadas(ByRef XMLS As Variant, ByRef dicChavesReferenciadas As Dictionary, Optional ByVal Msg As String)

Dim ChavesReferenciadas As Variant
Dim Doce As New DOMDocument60
Dim chDoce As String
Dim Comeco As Double
Dim XML As Variant
Dim b As Long
    
    b = 0
    Comeco = Timer
    For Each XML In XMLS
        
        Call Util.AntiTravamento(b, 50, Msg & "Identificando notas devolvidas.", XMLS.Count, Comeco)
        
        'Carrega dados do XML
        Set Doce = fnXML.RemoverNamespaces(XML)
    
        If fnXML.ValidarXML(Doce) Then
            
            ChavesReferenciadas = ExtrairChavesReferenciadas(Doce)
            If Not IsEmpty(ChavesReferenciadas) Then
                
                chDoce = VBA.Right(fnXML.ValidarTag(Doce, "//@Id"), 44)
                dicChavesReferenciadas(chDoce) = ChavesReferenciadas
             
            End If
            
        End If
        
    Next XML
    
End Sub

Private Function ExtrairChavesReferenciadas(ByRef Doce As DOMDocument60) As Variant

Dim ChavesReferenciadas As IXMLDOMNodeList
Dim arrReferenciadas As New ArrayList
Dim Chave As IXMLDOMNode
    
    'Carrega lista de chaves referenciadas
    Set ChavesReferenciadas = Doce.SelectNodes("//refNFe")
    
    For Each Chave In ChavesReferenciadas
        
        'Adiciona as chaves de acesso a lista de chaves devolvidas
        arrReferenciadas.Add Chave.text
        
    Next Chave
    
    If arrReferenciadas.Count > 0 Then ExtrairChavesReferenciadas = arrReferenciadas.toArray()
    
End Function

Public Function CriarRegistro0000(ByRef Doce As IXMLDOMNode, ByRef dic0000 As Dictionary, ByRef dic0001 As Dictionary, _
    ByRef dic0005 As Dictionary, ByRef dic0100 As Dictionary, ByRef dicC001 As Dictionary, ByVal Periodo As String, ByVal tpCont As String)

Dim Campos As Variant
Dim CHV_REG As String, CHV_0001$, CHV_0100$, CHV_C001$, DT_INI$, DT_FIN$, Ano$, Mes$, COD_VER$, COD_FIN$, NOME$, CNPJ$, UF$, COD_MUN$, SUFRAMA$, CPF$, IE$
    
    Ano = VBA.Right(Periodo, 4)
    Mes = VBA.Left(Periodo, 2)
    COD_VER = "'" & ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_COD_VER(Periodo)
    COD_FIN = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_COD_FIN(0)
    DT_INI = Ano & "-" & Mes & "-01"
    DT_FIN = VBA.Format(Application.WorksheetFunction.EoMonth(Ano & "-" & Mes & "-" & "01", 0), "yyyy-mm-dd")
    NOME = ValidarTag(Doce, "//" & tpCont & "/xNome")
    CNPJ = Util.FormatarCNPJ(CNPJContribuinte)
    CPF = ""
    UF = ValidarTag(Doce, "//" & tpCont & "//UF")
    IE = Util.FormatarValores(Util.ApenasNumeros(ValidarTag(Doce, "//" & tpCont & "/IE"))) * 1
    COD_MUN = ValidarTag(Doce, "//" & tpCont & "//cMun")
    SUFRAMA = ValidarTag(Doce, "//" & tpCont & "/SUFRAMA")
    ARQUIVO = VBA.Format(Periodo, "mm/yyyy") & "-" & CNPJ

    CHV_REG = fnSPED.GerarChaveRegistro("", VBA.Format(DT_INI, "ddmmyyyy"), VBA.Format(DT_FIN, "ddmmyyyy"), CNPJ, CPF, IE)
    Campos = Array("'0000", ARQUIVO, CHV_REG, "", "", COD_VER, COD_FIN, CDate(DT_INI), CDate(DT_FIN), NOME, "'" & CNPJ, CPF, UF, "'" & IE, COD_MUN, "", SUFRAMA, "", "")
    dic0000(ARQUIVO) = Campos
    
    CHV_0001 = fnSPED.GerarChaveRegistro(CHV_REG, "0001")
    Campos = Array("'0001", ARQUIVO, CHV_0001, CHV_REG, "", ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_MOV(0))
    dic0001(ARQUIVO) = Campos
    
    Call CriarRegistro0005(Doce, dic0005, dic0001(ARQUIVO)(2), Periodo, tpCont)
    
    CHV_0100 = fnSPED.GerarChaveRegistro(CHV_REG, "0100")
    If Not dic0100.Exists(ARQUIVO) Then
        Campos = Array("'0100", ARQUIVO, CHV_0100, CHV_0001, "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        dic0100(ARQUIVO) = Campos
    End If
    
    CHV_C001 = fnSPED.GerarChaveRegistro(CHV_REG, "C001")
    If Not dic0100.Exists(ARQUIVO) Then
        Campos = Array("C001", ARQUIVO, CHV_C001, CHV_REG, "", ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_MOV(0))
        dicC001(ARQUIVO) = Campos
    End If
    
End Function

Public Function CriarRegistro0000_Contr(ByRef Doce As IXMLDOMNode, ByRef dic0000 As Dictionary, _
    ByRef dic0001 As Dictionary, ByRef dic0110 As Dictionary, ByVal Periodo As String, ByVal tpCont As String)

Dim Campos As Variant
Dim CHV_REG As String, DT_INI$, DT_FIN$, Ano$, Mes$, CHV_0001$, CHV_0110$, COD_VER$, TIPO_ESCRIT$, NOME$, CNPJ$, UF$, COD_MUN$, SUFRAMA$
    
    Ano = VBA.Right(Periodo, 4)
    Mes = VBA.Left(Periodo, 2)
    COD_VER = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_COD_VER(Periodo)
    TIPO_ESCRIT = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_TIPO_ESCRIT(0)
    DT_INI = Ano & "-" & Mes & "-01"
    DT_FIN = VBA.Format(Application.WorksheetFunction.EoMonth(DT_INI, 0), "yyyy-mm-dd")
    NOME = ValidarTag(Doce, "//" & tpCont & "/xNome")
    CNPJ = CNPJContribuinte
    UF = ValidarTag(Doce, "//" & tpCont & "//UF")
    COD_MUN = ValidarTag(Doce, "//" & tpCont & "//cMun")
    SUFRAMA = ValidarTag(Doce, "//" & tpCont & "/SUFRAMA")
    ARQUIVO = Periodo & "-" & CNPJ
    CHV_REG = fnSPED.GerarChaveRegistro(DT_INI, DT_FIN, CNPJ)
    
    Campos = Array("'0000", ARQUIVO, CHV_REG, "", "", "'" & COD_VER, "'" & TIPO_ESCRIT, "", "", DT_INI, DT_FIN, NOME, "'" & CNPJ, UF, COD_MUN, "'" & SUFRAMA, "", "")
    dic0000(ARQUIVO) = Campos
    
    CHV_0001 = fnSPED.GerarChaveRegistro(CHV_REG, "0001")
    Campos = Array("'0001", ARQUIVO, CHV_0001, "", CHV_REG, "0001")
    If Not dic0001.Exists(ARQUIVO) Then dic0001(ARQUIVO) = Campos
    
    CHV_0110 = fnSPED.GerarChaveRegistro(CHV_0001, "0110")
    If Not dic0110.Exists(ARQUIVO) Then
        Campos = Array("'0110", ARQUIVO, CHV_0110, "", CHV_0001, "", "", "", "")
        dic0110(ARQUIVO) = Campos
    End If
    
End Function

Public Function CriarRegistro0005(ByRef Doce As IXMLDOMNode, ByRef dic0005 As Dictionary, ByVal CHV_PAI As String, ByVal Periodo As String, ByVal tpCont As String)

Dim Campos As Variant
Dim CHV_REG As String, FANTASIA$, CEP$, ENDER$, NUM$, COMPL$, BAIRRO$, FONE$, FAX$, EMAIL$
    
    FANTASIA = ValidarTag(Doce, "//" & tpCont & "/xFant")
    CEP = fnExcel.FormatarTexto(ValidarTag(Doce, "//" & tpCont & "//CEP"))
    ENDER = ValidarTag(Doce, "//" & tpCont & "//xLgr")
    NUM = fnExcel.FormatarTexto(ValidarTag(Doce, "//" & tpCont & "//nro"))
    COMPL = ValidarTag(Doce, "//" & tpCont & "//xCpl")
    BAIRRO = ValidarTag(Doce, "//" & tpCont & "//xBairro")
    FONE = fnExcel.FormatarTexto(ValidarTag(Doce, "//" & tpCont & "//fone"))
    FAX = fnExcel.FormatarTexto(ValidarTag(Doce, "//" & tpCont & "//fone"))
    EMAIL = ValidarTag(Doce, "//" & tpCont & "/email")
    CHV_REG = fnSPED.GerarChaveRegistro(CHV_PAI, "0005")
    
    If Not dic0005.Exists(ARQUIVO) Then
        Campos = Array("'0005", ARQUIVO, CHV_REG, CHV_PAI, "", FANTASIA, CEP, ENDER, NUM, COMPL, BAIRRO, FONE, FAX, EMAIL)
        dic0005(ARQUIVO) = Campos
    End If
    
End Function

Function ValidarNotaDevolucao(ByRef Doce As IXMLDOMNode) As Boolean

Dim finNFeNode As Object
    
    ' Verifica se o XML contém a tag finNFe
    Set finNFeNode = Doce.SelectSingleNode("//finNFe")
    
    ' Se finNFe for 4, é uma nota de devolução
    If Not finNFeNode Is Nothing Then If finNFeNode.text = "4" Then ValidarNotaDevolucao = True
    
End Function

Public Function ValidarXML(ByRef NFe As DOMDocument60) As Boolean
    
    If NFe.parseError.ErrorCode = 0 Then ValidarXML = True
    
End Function

Public Function ExtrairSituacaoDocumento(ByRef NFe As IXMLDOMNode, Optional ByVal SPEDContrib As Boolean) As String

Dim COD_SIT As String
    
    COD_SIT = ValidarTag(NFe, "//cStat")
    COD_SIT = fnXML.ValidarSituacao(COD_SIT)
    COD_SIT = fnSPED.GerarCodigoSituacao(COD_SIT)
    
    ExtrairSituacaoDocumento = EnumFiscal.ValidarEnumeracao_COD_SIT(COD_SIT)
    
End Function

Public Function ExtrairTipoFrete(ByRef NFe As IXMLDOMNode, ByVal DT_DOC As String, Optional ByVal SPEDContrib As Boolean) As String

Dim IND_FRT As String
    
    IND_FRT = ValidarTag(NFe, "//modFrete")
    
    Select Case True
        
        Case CDate(DT_DOC) > CDate("2018-01-01")
            ExtrairTipoFrete = EnumFiscal.ValidarEnumeracao_IND_FRT(IND_FRT)
            
        Case CDate(DT_DOC) > CDate("2012-01-01")
            ExtrairTipoFrete = EnumFiscal.ValidarEnumeracao_IND_FRT_2012(IND_FRT)
            
        Case Else
            ExtrairTipoFrete = EnumFiscal.ValidarEnumeracao_IND_FRT_INICIAL(IND_FRT)
            
    End Select
    
End Function

Public Sub IncluirRegistro0150(ByVal NFe As IXMLDOMNode, ByVal COD_PART As String, ByRef tpPart As String)

Dim ARQUIVO As Variant
    
    If SPEDFiscal.dicDados0150.Count = 0 Then Call DadosSPEDFiscal.CarregarDadosRegistro0150
    If SPEDFiscal.dicDados0001.Count = 0 Then Call DadosSPEDFiscal.CarregarDadosRegistro0001
    
    For Each ARQUIVO In SPEDFiscal.dicDados0001.Keys()
        
        If Not SPEDFiscal.dicDados0150.Exists(Util.UnirCampos(ARQUIVO, COD_PART)) Then
            
            With Campos0150
                
                .REG = "'0150"
                .COD_PART = COD_PART
                .NOME = VBA.Trim(VBA.Left(ValidarTag(NFe, "//" & tpPart & "/xNome"), 100))
                .COD_PAIS = ValidarTag(NFe, "//" & tpPart & "//cPais")
                .CNPJ = ValidarTag(NFe, "//" & tpPart & "/CNPJ")
                .CPF = ValidarTag(NFe, "//" & tpPart & "/CPF")
                .IE = fnExcel.FormatarTexto(ValidarTag(NFe, "//" & tpPart & "/IE"))
                .COD_MUN = fnExcel.FormatarTexto(ValidarTag(NFe, "//" & tpPart & "//cMun"))
                .SUFRAMA = fnExcel.FormatarTexto(ValidarTag(NFe, "//" & tpPart & "/SUFRAMA"))
                .END = VBA.Trim(VBA.Left(ValidarTag(NFe, "//" & tpPart & "//xLgr"), 60))
                .NUM = VBA.Trim(VBA.Left(fnExcel.FormatarTexto(ValidarTag(NFe, "//" & tpPart & "//nro")), 10))
                .COMPL = VBA.Trim(VBA.Left(ValidarTag(NFe, "//" & tpPart & "//xCpl"), 60))
                .BAIRRO = VBA.Trim(VBA.Left(ValidarTag(NFe, "//" & tpPart & "//xBairro"), 60))
                If .COD_PAIS = "" Then .COD_PAIS = "1058"
                
                .CHV_PAI = SPEDFiscal.dicDados0001(ARQUIVO)(SPEDFiscal.dicTitulos0001("CHV_REG"))
                .CHV_REG = fnSPED.GerarChaveRegistro(ARQUIVO, .CNPJ, .CPF)
                
                SPEDFiscal.dicDados0150.Add Util.UnirCampos(ARQUIVO, COD_PART), Array(.REG, ARQUIVO, .CHV_REG, .CHV_PAI, "", _
                    Util.FormatarTexto(.COD_PART), .NOME, Util.FormatarTexto(.COD_PAIS), fnExcel.FormatarTexto(.CNPJ), _
                    fnExcel.FormatarTexto(.CPF), .IE, .COD_MUN, .SUFRAMA, .END, .NUM, .COMPL, .BAIRRO)
                
            End With
            
        End If
        
    Next ARQUIVO
    
End Sub

Public Function ExtrairCEST(ByRef Produto As IXMLDOMNode) As String

Dim CEST As String
    
    CEST = fnXML.ValidarTag(Produto, "prod/CEST")
    If CEST <> "" Then
        
        If Not ValidacoesGerais.ValidarCEST(CEST) Then
                
            ExtrairCEST = 0
            Exit Function
        
        End If
        
    Else
    
        CEST = 0
    
    End If
    
    ExtrairCEST = CEST
    
End Function

Public Function RemoverNamespaces(ByVal XML As String, Optional Texto As Boolean = False) As DOMDocument60

Dim xDoc As New DOMDocument60
Dim xmlText As String
Dim regex As New RegExp

    If Texto Then Call xDoc.LoadXML(XML) Else Call xDoc.Load(XML)

    With regex

        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = "\s+xmlns(:\w+)?\s*=\s*""[^""]*"""

    End With

    xmlText = xDoc.XML
    xmlText = regex.Replace(xmlText, "")

    regex.Pattern = "(</?|\s+)\w+:"
    xmlText = regex.Replace(xmlText, "$1")

    xDoc.async = False
    xDoc.validateOnParse = False
    xDoc.LoadXML xmlText

    Set RemoverNamespaces = xDoc

End Function

