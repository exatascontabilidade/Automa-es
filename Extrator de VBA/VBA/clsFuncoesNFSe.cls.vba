Attribute VB_Name = "clsFuncoesNFSe"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function CriarRegistro0000(ByRef NFSe As IXMLDOMNode, ByRef dic0000 As Dictionary, ByVal Periodo As String)

Dim Campos As Variant
Dim CHV_REG As String, DT_INI$, DT_FIN$, Ano$, Mes$, CHV_0001$, CHV_0110$, COD_VER$, TIPO_ESCRIT$, NOME$, CNPJ$, UF$, COD_MUN$, SUFRAMA$, tpCont$
    
    Call fnNFSe.ExtrairCNPJContribunte(NFSe, tpCont)
    
    Ano = VBA.Right(Periodo, 4)
    Mes = VBA.Left(Periodo, 2)
    COD_VER = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_COD_VER(Periodo)
    TIPO_ESCRIT = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_TIPO_ESCRIT(0)
    DT_INI = Ano & "-" & Mes & "-01"
    DT_FIN = Application.WorksheetFunction.EoMonth(Ano & "-" & Mes & "-" & "01", 0)
    NOME = ValidarTag(NFSe, "Nfse//" & tpCont & "//RazaoSocial")
    CNPJ = CNPJContribuinte
    UF = ValidarTag(NFSe, "Nfse//" & tpCont & "//Uf")
    COD_MUN = ValidarTag(NFSe, "Nfse//" & tpCont & "//CodigoMunicipio")
    SUFRAMA = ""
    ARQUIVO = Periodo & "-" & CNPJ
    CHV_REG = fnSPED.GerarChaveRegistro(DT_INI, DT_FIN, CNPJ)
    
    Campos = Array("'0000", ARQUIVO, CHV_REG, "", "", "'" & COD_VER, "'" & TIPO_ESCRIT, "", "", DT_INI, DT_FIN, NOME, "'" & CNPJ, UF, COD_MUN, "'" & SUFRAMA, "", "")
    dic0000(ARQUIVO) = Campos
    
End Function

Public Sub CriarRegistro0140(ByVal NFSe As IXMLDOMNode, ByRef dicDados As Dictionary, ByVal ARQUIVO As String, ByVal CHV_PAI As String)

Dim tpCont As String, Chave$
    
    With Campos0140
        
        .CNPJ = ExtrairCNPJContribunte(NFSe, tpCont)
        .CHV_REG = fnSPED.GerarChaveRegistro(ARQUIVO, .CNPJ)
        If Not dicDados.Exists(.CHV_REG) And .CHV_REG <> "" Then
            
            .REG = "'0140"
            .COD_EST = ""
            .NOME = ValidarTag(NFSe, "Nfse//" & tpCont & "//RazaoSocial")
            .UF = ValidarTag(NFSe, "Nfse//" & tpCont & "//Uf")
            .IE = ""
            .COD_MUN = fnExcel.FormatarTexto(ValidarTag(NFSe, "Nfse//" & tpCont & "//CodigoMunicipio"))
            .IM = fnExcel.FormatarTexto(ValidarTag(NFSe, "Nfse//" & tpCont & "//InscricaoMunicipal"))
            .SUFRAMA = ""
            
            Chave = VBA.Join(Array(ARQUIVO, .CNPJ))
            dicDados(Chave) = Array(.REG, ARQUIVO, .CHV_REG, "", CHV_PAI, .COD_EST, .NOME, "'" & .CNPJ, .UF, .IE, .COD_MUN, .IM, .SUFRAMA)
            
        End If
        
    End With
    
End Sub

Public Sub CriarRegistro0150(ByVal Nota As IXMLDOMNode, ByRef dicDados As Dictionary, ByVal ARQUIVO As String, ByVal CHV_PAI As String)
            
Dim tpPart As String
    
    With Campos0150
        
        tpPart = ExtrairTagParticipante(Nota)
        .CNPJ = fnExcel.FormatarTexto(Util.ApenasNumeros(ValidarTag(Nota, "Nfse//" & tpPart & "//Cnpj")))
        .CPF = fnExcel.FormatarTexto(Util.ApenasNumeros(ValidarTag(Nota, "Nfse//" & tpPart & "//Cpf")))
        
        .CHV_REG = VBA.Join(Array(.CNPJ, .CPF))
        If Not dicDados.Exists(.CHV_REG) And .CHV_REG <> "" Then
            
            .REG = "'0150"
            .COD_PART = ExtrairCNPJParticipante(Nota)
            .NOME = ValidarTag(Nota, "Nfse//" & tpPart & "/RazaoSocial")
            .COD_PAIS = ""
            .CNPJ = fnExcel.FormatarTexto(Util.ApenasNumeros(ValidarTag(Nota, "Nfse//" & tpPart & "//Cnpj")))
            .CPF = fnExcel.FormatarTexto(Util.ApenasNumeros(ValidarTag(Nota, "Nfse//" & tpPart & "//Cpf")))
            .IE = ""
            .COD_MUN = fnExcel.FormatarTexto(ValidarTag(Nota, "Nfse//" & tpPart & "//CodigoMunicipio"))
            .SUFRAMA = ""
            .END = VBA.Left(ValidarTag(Nota, "Nfse//" & tpPart & "//Endereco"), 60)
            .NUM = fnExcel.FormatarTexto(ValidarTag(Nota, "Nfse//" & tpPart & "//Numero"))
            .COMPL = ValidarTag(Nota, "Nfse//" & tpPart & "//Complemento")
            .BAIRRO = ValidarTag(Nota, "Nfse//" & tpPart & "//Bairro")
            If .COD_PAIS = "" Then .COD_PAIS = "1058"
            
            .CHV_REG = fnSPED.GerarChaveRegistro(CHV_PAI, .COD_PART)
            dicDados(.CHV_REG) = Array(.REG, ARQUIVO, .CHV_REG, "", CHV_PAI, Util.FormatarTexto(.COD_PART), _
                .NOME, .COD_PAIS, .CNPJ, .CPF, .IE, .COD_MUN, .SUFRAMA, .END, .NUM, .COMPL, .BAIRRO)
                    
        End If

    End With

End Sub

Public Sub CriarRegistro0200(ByVal item As IXMLDOMNode, ByRef dicDados As Dictionary, ByVal ARQUIVO As String, ByVal CHV_PAI As String)

Dim Campos As Variant
    
    With Campos0200
        
        .COD_ITEM = ValidarTag(item, "Nfse//ItemListaServico")
        .CHV_REG = fnSPED.GerarChaveRegistro(CHV_PAI, .COD_ITEM)
        If Not dicDados.Exists(.CHV_REG) Then
            
            .REG = "'0200"
            .DESCR_ITEM = Util.LimparTexto(ValidarTag(item, "Nfse//Discriminacao"))
            .COD_BARRA = ""
            .COD_ANT_ITEM = ""
            .UNID_INV = ""
            .TIPO_ITEM = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_TIPO_ITEM("09")
            .COD_NCM = ""
            .EX_IPI = ""
            .COD_GEN = ""
            .COD_LST = "'" & .COD_ITEM
            .ALIQ_ICMS = ""
            .CEST = ""
            
            dicDados(.CHV_REG) = Array(.REG, ARQUIVO, .CHV_REG, "", CHV_PAI, "'" & .COD_ITEM, .DESCR_ITEM, .COD_BARRA, _
                .COD_ANT_ITEM, .UNID_INV, .TIPO_ITEM, .COD_NCM, .EX_IPI, .COD_GEN, .COD_LST, .ALIQ_ICMS, .CEST)
            
        End If
        
    End With
    
End Sub

Public Sub CriarRegistroA100(ByRef Nota As IXMLDOMNode, ByRef dicDadosA100 As Dictionary, _
    ByRef dicDadosA170 As Dictionary, ByRef dicTitulosA170 As Dictionary, ByVal CHV_PAI As String)
'ByRef dicDados0150 As Dictionary, ByRef dicTitulos0150 As Dictionary, ByRef dicDados0200 As Dictionary, ByRef dicTitulos0200 As Dictionary

Dim Campos As Variant
Dim VL_DEDUCOES As Double, VL_PRESTACAO#
Dim Emissao As String, Competencia$, CNPJPrest$, CNPJToma$, Chave$
    
    With CamposA100
        
        Call DefinirDadosA100(Nota)
        .CHV_NFSE = VBA.UCase(ValidarTag(Nota, "Nfse//CodigoVerificacao"))
        
        .CHV_REG = VBA.Join(Array(.IND_OPER, .IND_EMIT, .CHV_NFSE))
        If Not dicDadosA100.Exists(.CHV_REG) Then
            
            VL_PRESTACAO = fnXML.ValidarValores(Nota, "Nfse//ValorServicos")
            VL_DEDUCOES = fnXML.ValidarValores(Nota, "Nfse//ValorDeducoes")
            Competencia = fnExcel.FormatarData(VBA.Left(ValidarTag(Nota, "Nfse//Competencia"), 10))
            Emissao = fnExcel.FormatarData(VBA.Left(ValidarTag(Nota, "Nfse//DataEmissao"), 10))
            
            'Verifica se a nota foi emitida em outra competência
            If VBA.Month(Competencia) < VBA.Month(Emissao) Then Emissao = WorksheetFunction.EoMonth(Competencia, 0)
            If VBA.Month(Competencia) > VBA.Month(Emissao) Then Emissao = Competencia
                        
            .REG = "A100"
            .COD_SIT = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_COD_SIT("00")
            .SER = ""
            .SUB = ""
            .NUM_DOC = ValidarTag(Nota, "Nfse//Numero")
            .DT_DOC = Emissao
            .DT_EXE_SERV = Emissao
            .VL_DOC = fnXML.ValidarValores(Nota, "Nfse//ValorServicos") 'Valor do Serviço
            .IND_PGTO = ""
            .VL_DESC = fnXML.ValidarValores(Nota, "Nfse//DescontoIncondicionado") 'Valor do Desconto
            .VL_BC_PIS = .VL_DOC - VL_DEDUCOES
            .VL_PIS = 0
            .VL_BC_COFINS = .VL_BC_PIS
            .VL_COFINS = 0
            .VL_PIS_RET = fnXML.ValidarValores(Nota, "Nfse//ValorPis") 'Valor do PIS Retido
            .VL_COFINS_RET = fnXML.ValidarValores(Nota, "Nfse//ValorCofins") 'Valor da COFINS Retida
            .VL_ISS = fnXML.ValidarValores(Nota, "Nfse//ValorIss") 'Valor do ISS
            
            .CHV_REG = fnSPED.GerarChaveRegistro(CHV_PAI, .IND_OPER, .IND_EMIT, .COD_PART, .COD_SIT, .SER, .NUM_DOC, .CHV_NFSE)
            Campos = Array(.REG, .ARQUIVO, .CHV_REG, "", CHV_PAI, .IND_OPER, .IND_EMIT, "'" & .COD_PART, .COD_SIT, .SER, _
                .SUB, "'" & .NUM_DOC, .CHV_NFSE, .DT_DOC, .DT_EXE_SERV, CDbl(.VL_DOC), .IND_PGTO, CDbl(.VL_DESC), CDbl(.VL_BC_PIS), _
                CDbl(.VL_PIS), CDbl(.VL_BC_COFINS), CDbl(.VL_COFINS), CDbl(.VL_PIS_RET), CDbl(.VL_COFINS_RET), CDbl(.VL_ISS))
                
            Chave = VBA.Join(Array(.IND_OPER, .IND_EMIT, .CHV_NFSE))
            dicDadosA100(Chave) = Campos
            
            Call CriarRegistroA170(Nota, dicTitulosA170, dicDadosA170)
        
        End If
        
    End With
    
End Sub

Public Sub CriarRegistroA170(ByVal item As IXMLDOMNode, ByRef dicTitulos As Dictionary, ByRef dicDados As Dictionary)
    
    With CamposA170
       
       .REG = "A170"
       .ARQUIVO = CamposA100.ARQUIVO
       .NUM_ITEM = 1
       .COD_ITEM = fnExcel.FormatarTexto(fnXML.ValidarTag(item, "Nfse//ItemListaServico"))
       .DESCR_COMPL = Util.LimparTexto(fnXML.ValidarTag(item, "Nfse//Discriminacao"), 60)
       .VL_ITEM = fnXML.ValidarTag(item, "Nfse//ValorServicos")
       .VL_DESC = fnXML.ValidarTag(item, "Nfse//DescontoIncondicionado")
       .NAT_BC_CRED = ""
       .IND_ORIG_CRED = ""
       .CST_PIS = ""
       .VL_BC_PIS = .VL_ITEM - .VL_DESC
       .ALIQ_PIS = 0
       .VL_PIS = 0
       .CST_COFINS = ""
       .VL_BC_COFINS = .VL_BC_PIS
       .ALIQ_COFINS = 0
       .VL_COFINS = 0
       .COD_CTA = ""
       .COD_CCUS = ""
       .CHV_PAI = CamposA100.CHV_REG
       .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .NUM_ITEM)
       
       dicDados(.CHV_REG) = Array(.REG, .ARQUIVO, .CHV_REG, "", .CHV_PAI, "", .NUM_ITEM, .COD_ITEM, .DESCR_COMPL, _
        CDbl(.VL_ITEM), CDbl(.VL_DESC), .NAT_BC_CRED, .IND_ORIG_CRED, .CST_PIS, CDbl(.VL_BC_PIS), CDbl(.ALIQ_PIS), _
        CDbl(.VL_PIS), .CST_COFINS, CDbl(.VL_BC_COFINS), CDbl(.ALIQ_COFINS), CDbl(.VL_COFINS), .COD_CTA, .COD_CCUS)
        
    End With
    
End Sub

Public Function ExtrairCNPJContribunte(ByRef NFSe As IXMLDOMNode, Optional ByRef tpCont As String)

Dim CNPJPrest As String, CNPJToma$
    
    CNPJPrest = Util.ApenasNumeros(ExtrairPrestador(NFSe))
    CNPJToma = Util.ApenasNumeros(ExtrairTomador(NFSe))
    
    If CNPJPrest Like CNPJBase & "*" Then
        
        ExtrairCNPJContribunte = CNPJPrest
        tpCont = "PrestadorServico"
        
    End If
    
    If CNPJToma Like CNPJBase & "*" Then
        
        ExtrairCNPJContribunte = CNPJToma
        tpCont = "TomadorServico"
        
    End If
    
End Function

Public Function ExtrairCNPJParticipante(ByRef NFSe As IXMLDOMNode, Optional ByRef tpCont As String)

Dim CNPJPrest As String, CNPJToma$
    
    CNPJPrest = Util.ApenasNumeros(ExtrairPrestador(NFSe))
    CNPJToma = Util.ApenasNumeros(ExtrairTomador(NFSe))
    
    If CNPJPrest Like CNPJBase & "*" Then
        
        ExtrairCNPJParticipante = CNPJToma
        tpCont = "PrestadorServico"
        
    End If
    
    If CNPJToma Like CNPJBase & "*" Then
        
        ExtrairCNPJParticipante = CNPJPrest
        tpCont = "TomadorServico"
        
    End If
    
End Function

Public Function ExtrairTagParticipante(ByRef NFSe As IXMLDOMNode)

Dim CNPJPrest As String, CNPJToma$
    
    CNPJPrest = Util.ApenasNumeros(ExtrairPrestador(NFSe))
    CNPJToma = Util.ApenasNumeros(ExtrairTomador(NFSe))
    
    If CNPJPrest Like CNPJBase & "*" Then ExtrairTagParticipante = "TomadorServico"
    If CNPJToma Like CNPJBase & "*" Then ExtrairTagParticipante = "PrestadorServico"
        
End Function

Public Function DefinirDadosA100(ByRef NFSe As IXMLDOMNode)

Dim CNPJPrest As String, CNPJToma$

    CNPJPrest = Util.ApenasNumeros(ExtrairPrestador(NFSe))
    CNPJToma = Util.ApenasNumeros(ExtrairTomador(NFSe))
    
    With CamposA100
        
        If CNPJPrest Like CNPJBase & "*" Then
        
            .IND_OPER = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_A100_IND_OPER(1)
            .IND_EMIT = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_IND_EMIT(0)
            .COD_PART = CNPJToma
             
        End If
        
        If CNPJToma Like CNPJBase & "*" Then
        
            .IND_OPER = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_A100_IND_OPER(0)
            .IND_EMIT = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_IND_EMIT(1)
            .COD_PART = CNPJPrest
            
        End If
        
    End With
    
End Function

Public Function ExtrairPrestador(ByVal NFSe As IXMLDOMNode)
    
    If Not NFSe.SelectSingleNode("Nfse//PrestadorServico//Cnpj") Is Nothing Then
        ExtrairPrestador = NFSe.SelectSingleNode("Nfse//PrestadorServico//Cnpj").text
    ElseIf Not NFSe.SelectSingleNode("Nfse//PrestadorServico//Cpf") Is Nothing Then
        ExtrairPrestador = NFSe.SelectSingleNode("Nfse//PrestadorServico//Cpf").text
    End If
    
End Function

Public Function ExtrairTomador(ByVal NFSe As IXMLDOMNode)
    
    If Not NFSe.SelectSingleNode("Nfse//TomadorServico//Cnpj") Is Nothing Then
        ExtrairTomador = NFSe.SelectSingleNode("Nfse//TomadorServico//Cnpj").text
    ElseIf Not NFSe.SelectSingleNode("Nfse//TomadorServico//Cpf") Is Nothing Then
        ExtrairTomador = NFSe.SelectSingleNode("Nfse//TomadorServico//Cpf").text
    End If
    
End Function
