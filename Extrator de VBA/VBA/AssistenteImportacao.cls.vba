Attribute VB_Name = "AssistenteImportacao"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private EnumContrib As New clsEnumeracoesSPEDContribuicoes
Private ImportNFeNFCe As New AssistenteImportacaoNFeNFCe
Private EnumFiscal As New clsEnumeracoesSPEDFiscal
Private GerenciadorSPED As New clsRegistrosSPED
Private Doce As New DOMDocument60

Public Function CriarRegistro0000(ByRef NFe As DOMDocument60)

Dim tpCont As String, Ano$, Mes$
Dim Campos As Variant
    
    If dtoRegSPED.r0000 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0000("ARQUIVO")
    
    tpCont = DefinirTipoContribuinteNFe()
    Ano = VBA.Right(DadosXML.Periodo, 4)
    Mes = VBA.Left(DadosXML.Periodo, 2)
    
    With Campos0000
        
        .REG = "'0000"
        .ARQUIVO = DadosXML.ARQUIVO
        .COD_VER = fnExcel.FormatarTexto(EnumFiscal.ValidarEnumeracao_COD_VER(DadosXML.Periodo))
        .COD_FIN = EnumFiscal.ValidarEnumeracao_COD_FIN(0)
        .DT_INI = Ano & "-" & Mes & "-01"
        .DT_FIN = VBA.Format(Application.WorksheetFunction.EoMonth(Ano & "-" & Mes & "-" & "01", 0), "yyyy-mm-dd")
        .NOME = ValidarTag(NFe, "//" & tpCont & "/xNome")
        .CNPJ = Util.FormatarCNPJ(CNPJContribuinte)
        .CPF = ""
        .UF = ValidarTag(NFe, "//" & tpCont & "//UF")
        .IE = Util.FormatarValores(Util.ApenasNumeros(ValidarTag(NFe, "//" & tpCont & "/IE"))) * 1
        .COD_MUN = ValidarTag(NFe, "//" & tpCont & "//cMun")
        .SUFRAMA = ValidarTag(NFe, "//" & tpCont & "/SUFRAMA")
        .CHV_REG = fnSPED.GerarChaveRegistro(.DT_INI, .DT_FIN, .CNPJ, .CPF, .IE)
        
        Campos = Array(.REG, .ARQUIVO, .CHV_REG, "", "", .COD_VER, .COD_FIN, .DT_INI, .DT_FIN, .NOME, _
            fnExcel.FormatarTexto(.CNPJ), .CPF, .UF, fnExcel.FormatarTexto(.IE), .COD_MUN, "", .SUFRAMA, "", "")
            
        dtoRegSPED.r0000(.ARQUIVO) = Campos
        
    End With
    
End Function

Public Function CriarRegistro0000_Contr(ByRef NFe As DOMDocument60)

Dim tpCont As String, Ano$, Mes$
Dim Campos As Variant
    
    If dtoRegSPED.r0000_Contr Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0000_Contr("ARQUIVO")
    
    tpCont = DefinirTipoContribuinteNFe()
    Ano = VBA.Right(DadosXML.Periodo, 4)
    Mes = VBA.Left(DadosXML.Periodo, 2)
    
    With Campos0000_Contr
        
        .REG = "'0000"
        .ARQUIVO = DadosXML.ARQUIVO
        .COD_VER = fnExcel.FormatarTexto(EnumContrib.ValidarEnumeracao_COD_VER(DadosXML.Periodo))
        .TIPO_ESCRIT = fnExcel.FormatarTexto(EnumContrib.ValidarEnumeracao_TIPO_ESCRIT(0))
        .DT_INI = Ano & "-" & Mes & "-01"
        .DT_FIN = VBA.Format(Application.WorksheetFunction.EoMonth(.DT_INI, 0), "yyyy-mm-dd")
        .NOME = ValidarTag(NFe, "//" & tpCont & "/xNome")
        .CNPJ = fnExcel.FormatarTexto(Util.FormatarCNPJ(CNPJContribuinte))
        .UF = ValidarTag(NFe, "//" & tpCont & "//UF")
        .COD_MUN = ValidarTag(NFe, "//" & tpCont & "//cMun")
        .SUFRAMA = fnExcel.FormatarTexto(ValidarTag(NFe, "//" & tpCont & "/SUFRAMA"))
        .CHV_REG = fnSPED.GerarChaveRegistro(.DT_INI, .DT_FIN, .CNPJ)
        
        Campos = Array(.REG, .ARQUIVO, .CHV_REG, "", "", .COD_VER, .TIPO_ESCRIT, "", "", _
            .DT_INI, .DT_FIN, .NOME, fnExcel.FormatarTexto(.CNPJ), .UF, .COD_MUN, .SUFRAMA, "", "")
            
        dtoRegSPED.r0000_Contr(.ARQUIVO) = Campos
        
    End With
    
End Function

Public Sub CriarRegistro0001()

Dim tpCont As String, Chave$
Dim Campos As Variant
    
    If dtoRegSPED.r0001 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0001("ARQUIVO")
    
    With Campos0001
        
        If Not dtoRegSPED.r0001.Exists(DadosXML.ARQUIVO) Then
            
            .REG = "'0001"
            .ARQUIVO = DadosXML.ARQUIVO
            .IND_MOV = EnumFiscal.ValidarEnumeracao_IND_MOV("0")
            .CHV_PAI_FISCAL = dtoRegSPED.r0000(.ARQUIVO)(dtoTitSPED.t0000("CHV_REG") - 1)
            .CHV_PAI_CONTRIBUICOES = dtoRegSPED.r0000_Contr(.ARQUIVO)(dtoTitSPED.t0000_Contr("CHV_REG") - 1)
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI_FISCAL, "0001")
            
            Campos = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, .IND_MOV)
            
            dtoRegSPED.r0001(.ARQUIVO) = Campos
            
        End If
        
    End With
    
End Sub

Public Sub CriarRegistro0100()

Dim tpCont As String, Chave$
Dim Campos As Variant
    
    If dtoRegSPED.r0100 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0100("ARQUIVO")
    
    With Campos0100
        
        If Not dtoRegSPED.r0100.Exists(DadosXML.ARQUIVO) Then
            
            .REG = "'0100"
            .ARQUIVO = DadosXML.ARQUIVO
            .NOME = ""
            .CPF = ""
            .CRC = ""
            .CNPJ = ""
            .CEP = ""
            .END = ""
            .NUM = ""
            .COMPL = ""
            .BAIRRO = ""
            .FONE = ""
            .FAX = ""
            .EMAIL = ""
            .COD_MUN = ""
            .CHV_PAI_FISCAL = ExtrairChaveReg0001()
            .CHV_PAI_CONTRIBUICOES = ""
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, "0100")
            
            Campos = Array(.REG, .ARQUIVO, .CHV_REG, "", .CHV_PAI, .NOME, .CPF, .CRC, _
                .CNPJ, .CEP, .END, .NUM, .COMPL, .BAIRRO, .FONE, .FAX, .EMAIL, .COD_MUN)
                
            dtoRegSPED.r0100(.ARQUIVO) = Campos
            
        End If
        
    End With
    
End Sub

Public Sub CriarRegistro0005(ByRef NFe As DOMDocument60)

Dim tpCont As String
Dim Campos As Variant
    
    If dtoRegSPED.r0005 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0005("ARQUIVO")
    
    tpCont = DefinirTipoContribuinteNFe()
    
    With Campos0005
        
        If Not dtoRegSPED.r0005.Exists(DadosXML.ARQUIVO) Then
            
            .REG = "'0005"
            .ARQUIVO = DadosXML.ARQUIVO
            .FANTASIA = ValidarTag(NFe, "//" & tpCont & "/xFant")
            .CEP = fnExcel.FormatarTexto(ValidarTag(NFe, "//" & tpCont & "//CEP"))
            .END = ValidarTag(NFe, "//" & tpCont & "//xLgr")
            .NUM = fnExcel.FormatarTexto(ValidarTag(NFe, "//" & tpCont & "//nro"))
            .COMPL = ValidarTag(NFe, "//" & tpCont & "//xCpl")
            .BAIRRO = ValidarTag(NFe, "//" & tpCont & "//xBairro")
            .FONE = fnExcel.FormatarTexto(ValidarTag(NFe, "//" & tpCont & "//fone"))
            .FAX = fnExcel.FormatarTexto(ValidarTag(NFe, "//" & tpCont & "//fone"))
            .EMAIL = ValidarTag(NFe, "//" & tpCont & "/email")
            .CHV_PAI_FISCAL = ExtrairChaveReg0001()
            .CHV_PAI_CONTRIBUICOES = ""
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI_FISCAL, "0005")
            
            Campos = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, "", .FANTASIA, .CEP, .END, .NUM, .COMPL, .BAIRRO, .FONE, .FAX, .EMAIL)
                
            dtoRegSPED.r0005(.ARQUIVO) = Campos
            
        End If
        
    End With
    
End Sub

Public Sub CriarRegistro0110()

Dim tpCont As String, Chave$
Dim Campos As Variant
    
    If dtoRegSPED.r0110 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0110("ARQUIVO")
    
    With Campos0110

        If Not dtoRegSPED.r0110.Exists(DadosXML.ARQUIVO) Then

            .REG = "'0110"
            .ARQUIVO = DadosXML.ARQUIVO
            .COD_INC_TRIB = ""
            .IND_APRO_CRED = ""
            .COD_TIPO_CONT = ""
            .IND_REG_CUM = ""
            .CHV_PAI = ExtrairChaveReg0001()
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, "0110")

            Campos = Array(.REG, .ARQUIVO, .CHV_REG, "", .CHV_PAI, .COD_INC_TRIB, .IND_APRO_CRED, .COD_TIPO_CONT, .IND_REG_CUM)
            dtoRegSPED.r0110(DadosXML.ARQUIVO) = Campos

        End If

    End With

End Sub

Public Sub CriarRegistro0140(ByVal NFe As DOMDocument60)

Dim tpCont As String, Chave$
Dim Campos As Variant
    
    If dtoRegSPED.r0140 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0140("ARQUIVO", "CNPJ")
    
    tpCont = DefinirTipoContribuinteNFe()
    
    With Campos0140
        
        Chave = Util.UnirCampos(DadosXML.ARQUIVO, DadosXML.CNPJ_ESTABELECIMENTO)
        If Not dtoRegSPED.r0140.Exists(Chave) Then
            
            .REG = "'0140"
            .ARQUIVO = DadosXML.ARQUIVO
            .COD_EST = ""
            .NOME = fnXML.ValidarTag(NFe, "//" & tpCont & "/xNome")
            .CNPJ = DadosXML.CNPJ_ESTABELECIMENTO
            .UF = fnXML.ValidarTag(NFe, "//" & tpCont & "//UF")
            .IE = fnXML.ValidarTag(NFe, "//" & tpCont & "/IE")
            .COD_MUN = fnXML.ValidarTag(NFe, "//" & tpCont & "//cMun")
            .IM = ""
            .SUFRAMA = fnXML.ValidarTag(NFe, "//" & tpCont & "/SUFRAMA")
            .CHV_PAI = ExtrairChaveReg0001()
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .CNPJ)
            
            Campos = Array(.REG, .ARQUIVO, .CHV_REG, "", .CHV_PAI, .COD_EST, .NOME, "'" & .CNPJ, .UF, "'" & .IE, .COD_MUN, .IM, .SUFRAMA)
            
            dtoRegSPED.r0140(Chave) = Campos
            
        End If
        
    End With
    
End Sub

Public Sub CriarRegistro0150(ByVal NFe As DOMDocument60, ByVal COD_PART As String)

Dim Chave As String, tpPart$
    
    If dtoRegSPED.r0150 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0150("CHV_PAI_FISCAL", "COD_PART")
    
    tpPart = ExtrairTipoParticipante()
    
    With Campos0150
        
        .CHV_PAI_FISCAL = ExtrairChaveReg0001()
        Chave = Util.UnirCampos(.CHV_PAI_FISCAL, COD_PART)
        If Not dtoRegSPED.r0150.Exists(Chave) Then
            
            .REG = "'0150"
            .ARQUIVO = DadosXML.ARQUIVO
            .COD_PART = Util.FormatarTexto(COD_PART)
            .NOME = VBA.UCase(VBA.Trim(VBA.Left(ValidarTag(NFe, "//" & tpPart & "/xNome"), 100)))
            .COD_PAIS = ValidarTag(NFe, "//" & tpPart & "//cPais")
            .CNPJ = Util.FormatarTexto(ValidarTag(NFe, "//" & tpPart & "/CNPJ"))
            .CPF = Util.FormatarTexto(ValidarTag(NFe, "//" & tpPart & "/CPF"))
            .IE = fnExcel.FormatarTexto(ValidarTag(NFe, "//" & tpPart & "/IE"))
            .COD_MUN = fnExcel.FormatarTexto(ValidarTag(NFe, "//" & tpPart & "//cMun"))
            .SUFRAMA = fnExcel.FormatarTexto(ValidarTag(NFe, "//" & tpPart & "/SUFRAMA"))
            .END = VBA.Trim(VBA.Left(ValidarTag(NFe, "//" & tpPart & "//xLgr"), 60))
            .NUM = VBA.Trim(VBA.Left(fnExcel.FormatarTexto(ValidarTag(NFe, "//" & tpPart & "//nro")), 10))
            .COMPL = VBA.Trim(VBA.Left(ValidarTag(NFe, "//" & tpPart & "//xCpl"), 60))
            .BAIRRO = VBA.Trim(VBA.Left(ValidarTag(NFe, "//" & tpPart & "//xBairro"), 60))
            
            If .COD_PAIS = "" Or .COD_PAIS = "01058" Then .COD_PAIS = "'1058" Else .COD_PAIS = Util.FormatarTexto(.COD_PAIS)
                        
            .CHV_PAI_CONTRIBUICOES = ExtrairChaveReg0140()
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI_FISCAL, .COD_PART)
            
            dtoRegSPED.r0150(Chave) = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, _
                .COD_PART, .NOME, .COD_PAIS, .CNPJ, .CPF, .IE, .COD_MUN, .SUFRAMA, .END, .NUM, .COMPL, .BAIRRO)
            
        End If
        
    End With
    
End Sub

Public Sub CriarRegistro0190(ByVal UNID As String)

Dim Chave As String
    
    If dtoRegSPED.r0190 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0190("CHV_PAI_FISCAL", "UNID")
    
    With Campos0190
        
        .CHV_PAI_FISCAL = ExtrairChaveReg0001()
        Chave = Util.UnirCampos(.CHV_PAI_FISCAL, UNID)
        If Not dtoRegSPED.r0190.Exists(Chave) Then
            
            .REG = "'0190"
            .ARQUIVO = DadosXML.ARQUIVO
            .UNID = UNID
            .DESCR = VBA.UCase(.UNID)
            .CHV_PAI_CONTRIBUICOES = ExtrairChaveReg0140()
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .UNID)
            
            dtoRegSPED.r0190(Chave) = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, .UNID, .DESCR)
            
        End If
        
    End With
    
End Sub

Public Sub CriarRegistro0200(ByVal Produto As IXMLDOMNode)

Dim Chave As String
    
    If dtoRegSPED.r0200 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0200("CHV_PAI_FISCAL", "COD_ITEM")
    
    With Campos0200
        
        .CHV_PAI_FISCAL = ExtrairChaveReg0001()
        Chave = Util.UnirCampos(.CHV_PAI_FISCAL, .COD_ITEM)
        If Not dtoRegSPED.r0200.Exists(Chave) Then
            
            .REG = "'0200"
            .ARQUIVO = DadosXML.ARQUIVO
            .COD_BARRA = Util.FormatarTexto(fnXML.ExtrairCodigoBarrasProduto(Produto))
            .COD_ANT_ITEM = ""
            .TIPO_ITEM = ""
            .COD_NCM = fnExcel.FormatarTexto(VBA.Format(ValidarTag(Produto, "prod/NCM"), String(8, "0")))
            .EX_IPI = fnExcel.FormatarTexto(ValidarTag(Produto, "prod/EXTIPI"))
            .COD_GEN = fnExcel.FormatarTexto(VBA.Left(Util.ApenasNumeros(.COD_NCM), 2))
            .COD_LST = ""
            .ALIQ_ICMS = fnXML.ValidarPercentual(Produto, "imposto/ICMS//pICMS")
            .CEST = fnXML.ExtrairCEST(Produto)
            .CHV_PAI_CONTRIBUICOES = ExtrairChaveReg0140()
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI_FISCAL, .COD_ITEM)

            If CamposC100.COD_MOD = "65" Then .TIPO_ITEM = EnumContrib.ValidarEnumeracao_TIPO_ITEM("00")
            If VBA.UCase(.COD_BARRA) = "SEM GTIN" Then .COD_BARRA = ""
            If .CEST = "0" Then .CEST = ""
            
            dtoRegSPED.r0200(Chave) = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, "'" & .COD_ITEM, _
                .DESCR_ITEM, .COD_BARRA, .COD_ANT_ITEM, .UNID_INV, .TIPO_ITEM, .COD_NCM, .EX_IPI, .COD_GEN, .COD_LST, CDbl(.ALIQ_ICMS), "")
                
            Call DTO_EFDICMS.ResetarCampos0200
            
        End If
        
    End With
    
End Sub

Public Sub CriarRegistro0220(ByVal UNID_COM As String, ByVal FAT_CONV As Double)

Dim Chave As String
    
    If dtoRegSPED.r0150 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0150("CHV_PAI_FISCAL", "UNID_CONV")
    
    With Campos0220
        
        .CHV_PAI_FISCAL = ExtrairChaveReg0200(Campos0200.COD_ITEM)
        Chave = Util.UnirCampos(ARQUIVO, UNID_COM)
        If Not dtoRegSPED.r0220.Exists(Chave) Then
            
            .REG = "'0220"
            .ARQUIVO = DadosXML.ARQUIVO
            .UNID_CONV = UNID_COM
            .FAT_CONV = FAT_CONV
            .COD_BARRA = ""
            .CHV_PAI_CONTRIBUICOES = ""
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI_FISCAL, .UNID_CONV)
            
            dtoRegSPED.r0220(Chave) = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, .UNID_CONV, .FAT_CONV, .COD_BARRA)
            
        End If
        
    End With
    
End Sub
                            
Public Sub CriarRegistroC001()

Dim Campos As Variant
    
    If dtoRegSPED.rC001 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC001("ARQUIVO")
    
    With CamposC001
        
        If Not dtoRegSPED.rC001.Exists(DadosXML.ARQUIVO) Then
            
            .REG = "C001"
            .ARQUIVO = DadosXML.ARQUIVO
            .IND_MOV = EnumFiscal.ValidarEnumeracao_IND_MOV("0")
            .CHV_PAI_FISCAL = dtoRegSPED.r0000(.ARQUIVO)(dtoTitSPED.t0000("CHV_REG") - 1)
            .CHV_PAI_CONTRIBUICOES = dtoRegSPED.r0000_Contr(.ARQUIVO)(dtoTitSPED.t0000_Contr("CHV_REG") - 1)
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI_FISCAL, "C001")
            
            Campos = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, .IND_MOV)
            
            dtoRegSPED.rC001(.ARQUIVO) = Campos
            
        End If
        
    End With
    
End Sub

Public Sub CriarRegistroC010()

Dim Chave As String
Dim Campos As Variant
    
    If dtoRegSPED.rC010 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC010("ARQUIVO", "CNPJ")
    
    With CamposC010
        
        Chave = Util.UnirCampos(DadosXML.ARQUIVO, DadosXML.CNPJ_ESTABELECIMENTO)
        If Not dtoRegSPED.rC010.Exists(Chave) Then
            
            .REG = "C010"
            .ARQUIVO = DadosXML.ARQUIVO
            .CNPJ = DadosXML.CNPJ_ESTABELECIMENTO
            .IND_ESCRI = "2"
            .CHV_PAI = ExtrairChaveReg0001()
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .CNPJ)
            
            Campos = Array(.REG, .ARQUIVO, .CHV_REG, "", .CHV_PAI, fnExcel.FormatarTexto(.CNPJ), .IND_ESCRI)
            
            dtoRegSPED.rC010(Chave) = Campos
            
        End If
        
    End With
    
End Sub

Public Sub CriarRegistroD001()

Dim Campos As Variant
    
    If dtoRegSPED.rD001 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroD001("ARQUIVO")
    
    With CamposD001
        
        If Not dtoRegSPED.rD001.Exists(DadosXML.ARQUIVO) Then
            
            .REG = "D001"
            .ARQUIVO = DadosXML.ARQUIVO
            .IND_MOV = EnumFiscal.ValidarEnumeracao_IND_MOV("0")
            .CHV_PAI_FISCAL = dtoRegSPED.r0000(.ARQUIVO)(dtoTitSPED.t0000("CHV_REG") - 1)
            .CHV_PAI_CONTRIBUICOES = dtoRegSPED.r0000_Contr(.ARQUIVO)(dtoTitSPED.t0000_Contr("CHV_REG") - 1)
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI_FISCAL, "D001")
            
            Campos = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, .IND_MOV)
            
            dtoRegSPED.rD001(.ARQUIVO) = Campos
            
        End If
        
    End With
    
End Sub

Public Sub CriarRegistroD010()

Dim Chave As String
Dim Campos As Variant
    
    If dtoRegSPED.rD010 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroD010("ARQUIVO", "CNPJ")
    
    With CamposD010
        
        Chave = Util.UnirCampos(DadosXML.ARQUIVO, DadosXML.CNPJ_ESTABELECIMENTO)
        If Not dtoRegSPED.rD010.Exists(Chave) Then
            
            .REG = "D010"
            .ARQUIVO = DadosXML.ARQUIVO
            .CNPJ = DadosXML.CNPJ_ESTABELECIMENTO
            .CHV_PAI = ExtrairChaveReg0001()
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .CNPJ)
            
            Campos = Array(.REG, .ARQUIVO, .CHV_REG, "", .CHV_PAI, fnExcel.FormatarTexto(.CNPJ))
            
            dtoRegSPED.rD010(Chave) = Campos
            
        End If
        
    End With
    
End Sub

Public Function ExtrairChaveReg0000()

Dim i As Integer
Dim Campos As Variant
    
    With dtoRegSPED
        
        If .r0000 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0000("ARQUIVO")
        
        If Not .r0000.Exists(DadosXML.ARQUIVO) Then Call CriarRegistro0000(Doce)
        
        Campos = .r0000(DadosXML.ARQUIVO)
        i = Util.VerificarPosicaoInicialArray(Campos)
        
        ExtrairChaveReg0000 = Campos(dtoTitSPED.t0000("CHV_REG") - i)
        
    End With
        
End Function

Public Function ExtrairChaveReg0000_Contr()

Dim i As Integer
Dim Campos As Variant
    
    With dtoRegSPED
        
        If .r0000_Contr Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0000_Contr("ARQUIVO")
        
        If Not .r0000_Contr.Exists(DadosXML.ARQUIVO) Then Call CriarRegistro0000_Contr(Doce)
        
        Campos = .r0000_Contr(DadosXML.ARQUIVO)
        i = Util.VerificarPosicaoInicialArray(Campos)
        
        ExtrairChaveReg0000_Contr = Campos(dtoTitSPED.t0000_Contr("CHV_REG") - i)
        
    End With
        
End Function

Public Function ExtrairChaveReg0001()

Dim i As Integer
Dim Campos As Variant
    
    With dtoRegSPED
        
        If .r0001 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0001("ARQUIVO")
        
        If Not .r0001.Exists(DadosXML.ARQUIVO) Then Call CriarRegistro0001
        
        Campos = .r0001(DadosXML.ARQUIVO)
        i = Util.VerificarPosicaoInicialArray(Campos)
        
        ExtrairChaveReg0001 = Campos(dtoTitSPED.t0001("CHV_REG") - i)
        
    End With
        
End Function

Public Function ExtrairChaveReg0140() As String

Dim i As Byte
Dim Campos As Variant
Dim Chave As String, CHV_PAI$, CHV_REG$
    
    With dtoRegSPED
        
        If .r0140 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0140("ARQUIVO", "CNPJ")
        
        Chave = Util.UnirCampos(DadosXML.ARQUIVO, DadosXML.CNPJ_ESTABELECIMENTO)
        If Not .r0140.Exists(Chave) Then Call CriarRegistro0140(Doce)
        
        Campos = .r0140(Chave)
        i = Util.VerificarPosicaoInicialArray(Campos)
        
        ExtrairChaveReg0140 = Campos(dtoTitSPED.t0140("CHV_REG") - i)
        
    End With
    
End Function

Public Function ExtrairChaveReg0200(ByVal COD_ITEM As String) As String

Dim i As Byte
Dim Campos As Variant
Dim Chave As String
            
    With dtoRegSPED
        
        If .r0200 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0200("CHV_PAI_FISCAL", "COD_ITEM")
        
        Chave = Util.UnirCampos(DadosXML.ARQUIVO, COD_ITEM)
        If Not .r0200.Exists(Chave) Then Exit Function
        
        Campos = .r0200(Chave)
        i = Util.VerificarPosicaoInicialArray(Campos)
        
        ExtrairChaveReg0200 = Campos(dtoTitSPED.t0200("CHV_REG") - i)
        
    End With
    
End Function

Public Function ExtrairChaveRegC001()

Dim i As Integer
Dim Campos As Variant
    
    With dtoRegSPED
        
        If .rC001 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC001("ARQUIVO")
        
        If Not .rC001.Exists(DadosXML.ARQUIVO) Then Call CriarRegistroC001
        
        Campos = .rC001(DadosXML.ARQUIVO)
        i = Util.VerificarPosicaoInicialArray(Campos)
        
        ExtrairChaveRegC001 = Campos(dtoTitSPED.tC001("CHV_REG") - i)
        
    End With
    
End Function

Public Function ExtrairChaveRegC010() As String

Dim i As Byte
Dim Campos As Variant
Dim Chave As String, CHV_PAI$, CHV_REG$
    
    With dtoRegSPED
        
        If .rC010 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC010("ARQUIVO", "CNPJ")
        
        Chave = Util.UnirCampos(DadosXML.ARQUIVO, DadosXML.CNPJ_ESTABELECIMENTO)
        If Not .rC010.Exists(Chave) Then Call CriarRegistroC010
        
        Campos = .rC010(Chave)
        i = Util.VerificarPosicaoInicialArray(Campos)
        
        ExtrairChaveRegC010 = Campos(dtoTitSPED.tC010("CHV_REG") - i)
        
    End With
    
End Function

Public Function ExtrairChaveRegD001()

Dim i As Integer
Dim Campos As Variant
    
    With dtoRegSPED
        
        If .rD001 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroD001("ARQUIVO")
        
        If Not .rD001.Exists(DadosXML.ARQUIVO) Then Call CriarRegistroD001
        
        Campos = .rD001(DadosXML.ARQUIVO)
        i = Util.VerificarPosicaoInicialArray(Campos)
        
        ExtrairChaveRegD001 = Campos(dtoTitSPED.tD001("CHV_REG") - i)
        
    End With
    
End Function

Public Function ExtrairChaveRegD010() As String

Dim i As Byte
Dim Campos As Variant
Dim Chave As String, CHV_PAI$, CHV_REG$
    
    With dtoRegSPED
        
        If .rD010 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroD010("ARQUIVO", "CNPJ")
        
        Chave = Util.UnirCampos(DadosXML.ARQUIVO, DadosXML.CNPJ_ESTABELECIMENTO)
        If Not .rD010.Exists(Chave) Then Call CriarRegistroD010
        
        Campos = .rD010(Chave)
        i = Util.VerificarPosicaoInicialArray(Campos)
        
        ExtrairChaveRegD010 = Campos(dtoTitSPED.tD010("CHV_REG") - i)
        
    End With
    
End Function

Public Function DefinirPeriodoArquivo(ByRef NFe As DOMDocument60) As String

Dim Periodo As String
    
    Periodo = Util.ExtrairPeriodo(fnXML.ExtrairDataDocumento(NFe))
    If UsarPeriodo And PeriodoEspecifico <> "" Then Periodo = VBA.Format(PeriodoEspecifico, "00/0000")
    
    DefinirPeriodoArquivo = Periodo
    
End Function

Public Function DefinirCNPJEstabelecimento() As String
    
    With DadosXML
        
        If .CNPJ_EMITENTE Like CNPJBase & "*" Then
            
            DefinirCNPJEstabelecimento = .CNPJ_EMITENTE
            
        ElseIf .CNPJ_DESTINATARIO Like CNPJBase & "*" Then
            
            DefinirCNPJEstabelecimento = .CNPJ_DESTINATARIO
            
        End If
        
    End With
    
End Function

Public Function ValidarParticipante() As Boolean
    
    With DadosXML
        
        If (.CNPJ_EMITENTE Like CNPJBase & "*") Or (.CNPJ_DESTINATARIO Like CNPJBase & "*") Then ValidarParticipante = True
        
    End With
    
End Function

Public Function ExtrairCodigoParticipante()

Dim i As Byte
Dim Campos As Variant
Dim Chave As String, COD_PART$

    If dtoRegSPED.r0150 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0150("ARQUIVO")

    COD_PART = IdentificarParticipante()

    With dtoRegSPED

        Chave = Util.UnirCampos(ARQUIVO, COD_PART)
        If .r0150.Exists(Chave) Then

            Campos = .r0150(Chave)
            i = Util.VerificarPosicaoInicialArray(Campos)

            COD_PART = Campos(dtoTitSPED.t0150("COD_PART") - i)

        End If

    End With

    ExtrairCodigoParticipante = COD_PART

End Function

Public Function IdentificarParticipante()

Dim tpNF As String, tpEmit$

    With CamposC100

        tpNF = VBA.Left(.IND_OPER, 1)
        tpEmit = VBA.Left(.IND_EMIT, 1)

    End With
    
    With DadosXML
        
        Select Case True
            
            Case (tpNF = "1" And tpEmit = "0") Or (tpNF = "1" And tpEmit = "1") Or (tpNF = "0" And tpEmit = "0")
                IdentificarParticipante = .CNPJ_DESTINATARIO
                
            Case tpNF = "0" And tpEmit = "1"
                IdentificarParticipante = .CNPJ_EMITENTE
                
        End Select
        
    End With
    
End Function

Public Function ValidarModeloDocumento(ByVal NFe As IXMLDOMNode)

Dim COD_MOD As String, CHV_NFE$
    
    With DadosXML
        
        COD_MOD = fnXML.ValidarTag(NFe, "//mod")
        CHV_NFE = fnXML.ExtrairChaveAcessoNFe(NFe)
        
        Select Case True
            
            Case COD_MOD Like "55*"
                ValidarModeloDocumento = True
                
            Case CHV_NFE Like "*" & .CNPJ_ESTABELECIMENTO & "*" And COD_MOD Like "65*"
                ValidarModeloDocumento = True
                
        End Select
        
    End With
    
End Function

Public Function DefinirTipoContribuinteNFe() As String
    
    With DadosXML
        
        If .CNPJ_EMITENTE Like CNPJBase & "*" Then
            
            DefinirTipoContribuinteNFe = "emit"
            
        ElseIf .CNPJ_DESTINATARIO Like CNPJBase & "*" Then
            
            DefinirTipoContribuinteNFe = "dest"
            
        End If
        
    End With
    
End Function

Public Function ExtrairDadosXML(ByVal XML As Variant) As DOMDocument60
    
    Call DTO_DadosXML.ResetarDadosXML
    
    Set Doce = fnXML.RemoverNamespaces(XML)
    If Not fnXML.ValidarXML(Doce) Then Exit Function
    
    With DadosXML
                
        .CNPJ_EMITENTE = fnXML.ExtrairCNPJEmitente(Doce)
        .CNPJ_DESTINATARIO = fnXML.ExtrairCNPJDestinatario(Doce)
        .CNPJ_ESTABELECIMENTO = DefinirCNPJEstabelecimento()
        
        If .CNPJ_ESTABELECIMENTO = "" Then Exit Function
        
        .Periodo = DefinirPeriodoArquivo(Doce)
        .ARQUIVO = .Periodo & "-" & CNPJContribuinte
        .TIPO_NF = fnXML.ValidarTag(Doce, "//tpNF")
        
    End With
    
    If Not ValidarParticipante() Or Not ValidarModeloDocumento(Doce) Then Exit Function
    
    Set ExtrairDadosXML = Doce
    
End Function

Public Function ExtrairTipoParticipante() As String
    
    With DadosXML
        
        If .CNPJ_EMITENTE Like CNPJBase & "*" Then
            
            ExtrairTipoParticipante = "dest"
            
        ElseIf .CNPJ_DESTINATARIO Like CNPJBase & "*" Then
            
            ExtrairTipoParticipante = "emit"
            
        End If
        
    End With
    
End Function

Public Function IdentificarTipoOperacao() As String
    
    With DadosXML
        
        Select Case True
            
            Case (.TIPO_NF = "1" And .CNPJ_EMITENTE = .CNPJ_DESTINATARIO And .CNPJ_EMITENTE Like CNPJBase & "*")
                IdentificarTipoOperacao = EnumContrib.ValidarEnumeracao_IND_OPER("1")
                
            Case (.TIPO_NF = "1" And .CNPJ_EMITENTE Like CNPJBase & "*")
                IdentificarTipoOperacao = EnumContrib.ValidarEnumeracao_IND_OPER("1")
                
            Case (.TIPO_NF = "0" And .CNPJ_EMITENTE = .CNPJ_DESTINATARIO And .CNPJ_EMITENTE Like CNPJBase & "*")
                IdentificarTipoOperacao = EnumContrib.ValidarEnumeracao_IND_OPER("0")
                
            Case Else
                IdentificarTipoOperacao = EnumContrib.ValidarEnumeracao_IND_OPER("0")
                
        End Select
        
    End With
    
End Function

Public Function IdentificarTipoEmissao(ByRef NFe As IXMLDOMNode) As String
    
    With DadosXML
        
        Select Case True
            
            Case (.CNPJ_DESTINATARIO Like CNPJBase & "*" And .TIPO_NF = "0")
                IdentificarTipoEmissao = EnumContrib.ValidarEnumeracao_IND_EMIT("1")
                
            Case (.CNPJ_EMITENTE Like CNPJBase & "*")
                IdentificarTipoEmissao = EnumContrib.ValidarEnumeracao_IND_EMIT("0")
                
            Case Else
                IdentificarTipoEmissao = EnumContrib.ValidarEnumeracao_IND_EMIT("1")
                
        End Select
        
    End With
    
End Function

Public Function IdentificarTipoOperacao_C113(ByVal TIPO_NF As String, ByVal CNPJ_EMITENTE As String) As String
    
    With DadosXML
        
        Select Case True
            
            Case (TIPO_NF = "1" And CNPJ_EMITENTE = .CNPJ_DESTINATARIO And CNPJ_EMITENTE Like CNPJBase & "*")
                IdentificarTipoOperacao_C113 = EnumContrib.ValidarEnumeracao_IND_OPER("1")
                
            Case (TIPO_NF = "1" And CNPJ_EMITENTE Like CNPJBase & "*")
                IdentificarTipoOperacao_C113 = EnumContrib.ValidarEnumeracao_IND_OPER("1")
                
            Case (TIPO_NF = "0" And CNPJ_EMITENTE = .CNPJ_DESTINATARIO And CNPJ_EMITENTE Like CNPJBase & "*")
                IdentificarTipoOperacao_C113 = EnumContrib.ValidarEnumeracao_IND_OPER("0")
                
            Case Else
                IdentificarTipoOperacao_C113 = EnumContrib.ValidarEnumeracao_IND_OPER("0")
                
        End Select
        
    End With
    
End Function

Public Function IdentificarTipoEmissao_C113(ByVal TIPO_NF As String, ByVal CNPJ_EMITENTE As String) As String
    
    With DadosXML
        
        Select Case True
            
            Case (CNPJ_EMITENTE Like CNPJBase & "*")
                IdentificarTipoEmissao_C113 = EnumContrib.ValidarEnumeracao_IND_EMIT("0")
                
            Case Else
                IdentificarTipoEmissao_C113 = EnumContrib.ValidarEnumeracao_IND_EMIT("1")
                
        End Select
        
    End With
    
End Function

Public Function AjustarCST_IPI(ByVal CST_IPI As String) As String
    
    Select Case True
        
        Case CST_IPI Like "99*"
            AjustarCST_IPI = EnumContrib.ValidarEnumeracao_CST_IPI(49)
            
        Case CST_IPI Like "5*"
            AjustarCST_IPI = EnumContrib.ValidarEnumeracao_CST_IPI("0" & VBA.Right(Util.ApenasNumeros(CST_IPI), 1))
            
    End Select
    
End Function

Public Function CadastrarItensTerceirosComoProprios(ByVal Produto As IXMLDOMNode)

    With CamposC170

        .COD_ITEM = DadosXML.CNPJ_EMITENTE & " - " & .COD_ITEM
        Call CriarRegistro0190(.UNID)

    End With
    
    With Campos0200

        .COD_ITEM = CamposC170.COD_ITEM
        .DESCR_ITEM = Util.RemoverPipes(fnXML.ValidarTag(Produto, "prod/xProd"))
        .UNID_INV = ValidarTag(Produto, "prod/uCom")

        Call CriarRegistro0200(Produto)

    End With

End Function

Public Function ExtrairCST_PIS_COFINS_AquisicaoFrete(ByVal ARQUIVO As String) As String

Dim i As Byte
Dim Campos As Variant
Dim COD_INC_TRIB As String
    
    Campos = dtoRegSPED.r0110(ARQUIVO)
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    COD_INC_TRIB = VBA.Left(Util.ApenasNumeros(Campos(dtoTitSPED.t0110("COD_INC_TRIB") - i)), 1)
    
    Select Case COD_INC_TRIB
        
        Case "1"
            ExtrairCST_PIS_COFINS_AquisicaoFrete = EnumContrib.ValidarEnumeracao_CST_PIS_COFINS("50")
            
        Case "2"
            ExtrairCST_PIS_COFINS_AquisicaoFrete = EnumContrib.ValidarEnumeracao_CST_PIS_COFINS("70")
            
    End Select
    
End Function

Public Function ExtrairALIQ_PIS_AquisicaoFrete(ByVal ARQUIVO As String) As Double

Dim i As Byte
Dim Campos As Variant
Dim COD_INC_TRIB As String
        
    Campos = dtoRegSPED.r0110(ARQUIVO)
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    COD_INC_TRIB = VBA.Left(Util.ApenasNumeros(Campos(dtoTitSPED.t0110("COD_INC_TRIB") - i)), 1)
    
    Select Case COD_INC_TRIB
        
        Case "1"
            ExtrairALIQ_PIS_AquisicaoFrete = 0.0165
            
        Case "2"
            ExtrairALIQ_PIS_AquisicaoFrete = 0
            
    End Select
    
End Function

Public Function ExtrairALIQ_COFINS_AquisicaoFrete(ByVal ARQUIVO As String) As Double

Dim i As Byte
Dim Campos As Variant
Dim COD_INC_TRIB As String
        
    Campos = dtoRegSPED.r0110(ARQUIVO)
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    COD_INC_TRIB = VBA.Left(Util.ApenasNumeros(Campos(dtoTitSPED.t0110("COD_INC_TRIB") - i)), 1)
    
    Select Case COD_INC_TRIB
        
        Case "1"
            ExtrairALIQ_COFINS_AquisicaoFrete = 0.076
            
        Case "2"
            ExtrairALIQ_COFINS_AquisicaoFrete = 0
            
    End Select
    
End Function

Public Function ValidarXML(ByRef NFe As DOMDocument60) As Boolean
    If NFe.parseError.ErrorCode = 0 Then ValidarXML = True
End Function

Public Function ValidarTag(ByVal NFe As IXMLDOMNode, Tag As String) As String
    If Not NFe.SelectSingleNode(Tag) Is Nothing Then ValidarTag = NFe.SelectSingleNode(Tag).text
End Function

Public Function ValidarValores(ByVal NFe As IXMLDOMNode, Tag As String) As Double
    If Not NFe.SelectSingleNode(Tag) Is Nothing Then ValidarValores = Replace(NFe.SelectSingleNode(Tag).text, ".", ",")
End Function

Public Function ValidarPercentuais(ByRef NFe As IXMLDOMNode, ByVal Tag As String) As Double
    If Not NFe.SelectSingleNode(Tag) Is Nothing Then ValidarPercentuais = Replace(NFe.SelectSingleNode(Tag).text, ".", ",") / 100
End Function

