Attribute VB_Name = "AssistenteDivergencias"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Campo As Variant
Public TipoRegistro As String
Public TipoRelatorio As String
Public CampoRelatorio As Variant
Public dicOperacoesXML As New Dictionary
Public dicOperacoesSPED As New Dictionary
Public arrTitulosRelatorio As New ArrayList
Public GerenciadorSPED As New clsRegistrosSPED
Public dicTitulosDivergencias As New Dictionary
Public dicTitulosRelatorioNotas As New Dictionary
Public dicTitulosRelatorioProdutos As New Dictionary

Public Sub CarregarDadosRegistro0000(Optional Contrib As Boolean = False)
    
    Select Case Contrib
        
        Case True
            GerenciadorSPED.Contrib = Contrib
            Call GerenciadorSPED.CarregarDadosRegistro0000("ARQUIVO")
            
        Case Else
            GerenciadorSPED.Contrib = Contrib
            Call GerenciadorSPED.CarregarDadosRegistro0000("ARQUIVO")
            
    End Select
    
End Sub

Private Sub CarregarDadosRegistro0150(Optional Contrib As Boolean = False)
    
    Select Case Contrib
        
        Case True
            Call GerenciadorSPED.CarregarDadosRegistro0150("CHV_PAI_CONTRIBUICOES", "COD_PART")
            
        Case Else
            Call GerenciadorSPED.CarregarDadosRegistro0150("CHV_PAI_FISCAL", "COD_PART")
            
    End Select
        
End Sub

Private Sub CarregarDadosRegistro0200(Optional Contrib As Boolean = False)
    
    Select Case Contrib
        
        Case True
            Call GerenciadorSPED.CarregarDadosRegistro0200("CHV_PAI_CONTRIBUICOES", "COD_ITEM")
            
        Case Else
            Call GerenciadorSPED.CarregarDadosRegistro0200("CHV_PAI_FISCAL", "COD_ITEM")
            
    End Select
        
End Sub

Public Function ExtrairCHV_0001(ByVal ARQUIVO As String) As String
    
    With dtoRegSPED
        
        If .r0001 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0001("ARQUIVO")
        If .r0001.Exists(ARQUIVO) Then ExtrairCHV_0001 = .r0001(ARQUIVO)(dtoTitSPED.t0001("CHV_REG"))
        
    End With
    
End Function

Public Function ExtrairCHV_0140(ByVal Chave As String) As String
    
    With dtoRegSPED
        
        If .r0140 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0140("ARQUIVO", "CNPJ")
        If .r0140.Exists(Chave) Then ExtrairCHV_0140 = .r0140(Chave)(dtoTitSPED.t0140("CHV_REG"))
        
    End With
    
End Function

Public Function ExtrairCHV_C001(ByVal ARQUIVO As String) As String
        
    With dtoRegSPED
        
        If .rC001 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC001("ARQUIVO")
        If .rC001.Exists(ARQUIVO) Then ExtrairCHV_C001 = .rC001(ARQUIVO)(dtoTitSPED.tC001("CHV_REG"))
        
    End With
    
End Function

Public Function ExtrairCHV_C010(ByVal Chave As String) As String
        
    With dtoRegSPED
        
        If .rC010 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC010("ARQUIVO", "CNPJ")
        If .rC010.Exists(Chave) Then ExtrairCHV_C010 = .rC010(Chave)(dtoTitSPED.tC010("CHV_REG"))
        
    End With
    
End Function

Public Sub ExtrairDados0150(ByVal CHV_PAI As String, ByVal COD_PART As String)

Dim CHV_REG As String, CNPJ$, CPF$, NOME$, IE$
    
    With dtoRegSPED
        
        CHV_REG = Util.UnirCampos(CHV_PAI, COD_PART)
        If .r0150.Exists(CHV_REG) Then
            
            CNPJ = .r0150(CHV_REG)(dtoTitSPED.t0150("CNPJ"))
            CPF = .r0150(CHV_REG)(dtoTitSPED.t0150("CPF"))
            NOME = .r0150(CHV_REG)(dtoTitSPED.t0150("NOME"))
            IE = .r0150(CHV_REG)(dtoTitSPED.t0150("IE"))
            
            AtribuirValor "COD_PART", fnExcel.FormatarTexto(CNPJ & CPF)
            AtribuirValor "NOME_RAZAO", NOME
            AtribuirValor "INSC_EST", fnExcel.FormatarTexto(IE)
            
        End If
        
    End With
    
End Sub

Public Sub ExtrairUnidade0190(ByVal CHV_PAI As String, ByVal UNID As String)

Dim CHV_REG As String, DESCR$
    
    With dtoRegSPED
                        
        CHV_REG = Util.UnirCampos(CHV_PAI, UNID)
        If .r0190.Exists(CHV_REG) Then
                    
            UNID = .r0190(CHV_REG)(dtoTitSPED.t0190("UNID"))
            DESCR = .r0190(CHV_REG)(dtoTitSPED.t0190("DESCR"))
    
            If Util.ApenasNumeros(UNID) <> "" Then _
                AtribuirValor "UNID", UNID & " - " & DESCR _
                    Else AtribuirValor "UNID", UNID
            
        End If
    
    End With
    
End Sub

Public Sub ExtrairDados0200(ByVal CHV_PAI As String, ByVal COD_ITEM As String)

Dim i As Byte
Dim CHV_REG As String
    
    With dtoRegSPED
                
        CHV_REG = Util.UnirCampos(CHV_PAI, COD_ITEM)
        If .r0200.Exists(CHV_REG) Then
            
            AtribuirValor "DESCR_ITEM", .r0200(CHV_REG)(dtoTitSPED.t0200("DESCR_ITEM"))
            AtribuirValor "COD_BARRA", fnExcel.FormatarTexto(.r0200(CHV_REG)(dtoTitSPED.t0200("COD_BARRA")))
            AtribuirValor "COD_NCM", fnExcel.FormatarTexto(.r0200(CHV_REG)(dtoTitSPED.t0200("COD_NCM")))
            AtribuirValor "EX_IPI", fnExcel.FormatarTexto(.r0200(CHV_REG)(dtoTitSPED.t0200("EX_IPI")))
            AtribuirValor "CEST", fnExcel.FormatarTexto(.r0200(CHV_REG)(dtoTitSPED.t0200("CEST")))
            
        Else
            
            AtribuirValor "DESCR_ITEM", "ITEM NÃO IDENTIFICADO"
            
        End If
    
    End With
    
End Sub

Public Sub ExtrairDadosC100(ByVal CHV_REG As String)

Dim i As Byte
    
    With dtoRegSPED
        
        If .rC100 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC100
        If Not Util.ValidarDicionario(.rC100) Then Exit Sub
        
        If .rC100.Exists(CHV_REG) Then
            
            AtribuirValor "CHV_NFE", fnExcel.FormatarTexto(.rC100(CHV_REG)(dtoTitSPED.tC100("CHV_NFE")))
            AtribuirValor "NUM_DOC", fnExcel.FormatarTexto(.rC100(CHV_REG)(dtoTitSPED.tC100("NUM_DOC")))
            AtribuirValor "SER", fnExcel.FormatarTexto(.rC100(CHV_REG)(dtoTitSPED.tC100("SER")))
            
        End If
    
    End With
    
End Sub

Public Sub ExtrairIdentificacaoNFe(ByRef NFe As DOMDocument60)

Dim i As Byte
    
    AtribuirValor "CHV_NFE", fnExcel.FormatarTexto(fnXML.ExtrairChaveAcessoNFe(NFe))
    AtribuirValor "NUM_DOC", fnExcel.FormatarTexto(fnXML.ExtrairNumeroDocumentoNFe(NFe))
    AtribuirValor "SER", fnExcel.FormatarTexto(fnXML.ExtrairSerieDocumentoNFe(NFe))
    
End Sub

Public Sub ExtrairCadastroProduto(ByRef Produto As IXMLDOMNode)

Dim i As Byte
Dim CEST As String

    AtribuirValor "COD_ITEM", fnExcel.FormatarTexto(fnXML.ValidarTag(Produto, "prod/cProd"))
    AtribuirValor "DESCR_ITEM", fnExcel.FormatarTexto(fnXML.ValidarTag(Produto, "prod/xProd"))
    AtribuirValor "COD_BARRA", fnXML.ExtrairCodigoBarrasProduto(Produto)
    AtribuirValor "COD_NCM", fnExcel.FormatarTexto(fnXML.ValidarTag(Produto, "prod/NCM"))
    AtribuirValor "EX_IPI", fnExcel.FormatarTexto(fnXML.ValidarTag(Produto, "prod/EXTIPI"))
    
    CEST = fnXML.ExtrairCEST(Produto)
    If CEST <> 0 Then AtribuirValor "CEST", fnExcel.FormatarTexto(CEST) Else AtribuirValor "CEST", ""
    
End Sub

Public Function ExtrairCNPJ_CONTRIBUINTE(ByVal ARQUIVO As String) As String
    
    ExtrairCNPJ_CONTRIBUINTE = VBA.Split(ARQUIVO, "-")(1)
    
End Function

Public Function ExtrairCNPJ_ESTABELECIMENTO_C100(ByVal CHV_C100 As String, Optional ARQUIVO As String) As String

Dim CHV_C010 As String
    
    With dtoRegSPED
        
        If .rC010 Is Nothing Then Call CarregarDadosRegistroC010
        If Not Util.ValidarDicionario(.rC010) Then
            ExtrairCNPJ_ESTABELECIMENTO_C100 = ExtrairCNPJ_CONTRIBUINTE(ARQUIVO)
            Exit Function
        End If
        
        If .rC100 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC100
        If Not Util.ValidarDicionario(.rC100) Then Exit Function
        
        If .rC100.Exists(CHV_C100) Then
            
            CHV_C010 = .rC100(CHV_C100)(dtoTitSPED.tC100("CHV_PAI_FISCAL"))
            If .rC010.Exists(CHV_C010) Then ExtrairCNPJ_ESTABELECIMENTO_C100 = .rC010(CHV_C010)(dtoTitSPED.tC010("CNPJ"))
            
        End If
    
    End With
    
End Function

Public Function Extrair_CHV_NFE(ByVal CHV_REG As String) As String
    
    With dtoRegSPED
        
        If .rC100 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC100
        If Not Util.ValidarDicionario(.rC100) Then Exit Function
        
        If .rC100.Exists(CHV_REG) Then
            Extrair_CHV_NFE = fnExcel.FormatarTexto(.rC100(CHV_REG)(dtoTitSPED.tC100("CHV_NFE")))
        End If
    
    End With
    
End Function

Public Function Extrair_VL_OPER(ByRef dicTitulos As Dictionary, ByRef Campos As Variant) As Double

Dim VL_ITEM As Double, VL_DESC#, VL_ICMS_ST#, VL_IPI#
    
    VL_ITEM = fnExcel.ConverterValores(Campos(dicTitulos("VL_ITEM")), True, 2)
    VL_DESC = fnExcel.ConverterValores(Campos(dicTitulos("VL_DESC")), True, 2)
    VL_ICMS_ST = fnExcel.ConverterValores(Campos(dicTitulos("VL_ICMS_ST")), True, 2)
    VL_IPI = fnExcel.ConverterValores(Campos(dicTitulos("VL_IPI")), True, 2)
    
    Extrair_VL_OPER = VBA.Round(VL_ITEM - VL_DESC + VL_ICMS_ST + VL_IPI, 2)
    
End Function

Public Function Extrair_VL_OPER_XML(ByRef Produto As IXMLDOMNode) As Double

Dim vProd As Double, vDesc#, vFrete#, vSeg#, vOutro#, vDesp#, vFCPST#, vICMSST#, vIPI#
    
    vProd = fnExcel.FormatarValores(fnXML.ValidarValores(Produto, "prod/vProd"), True, 2)
    vDesc = fnExcel.FormatarValores(fnXML.ValidarValores(Produto, "prod/vDesc"), True, 2)
    vFrete = fnExcel.FormatarValores(fnXML.ValidarValores(Produto, "prod/vFrete"), True, 2)
    vSeg = fnExcel.FormatarValores(fnXML.ValidarValores(Produto, "prod/vSeg"), True, 2)
    vOutro = fnExcel.FormatarValores(fnXML.ValidarValores(Produto, "prod/vOutro"), True, 2)
    vICMSST = fnExcel.FormatarValores(fnXML.ExtrairValorICMSST(Produto), True, 2)
    vIPI = fnExcel.FormatarValores(fnXML.ValidarValores(Produto, "imposto/IPI//vIPI"), True, 2)
    
    vDesp = vFrete + vSeg + vOutro
    vICMSST = vICMSST + vFCPST
    
    Extrair_VL_OPER_XML = fnExcel.FormatarValores(vProd - vDesc + vICMSST + vIPI, True, 2)
    
End Function

Public Function AtribuirValor(ByVal Titulo As String, ByVal Valor As Variant)

Dim Posicao As Byte
    
    Posicao = RetornarPosicaoTitulo(Titulo)
    Campo(Posicao) = Valor
    
End Function

Public Function AtribuirCorrelacao(ByVal Titulo As String, ByVal Valor As Variant)

Dim Posicao As Byte
    
    Posicao = RetornarPosicaoTitulo(Titulo)
    CampoRelatorio(Posicao) = Valor
    
End Function

Public Function RetornarPosicaoTitulo(ByVal Titulo As String) As Byte
    
    Select Case TipoRelatorio
        
        Case "Notas"
            RetornarPosicaoTitulo = RetornarPosicaoTituloRelatorioNotas(Titulo)
            
        Case "Produtos"
            RetornarPosicaoTitulo = RetornarPosicaoTituloRelatorioProdutos(Titulo)
            
        Case "SPED", "XML"
            If arrTitulosRelatorio.Count = 0 Then Call CarregarTitulosRelatorio
            RetornarPosicaoTitulo = arrTitulosRelatorio.IndexOf(Titulo, 0) + 1
            
    End Select
    
End Function

Public Function RetornarPosicaoTituloRelatorioNotas(ByVal Titulo As String) As Byte
    
    If dicTitulosRelatorioNotas.Count = 0 Then Set dicTitulosRelatorioNotas = Util.MapearTitulos(relDivergenciasNotas, 3)
    RetornarPosicaoTituloRelatorioNotas = dicTitulosRelatorioNotas(Titulo)
    
End Function

Public Function RetornarPosicaoTituloRelatorioProdutos(ByVal Titulo As String) As Byte
    
    If dicTitulosRelatorioProdutos.Count = 0 Then Set dicTitulosRelatorioProdutos = Util.MapearTitulos(relDivergenciasProdutos, 3)
    RetornarPosicaoTituloRelatorioProdutos = dicTitulosRelatorioProdutos(Titulo)
    
End Function

Public Function RedimensionarArray(ByVal NumCampos As Long)
    
    ReDim Campo(1 To NumCampos) As Variant
    
End Function

Public Function RedimensionarCampoRelatorioProdutos()
    
    If dicTitulosRelatorioProdutos.Count = 0 Then Set dicTitulosRelatorioProdutos = Util.MapearTitulos(relDivergenciasProdutos, 3)
    ReDim CampoRelatorio(1 To dicTitulosRelatorioProdutos.Count) As Variant
    
End Function

Public Function RedimensionarCampoRelatorioNotas()
    
    If dicTitulosRelatorioNotas.Count = 0 Then Set dicTitulosRelatorioNotas = Util.MapearTitulos(relDivergenciasNotas, 3)
    ReDim CampoRelatorio(1 To dicTitulosRelatorioNotas.Count) As Variant
    
End Function

Public Function ExtrairValorCampo(ByVal Titulo As String)

Dim Posicao As Byte
    
    Posicao = RetornarPosicaoTitulo(Titulo)
    ExtrairValorCampo = Campo(Posicao)
    
End Function

Public Function RegistrarNotaSPED()

Dim CHV_NFE As String
    
    CHV_NFE = Util.ApenasNumeros(ExtrairValorCampo("CHV_NFE"))
    
    If Not Util.VerificarStringVazia(CHV_NFE) Then
        
        If Not dicOperacoesSPED.Exists(CHV_NFE) Then dicOperacoesSPED(CHV_NFE) = Campo
        
    End If
    
End Function

Public Function RegistrarOperacaoSPED()

Dim CHV_NFE As String
Dim NUM_ITEM As Integer
    
    CHV_NFE = Util.ApenasNumeros(ExtrairValorCampo("CHV_NFE"))
    NUM_ITEM = Util.ApenasNumeros(ExtrairValorCampo("NUM_ITEM"))
    
    If Not Util.VerificarStringVazia(CHV_NFE) Then
        
        If Not dicOperacoesSPED.Exists(CHV_NFE) Then Set dicOperacoesSPED(CHV_NFE) = New Dictionary
        dicOperacoesSPED(CHV_NFE)(NUM_ITEM) = Campo
        
    End If
    
End Function

Public Function RegistrarNotaXML()

Dim CHV_NFE As String
    
    CHV_NFE = Util.ApenasNumeros(ExtrairValorCampo("CHV_NFE"))
    
    If Not Util.VerificarStringVazia(CHV_NFE) Then
        
        If Not dicOperacoesXML.Exists(CHV_NFE) Then dicOperacoesXML(CHV_NFE) = Campo
        
    End If
    
End Function

Public Function RegistrarOperacaoXML()

Dim CHV_NFE As String
Dim NUM_ITEM As Integer
    
    CHV_NFE = Util.ApenasNumeros(ExtrairValorCampo("CHV_NFE"))
    NUM_ITEM = Util.ApenasNumeros(ExtrairValorCampo("NUM_ITEM"))
    
    If Not Util.VerificarStringVazia(CHV_NFE) Then
        
        If Not dicOperacoesXML.Exists(CHV_NFE) Then Set dicOperacoesXML(CHV_NFE) = New Dictionary
        dicOperacoesXML(CHV_NFE)(NUM_ITEM) = Campo
        
    End If
    
End Function

Public Function VerificarExistenciaSPED() As Boolean

Dim CHV_NFE As String
    
    CHV_NFE = Util.ApenasNumeros(ExtrairValorCampo("CHV_NFE"))
    
    If dicOperacoesSPED.Exists(CHV_NFE) Then VerificarExistenciaSPED = True
    
End Function

Public Sub CarregarTitulosRelatorio()
    
    Select Case TipoRelatorio
        
        Case "Notas"
            Call CarregarTitulosRelatorioNotas
            
        Case "Produtos"
            Call CarregarTitulosRelatorioProdutos
            
    End Select
    
End Sub

Public Sub CarregarTitulosRelatorioProdutos()
    
    Call arrTitulosRelatorio.Clear
    
    arrTitulosRelatorio.Add "REG"
    arrTitulosRelatorio.Add "ARQUIVO"
    arrTitulosRelatorio.Add "CHV_PAI_FISCAL"
    arrTitulosRelatorio.Add "CHV_PAI_CONTRIBUICOES"
    arrTitulosRelatorio.Add "CHV_REG"
    arrTitulosRelatorio.Add "CHV_NFE"
    arrTitulosRelatorio.Add "NUM_DOC"
    arrTitulosRelatorio.Add "SER"
    arrTitulosRelatorio.Add "NUM_ITEM"
    arrTitulosRelatorio.Add "COD_ITEM"
    arrTitulosRelatorio.Add "DESCR_ITEM"
    arrTitulosRelatorio.Add "COD_BARRA"
    arrTitulosRelatorio.Add "COD_NCM"
    arrTitulosRelatorio.Add "EX_IPI"
    arrTitulosRelatorio.Add "CEST"
    arrTitulosRelatorio.Add "QTD"
    arrTitulosRelatorio.Add "UNID"
    arrTitulosRelatorio.Add "CFOP"
    arrTitulosRelatorio.Add "CST_ICMS"
    arrTitulosRelatorio.Add "VL_ITEM"
    arrTitulosRelatorio.Add "VL_DESC"
    arrTitulosRelatorio.Add "VL_BC_ICMS"
    arrTitulosRelatorio.Add "ALIQ_ICMS"
    arrTitulosRelatorio.Add "VL_ICMS"
    arrTitulosRelatorio.Add "VL_BC_ICMS_ST"
    arrTitulosRelatorio.Add "ALIQ_ST"
    arrTitulosRelatorio.Add "VL_ICMS_ST"
    arrTitulosRelatorio.Add "CST_IPI"
    arrTitulosRelatorio.Add "VL_BC_IPI"
    arrTitulosRelatorio.Add "ALIQ_IPI"
    arrTitulosRelatorio.Add "VL_IPI"
    arrTitulosRelatorio.Add "CST_PIS"
    arrTitulosRelatorio.Add "VL_BC_PIS"
    arrTitulosRelatorio.Add "ALIQ_PIS"
    arrTitulosRelatorio.Add "QUANT_BC_PIS"
    arrTitulosRelatorio.Add "ALIQ_PIS_QUANT"
    arrTitulosRelatorio.Add "VL_PIS"
    arrTitulosRelatorio.Add "CST_COFINS"
    arrTitulosRelatorio.Add "VL_BC_COFINS"
    arrTitulosRelatorio.Add "ALIQ_COFINS"
    arrTitulosRelatorio.Add "QUANT_BC_COFINS"
    arrTitulosRelatorio.Add "ALIQ_COFINS_QUANT"
    arrTitulosRelatorio.Add "VL_COFINS"
    arrTitulosRelatorio.Add "VL_OPER"
    
End Sub

Public Sub CarregarTitulosRelatorioNotas()
    
    Call arrTitulosRelatorio.Clear
    
    arrTitulosRelatorio.Add "REG"
    arrTitulosRelatorio.Add "ARQUIVO"
    arrTitulosRelatorio.Add "CHV_PAI_FISCAL"
    arrTitulosRelatorio.Add "CHV_PAI_CONTRIBUICOES"
    arrTitulosRelatorio.Add "CHV_REG"
    arrTitulosRelatorio.Add "CHV_NFE"
    arrTitulosRelatorio.Add "COD_MOD"
    arrTitulosRelatorio.Add "NUM_DOC"
    arrTitulosRelatorio.Add "SER"
    arrTitulosRelatorio.Add "IND_OPER"
    arrTitulosRelatorio.Add "IND_EMIT"
    arrTitulosRelatorio.Add "COD_SIT"
    arrTitulosRelatorio.Add "COD_PART"
    arrTitulosRelatorio.Add "NOME_RAZAO"
    arrTitulosRelatorio.Add "INSC_EST"
    arrTitulosRelatorio.Add "DT_DOC"
    arrTitulosRelatorio.Add "DT_E_S"
    arrTitulosRelatorio.Add "IND_PGTO"
    arrTitulosRelatorio.Add "VL_DOC"
    arrTitulosRelatorio.Add "VL_DESC"
    arrTitulosRelatorio.Add "VL_ABAT_NT"
    arrTitulosRelatorio.Add "VL_MERC"
    arrTitulosRelatorio.Add "IND_FRT"
    arrTitulosRelatorio.Add "VL_FRT"
    arrTitulosRelatorio.Add "VL_SEG"
    arrTitulosRelatorio.Add "VL_OUT_DA"
    arrTitulosRelatorio.Add "VL_BC_ICMS"
    arrTitulosRelatorio.Add "VL_ICMS"
    arrTitulosRelatorio.Add "VL_BC_ICMS_ST"
    arrTitulosRelatorio.Add "VL_ICMS_ST"
    arrTitulosRelatorio.Add "VL_IPI"
    arrTitulosRelatorio.Add "VL_PIS"
    arrTitulosRelatorio.Add "VL_COFINS"
    arrTitulosRelatorio.Add "VL_PIS_ST"
    arrTitulosRelatorio.Add "VL_COFINS_ST"
    
End Sub
