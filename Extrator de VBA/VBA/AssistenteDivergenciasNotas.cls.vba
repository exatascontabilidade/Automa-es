Attribute VB_Name = "AssistenteDivergenciasNotas"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private XMLS As New ArrayList
Private XMLsAusentes As Integer
Private CamposRelatorio As Variant
Private arrRelatorio As New ArrayList
Public dicDados0150 As New Dictionary
Private dicTitulos0150 As New Dictionary
Private arrTitulosRelatorio As New ArrayList
Private dicTitulosRelatorio As New Dictionary
Private GerenciadorSPED As clsRegistrosSPED
Private EnumFiscal As New clsEnumeracoesSPEDFiscal
Private Divergencias As AssistenteDivergencias
Private ValidacoesNotas As New clsRegrasDivergenciasNotas
Private Const TitulosIgnorar As String = "REG, ARQUIVO, CHV_REG, CHV_PAI_FISCAL, CHV_PAI_CONTRIBUICOES, CHV_NFE"

Public Sub GerarComparativoXMLSPED()

Dim Msg As String
Dim SPEDS As Variant
Dim Status As Boolean
Dim Result As VbMsgBoxResult
Dim RemovidaDuplicatas As Boolean
Dim TentavivaRemocaoDuplicidade As Boolean
    
    Status = CarregarDadosContribuinte
    If Not Status Then Exit Sub
    
    If Util.ChecarAusenciaDados(regC100, , "Dados ausentes no registro C100") Then Exit Sub
    
    If Not ProcessarDocumentos.CarregarXMLS("Lote") Then Exit Sub
    
    XMLsAusentes = 0
    Call InicializarObjetos
    
Reimportar:
    Call CarregarDadosC100
    Call CarregarDadosXML
    Call CorrelacionarNotas
    
    If arrRelatorio.Count > 0 Then
        
        Call Util.AtualizarBarraStatus("Exportando dados para o relatório...")
        
        Call ValidacoesNotas.IdentificarDivergenciasNotas(arrRelatorio)
        
        If relDivergenciasNotas.AutoFilter.FilterMode Then relDivergenciasNotas.ShowAllData
        
        Call Util.LimparDados(relDivergenciasNotas, 4, False)
        Call Util.ExportarDadosArrayList(relDivergenciasNotas, arrRelatorio)
        
        Call FuncoesFormatacao.DestacarInconsistencias(relDivergenciasNotas)
        
        Call Util.MsgInformativa("Relatório gerado com sucesso", "Relatório de Divergências de Notas", Inicio)
        
    Else
        
        Msg = "Nenhum dado encontrado para geração do relatório." & vbCrLf & vbCrLf
        Msg = Msg & "Por favor verifique se o SPED e/ou XMLs foram importados e tente novamente."
        Call Util.MsgAlerta(Msg, "Relatório de Divergências de Notas")
        
    End If
    
    Call ResetarObjetos
    
End Sub

Private Sub InicializarObjetos()
    
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    
    Call dicInconsistenciasIgnoradas.RemoveAll
    Call arrRelatorio.Clear
    
    Set dicTitulosRelatorio = Util.MapearTitulos(relDivergenciasNotas, 3)
    Set Divergencias = New AssistenteDivergencias
    Set GerenciadorSPED = New clsRegistrosSPED
    
    Call Divergencias.arrTitulosRelatorio.Clear

    With GerenciadorSPED
        
        Call .CarregarDadosRegistro0000("ARQUIVO")
        Call .CarregarDadosRegistro0000_Contr("ARQUIVO")
        Call .CarregarDadosRegistro0001("ARQUIVO")
        Call .CarregarDadosRegistro0140("CNPJ")
        Call .CarregarDadosRegistroC010
        Call .CarregarDadosRegistroC100
        
        If dtoRegSPED.r0000.Count = 0 Then _
            Call .CarregarDadosRegistro0150("CHV_PAI_CONTRIBUICOES", "COD_PART") _
                Else Call .CarregarDadosRegistro0150("CHV_PAI_FISCAL", "COD_PART")
            
    End With
    
    Call Util.DesabilitarControles
    
End Sub

Private Sub ResetarObjetos()
    
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    
    Call dicInconsistenciasIgnoradas.RemoveAll
    Call arrRelatorio.Clear
    
    Set dicTitulosRelatorio = Nothing
    Set GerenciadorSPED = Nothing
    Set Divergencias = Nothing
    
    Call Util.AtualizarBarraStatus(False)
    
    Call Util.HabilitarControles
    
End Sub

Private Function EmitirMensagemXMLsAusentes()

Dim Msg As String
    
    Select Case XMLsAusentes
        
        Case Is = 1
            Msg = "Foi identificado " & XMLsAusentes & "XML ausente na movimentação." & vbCrLf & vbCrLf
            Msg = Msg & "Identifique o XML ausente e faça a importação novamente"
            
        Case Is > 1
            Msg = "Foram identificados " & XMLsAusentes & "XMLS ausentes na movimentação." & vbCrLf & vbCrLf
            Msg = Msg & "Identifique os XMLs ausentes e faça a importação novamente"
            
    End Select
    
    Call Util.MsgInformativa(Msg, "Relatório de Divergências de Notas", Inicio)
    
End Function

Public Sub CarregarDadosC100()

Dim Campos As Variant, Campo
Dim Comeco As Double
Dim Msg As String
Dim a As Long
    
    Call GerenciadorSPED.CarregarDadosRegistroC100
    
    a = 0
    Comeco = Timer()
    With Divergencias
        
        .TipoRelatorio = "SPED"
        Call .dicOperacoesSPED.RemoveAll
        Call .CarregarTitulosRelatorioNotas
        Call .RedimensionarArray(.arrTitulosRelatorio.Count)
        
        For Each Campos In dtoRegSPED.rC100.Items()
            
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                For Each Campo In .arrTitulosRelatorio
                    
                    Call Util.AntiTravamento(a, 50, Msg & "Processando dados do registro C100", dtoRegSPED.rC100.Count, Comeco)
                    Call TratarCamposC100(Campos, Campo)
                    
                Next Campo
                
                Call .RegistrarNotaSPED
                
            End If
Prx:
        Next Campos
        
    End With
    
End Sub

Private Sub TratarCamposC100(ByVal Campos As Variant, ByVal Campo As String)

Const CamposIgnorar As String = "INSC_EST, NOME_RAZAO"
Dim ARQUIVO As String, CHV_REG$, CHV_PAI$, CHV_PAI_FISCAL$, CHV_PAI_CONTRIBUICOES$, COD_PART$, COD_MOD$, CNPJ$
    
    With Divergencias
        
        Select Case True
            
            Case CamposIgnorar Like "*" & Campo & "*"
                Exit Sub
                
            Case Campo Like "CHV_PAI_FISCAL"
                CHV_PAI_FISCAL = Campos(dtoTitSPED.tC100(Campo))
                .AtribuirValor Campo, CHV_PAI_FISCAL
                Call .ExtrairDadosC100(CHV_PAI_FISCAL)
                
            Case Campo Like "CHV_PAI_CONTRIBUICOES"
                CHV_PAI_CONTRIBUICOES = Campos(dtoTitSPED.tC100(Campo))
                .AtribuirValor Campo, CHV_PAI_CONTRIBUICOES
                Call .ExtrairDadosC100(CHV_PAI_CONTRIBUICOES)
                
            Case Campo Like "ARQUIVO"
                ARQUIVO = Campos(dtoTitSPED.tC100(Campo))
                .AtribuirValor Campo, ARQUIVO
                
            Case Campo Like "COD_PART"
                .AtribuirValor Campo, fnExcel.FormatarTexto(ExtrairCampoParticipante(dtoTitSPED.tC100, Campos, False, "CNPJ", "CPF"))
                .AtribuirValor "NOME_RAZAO", ExtrairCampoParticipante(dtoTitSPED.tC100, Campos, False, "NOME")
                .AtribuirValor "INSC_EST", fnExcel.FormatarTexto(ExtrairCampoParticipante(dtoTitSPED.tC100, Campos, False, "IE"))
                
            Case Campo Like "*SER*"
                .AtribuirValor Campo, VBA.Format(Campos(dtoTitSPED.tC100(Campo)), "000")
                
            Case Campo Like "*NUM_DOC*"
                .AtribuirValor Campo, CLng(Campos(dtoTitSPED.tC100(Campo)))
                
            Case Campo Like "DT_*"
                .AtribuirValor Campo, fnExcel.FormatarData(Campos(dtoTitSPED.tC100(Campo)))
                
            Case Campo Like "VL_*"
                .AtribuirValor Campo, fnExcel.ConverterValores(Campos(dtoTitSPED.tC100(Campo)))
                
            Case Campo Like "REG", Campo Like "CHV_REG"
                .AtribuirValor Campo, Campos(dtoTitSPED.tC100(Campo))
                
            Case Else
                .AtribuirValor Campo, fnExcel.FormatarTexto(Campos(dtoTitSPED.tC100(Campo)))
                
        End Select
        
    End With
    
End Sub

Private Function FormatarCamposC100(ByVal Campos As Variant, ByVal Campo As String) As Variant

Dim Campos0150 As Variant
Dim ARQUIVO As String, CHV_REG$, CHV_PAI$, CNPJ$, CHV_PAI_FISCAL$, COD_PART$, CHV_PAI_CONTRIBUICOES$, COD_MOD$, COD_SIT$
    
    With Divergencias
        
        Select Case True
            
            Case Campo Like "*COD_PART_SPED*"
                FormatarCamposC100 = ExtrairCampoParticipante(dicTitulosRelatorio, Campos, True, "COD_PART")
                
            Case Campo Like "*NOME_RAZAO_SPED*"
                Call IncluirCampoParticipante(Campos, "NOME", "NOME_RAZAO_SPED")
                
            Case Campo Like "*INSC_EST_SPED*"
                Call IncluirCampoParticipante(Campos, "IE", "INSC_EST_SPED")
                
            Case Campo Like "*SER*"
                FormatarCamposC100 = VBA.Format(Campos(dicTitulosRelatorio(Campo)), "000")
                
            Case Campo Like "*NUM_DOC*"
                FormatarCamposC100 = CLng(Campos(dicTitulosRelatorio(Campo)))
                
            Case Campo Like "DT_*"
                If Campos(dicTitulosRelatorio(Campo)) <> "" Then _
                    FormatarCamposC100 = fnExcel.FormatarData(Campos(dicTitulosRelatorio(Campo)))
                    
            Case Campo Like "VL_*"
                FormatarCamposC100 = fnExcel.ConverterValores(Campos(dicTitulosRelatorio(Campo)))
                
            Case Campo Like "REG", Campo Like "CHV_REG", Campo Like "CHV_PAI_FISCAL", Campo Like "ARQUIVO"
                FormatarCamposC100 = Campos(dicTitulosRelatorio(Campo))
                
            Case Else
                FormatarCamposC100 = fnExcel.FormatarTexto(Campos(dicTitulosRelatorio(Campo)))
                
        End Select
        
    End With
    
End Function

Private Function ExtrairCampoParticipante(ByRef dicTitulos As Dictionary, ByVal Campos As Variant, ByVal CamposSPED As Boolean, ParamArray CamposChave()) As String

Dim i As Integer
Dim Chave As String
Dim Campo As Variant, Valor
Dim Campos0001 As Variant, Campos0140, CamposC010, Campos0150
Dim ARQUIVO As String, CHV_REG$, CHV_PAI$, CNPJ$, CHV_PAI_FISCAL$, COD_PART$, CHV_PAI_CONTRIBUICOES$, COD_MOD$, COD_SIT$
    
    i = Util.VerificarPosicaoInicialArray(Campos)
        
    ARQUIVO = Campos(dicTitulos("ARQUIVO") - i)
    COD_PART = Campos(dicTitulos(IIf(CamposSPED = True, "COD_PART_SPED", "COD_PART")) - i)
    COD_SIT = Campos(dicTitulos(IIf(CamposSPED = True, "COD_SIT_SPED", "COD_SIT")) - i)
    COD_MOD = Campos(dicTitulos(IIf(CamposSPED = True, "COD_MOD_SPED", "COD_MOD")) - i)
    
    CHV_PAI_FISCAL = Campos(dtoTitSPED.tC100("CHV_PAI_FISCAL") - i)
    CHV_PAI_CONTRIBUICOES = Campos(dtoTitSPED.tC100("CHV_PAI_CONTRIBUICOES") - i)
    
    If Not COD_SIT Like "*02*" And Not COD_MOD Like "*65*" Then
        
        If CHV_PAI_FISCAL <> "" And dtoRegSPED.r0000.Count > 0 Then
            
            Campos0001 = dtoRegSPED.r0001(ARQUIVO)
            i = Util.VerificarPosicaoInicialArray(Campos0001)
            CHV_PAI = Campos0001(dtoTitSPED.t0001("CHV_REG") - i)
            
        ElseIf CHV_PAI_CONTRIBUICOES <> "" Then
            
            CamposC010 = dtoRegSPED.rC010(CHV_PAI_CONTRIBUICOES)
            i = Util.VerificarPosicaoInicialArray(CamposC010)
            CNPJ = CamposC010(dtoTitSPED.tC010("CNPJ") - i)
            
            Campos0140 = dtoRegSPED.r0140(CNPJ)
            i = Util.VerificarPosicaoInicialArray(Campos0140)
            CHV_PAI = Campos0140(dtoTitSPED.t0140("CHV_REG") - i)
            
        End If
        
        CHV_REG = Util.UnirCampos(CHV_PAI, COD_PART)
        If dtoRegSPED.r0150.Exists(CHV_REG) Then
            
            Campos0150 = dtoRegSPED.r0150(CHV_REG)
            i = Util.VerificarPosicaoInicialArray(Campos0150)
            
            For Each Campo In CamposChave
                
                Valor = Campos0150(dtoTitSPED.t0150(Campo) - i)
                Chave = Chave & Valor
                
            Next Campo
            
            ExtrairCampoParticipante = Chave
            
        End If
        
    End If
    
End Function

Private Sub IncluirCampoParticipante(ByVal Campos As Variant, ByVal Campo, ByVal CampoIncluir As String)

Dim i As Integer
Dim ValorIncluir As Variant
Dim Campos0001 As Variant, Campos0140, CamposC010, Campos0150
Dim ARQUIVO As String, CHV_REG$, CHV_PAI$, CNPJ$, CHV_PAI_FISCAL$, COD_PART$, CHV_PAI_CONTRIBUICOES$, COD_MOD$, COD_SIT$
    
    i = Util.VerificarPosicaoInicialArray(Campos)
    
    ARQUIVO = Campos(dicTitulosRelatorio("ARQUIVO") - i)
    COD_PART = Campos(dicTitulosRelatorio("COD_PART_SPED") - i)
    COD_SIT = Campos(dicTitulosRelatorio("COD_SIT_SPED") - i)
    COD_MOD = Campos(dicTitulosRelatorio("COD_MOD_SPED") - i)
    ValorIncluir = Campos(dicTitulosRelatorio(CampoIncluir) - i)
    
    CHV_PAI_FISCAL = Campos(dtoTitSPED.tC100("CHV_PAI_FISCAL") - i)
    CHV_PAI_CONTRIBUICOES = Campos(dtoTitSPED.tC100("CHV_PAI_CONTRIBUICOES") - i)
    
    If Not COD_SIT Like "*02*" And Not COD_MOD Like "*65*" Then
        
        If CHV_PAI_FISCAL <> "" And dtoRegSPED.r0000.Count > 0 Then
            
            Campos0001 = dtoRegSPED.r0001(ARQUIVO)
            i = Util.VerificarPosicaoInicialArray(Campos0001)
            CHV_PAI = Campos0001(dtoTitSPED.t0001("CHV_REG") - i)
            
        ElseIf CHV_PAI_CONTRIBUICOES <> "" Then
            
            CamposC010 = dtoRegSPED.rC010(CHV_PAI_CONTRIBUICOES)
            i = Util.VerificarPosicaoInicialArray(CamposC010)
            CNPJ = CamposC010(dtoTitSPED.tC010("CNPJ") - i)
            
            Campos0140 = dtoRegSPED.r0140(CNPJ)
            i = Util.VerificarPosicaoInicialArray(Campos0140)
            CHV_PAI = Campos0140(dtoTitSPED.t0140("CHV_REG") - i)
            
        End If
        
        CHV_REG = Util.UnirCampos(CHV_PAI, COD_PART)
        If dtoRegSPED.r0150.Exists(CHV_REG) Then
            
            Campos0150 = dtoRegSPED.r0150(CHV_REG)
                
                i = Util.VerificarPosicaoInicialArray(Campos0150)
                Campos0150(dtoTitSPED.t0150(Campo) - i) = ValorIncluir
                
            dtoRegSPED.r0150(CHV_REG) = Campos0150
            
        End If
        
    End If
    
End Sub

Public Sub CarregarDadosXML()

Dim DT_DOC As String, CHV_NFE$, COD_SIT$, COD_MOD$, tpEmit$
Dim NFe As New DOMDocument60
Dim Comeco As Double
Dim XML As Variant
Dim Msg As String
Dim a As Long
    
    a = 0
    Comeco = Timer
    
    Call SPEDFiscal.dicDados0001.RemoveAll
    Call SPEDFiscal.dicDados0150.RemoveAll
    
    With Divergencias
        
        .TipoRelatorio = "XML"
        Call .dicOperacoesXML.RemoveAll
        Call .CarregarTitulosRelatorioNotas
        
        Call ProcessarDocumentos.ListarTodosDocumentos
        
        For Each XML In DocsFiscais.arrTodos
            
            Call Util.AntiTravamento(a, 10, Msg & "Processando dados dos XMLS", DocsFiscais.arrTodos.Count, Comeco)
            Call .RedimensionarArray(.arrTitulosRelatorio.Count)
            Set NFe = fnXML.RemoverNamespaces(XML)
            
            If fnXML.ValidarNFe(NFe) Then Call InformarDadosXML(NFe)
            
        Next XML
        
    End With
    
    Call Util.LimparDados(reg0150, 4, False)
    Call Util.ExportarDadosDicionario(reg0150, SPEDFiscal.dicDados0150)
    
End Sub

Private Function InformarDadosXML(ByRef NFe As DOMDocument60)

Dim DT_DOC As String, CHV_NFE$, COD_SIT$, COD_MOD$, NOME_RAZAO$, tpEmit$
    
    With Divergencias
        
        tpEmit = fnXML.ExtrairTipoEmissao(NFe)
        DT_DOC = fnXML.ExtrairDataDocumento(NFe)
        CHV_NFE = fnXML.ExtrairChaveAcessoNFe(NFe)
        COD_SIT = fnXML.ExtrairSituacaoDocumento(NFe)
        COD_MOD = fnXML.ValidarTag(NFe, "//mod")
        
        tpEmit = fnXML.ExtrairTipoEmissao(NFe)
        DT_DOC = fnXML.ExtrairDataDocumento(NFe)
        CHV_NFE = fnXML.ExtrairChaveAcessoNFe(NFe)
        COD_SIT = fnXML.ExtrairSituacaoDocumento(NFe)
        COD_MOD = fnXML.ValidarTag(NFe, "//mod")
        
        Call VerificarCancelamento(CHV_NFE, COD_SIT)
        
        .AtribuirValor "CHV_NFE", CHV_NFE
        .AtribuirValor "COD_MOD", COD_MOD
        .AtribuirValor "NUM_DOC", CLng(fnXML.ValidarTag(NFe, "//nNF"))
        .AtribuirValor "SER", VBA.Format(fnXML.ValidarTag(NFe, "//serie"), "000")
        .AtribuirValor "IND_OPER", fnXML.ExtrairTipoOperacao(NFe)
        .AtribuirValor "IND_EMIT", tpEmit
        .AtribuirValor "COD_SIT", COD_SIT
        .AtribuirValor "DT_DOC", DT_DOC
        .AtribuirValor "DT_E_S", fnXML.ExtrairDataEntradaSaida(NFe)
                        
        If Not COD_SIT Like "*02*" Then
            
            If COD_MOD <> "65" Then
                
                NOME_RAZAO = fnXML.ExtrairNomeRazaoParticipante(NFe, COD_MOD, COD_SIT)
                
                .AtribuirValor "COD_PART", fnExcel.FormatarTexto(fnXML.ExtrairParticipante(NFe))
                .AtribuirValor "NOME_RAZAO", NOME_RAZAO
                .AtribuirValor "INSC_EST", fnXML.ExtrairInscricaoParticipante(NFe)
            
            End If
            
            .AtribuirValor "IND_PGTO", fnXML.ExtrairTipoPagamento(NFe, DT_DOC)
            .AtribuirValor "VL_DOC", fnXML.ValidarValores(NFe, "//vNF")
            .AtribuirValor "VL_DESC", fnXML.ValidarValores(NFe, "//ICMSTot/vDesc")
            .AtribuirValor "VL_ABAT_NT", fnXML.ValidarValores(NFe, "//ICMSTot/vICMSDeson")
            .AtribuirValor "VL_MERC", fnXML.ValidarValores(NFe, "//ICMSTot/vProd")
            .AtribuirValor "IND_FRT", fnXML.ExtrairTipoFrete(NFe, DT_DOC)
            .AtribuirValor "VL_FRT", fnXML.ValidarValores(NFe, "//ICMSTot/vFrete")
            .AtribuirValor "VL_SEG", fnXML.ValidarValores(NFe, "//ICMSTot/vSeg")
            .AtribuirValor "VL_OUT_DA", fnXML.ValidarValores(NFe, "//ICMSTot/vOutro")
            .AtribuirValor "VL_BC_ICMS", fnXML.ValidarValores(NFe, "//ICMSTot/vBC")
            .AtribuirValor "VL_ICMS", fnXML.ExtrairTotaisICMS(NFe)
            .AtribuirValor "VL_BC_ICMS_ST", fnXML.ValidarValores(NFe, "//ICMSTot/vBCST")
            .AtribuirValor "VL_ICMS_ST", fnXML.ExtrairTotaisICMSST(NFe)
            .AtribuirValor "VL_IPI", fnXML.ValidarValores(NFe, "//ICMSTot/vIPI")
            .AtribuirValor "VL_PIS", fnXML.ValidarValores(NFe, "//ICMSTot/vPIS")
            .AtribuirValor "VL_COFINS", fnXML.ValidarValores(NFe, "//ICMSTot/vCOFINS")
            .AtribuirValor "VL_PIS_ST", fnXML.ValidarValores(NFe, "//ICMSTot/vPISST")
            .AtribuirValor "VL_COFINS_ST", fnXML.ValidarValores(NFe, "//ICMSTot/vCOFINSST")
            
        Else
            
            .AtribuirValor "VL_DOC", 0
            .AtribuirValor "VL_DESC", 0
            .AtribuirValor "VL_ABAT_NT", 0
            .AtribuirValor "VL_MERC", 0
            .AtribuirValor "IND_FRT", 0
            .AtribuirValor "VL_FRT", 0
            .AtribuirValor "VL_SEG", 0
            .AtribuirValor "VL_OUT_DA", 0
            .AtribuirValor "VL_BC_ICMS", 0
            .AtribuirValor "VL_ICMS", 0
            .AtribuirValor "VL_BC_ICMS_ST", 0
            .AtribuirValor "VL_ICMS_ST", 0
            .AtribuirValor "VL_IPI", 0
            .AtribuirValor "VL_PIS", 0
            .AtribuirValor "VL_COFINS", 0
            .AtribuirValor "VL_PIS_ST", 0
            .AtribuirValor "VL_COFINS_ST", 0
            
        End If
        
        If .VerificarExistenciaSPED Then Call .RegistrarNotaXML
        
    End With
    
End Function

Private Function CorrelacionarNotas()

Dim chNFe As Variant
    
    XMLsAusentes = 0
    With Divergencias
         
        .TipoRelatorio = "Notas"
        For Each chNFe In .dicOperacoesSPED.Keys()
            
            If .dicOperacoesXML.Exists(chNFe) Then
                
                Call RegistrarCorrelacao(chNFe)
                
            Else
                
                Call RegistrarNotasSemCorrelacao(chNFe)
                
            End If
            
        Next chNFe
        
    End With
    
End Function

Private Function RegistrarCorrelacao(ByVal chNFe As String)

Dim CamposSPED As Variant, CamposXML, Titulo, TituloSPED, TituloXML
Dim Posicao As Byte
    
    With Divergencias
                
        Call .RedimensionarCampoRelatorioNotas
        
        CamposSPED = Empty
        CamposSPED = .dicOperacoesSPED(chNFe)
        
        CamposXML = Empty
        CamposXML = .dicOperacoesXML(chNFe)
        
        For Each Titulo In .arrTitulosRelatorio
            
            If Not TitulosIgnorar Like "*" & Titulo & "*" Then TituloSPED = Titulo & "_SPED" Else TituloSPED = Titulo
            
            .TipoRelatorio = "SPED"
            Posicao = .RetornarPosicaoTitulo(Titulo)
            
            .TipoRelatorio = "Notas"
            .AtribuirCorrelacao TituloSPED, CamposSPED(Posicao)
            
            If Not IgnorarCampoNota(Titulo) Then
                
                TituloXML = Titulo & "_NF"
                
                .TipoRelatorio = "XML"
                Posicao = .RetornarPosicaoTitulo(Titulo)
                
                .TipoRelatorio = "Notas"
                .AtribuirCorrelacao TituloXML, CamposXML(Posicao)
                
            End If
            
        Next Titulo
        
        arrRelatorio.Add .CampoRelatorio
        
    End With
    
End Function

Private Function RegistrarNotasSemCorrelacao(ByVal chNFe As String)

Dim Titulo As Variant, TituloSPED, CamposSPED
Dim Posicao As Byte
    
    With Divergencias
        
        CamposSPED = .dicOperacoesSPED(chNFe)
        Call .RedimensionarCampoRelatorioNotas
        
        For Each Titulo In .arrTitulosRelatorio
                        
            If Not TitulosIgnorar Like "*" & Titulo & "*" Then TituloSPED = Titulo & "_SPED" Else TituloSPED = Titulo
            
            .TipoRelatorio = "SPED"
            Posicao = .RetornarPosicaoTitulo(Titulo)
            
            .TipoRelatorio = "Notas"
            .AtribuirCorrelacao TituloSPED, CamposSPED(Posicao)
            
        Next Titulo
        
        Call InformarXMLNaoIdentificado
        arrRelatorio.Add .CampoRelatorio
        
    End With
    
End Function

Private Sub InformarXMLNaoIdentificado()
    
    XMLsAusentes = XMLsAusentes + 1
    
    With Divergencias
        
        .AtribuirCorrelacao "INCONSISTENCIA", "O XML dessa operação não foi importado"
        .AtribuirCorrelacao "SUGESTAO", "Inclua o XML dessa operação na pasta e gere o relatório novamente"
        
    End With
    
End Sub

Private Function IgnorarCampoNota(ByVal Titulo As String) As Boolean
    
    If TitulosIgnorar Like "*" & Titulo & "*" Then IgnorarCampoNota = True
    
End Function

Public Function ReprocessarSugestoes()

Dim Dados As Range, Linha As Range
Dim arrRelatorio As New ArrayList
Dim dicTitulos As New Dictionary
Dim Campos As Variant
    
    If relDivergenciasNotas.AutoFilterMode Then relDivergenciasNotas.AutoFilter.ShowAllData
    
    Set dicTitulosRelatorio = Util.MapearTitulos(relDivergenciasNotas, 3)
    Set Dados = Util.DefinirIntervalo(relDivergenciasNotas, 4, 3)
    If Dados Is Nothing Then Exit Function
    
    a = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 10, "Reprocessando sugestões, por favor aguarde...", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            If Campos(dicTitulosRelatorio("INCONSISTENCIA")) <> "O XML dessa operação não foi importado" Then
                
                Campos(dicTitulosRelatorio("INCONSISTENCIA")) = Empty
                Campos(dicTitulosRelatorio("SUGESTAO")) = Empty
                
            End If
            
            arrRelatorio.Add Campos
            
        End If
        
    Next Linha
    
    Call ValidacoesNotas.IdentificarDivergenciasNotas(arrRelatorio)
    
    Call Util.LimparDados(relDivergenciasNotas, 4, False)
    Call Util.ExportarDadosArrayList(relDivergenciasNotas, arrRelatorio)
    Call FuncoesFormatacao.DestacarInconsistencias(relDivergenciasNotas)
    
    Call Util.AtualizarBarraStatus("Processamento Concluído!")
    
End Function

Public Function AceitarSugestoesNotas()

Dim Dados As Range, Linha As Range
Dim i As Long
    
    Set Dados = relDivergenciasNotas.Range("A4").CurrentRegion
    Set arrRelatorio = New ArrayList
    
    If Dados Is Nothing Then
        Call Util.MsgAlerta("Não foi encontrada nenhuma inconsistência!", "Inconsistências Fiscais")
        Exit Function
    End If
    
    For Each Linha In Dados.Rows
        
        With CamposNota
            
            CamposRelatorio = Application.index(Linha.Value2, 0, 0)
            
            Call DadosDivergenciasNotas.ResetarCamposNota
            If Util.ChecarCamposPreenchidos(CamposRelatorio) Then
                
                If Linha.Row > 3 Then Call DadosDivergenciasNotas.CarregarDadosRegistroDivergenciaNotas(CamposRelatorio)
                
                If Linha.EntireRow.Hidden = False And .SUGESTAO <> "" And Linha.Row > 3 Then
                    
                    Call AplicarSugestaoNota
                    
                    If .INCONSISTENCIA <> "O XML dessa operação não foi importado" Then
                        
                        .INCONSISTENCIA = Empty
                        .SUGESTAO = Empty
                        
                    End If
                    
                    CamposRelatorio = DadosDivergenciasNotas.AtribuirCamposNota(CamposRelatorio)
                    
                End If
                
                If Linha.Row > 3 Then arrRelatorio.Add CamposRelatorio
                
            End If
            
        End With
        
    Next Linha
    
    Call Util.AtualizarBarraStatus("Reanalisando inconsistências, por favor aguarde...")
    Call ValidacoesNotas.IdentificarDivergenciasNotas(arrRelatorio)
    
    If relDivergenciasNotas.AutoFilter.FilterMode Then relDivergenciasNotas.ShowAllData
    Call Util.LimparDados(relDivergenciasNotas, 4, False)
    
    Call Util.ExportarDadosArrayList(relDivergenciasNotas, arrRelatorio)
    Call FuncoesFormatacao.DestacarInconsistencias(relDivergenciasNotas)
    
    Application.StatusBar = False
    
End Function

Private Function AplicarSugestaoNota()
    
    With CamposNota
        
        Select Case .SUGESTAO
            
            Case "Informar o mesmo modelo do XML para o SPED"
                .COD_MOD_SPED = .COD_MOD_NF
                
            Case "Informar o mesmo código de situação do XML para o SPED"
                .COD_SIT_SPED = .COD_SIT_NF
                
            Case "Informar o mesmo número de documento do XML para o SPED"
                .NUM_DOC_SPED = .NUM_DOC_NF
                
            Case "Informar a mesma série do XML para o SPED"
                .SER_SPED = .SER_NF
                
            Case "Informar o mesmo tipo de operação do XML para o SPED"
                .IND_OPER_SPED = .IND_OPER_NF
                
            Case "Informar o mesmo tipo emissão do XML para o SPED"
                .IND_EMIT_SPED = .IND_EMIT_NF
                
            Case "Informar o mesmo participante do XML para o SPED"
                .COD_PART_SPED = .COD_PART_NF
                
            Case "Informar a mesma razão do participante do XML para o SPED"
                .NOME_RAZAO_SPED = .NOME_RAZAO_NF
                
            Case "Informar a mesma inscrição estadual do XML para o SPED", "Informar a inscrição estadual do XML para o SPED"
                .INSC_EST_SPED = .INSC_EST_NF
                
            Case "Informar a mesma data de emissão do XML para o SPED"
                .DT_DOC_SPED = .DT_DOC_NF
                
            Case "Informar a mesma data de entrada/saída do XML para o SPED"
                .DT_E_S_SPED = .DT_E_S_NF
                
            Case "Informar o mesmo tipo de pagamento do XML para o SPED"
                .IND_PGTO_SPED = .IND_PGTO_NF
                
            Case "Informar o mesmo tipo de frete do XML para o SPED"
                .IND_FRT_SPED = .IND_FRT_NF
                
            Case "Informar o mesmo valor de documento do XML para o SPED"
                .VL_DOC_SPED = .VL_DOC_NF
                
            Case "Informar o mesmo valor de desconto do XML para o SPED"
                .VL_DESC_SPED = .VL_DESC_NF
                
            Case "Informar o mesmo valor de abatimento do XML para o SPED"
                .VL_ABAT_NT_SPED = .VL_ABAT_NT_NF
                
            Case "Informar o mesmo valor de mercadorias do XML para o SPED"
                .VL_MERC_SPED = .VL_MERC_NF
                
            Case "Informar o mesmo valor de frete do XML para o SPED"
                .VL_FRT_SPED = .VL_FRT_NF
                
            Case "Informar o mesmo valor de seguro do XML para o SPED"
                .VL_SEG_SPED = .VL_SEG_NF
                
            Case "Informar o mesmo valor de outras despesas do XML para o SPED"
                .VL_OUT_DA_SPED = .VL_OUT_DA_NF
                
            Case "Zerar valor do campo VL_BC_ICMS_SPED"
                .VL_BC_ICMS_SPED = 0
                
            Case "Informar o mesmo valor de base do ICMS do XML para o SPED"
                .VL_BC_ICMS_SPED = .VL_BC_ICMS_NF
                
            Case "Zerar valor do campo VL_ICMS_SPED"
                .VL_ICMS_SPED = 0
                
            Case "Informar o mesmo valor de ICMS do XML para o SPED"
                .VL_ICMS_SPED = .VL_ICMS_NF
                
            Case "Informar o mesmo valor de base do ICMS-ST do XML para o SPED"
                .VL_BC_ICMS_ST_SPED = .VL_BC_ICMS_ST_NF
                
            Case "Informar o mesmo valor de ICMS-ST do XML para o SPED"
                .VL_ICMS_ST_SPED = .VL_ICMS_ST_NF
                
            Case "Informar o mesmo valor de IPI do XML para o SPED"
                .VL_IPI_SPED = .VL_IPI_NF
                
            Case "Informar o mesmo valor do PIS do XML para o SPED"
                .VL_PIS_SPED = .VL_PIS_NF
                
            Case "Informar o mesmo valor da COFINS do XML para o SPED"
                .VL_COFINS_SPED = .VL_COFINS_NF
                
            Case "Informar o mesmo valor do PIS-ST do XML para o SPED"
                .VL_PIS_ST_SPED = .VL_PIS_ST_NF
                
            Case "Informar o mesmo valor da COFINS-ST do XML para o SPED"
                .VL_COFINS_ST_SPED = .VL_COFINS_ST_NF
                
        End Select
        
    End With
    
End Function

'Private Function AtribuirParticipante()
'
'Dim CHV_REG As String
'
'    If dicDados0150.Count = 0 Then Set dicDados0150 = Util.CriarDicionarioRegistro(reg0150, "ARQUIVO", "CNPJ", "CPF")
'    If dicTitulos0150.Count = 0 Then Set dicTitulos0150 = Util.MapearTitulos(reg0150, 3)
'
'    With CamposNota
'
'        CHV_REG = VBA.Join(Array(.ARQUIVO, .COD_PART_SPED, ""))
'
'        If dicDados0150.Exists(CHV_REG) Then
'
'            .COD_PART_SPED = .COD_PART_NF
'
'        End If
'
'    End With
'
'End Function

Public Function IgnorarInconsistencias()

Dim Dados As Range, Linha As Range
Dim CHV_REG As String, INCONSISTENCIA$
Dim dicTitulos As New Dictionary
Dim dicDados As New Dictionary
Dim Resposta As VbMsgBoxResult
Dim Campos As Variant
    
    Resposta = MsgBox("Tem certeza que deseja ignorar as inconsistências selecionadas?" & vbCrLf & _
                      "Essa operação NÃO pode ser desfeita.", vbExclamation + vbYesNo, "Ignorar Inconsistências")
    
    If Resposta = vbNo Then Exit Function
    
    Inicio = Now()
    Application.StatusBar = "Ignorando as sugestões selecionadas, por favor aguarde..."
    
    Set dicTitulos = Util.MapearTitulos(relDivergenciasNotas, 3)
    Set Dados = relDivergenciasNotas.Range("A4").CurrentRegion
    If Dados Is Nothing Then
        Call Util.MsgAlerta("Não foi encontrada nenhuma inconsistência!", "Inconsistências Fiscais")
        Exit Function
    End If
    
    For Each Linha In Dados.Rows
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            If Linha.EntireRow.Hidden = False And Campos(dicTitulos("INCONSISTENCIA")) <> "" And Linha.Row > 3 Then
                
                CHV_REG = Campos(dicTitulos("CHV_REG"))
                INCONSISTENCIA = Campos(dicTitulos("INCONSISTENCIA"))
                
                'Verifica se o registro já possui inconsistências ignoradas, caso não exista cria
                If Not dicInconsistenciasIgnoradas.Exists(CHV_REG) Then Set dicInconsistenciasIgnoradas(CHV_REG) = New ArrayList
                
                'Verifica se a inconsistência já foi ignorada e caso contrário adiciona ela na lista
                If Not dicInconsistenciasIgnoradas(CHV_REG).contains(INCONSISTENCIA) Then _
                    dicInconsistenciasIgnoradas(CHV_REG).Add INCONSISTENCIA
                    
                Campos(dicTitulos("INCONSISTENCIA")) = Empty
                Campos(dicTitulos("SUGESTAO")) = Empty
                
            End If
            
            If Linha.Row > 3 Then arrRelatorio.Add Campos
            
        End If
        
    Next Linha
    
    If dicInconsistenciasIgnoradas.Count = 0 Then
        Call Util.MsgAlerta("Não existem Inconsistêncais a ignorar!", "Ignorar Inconsistências")
        Exit Function
    End If
    
    Call ReprocessarSugestoes
    
    Call Util.MsgInformativa("Inconsistências ignoradas com sucesso!", "Ignorar Inconsistências", Inicio)
    Application.StatusBar = False
    
End Function

Public Sub AtualizarRegistros()

Const CamposIgnorar As String = "NOME_RAZAO, INSC_EST"
Dim CHV_REG As String, TituloC100$
Dim dicTitulos As New Dictionary
Dim arrDados As New ArrayList
Dim dicCampos As Variant, Titulo, Campo
    
    If Util.ChecarAusenciaDados(relDivergenciasNotas) Then Exit Sub
    
    Inicio = Now()
    
    Call Util.DesabilitarControles
    
    Application.StatusBar = "Atualizando dados no SPED, por favor aguarde..."
    If relDivergenciasNotas.AutoFilterMode Then relDivergenciasNotas.AutoFilter.ShowAllData
    
    Set arrDados = Util.CriarArrayListRegistro(relDivergenciasNotas)
    
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    Set GerenciadorSPED = New clsRegistrosSPED
    
    With GerenciadorSPED
        
        Call .CarregarDadosRegistro0000("ARQUIVO")
        Call .CarregarDadosRegistro0000_Contr("ARQUIVO")
        Call .CarregarDadosRegistro0001("ARQUIVO")
        Call .CarregarDadosRegistro0140("CNPJ")
        Call .CarregarDadosRegistroC010
        Call .CarregarDadosRegistroC100
        
        If dtoRegSPED.r0000.Count = 0 Then Call .CarregarDadosRegistro0150("CHV_PAI_CONTRIBUICOES", "CNPJ", "CPF") _
            Else Call .CarregarDadosRegistro0150("CHV_PAI_FISCAL", "CNPJ", "CPF")
            
    End With
    
    Set dicTitulosRelatorio = Util.MapearTitulos(relDivergenciasNotas, 3)
    
    With dtoRegSPED
        
        a = 0
        Comeco = Timer
        For Each Campos In arrDados
            
            Call Util.AntiTravamento(a, 100, "atualizando dados do SPED, por favor aguarde...", arrDados.Count, Comeco)
            
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                CHV_REG = Campos(dicTitulosRelatorio("CHV_REG"))
                If .rC100.Exists(CHV_REG) Then
                    
                    dicCampos = .rC100(CHV_REG)
                    For Each Titulo In dicTitulosRelatorio.Keys()
                        
                        If Titulo Like "*_SPED" Then
                            
                            TituloC100 = VBA.Replace(Titulo, "_SPED", "")
                            Campo = FormatarCamposC100(Campos, Titulo)
                            If Not CamposIgnorar Like "*" & TituloC100 & "*" Then _
                                dicCampos(dtoTitSPED.tC100(TituloC100)) = Campo
                            
                        End If
                        
                    Next Titulo
                    
                    .rC100(CHV_REG) = dicCampos
                    
                End If
                
            End If
            
        Next Campos
    
    End With
    
    Application.StatusBar = "Atualizando dados do registro C100, por favor aguarde..."
    Call Util.LimparDados(regC100, 4, False)
    Call Util.ExportarDadosDicionario(regC100, dtoRegSPED.rC100)
    
    Call Util.LimparDados(reg0150, 4, False)
    Call Util.ExportarDadosDicionario(reg0150, dtoRegSPED.r0150)
    
    Call FuncoesFormatacao.AplicarFormatacao(relDivergenciasNotas)
    Call FuncoesFormatacao.DestacarInconsistencias(relDivergenciasNotas)
    
    Application.StatusBar = "Atualização concluída com sucesso!"
    Call Util.MsgInformativa("Registros atualizados com sucesso!", "Atualização de dados", Inicio)
    
    Call Util.HabilitarControles
    
    Call Util.AtualizarBarraStatus(False)
    
End Sub

Private Sub VerificarCancelamento(ByRef CHV_NFE As String, ByRef COD_SIT As String)
    
    CHV_NFE = Util.ApenasNumeros(CHV_NFE)
    If DocsFiscais.arrChavesCanceladas.contains(CHV_NFE) Then COD_SIT = EnumFiscal.ValidarEnumeracao_COD_SIT("02")
    
End Sub
