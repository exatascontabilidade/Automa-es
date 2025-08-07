Attribute VB_Name = "Acionadores"
Option Explicit

Private ImportNFeNFCe As New AssistenteImportacaoNFeNFCe

Public Sub AcionarBotao(control As IRibbonControl, Optional ByRef Valor)

Dim vbResult As VbMsgBoxResult
Dim Msg As String, Plano$
Dim NivelPlano As Byte
Dim Status As String
    
    Call Util.DesabilitarControles
    If FuncoesControlDocs.ObterUuidComputador = "8A1DD300-CCB5-11EC-B9D4-478D28DC6B00" Then Debug.Print control.id
    
    If Not VerificacoesControlDocs.VerificarConfiguracoesControlDocs Then GoTo Sair:
    
    If ControlPressionado Then
        Call DocumentacaoControlDocs.AcessarDocumentacao(control)
        Exit Sub
    End If
    
    'Registra acionamento de botão do usuário
    Call RegistrarAcionamento(control.id)
    
    'Verifica quantiade de acionamentos para envio dos dados
    If ContarAcionamentos > 100 Then FuncoesControlDocs.EnviarAcionamentos ("RELATORIO_ACIONAMENTOS")
    
    If control.id Like "btnExperimentarControlDocs" Then
        
        Call FuncoesControlDocs.ConsultarStatusAssinatura("ASSINATURA_EXPERIMENTAL", True)
        GoTo Sair:
        
    End If
    
    If control.id Like "btnIndividual*" Or control.id Like "btnEmpresarial*" Then
        
        Call Rotinas.AssinaturaControlDocs(control)
        GoTo Sair:
        
    ElseIf VBA.Left(control.id, 6) = "btnReg" Then
        
        Call FaixaOpcoes.IrPara(control)
        GoTo Sair:
        
    End If
    
    Select Case control.id
        
        Case "btnLimparDadosAssinatura"
            With relGestaoAssinatura
                
                .Range("E2").value = ""
                .Range("G2").value = ""
                .Range("I2").value = ""
                .Range("K2").value = ""
                
            End With
            Call Util.LimparDados(relGestaoAssinatura, 4, False)
            
        Case "btnRecursosControlDocs"
            Call FaixaOpcoes.MostrarGrupos(control, ControlDocs)
            GoTo Sair:
            
        Case "btnAssinaturaControlDocs"
            Call FaixaOpcoes.MostrarGrupos(control, relGestaoAssinatura)
            GoTo Sair:
            
        Case "btnConfiguracoesControlDocs"
            Call FaixaOpcoes.MostrarGrupos(control, ControlDocs)
            GoTo Sair:
            
        Case "btnControlDocs"
            Call FuncoesLinks.AbrirUrl(urlTutoriais)
            GoTo Sair:
            
        Case "btnDonwloadControlDocs"
            Call FuncoesLinks.AbrirUrl(DownloadControlDocs)
            GoTo Sair:
            
        Case "btnSuporte"
            Call FuncoesLinks.AbrirUrl(urlSuporte)
            GoTo Sair:
            
        Case "btnSugestoes"
            Call FuncoesLinks.AbrirUrl(urlSugestoes)
            GoTo Sair:
        
        Case "btnDocControlDocs"
            Call FuncoesLinks.AbrirUrl(urlDocumentacao)
            GoTo Sair:
            
        Case "btnAutenticarUsuario"
            If EmailAssinante <> "" And Not Funcoes.ValidarEmail(EmailAssinante) Then
                Call FuncoesControlDocs.ResetarAssinatura
                MsgBox "Informe um e-mail válido para autenticar sua assinatura.", vbExclamation, "Email não informado ou inválido"
                relGestaoAssinatura.Range("email_cliente").Activate
                GoTo Sair:
            End If
            Call FuncoesControlDocs.ConsultarStatusAssinatura("REGISTRAR_MAQUINA", True)
            GoTo Sair:
            
        Case "btnAutenticar"
            Call FaixaOpcoes.IrPara(control)
            GoTo Sair:
            
        Case "btnResetarControlDocs"
            If fnSeguranca.VerificarDadosTributarios("resetar") Then Call FuncoesControlDocs.ResetarControlDocs
            GoTo Sair:
            
        Case "btnCadContrib", "btnEntNFe", "btnEntCTe", "btnSaiNFe", "btnSaiNFCe", "btnSaiCTe", "btnSaiCFe", "btnDocsAusentes", "btnQuebraSequencia"
            Call FaixaOpcoes.IrPara(control)
            GoTo Sair:
            
        Case "btnImportarExcel"
            Call FuncoesControlDocs.CarregarDadosPlanilha
            GoTo Sair:
        
        Case "btnImportSefazBA"
            If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
            Call Rotinas.ImportarDadosSEFAZBA
            GoTo Sair:
        
        Case "btnImportDoce"
            If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
            Call FuncoesXML.ImportarDocumentosEletronicos
            GoTo Sair:
            
        Case "btnImportProtCancel"
            Call FuncoesXML.ImportarProtocolosCancelamento
            GoTo Sair:
            
        Case "btnImportSPED"
            Call Rotinas.CruzarDadosComSPED
            GoTo Sair:
            
        Case "btnDocsSemLancar"
            Call FuncoesFiltragem.ListarNotasSemLancar
            GoTo Sair:
            
        Case "btnExtCadWeb"
            Call WebScraping.ExtrairCadastroContribuinteWeb
            GoTo Sair:
            
        Case "btnExtCadSPED"
            Call fnSPED.ExtrairCadastroSPEDFiscal
            GoTo Sair:
            
        Case "btnConsultarAssinatura"
            Call FuncoesControlDocs.ListarDispositivosControlDocs
            GoTo Sair:
            
    End Select
    
    Call VerificarAssinatura(NivelPlano, Status, Plano)
    
    If NivelPlano >= 1 Then
        
        Select Case control.id
            
            Case "btnDefinirPeriodo"
                UsarPeriodo = Valor
                Call FaixaOpcoes.getPressed(control, Valor)
                GoTo Sair:
                
            Case "btnICMS", "btnDivergencias", "btnCorrelacoes", "btnIPI", "btnPISCOFINS", "btnTributacao"
                Call FaixaOpcoes.IrPara(control)
                GoTo Sair:
                
            Case "btnGerarLivroICMS"
                Call FuncoesLivrosFiscais.GerarLivroICMS
                GoTo Sair:
                
            Case "btnFiltrarEntradas"
                Call FuncoesFiltragem.FiltrarEntradas
                GoTo Sair:
                
            Case "btnFiltrarSaidas"
                Call FuncoesFiltragem.FiltrarSaidas
                GoTo Sair:
                
            Case "btnAcessarEnfoqueDeclarante"
                Call FuncoesFiltragem.AcessarEnfoqueDeclarante
                GoTo Sair:
                
            Case "btnListarIncosnsistencias"
                Call FuncoesFiltragem.ListarDivergencias
                GoTo Sair:
                
            Case "btnImportarSPEDFiscal"
                Call FuncoesSPEDFiscal.ImportarSPED
                GoTo Sair:
                
            Case "btnCentralizarSPEDFiscal"
                Call FuncoesSPEDFiscal.ImportarSPED(Centralizar:=True)
                GoTo Sair:
                
            Case "btnImportarCadastroContador"
                Call r0100.ImportarDadosContador
                GoTo Sair:
                
            Case "btnImportRegAtual"
                Call FuncoesSPEDFiscal.ImportarSPED(VBA.Left(ActiveSheet.name, 4), PeriodoImportacao)
                GoTo Sair:
                
            Case "btnLimparRegistrosSPED"
                Call fnSPED.LimparRegistrosEFD
                GoTo Sair:
            
            'Filtros 0150
            Case "btnListarNotasC100"
                Call Util.FiltrarRegistros(reg0150, regC100, "COD_PART", "COD_PART")
                GoTo Sair:
                
            'Filtros 0200
            Case "btnListarFat0220"
                Call Util.FiltrarRegistros(reg0200, reg0220, "CHV_REG", "CHV_PAI_FISCAL")
                GoTo Sair:
                
            'Filtros 0220
            Case "btnListarItem0200"
                Call Util.FiltrarRegistros(reg0220, reg0200, "CHV_PAI_FISCAL", "CHV_REG")
                GoTo Sair:
                
'Controles A100
            Case "btnImportarPlanilhaA100Filhos"
                rA100.ImportarPlanilhaA100_A170
                GoTo Sair:
                
            Case "btnGerarModeloA100A170"
                Call FuncoesExcel.GerarModeloA100Filhos
                GoTo Sair:
                
            Case "btnListarItensA170"
                Call Util.FiltrarRegistros(regA100, regA170, "CHV_REG", "CHV_PAI_CONTRIBUICOES")
                GoTo Sair:
                
            'Filtros A170
            Case "btnListarNotasA100"
                Call Util.FiltrarRegistros(regA170, regA100, "CHV_PAI_CONTRIBUICOES", "CHV_REG")
                GoTo Sair:
                
            'Filtros C100
            Case "btnListarItensC170"
                 Call Util.FiltrarRegistros(regC100, regC170, "CHV_REG", "CHV_PAI_FISCAL")
                GoTo Sair:
                
            Case "btnListarResumosC190"
                 Call Util.FiltrarRegistros(regC100, regC190, "CHV_REG", "CHV_PAI_FISCAL")
                GoTo Sair:
                
            'Filtros C170
            Case "btnListarNotasC170"
                Call Util.FiltrarRegistros(regC170, regC100, "CHV_PAI_FISCAL", "CHV_REG")
                GoTo Sair:
                
            Case "btnListarResumosC190C170"
                Call Util.FiltrarRegistros(regC170, regC190, "CHV_PAI_FISCAL", "CHV_PAI_FISCAL")
                GoTo Sair:
                
            'Filtros C190
            Case "btnListarItensC170C190"
                 Call Util.FiltrarRegistros(regC190, regC170, "CHV_PAI_FISCAL", "CHV_PAI_FISCAL")
                GoTo Sair:
                
            Case "btnListarNotasC190"
                Call Util.FiltrarRegistros(regC190, regC100, "CHV_PAI_FISCAL", "CHV_REG")
                GoTo Sair:
                
            'Filtros C191
            Case "btnListarResumosC191C190"
                 Call Util.FiltrarRegistros(regC191, regC190, "CHV_PAI_FISCAL", "CHV_REG")
                GoTo Sair:
                
            Case "btnListarNotasDivergencias"
                Call Util.FiltrarRegistros(regC190, relDivergenciasNotas, "CHV_PAI_FISCAL", "CHV_REG")
                On Error Resume Next
                Rib.ActivateTab "tbAssistentesFiscais"
                GoTo Sair:
                
            'Filtros C500
            Case "btnListarResumosC500C510"
                Call Util.FiltrarRegistros(regC500, regC510, "CHV_REG", "CHV_PAI_FISCAL")
                GoTo Sair:
                
            Case "btnListarResumosC500C590"
                Call Util.FiltrarRegistros(regC500, regC590, "CHV_REG", "CHV_PAI_FISCAL")
                GoTo Sair:
                
            'Filtros C510
            Case "btnListarNotasC510C500"
                Call Util.FiltrarRegistros(regC510, regC500, "CHV_PAI_FISCAL", "CHV_REG")
                GoTo Sair:
                
            Case "btnListarResumosC510C590"
                Call Util.FiltrarRegistros(regC510, regC590, "CHV_PAI_FISCAL", "CHV_PAI_FISCAL")
                GoTo Sair:
                
            'Filtros C590
            Case "btnListarNotasC590C500"
                Call Util.FiltrarRegistros(regC590, regC500, "CHV_PAI_FISCAL", "CHV_REG")
                GoTo Sair:
                
            Case "btnListarResumosC590C510"
                Call Util.FiltrarRegistros(regC590, regC510, "CHV_PAI_FISCAL", "CHV_PAI_FISCAL")
                GoTo Sair:
                
            'Filtros C800
            Case "btnListarItensC810"
                 Call Util.FiltrarRegistros(regC800, regC810, "CHV_REG", "CHV_PAI_FISCAL")
                GoTo Sair:
                
            Case "btnListarResumosC850"
                Call Util.FiltrarRegistros(regC800, regC850, "CHV_REG", "CHV_PAI_FISCAL")
                GoTo Sair:
            
            'Filtros C810
            Case "btnListarResumosC850C810"
                Call Util.FiltrarRegistros(regC810, regC850, "CHV_PAI_FISCAL", "CHV_PAI_FISCAL")
                GoTo Sair:
                
            Case "btnListarNotasC810"
                Call Util.FiltrarRegistros(regC810, regC800, "CHV_PAI_FISCAL", "CHV_REG")
                GoTo Sair:
                
            'Filtros C850
            Case "btnListarItensC810C850"
                Call Util.FiltrarRegistros(regC850, regC810, "CHV_PAI_FISCAL", "CHV_PAI_FISCAL")
                GoTo Sair:
                
            Case "btnListarNotasC850"
                Call Util.FiltrarRegistros(regC850, regC800, "CHV_PAI_FISCAL", "CHV_REG")
                GoTo Sair:
            
            'Filtros D100
            Case "btnListarResumosD190"
                Call Util.FiltrarRegistros(regD100, regD190, "CHV_REG", "CHV_PAI_FISCAL")
                GoTo Sair:
                
            Case "btnListarRegD100D101Contrib"
                Call Util.FiltrarRegistros(regD100, regD101_Contr, "CHV_REG", "CHV_PAI_FISCAL")
                GoTo Sair:
                
            Case "btnListarRegD100D105"
                Call Util.FiltrarRegistros(regD100, regD105, "CHV_REG", "CHV_PAI_FISCAL")
                GoTo Sair:
                
            Case "btnListarRegD100D190"
                Call Util.FiltrarRegistros(regD100, regD190, "CHV_REG", "CHV_PAI_FISCAL")
                GoTo Sair:
            
            'Filtros D101 Contribuições
            Case "btnListarRegD101ContrD100"
                Call Util.FiltrarRegistros(regD101_Contr, regD100, "CHV_PAI_FISCAL", "CHV_REG")
                GoTo Sair:
                
            Case "btnListarRegD101ContrD105"
                Call Util.FiltrarRegistros(regD101_Contr, regD105, "CHV_PAI_FISCAL", "CHV_PAI_FISCAL")
                GoTo Sair:
                
            Case "btnListarRegD101ContrD190"
                Call Util.FiltrarRegistros(regD101_Contr, regD190, "CHV_PAI_FISCAL", "CHV_PAI_FISCAL")
                GoTo Sair:
            
            'Filtros D105
            Case "btnListarRegD105D100"
                Call Util.FiltrarRegistros(regD105, regD100, "CHV_PAI_FISCAL", "CHV_REG")
                GoTo Sair:
                
            Case "btnListarRegD105D101"
                Call Util.FiltrarRegistros(regD105, regD101_Contr, "CHV_PAI_FISCAL", "CHV_PAI_FISCAL")
                GoTo Sair:
                
            Case "btnListarRegD105D190"
                Call Util.FiltrarRegistros(regD105, regD190, "CHV_PAI_FISCAL", "CHV_PAI_FISCAL")
                GoTo Sair:
                
            'Filtros D190
            Case "btnListarNotasD100"
                Call Util.FiltrarRegistros(regD190, regD100, "CHV_PAI_FISCAL", "CHV_REG")
                GoTo Sair:
                
            Case "btnListarRegD190D101Contrib"
                Call Util.FiltrarRegistros(regD190, regD101_Contr, "CHV_PAI_FISCAL", "CHV_PAI_FISCAL")
                GoTo Sair:
                
            Case "btnListarRegD190D105"
                Call Util.FiltrarRegistros(regD190, regD105, "CHV_PAI_FISCAL", "CHV_PAI_FISCAL")
                GoTo Sair:
                
            Case "btnGerarModelo"
                Call FuncoesExcel.GerarModelo0200
                GoTo Sair:
                
            Case "btnAjustarC190peloC100"
                Call rC100.RatearDivergenciasC100ParaC190
                GoTo Sair:
                
            Case "btnImportarXMLSAnalise"
                Call FuncoesXML.ImportarXMLSParaAnalise("Arquivo")
                GoTo Sair:
                
            Case "btnGerarC175peloC170"
                Call rC170.GerarC175
                GoTo Sair:
                
            Case "btnGerarC190peloC170"
                Call rC170.GerarC190
                GoTo Sair:
                
            Case "btnSomarST", "btnSomarIPIeST"
                If control.id = "btnSomarST" Then Call rC170.SomarIPIeSTaosItens("ST")
                If control.id = "btnSomarIPIeST" Then Call rC170.SomarIPIeSTaosItens("IPI-ST")
                GoTo Sair:
            
            Case "btnAtualizarC100"
                Call rC170.AtualizarImpostosC100
                GoTo Sair:
            
            Case "btnAtualizarC100C190"
                Call rC190.AtualizarImpostosC100
                GoTo Sair:
            
            Case "btnExportarSPEDFiscal"
                Call FuncoesSPEDFiscal.GerarEFDICMSIPI
                GoTo Sair:
            
            Case "btnAgruparRegistrosC170"
                rC170.AgruparRegistros
                GoTo Sair:
                
            Case "btnImportarICMSC170SPEDFiscal"
                Call rC170.ImportDadosICMS.ImportarDadosICMS_SPED_Fiscal("Arquivo")
                GoTo Sair:
                 
            Case "btnAgruparC190"
                rC190.AgruparRegistros
                GoTo Sair:
                
            Case "btnImportarLoteC100Fiscal"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call FuncoesXML.ImportarXMLsC100("Lote")
                GoTo Sair:
                
            Case "btnImportarArqC100Fiscal"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call FuncoesXML.ImportarXMLsC100("Arquivo")
                GoTo Sair:
                
            Case "btnImportarLoteC800Fiscal"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call FuncoesXML.ImportarRegistrosCFeXML("Lote")
                GoTo Sair:
                
            Case "btnImportarArqC800Fiscal"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call FuncoesXML.ImportarRegistrosCFeXML("Arquivo")
                GoTo Sair:
                
            Case "btnImportarLoteD100Fiscal"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call FuncoesXML.ImportarRegistrosCTeXML("Lote")
                GoTo Sair:
                
            Case "btnImportarArqD100Fiscal"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call FuncoesXML.ImportarRegistrosCTeXML("Arquivo")
                GoTo Sair:
                
            Case "btnAtualizarC800C850"
                Call rC850.AtualizarImpostosC800
                GoTo Sair:
                
            Case "btnAtualizarCodGenero"
                Call r0200.AtualizarCodigoGenero
                GoTo Sair:
                
            Case "btnAgruparC850"
                Call rC850.AgruparRegistros
                GoTo Sair:
                
            Case "btnCalcBasePISCOFINS"
                Call rC170.CalcularPISCOFINS
                GoTo Sair:
                
            Case "btnExcluirICMS"
                Call rC170.CalcularPISCOFINS(ExcluirICMS:=True)
                GoTo Sair:
                
            Case "btnExcluirICMSST"
                Call rC170.CalcularPISCOFINS(ExcluirICMS_ST:=True)
                GoTo Sair:
                
            Case "btnGerarCreditoSIMPLESNACIONAL"
                Call rE111.GerarCreditoSIMPLESNACIONAL
                GoTo Sair:
                
            Case "btnProdutosFornecedorArquivo"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call FuncoesXML.ImportarProdutosFornecedor("Arquivo")
                GoTo Sair:
                
            Case "btnProdutosFornecedorLote"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call FuncoesXML.ImportarProdutosFornecedor("Lote")
                GoTo Sair:
                
            Case "btnCorrelacionarSPEDXML"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call Correlacionamentos.GerarCorrelacoesSPEDXML
                GoTo Sair:
                
            Case "btnImportarEFDAnalise"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call FuncoesSPEDFiscal.ImportarSPEDFiscalparaAnalise
                Call FuncoesFormatacao.AplicarFormatacao(relDivergencias)
                GoTo Sair:
                
            Case "btnImportarXMLAnaliseLote"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call FuncoesXML.ImportarXMLSParaAnalise("Lote")
                GoTo Sair:
                
            Case "btnImportarXMLAnaliseArquivo"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call FuncoesXML.ImportarXMLSParaAnalise("Arquivo")
                GoTo Sair:
                
            Case "btnListarXMLSAusentes"
                Call Rotinas.ListarXMLsAusentes
                GoTo Sair:
                
            Case "btnExportarRelatorioCorrelacao"
                Call FuncoesAssistentesInteligentes.EnviarDados
                GoTo Sair:
                
            Case "btnCalcRedBCICMSC190"
                Call rC190.CalcularReducaoBaseICMS
                GoTo Sair:
                
            Case "btnAcessarNotaSelecionada"
                Call Util.FiltrarRegistros(relDivergencias, regC100, "CHV_NFE", "CHV_NFE")
                GoTo Sair:
                
            Case "btnImportarItensSPED"
                Call FuncoesTributacao.ImportarItensSPED
                GoTo Sair:
                
            Case "btnAnalisarTributacao"
                Call FuncoesTributacao.VerificarTributacao
                GoTo Sair:
                
            Case "btnImportarCadastroItens"
                Call FuncoesTributacao.ImportarCadastroTributacao
                GoTo Sair:
                
            Case "btnExportarCadastroItens"
                Call FuncoesTXT.ExportarParaTxt(Tributacao, "CODIGO", "CFOP")
                GoTo Sair:
                
            Case "btnImportarCadastroCorrelacoes"
                Call FuncoesExcel.ImportarCadastroCorrelacoes
                GoTo Sair:
                
            Case "btnListarDivergencias"
                Call FuncoesFiltragem.ListarDivergencias
                GoTo Sair:
                
            Case "btnEstruturarSPED"
                Call FuncoesSPEDFiscal.EstruturarSPED
                GoTo Sair:
                
            Case "btnConsultarCEP0005", "btnConsultarCEP0100", "btnConsultarCEP0150"
                Call FuncoesAPI.ConsultarCEP(ActiveSheet)
                GoTo Sair:
                
            Case "btnImportarProd0200"
                Call FuncoesExcel.ImportarCadastro0200
                GoTo Sair:
                
            Case "btnExportarProd0200"
                Call FuncoesExcel.ExportarCadastro0200
                GoTo Sair:
                
            Case "btnImportar0200Lote"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call FuncoesXML.ImportarCadastro0200XML("Lote")
                GoTo Sair:
                
            Case "btnImportar0200Arquivo"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call FuncoesXML.ImportarCadastro0200XML("Arquivo")
                GoTo Sair:
                
           Case "btnAgruparD190"
                rD190.AgruparRegistros
                GoTo Sair:
                
            Case "btnAtualizarD100D190"
                Call rD190.AtualizarImpostosD100
                GoTo Sair:
                
            Case "btnCalcRedBCICMSD190"
                Call rD190.CalcularReducaoBaseICMS
                GoTo Sair:
                
            Case "btnRemoverDuplicatas"
                If ActiveSheet.CodeName Like "reg*" Then Call FuncoesPlanilha.RemoverDuplicatas(ActiveSheet, "CHV_REG", "CHV_PAI_FISCAL")
                GoTo Sair:
                
            Case "btnGerarReceitaMunicipioD100"
                Call r1400.GerarSaidasCTeMunicipio
                GoTo Sair:
                
            Case "btnImportarK200"
                Call FuncoesExcel.ImportarCadastroK200
                GoTo Sair:
                
            Case "btnExportarProdK200"
                Call Util.ExportarDadosRelatorio(regK200, "Registro K200 - Saldo Estoque")
                GoTo Sair:
                
            Case "btnGerarModeloK200"
                Call Util.GerarModeloRelatorio(regK200, "Estoque")
                GoTo Sair:
                
            Case "btnIdentificarProdutosAusentesK200"
                Dim fnK200 As New clsK200
                Call fnK200.ListarProdutosAusentes
                GoTo Sair:
                
            Case "btnGerarQuebraSequenciaXML"
                Call fnExcel.GerarRelatorioQuebraSequenciaXML
                GoTo Sair:
                
            Case "btnGerarQuebraSequenciaSPED"
                Call fnExcel.GerarRelatorioQuebraSequenciaSPED
                GoTo Sair:
                
        End Select
        
    Else
        
        Msg = "O seu nível de assinatura não dá acesso a esta funcionalidade." & vbCrLf & vbCrLf
        Msg = Msg & "Com o plano Básico, você desbloqueia o acesso aos recursos para trabalhar com os arquivos do SPED Fiscal." & vbCrLf & vbCrLf
        Msg = Msg & "Clique em SIM para fazer o upgrade do seu plano agora mesmo!"
        
        vbResult = MsgBox(Msg, vbExclamation + vbYesNo, "Necessário Upgrade de Plano")
        If vbResult = vbYes Then Call FuncoesLinks.AbrirUrl(urlAssinaturaEmpresarialMensal)
        
        GoTo Sair:
        
    End If
    
    If NivelPlano >= 2 Then
        
        Select Case control.id
            
            Case "btnIdentReferenciadas"
                Call FuncoesXML.IdentificarNotasReferenciadas
                GoTo Sair:
                
            Case "btnImportarSPEDContribuicoes"
                Call FuncoesSPEDContribuicoes.ImportarSPEDContribuicoes
                GoTo Sair:
                
            Case "btnUnificarSPEDContribuicoes"
                Call FuncoesSPEDContribuicoes.ImportarSPEDContribuicoes(Unificar:=True)
                GoTo Sair:
                
            Case "btnExportarSPEDContribuicoes"
                Call FuncoesSPEDContribuicoes.GerarEFDContribuicoes
                GoTo Sair:
                
            Case "btnImportarLoteC100"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call ImportNFeNFCe.ImportarNFeNFCe("Lote")
                GoTo Sair:
                
            Case "btnImportarArquivoC100"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call ImportNFeNFCe.ImportarNFeNFCe("Arquivo")
                GoTo Sair:
                
            Case "btnImportarLoteA100Contribuicoes"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call FuncoesXML.ImportarXMLsA100("Lote")
                GoTo Sair:
                
            Case "btnImportarArqA100Contribuicoes"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call FuncoesXML.ImportarXMLsA100("Arquivo")
                GoTo Sair:
                
            Case "btnImportarLoteC100Contribuicoes"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call ImportNFeNFCe.ImportarNFeNFCe("Lote")
                GoTo Sair:
                
            Case "btnImportarArqC100Contribuicoes"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call ImportNFeNFCe.ImportarNFeNFCe("Arquivo")
                GoTo Sair:
                
            Case "btnAgruparRegistrosC180Contr"
                Call fnC180Contrib.AgruparRegistros
                GoTo Sair:
                
            Case "btnAtualizarNCMC180Contrib"
                Call fnC180Contrib.AtualizarNCM_C180
                GoTo Sair:
                
            Case "btnAgruparRegistrosC190Contr"
                Call fnC190Contrib.AgruparRegistros
                GoTo Sair:
                
            Case "btnAtualizarNCMC190Contrib"
                Call fnC190Contrib.AtualizarNCM_C190
                GoTo Sair:
                
            Case "btnImportarLoteC800Contribuicoes"
                'If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                'Call FuncoesXML.ImportarRegistrosCFeXML("Lote", True)
                GoTo Sair:
                
            Case "btnImportarArqC800Contribuicoes"
                'If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                'Call FuncoesXML.ImportarRegistrosCFeXML("Arquivo", True)
                GoTo Sair:
            
            Case "btnImportarLoteD100Contribuicoes"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call FuncoesXML.ImportarRegistrosCTeXML("Lote", True)
                GoTo Sair:
                        
            Case "btnImportarArqD100Contribuicoes"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call FuncoesXML.ImportarRegistrosCTeXML("Arquivo", True)
                GoTo Sair:
                                
            Case "btnListarResumosC175Contr"
                Call Util.FiltrarRegistros(regC100, regC175_Contr, "CHV_REG", "CHV_PAI_CONTRIBUICOES")
                GoTo Sair:
                
            Case "btnListarResumosC175ContrC170"
                 Call Util.FiltrarRegistros(regC170, regC175_Contr, "CHV_PAI_FISCAL", "CHV_PAI_CONTRIBUICOES")
                GoTo Sair:
                
            Case "btnListarNotasC175Contr"
                Call Util.FiltrarRegistros(regC175_Contr, regC100, "CHV_PAI_CONTRIBUICOES", "CHV_REG")
                GoTo Sair:
                
            Case "btnListarItensC170C175Contr"
                Call Util.FiltrarRegistros(regC175_Contr, regC170, "CHV_PAI_CONTRIBUICOES", "CHV_PAI_CONTRIBUICOES")
                GoTo Sair:
                
            Case "btnAgruparC175Contrib"
                Call rC175Contr.AgruparRegistros
                GoTo Sair:
            
            Case "btnExcluirBasePISCOFINS"
                Call rC175Contr.ExcluirICMSBasePIS_COFINS
                GoTo Sair:
                
            Case "btnAtualizarC100C175Contr"
                Call rC175Contr.AtualizarImpostosC100
                GoTo Sair:
            
            Case "btnGerarD200"
                Call rD100.GerarRegistroD200
                GoTo Sair:
            
            Case "btnGerarD101D105"
                Call rD100.GerarRegistrosD101_D105_PISCOFINS
                GoTo Sair:
                
            Case "btnAgruparC181Contrib"
                Call rC185.AgruparRegistros(True)
                Call rC181.AgruparRegistros
                GoTo Sair:
                
            Case "btnAgruparC185Contrib"
                Call rC181.AgruparRegistros(True)
                Call rC185.AgruparRegistros
                GoTo Sair:
                
            Case "btnListarNotasC181Contr"
                Call Util.FiltrarRegistros(regC181_Contr, regC180_Contr, "CHV_PAI_FISCAL", "CHV_REG")
                GoTo Sair:
                
            Case "btnListarNotasC185Contr"
                Call Util.FiltrarRegistros(regC185_Contr, regC180_Contr, "CHV_PAI_FISCAL", "CHV_REG")
                GoTo Sair:
                
            Case "btnListarItensC181C185Contr"
                Call Util.FiltrarRegistros(regC181_Contr, regC185_Contr, "CHV_PAI_FISCAL", "CHV_PAI_FISCAL")
                GoTo Sair:
                
            Case "btnListarItensC185C181Contr"
                Call Util.FiltrarRegistros(regC185_Contr, regC181_Contr, "CHV_PAI_FISCAL", "CHV_PAI_FISCAL")
                GoTo Sair:
                
        End Select
        
    Else
        
        Msg = "O seu nível de assinatura não dá acesso a esta funcionalidade." & vbCrLf & vbCrLf
        Msg = Msg & "Com o plano Plus, você desbloqueia o acesso aos recursos para trabalhar com os arquivos do SPED Contribuições." & vbCrLf & vbCrLf
        Msg = Msg & "Clique em SIM para fazer o upgrade do seu plano agora mesmo!"
        
        vbResult = MsgBox(Msg, vbExclamation + vbYesNo, "Necessário Upgrade de Plano")
        If vbResult = vbYes Then Call FuncoesLinks.AbrirUrl(urlAssinaturaEmpresarialMensal)
        GoTo Sair:
        
    End If
    
    If NivelPlano >= 3 Then
        
'##----->> Botões de acesso aos Relatórios Inteligentes
        If control.id Like "btnAssistenteFiscal*" Or control.id Like "btnAssistenteApuracao*" _
            Or control.id Like "btnTributacao*" Or control.id Like "btnDivergencias*" Then
            
            Call FaixaOpcoes.IrPara(control)
            GoTo Sair:
            
        ElseIf control.id Like "btnListarInconsistencias*" Then
            
            Call FuncoesFiltragem.FiltrarInconsistencias(ActiveSheet)
            GoTo Sair:
            
        ElseIf control.id Like "btnResetarInconsistencias*" Then
            
            Call Assistente.ResetarInconsistencias
            GoTo Sair:
            
        End If
        
        Select Case control.id
            
'##----->> Funcionalidades do Assistente de Apuração do ICMS
            Case "btnGerarRelatorioICMS"
                Call FuncoesAssistentesInteligentes.GerarRelatorioApuracaoICMS
                GoTo Sair:
                
            Case "btnProcessarInconsistenciasICMS"
                Call FuncoesAssistentesInteligentes.ReprocessarSugestoesApuracaoICMS
                GoTo Sair:
                
            Case "btnAceitarSugestoesICMS"
                Call FuncoesAssistentesInteligentes.AceitarSugestoesApuracaoICMS
                GoTo Sair:
                
            Case "btnIgnorarInconsistenciasICMS"
                Call FuncoesAssistentesInteligentes.IgnorarInconsistenciasApuracaoICMS
                GoTo Sair:
                
            Case "btnAnalisarApuracaoICMS"
                resICMS.Activate
                Call AnalistaICMS.GerarResumoApuracaoICMS
                GoTo Sair:
                
            Case "btnAtualizarRegistrosICMS"
                Call FuncoesAssistentesInteligentes.AtualizarRegistrosApuracaoICMS
                GoTo Sair:
                
            Case "btnVerificarTributacaoICMS"
                Call Assistente.Tributario.ICMS.VerificarTributacaoICMS
                GoTo Sair:
                
            Case "btnSalvarTributacaoICMS"
                Call Assistente.Tributario.ICMS.SalvarTributacaoICMS
                GoTo Sair:
                
            Case "btnAplicarTributacaoICMS"
                Call Assistente.Tributario.ICMS.AplicarTributacaoICMS
                GoTo Sair:
                
            Case "btnListarTributacoesICMS"
                Call Assistente.Fiscal.Apuracao.ICMS.ListarTributacoesICMS
                GoTo Sair:
            
            'TODO: INCLUIR FILTROS ABAIXO NO SISTEMA DE FILTROS
            Case "btnListarNotasICMS"
                Call Util.FiltrarRegistros(assApuracaoICMS, relInteligenteDivergencias, "CHV_NFE", "CHV_NFE")
                GoTo Sair:
                
            Case "btnListarItensICMS"
                Call Util.FiltrarRegistros(assApuracaoICMS, relInteligenteDivergencias, Array("CHV_NFE", "CFOP"), Array("CHV_NFE", "CFOP_SPED"))
                GoTo Sair:
            
            Case "btnCalcularDIFALNaoContribuinte"
                Call Assistente.Fiscal.Apuracao.ICMS.CalcularDifalNaoContribuinte
                GoTo Sair:
            
            Case "btnFiltrarRegistroC101"
                Call Util.FiltrarRegistros(assApuracaoICMS, regC101, "CHV_PAI_FISCAL", "CHV_PAI_FISCAL")
                GoTo Sair:
                
'##----->> Funcionalidades do Assistente de Apuração do IPI
            Case "btnGerarRelatorioIPI"
                Call FuncoesAssistentesInteligentes.GerarRelatorioApuracaoIPI
                GoTo Sair:
                
            Case "btnProcessarInconsistenciasIPI"
                Call FuncoesAssistentesInteligentes.ReprocessarSugestoesApuracaoIPI
                GoTo Sair:
                
            Case "btnAceitarSugestoesIPI"
                Call FuncoesAssistentesInteligentes.AceitarSugestoesApuracaoIPI
                GoTo Sair:
                
            Case "btnIgnorarInconsistenciasIPI"
                Call FuncoesAssistentesInteligentes.IgnorarInconsistenciasApuracaoIPI
                GoTo Sair:
                
            Case "btnAtualizarRegistrosIPI"
                Call FuncoesAssistentesInteligentes.AtualizarRegistrosApuracaoIPI
                GoTo Sair:
                
            Case "btnVerificarTributacaoIPI"
                Call Assistente.Tributario.IPI.VerificarTributacaoIPI
                GoTo Sair:
                
            Case "btnSalvarTributacaoIPI"
                Call Assistente.Tributario.IPI.SalvarTributacaoIPI
                GoTo Sair:
                
            Case "btnAplicarTributacaoIPI"
                Call Assistente.Tributario.IPI.AplicarTributacaoIPI
                GoTo Sair:
                
            Case "btnListarTributacoesIPI"
                Call Assistente.Fiscal.Apuracao.IPI.ListarTributacoesIPI
                GoTo Sair:
                
            Case "btnListarNotasIPI"
                Call Util.FiltrarRegistros(assApuracaoIPI, relInteligenteDivergencias, "CHV_NFE", "CHV_NFE")
                GoTo Sair:
                
            Case "btnListarItensIPI"
                Call Util.FiltrarRegistros(assApuracaoIPI, relInteligenteDivergencias, Array("CHV_NFE", "CFOP"), Array("CHV_NFE", "CFOP_SPED"))
                GoTo Sair:
                
'##----->> Funcionalidades do Assistente de Apuração do PIS e COFINS
            Case "btnGerarRelatorioPISCOFINS"
                Call FuncoesAssistentesInteligentes.GerarRelatorioApuracaoPISCOFINS
                GoTo Sair:
                
            Case "btnProcessarInconsistenciasPISCOFINS"
                Call Assistente.Fiscal.Apuracao.PISCOFINS.ReprocessarSugestoes
                GoTo Sair:
                
            Case "btnAceitarSugestoesPISCOFINS"
                Call Assistente.Fiscal.Apuracao.PISCOFINS.AceitarSugestoes
                GoTo Sair:
                
            Case "btnIgnorarInconsistenciasPISCOFINS"
                Call FuncoesAssistentesInteligentes.IgnorarInconsistenciasPISCOFINS
                GoTo Sair:
            
            Case "btnAnalisarApuracaoPISCOFINS"
                resICMS.Activate
                Call AnalistaPISCOFINS.GerarResumoApuracaoPISCOFINS
                GoTo Sair:
                
            Case "btnAtualizarRegistrosPISCOFINS"
                Call FuncoesAssistentesInteligentes.AtualizarRegistrosApuracaoPIS_COFINS
                GoTo Sair:
                
            Case "btnVerificarTributacaoPISCOFINS"
                Call Assistente.Tributario.PIS_COFINS.VerificarTributacaoPISCOFINS
                GoTo Sair:
                
            Case "btnAplicarTributacaoPISCOFINS"
                Call Assistente.Tributario.PIS_COFINS.AplicarTributacaoPISCOFINS
                GoTo Sair:
                
            Case "btnSalvarTributacaoPISCOFINS"
                Call Assistente.Tributario.PIS_COFINS.SalvarTributacaoPISCOFINS
                GoTo Sair:
                
            Case "btnListarTributacoesPISCOFINS"
                Call Assistente.Fiscal.Apuracao.PISCOFINS.ListarTributacoesPISCOFINS
                GoTo Sair:
                
            Case "btnListarNotasPISCOFINS"
                Call Util.FiltrarRegistros(assApuracaoPISCOFINS, relInteligenteDivergencias, "CHV_NFE", "CHV_NFE")
                GoTo Sair:
                
'##----->> Funcionalidades do Assistente de Divergências de Notas
            Case "btnGerarRelatorioDivergenciasNotas"
                Call DivergenciasNotas.GerarComparativoXMLSPED
                GoTo Sair:
                
            Case "btnProcessarInconsistenciasNotas"
                Call DivergenciasNotas.ReprocessarSugestoes
                GoTo Sair:
                
            Case "btnAceitarSugestoesNotas"
                Call DivergenciasNotas.AceitarSugestoesNotas
                GoTo Sair:
                
            Case "btnIgnorarInconsistenciasNotas"
                Call DivergenciasNotas.IgnorarInconsistencias
                GoTo Sair:
                
            Case "btnAtualizarRegistrosNotas"
                Call DivergenciasNotas.AtualizarRegistros
                GoTo Sair:
                
            Case "btnListarResumosC190Notas"
                 Call Util.FiltrarRegistros(relDivergenciasNotas, regC190, "CHV_REG", "CHV_PAI_FISCAL")
                GoTo Sair:

            Case "btnFiltrarDivergenciasProdutos"
                Call Util.FiltrarRegistros(relDivergenciasNotas, relDivergenciasProdutos, "CHV_NFE", "CHV_NFE")
                GoTo Sair:
                
'##----->> Funcionalidades do Assistente de Divergências de Produtos
            Case "btnGerarRelatorioDivergenciasProdutos"
                Call DivergenciasProd.GerarComparativoXMLSPED
                GoTo Sair:

            Case "btnProcessarInconsistenciasProdutos"
                Call DivergenciasProd.ReprocessarSugestoes
                GoTo Sair:
            
            Case "btnAceitarSugestoesProdutos"
                Call DivergenciasProd.AceitarSugestoesProdutos
                GoTo Sair:
            
            Case "btnIgnorarInconsistenciasProdutos"
                Call DivergenciasProd.IgnorarInconsistencias
                GoTo Sair:
                
            Case "btnAtualizarRegistrosProdutos"
                Call DivergenciasProd.AtualizarRegistros
                GoTo Sair:
                
            Case "btnFiltrarDivergenciasNotas"
                Call Util.FiltrarRegistros(relDivergenciasProdutos, relDivergenciasNotas, "CHV_NFE", "CHV_NFE")
                GoTo Sair:
                
'##----->> Funcionalidades do Assistente de Tributação do ICMS
                
            Case "btnImportarTributacaoICMS"
                Call Assistente.Tributario.ICMS.ImportarTributacaoICMS
                GoTo Sair:
                
            Case "btnExportarTributacaoICMS"
                Call Util.ExportarDadosRelatorio(assTributacaoICMS, "Tributação ICMS")
                GoTo Sair:
                
            Case "btnGerarModTributacaoICMS"
                Call Util.GerarModeloRelatorio(assTributacaoICMS, "Tributação ICMS")
                GoTo Sair:
                
            Case "btnGerarRelatorioTributacaoICMS"
                Call Assistente.Tributario.ICMS.ReprocessarSugestoes
                GoTo Sair:

            Case "btnProcessarInconsistenciasTributacaoICMS"
                Call Assistente.Tributario.ICMS.ReprocessarSugestoes
                GoTo Sair:

            Case "btnAceitarSugestoesTributacaoICMS"
                Call Assistente.Tributario.ICMS.AceitarSugestoes
                GoTo Sair:

            Case "btnIgnorarInconsistenciasTributacaoICMS"
                Call FuncoesAssistentesInteligentes.IgnorarInconsistenciasTributacao
                GoTo Sair:

            Case "btnResetarInconsistenciasTributacaoICMS"
                Call FuncoesAssistentesInteligentes.IgnorarInconsistenciasTributacao
                GoTo Sair:
                
            Case "btnAtualizarRegistrosTributacaoICMS"
                Call FuncoesAssistentesInteligentes.AtualizarRegistrosTributacao
                GoTo Sair:
                


'##----->> Funcionalidades do Assistente de Tributação do IPI
                
            Case "btnImportarTributacaoIPI"
                Call Assistente.Tributario.IPI.ImportarTributacaoIPI
                GoTo Sair:
                
            Case "btnExportarTributacaoIPI"
                Call Util.ExportarDadosRelatorio(assTributacaoIPI, "Tributação PIS-COFINS")
                GoTo Sair:
                
            Case "btnGerarModTributacaoIPI"
                Call Util.GerarModeloRelatorio(assTributacaoIPI, "Tributação PIS-COFINS")
                GoTo Sair:
                
            Case "btnGerarRelatorioTributacaoIPI"
                Call Assistente.Tributario.IPI.ReprocessarSugestoes
                GoTo Sair:

            Case "btnProcessarInconsistenciasTributacaoIPI"
                Call Assistente.Tributario.IPI.ReprocessarSugestoes
                GoTo Sair:

            Case "btnAceitarSugestoesTributacaoIPI"
                Call Assistente.Tributario.IPI.AceitarSugestoes
                GoTo Sair:

            Case "btnIgnorarInconsistenciasTributacaoIPI"
                Call FuncoesAssistentesInteligentes.IgnorarInconsistenciasTributacao
                GoTo Sair:

            Case "btnResetarInconsistenciasTributacaoIPI"
                Call FuncoesAssistentesInteligentes.IgnorarInconsistenciasTributacao
                GoTo Sair:
                
            Case "btnAtualizarRegistrosTributacaoIPI"
                Call FuncoesAssistentesInteligentes.AtualizarRegistrosTributacao
                GoTo Sair:
                
                
'##----->> Funcionalidades do Assistente de Tributação do PIS/COFINS
                
            Case "btnImportarTributacaoPISCOFINS"
                Call Assistente.Tributario.PIS_COFINS.ImportarTributacaoPISCOFINS
                GoTo Sair:
                
            Case "btnExportarTributacaoPISCOFINS"
                Call Util.ExportarDadosRelatorio(assTributacaoPISCOFINS, "Tributação PIS-COFINS")
                GoTo Sair:
                
            Case "btnGerarModTributacaoPISCOFINS"
                Call Util.GerarModeloRelatorio(assTributacaoPISCOFINS, "Tributação PIS-COFINS")
                GoTo Sair:
                
            Case "btnGerarRelatorioTributacaoPISCOFINS"
                Call Assistente.Tributario.PIS_COFINS.ReprocessarSugestoes
                GoTo Sair:
                
            Case "btnImporTribNCM_PISCOFINS"
                Call impTributarioNCM.AtualizarTributacaoNCM(assTributacaoPISCOFINS)
                GoTo Sair:
                
            Case "btnExporTribNCM_PISCOFINS"
                Call impTributarioNCM.GerarPlanilhaTributacaoNCM(assTributacaoPISCOFINS)
                GoTo Sair:
                
            Case "btnGerarModeloNCM_PISCOFINS"
                Call impTributarioNCM.GerarModeloTributacaoNCM_PISCOFINS
                GoTo Sair:
                
            Case "btnProcessarInconsistenciasTributacaoPISCOFINS"
                Call Assistente.Tributario.PIS_COFINS.ReprocessarSugestoes
                GoTo Sair:
                
            Case "btnAceitarSugestoesTributacaoPISCOFINS"
                Call Assistente.Tributario.PIS_COFINS.AceitarSugestoes
                GoTo Sair:
                
            Case "btnIgnorarInconsistenciasTributacaoPISCOFINS"
                Call FuncoesAssistentesInteligentes.IgnorarInconsistenciasTributacao
                GoTo Sair:
                
            Case "btnResetarInconsistenciasTributacaoPISCOFINS"
                Call FuncoesAssistentesInteligentes.IgnorarInconsistenciasTributacao
                GoTo Sair:
                
            Case "btnAtualizarRegistrosTributacaoPISCOFINS"
                Call FuncoesAssistentesInteligentes.AtualizarRegistrosTributacao
                GoTo Sair:
                

'##----->> Funcionalidades do Assistente de Apuração de Assistente de Custos e Preços
            Case "btnGerarCustosPrecos"
                Call FuncoesAssistentesInteligentes.GerarRelatorioCustosPrecos
                GoTo Sair:
                                    
            
'##----->> Funcionalidades do Assistente de Apuração de Estoque
            Case "btnGerarEstoque"
                Call FuncoesAssistentesInteligentes.GerarRelatorioEstoque
                GoTo Sair:
            
            Case "btnProcessarInconsistenciasEstoque"
                Call FuncoesAssistentesInteligentes.ReprocessarSugestoesEstoque
                GoTo Sair:
            
            Case "btnAceitarSugestoesEstoque"
                Call FuncoesAssistentesInteligentes.AceitarSugestoesEstoque
                GoTo Sair:
            
            Case "btnIgnorarInconsistenciasEstoque"
                Call FuncoesAssistentesInteligentes.IgnorarInconsistenciasEstoque
                GoTo Sair:
                
            Case "btnAtualizarRegistrosEstoque"
                Call FuncoesAssistentesInteligentes.AtualizarRegistrosEstoque
                GoTo Sair:
            

'##----->> Funcionalidades do Assistente de Apuração de Contas
            Case "btnImportarC140Lote"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call FuncoesXML.ImportarRegistrosC140eFilhos("Lote")
                GoTo Sair:
            
            Case "btnImportarC140Arquivo"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call FuncoesXML.ImportarRegistrosC140eFilhos("Arquivo")
                GoTo Sair:
            
            Case "btnGerarRelatorioContas"
                Call FuncoesAssistentesInteligentes.GerarRelatorioContasPagarReceber
                GoTo Sair:
            
            Case "btnProcessarInconsistenciasContas"
                Call FuncoesAssistentesInteligentes.ReprocessarSugestoesContas
                GoTo Sair:
            
            Case "btnAceitarSugestoesContas"
                Call FuncoesAssistentesInteligentes.AceitarSugestoesContas
                GoTo Sair:
            
            Case "btnAtualizarRegistrosContas"
                Call FuncoesAssistentesInteligentes.AtualizarContasPagarReceber
                GoTo Sair:
            

'##----->> Funcionalidades do Assistente de Inventário
            Case "btnImportarInventario"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call FuncoesAssistentesInteligentes.ImportarInventarioFisico
                GoTo Sair:
                
            Case "btnGerarModeloInv"
                Call FuncoesExcel.GerarModeloInventario
                GoTo Sair:
            
            Case "btnGerarRelatorioInventario"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call FuncoesAssistentesInteligentes.GerarRelatorioInventario
                GoTo Sair:
                
            Case "btnProcessarInconsistenciasInventario"
                Call FuncoesAssistentesInteligentes.ReprocessarSugestoesInventario
                GoTo Sair:
                
            Case "btnAceitarSugestoesInventario"
                Call FuncoesAssistentesInteligentes.AceitarSugestoesInventario
                GoTo Sair:
                
            Case "btnAtualizarRegistrosInventario"
                If Not Funcoes.CarregarDadosContribuinte Then GoTo Sair:
                Call FuncoesAssistentesInteligentes.AtualizarInventario
                GoTo Sair:
                
            Case "btnListarInconsistenciasInventario"
                Call FuncoesFiltragem.FiltrarInconsistencias(ActiveSheet)
                GoTo Sair:
                
'##----->> Funcionalidades do Analista de Apuração do ICMS
            Case "btnGerarAnaliseICMS"
                Call AnalistaICMS.GerarResumoApuracaoICMS
                GoTo Sair:
                
            Case "btnFiltrarAnaliseICMS"
                Call AnalistaICMS.FiltrarRegistros
                GoTo Sair:
                
'##----->> Funcionalidades do Analista de Apuração do PIS e COFINS
            Case "btnGerarAnalisePISCOFINS"
                Call AnalistaPISCOFINS.GerarResumoApuracaoPISCOFINS
                GoTo Sair:
                
            Case "btnFiltrarAnalisePISCOFINS"
                Call AnalistaPISCOFINS.FiltrarRegistros
                GoTo Sair:
                
        End Select
        
    Else
        
        Msg = "O seu nível de assinatura não dá acesso a esta funcionalidade." & vbCrLf & vbCrLf
        Msg = Msg & "Com o plano Premium, você desbloqueia o acesso a todas as nossas funcionalidades de relatórios inteligentes e maximiza sua produtividade." & vbCrLf & vbCrLf
        Msg = Msg & "Clique em SIM para fazer o upgrade do seu plano agora mesmo!"
        
        vbResult = MsgBox(Msg, vbExclamation + vbYesNo, "Necessário Upgrade de Plano")
        If vbResult = vbYes Then Call FuncoesLinks.AbrirUrl(urlAssinaturaEmpresarialMensal)
        GoTo Sair:
        
    End If
        
    If NivelPlano >= 4 Then
        
'##----->> Botões de acesso aos recursos do plano Enterprise
        Select Case True
        
            Case control.id Like "btnOportunidades*", control.id Like "btnInventario*"
                Call FaixaOpcoes.IrPara(control)
                GoTo Sair:
        
        End Select
                
        Select Case control.id
                
                
'##----->> Funcionalidades do Assistente de Oportunidades do IPI

            Case "btnImportarSPEDsOriginaisIPI"
                Call Oportunidades.GerarRelatorioOportunidades(True, "IPI")
                GoTo Sair:
                
            Case "btnImportarSPEDsCorrigidosIPI"
                Call Oportunidades.GerarRelatorioOportunidades(False, "IPI")
                GoTo Sair:
                
'##----->> Funcionalidades do Assistente de Oportunidades do ICMS
            Case "btnImportarSPEDsOriginaisICMS"
                Call Oportunidades.GerarRelatorioOportunidades(True, "ICMS")
                GoTo Sair:
                
            Case "btnImportarSPEDsCorrigidosICMS"
                Call Oportunidades.GerarRelatorioOportunidades(False, "ICMS")
                GoTo Sair:
                
'##----->> Funcionalidades do Assistente de Oportunidades do PIS e COFINS
            Case "btnImportarSPEDsOriginaisPISCOFINS"
                Call Oportunidades.GerarRelatorioOportunidades(True, "PISCOFINS")
                GoTo Sair:
                
            Case "btnImportarSPEDsCorrigidosPISCOFINS"
                Call Oportunidades.GerarRelatorioOportunidades(False, "PISCOFINS")
                GoTo Sair:
                
'##----->> Funcionalidades do Assistente Analítico de Movimentação de Estoque
            Case "btnGerarAnaliseEstoque"
                Call Estoque.GerarRelatorioMovimentacaoEstoque
                GoTo Sair:
                
'##----->> Funcionalidades do Auditor de Inventário
            Case "btnGerarSaldoInventario"
                Call Inventario.GerarRelatorioSaldoInventario
                GoTo Sair:
                
        End Select
        
    Else
        
        Msg = "O seu nível de assinatura não dá acesso a esta funcionalidade." & vbCrLf & vbCrLf
        Msg = Msg & "Com o plano Enterprise, você desbloqueia o acesso a todas as nossas funcionalidades para maximizar sua produtividade, diversificar seus serviços e aumentar sua lucratividade." & vbCrLf & vbCrLf
        Msg = Msg & "Clique em SIM para fazer o upgrade do seu plano agora mesmo!"
        
        vbResult = MsgBox(Msg, vbExclamation + vbYesNo, "Necessário Upgrade de Plano")
        If vbResult = vbYes Then Call FuncoesLinks.AbrirUrl(urlAssinaturaEmpresarialMensal)
        GoTo Sair:
        
    End If
    
Sair:
    
    'Application.StatusBar = False
    Call Util.HabilitarControles
    
End Sub

Private Function VerificarAssinatura(ByRef NivelPlano As Byte, ByRef Status As String, ByVal Plano As String) As Boolean

Dim Msg As String
    
    EmailAssinante = relGestaoAssinatura.Range("email_cliente").value
    If EmailAssinante = "" Then Call FuncoesControlDocs.ResetarAssinatura

    Status = Util.FormatarNomePersonalizado("status")
    If Status <> "ACTIVE" Then
        Call FuncoesControlDocs.ResetarAssinatura
        
        Select Case True
        
            Case Status = ""
                Msg = "Assinatura não identificada!" & vbCrLf & vbCrLf
                Msg = Msg & "Por favor, faça a autenticação para continuar aproveitando uma rotina mais rápida, prática e segura."
                Call Util.MsgAlerta(Msg, "Assinatura Não Identificada")
        
            Case Status = "INACTIVE"
                Msg = "Assinatura Inativa!" & vbCrLf & vbCrLf
                Msg = Msg & "A sua assinatura está inativa, renove sua assinatura para continuar aproveitando uma rotina mais rápida, prática e segura."
                Call Util.MsgAlerta(Msg, "Assinatura Inativa")
        
            Case Status = "DELAYED"
                Msg = "Assinatura Atrasada!" & vbCrLf & vbCrLf
                Msg = Msg & "A sua assinatura está atrasada, renove sua assinatura para continuar aproveitando uma rotina mais rápida, prática e segura."
                Call Util.MsgAlerta(Msg, "Assinatura Atrasada")
            
            Case Status = "FINISH"
                Msg = "Assinatura Finalizada!" & vbCrLf & vbCrLf
                Msg = Msg & "A sua assinatura está finalizada, renove sua assinatura para continuar aproveitando uma rotina mais rápida, prática e segura."
                Call Util.MsgAlerta(Msg, "Assinatura Inativa")
            
            Case Status Like "CANCELLED"
                Msg = "Assinatura Cancelada!" & vbCrLf & vbCrLf
                Msg = Msg & "A sua assinatura está Cancelada, faça a renovação para continuar aproveitando uma rotina mais rápida, prática e segura."
                Call Util.MsgAlerta(Msg, "Assinatura Cancelada")
            
        End Select
        
        Exit Function
        
    End If
    
    If Application.EnableEvents = False Then Application.EnableEvents = True
    
    Plano = Util.FormatarNomePersonalizado("plano")
    Select Case True
        
        Case Plano Like "*ultra*", Plano Like "*enterprise*"
            NivelPlano = 4
            
        Case Plano Like "*premium*" Or Plano Like "*cliente*"
            NivelPlano = 3
            
        Case Plano Like "*plus*"
            NivelPlano = 2
            
        Case Plano Like "*basico*" Or Plano Like "*vitalicio*" Or Plano Like "*lancamento*"
            NivelPlano = 1
            
        Case Else
            NivelPlano = 0
            
    End Select
    
    VerificarAssinatura = True
    
End Function

