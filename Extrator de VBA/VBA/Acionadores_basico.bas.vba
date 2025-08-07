Attribute VB_Name = "Acionadores_basico"
Option Explicit

Public Function AcionadoresBasico(control As IRibbonControl, Optional ByRef Valor)

    Select Case control.id
        
        Case "btnDefinirPeriodo"
            UsarPeriodo = Valor
            Call FaixaOpcoes.getPressed(control, Valor)
            Exit Function
        
        Case "btnICMS", "btnDivergencias", "btnCorrelacoes", "btnIPI", "btnPISCOFINS", "btnTributacao"
            Call FaixaOpcoes.IrPara(control)
            Exit Function
        
        Case "btnGerarLivroICMS"
            Call FuncoesLivrosFiscais.GerarLivroICMS
            Exit Function
        
        Case "btnFiltrarEntradas"
            Call FuncoesFiltragem.FiltrarEntradas
            Exit Function
        
        Case "btnFiltrarSaidas"
            Call FuncoesFiltragem.FiltrarSaidas
            Exit Function
        
        Case "btnAcessarEnfoqueDeclarante"
            Call FuncoesFiltragem.AcessarEnfoqueDeclarante
            Exit Function
        
        Case "btnListarIncosnsistencias"
            Call FuncoesFiltragem.ListarDivergencias
            Exit Function
        
        Case "btnImportarSPEDFiscal"
            Call FuncoesSPEDFiscal.ImportarSPED
            Exit Function
        
        Case "btnImportRegAtual"
            Call FuncoesSPEDFiscal.ImportarSPED(VBA.Left(ActiveSheet.name, 4), PeriodoImportacao)
            Exit Function
        
        Case "btnLimparRegistrosSPED"
            Call fnSPED.LimparRegistrosEFD
            Exit Function
        
       Case "btnListarNotasC170"
             Call Util.FiltrarRegistros(regC170, regC100, "CHV_PAI_FISCAL", "CHV_REG")
            Exit Function
        
        Case "btnListarItensC170"
             Call Util.FiltrarRegistros(regC100, regC170, "CHV_REG", "CHV_PAI_FISCAL")
            Exit Function
        
        Case "btnListarResumosC190"
             Call Util.FiltrarRegistros(regC100, regC190, "CHV_REG", "CHV_PAI_FISCAL")
            Exit Function
        
        Case "btnListarNotasC100"
             Call Util.FiltrarRegistros(reg0150, regC100, "COD_PART", "COD_PART")
            Exit Function
        
        Case "btnListarItensC170C190"
             Call Util.FiltrarRegistros(regC190, regC170, "CHV_PAI_FISCAL", "CHV_PAI_FISCAL")
            Exit Function
        
        Case "btnListarResumosC190C170"
             Call Util.FiltrarRegistros(regC170, regC190, "CHV_PAI_FISCAL", "CHV_PAI_FISCAL")
            Exit Function
        
        Case "btnListarItensC810"
             Call Util.FiltrarRegistros(regC800, regC810, "CHV_REG", "CHV_PAI_FISCAL")
            Exit Function
        
        Case "btnListarResumosC850"
            Call Util.FiltrarRegistros(regC800, regC850, "CHV_REG", "CHV_PAI_FISCAL")
            Exit Function
        
        Case "btnListarItensC810C850"
            Call Util.FiltrarRegistros(regC850, regC810, "CHV_PAI_FISCAL", "CHV_PAI_FISCAL")
            Exit Function
        
        Case "btnListarResumosC850C810"
            Call Util.FiltrarRegistros(regC810, regC850, "CHV_PAI_FISCAL", "CHV_PAI_FISCAL")
            Exit Function
                         
        Case "btnListarNotasC810"
            Call Util.FiltrarRegistros(regC810, regC800, "CHV_PAI_FISCAL", "CHV_REG")
            Exit Function
        
        Case "btnListarNotasC850"
            Call Util.FiltrarRegistros(regC850, regC800, "CHV_PAI_FISCAL", "CHV_REG")
            Exit Function
        
        Case "btnListarResumosD190"
            Call Util.FiltrarRegistros(regD100, regD190, "CHV_REG", "CHV_PAI_FISCAL")
            Exit Function
        
        Case "btnListarNotasD100"
            Call Util.FiltrarRegistros(regD190, regD100, "CHV_PAI_FISCAL", "CHV_REG")
            Exit Function
        
        Case "btnListarNotasC190"
            Call Util.FiltrarRegistros(regC190, regC100, "CHV_PAI_FISCAL", "CHV_REG")
            Exit Function
        
        Case "btnGerarModelo"
            Call FuncoesExcel.GerarModelo0200
            Exit Function
        
        Case "btnAjustarC190peloC100"
            Call rC100.RatearDivergenciasC100ParaC190
            Exit Function
        
        Case "btnImportarXMLSAnalise"
            Call FuncoesXML.ImportarXMLSParaAnalise("Arquivo")
            Exit Function
        
        Case "btnGerarC175peloC170"
            Call rC170.GerarC175
            Exit Function
        
        Case "btnGerarC190peloC170"
            Call rC170.GerarC190
            Exit Function
        
        Case "btnSomarST", "btnSomarIPIeST"
            If control.id = "btnSomarST" Then Call rC170.SomarIPIeSTaosItens("ST")
            If control.id = "btnSomarIPIeST" Then Call rC170.SomarIPIeSTaosItens("IPI-ST")
            Exit Function
        
        Case "btnAtualizarC100"
            Call rC170.AtualizarImpostosC100
            Exit Function
        
        Case "btnAtualizarC100C190"
            Call rC190.AtualizarImpostosC100
            Exit Function
        
        Case "btnExportarSPEDFiscal"
            Call FuncoesSPEDFiscal.GerarEFDICMSIPI
            Exit Function
        
        Case "btnAgruparC190"
            rC190.AgruparRegistros
            Exit Function
        
        Case "btnImportarLoteC100Fiscal"
            Call FuncoesXML.ImportarXMLsC100("Lote")
            Exit Function
        
        Case "btnImportarArqC100Fiscal"
            Call FuncoesXML.ImportarXMLsC100("Arquivo")
            Exit Function
        
        Case "btnImportarLoteC800Fiscal"
            Call FuncoesXML.ImportarRegistrosCFeXML("Lote")
            Exit Function
        
        Case "btnImportarArqC800Fiscal"
            Call FuncoesXML.ImportarRegistrosCFeXML("Arquivo")
            Exit Function
        
        Case "btnImportarLoteD100Fiscal"
            Call FuncoesXML.ImportarRegistrosCTeXML("Lote")
            Exit Function
        
        Case "btnImportarArqD100Fiscal"
            Call FuncoesXML.ImportarRegistrosCTeXML("Arquivo")
            Exit Function
    
        Case "btnAtualizarC800C850"
            Call rC850.AtualizarImpostosC800
            Exit Function
        
        Case "btnAtualizarCodGenero"
            Call r0200.AtualizarCodigoGenero
            Exit Function
        
        Case "btnAgruparC850"
            Call rC850.AgruparRegistros
            Exit Function
        
        Case "btnCalcBasePISCOFINS"
            Call rC170.CalcularPISCOFINS(False)
            Exit Function
        
        Case "btnExcluirICMS"
            Call rC170.CalcularPISCOFINS(True)
            Exit Function
    
        Case "btnGerarCreditoSIMPLESNACIONAL"
            Call rE111.GerarCreditoSIMPLESNACIONAL
            Exit Function
        
        Case "btnProdutosFornecedorArquivo"
            Call FuncoesXML.ImportarProdutosFornecedor("Arquivo")
            Exit Function
        
        Case "btnProdutosFornecedorLote"
            Call FuncoesXML.ImportarProdutosFornecedor("Lote")
            Exit Function
        
        Case "btnCorrelacionarSPEDXML"
            Call FuncoesXML.ImportarProdutosFornecedor("Correlacionar")
            Exit Function
        
        Case "btnImportarEFDAnalise"
            Call FuncoesSPEDFiscal.ImportarSPEDFiscalparaAnalise
            Call FuncoesFormatacao.AplicarFormatacao(relInteligenteDivergencias)
            Exit Function
        
        Case "btnImportarXMLAnaliseLote"
            Call FuncoesXML.ImportarXMLSParaAnalise("Lote")
            Exit Function
        
        Case "btnImportarXMLAnaliseArquivo"
            Call FuncoesXML.ImportarXMLSParaAnalise("Arquivo")
            Exit Function
        
        Case "btnListarXMLSAusentes"
            Call Rotinas.ListarXMLsAusentes
            Exit Function
        
        Case "btnExportarRelatorioCorrelacao"
            Call FuncoesAssistentesInteligentes.EnviarDados
            Exit Function
        
        Case "btnCalcRedBCICMSC190"
            Call rC190.CalcularReducaoBaseICMS
            Exit Function
        
        Case "btnAcessarNotaSelecionada"
            Call Util.FiltrarRegistros(relInteligenteDivergencias, regC100, "CHV_NFE", "CHV_NFE")
            Exit Function
        
        Case "btnImportarItensSPED"
            Call FuncoesTributacao.ImportarItensSPED
            Exit Function
        
        Case "btnAnalisarTributacao"
            Call FuncoesTributacao.VerificarTributacao
            Exit Function
        
        Case "btnImportarCadastroItens"
            Call FuncoesTributacao.ImportarCadastroTributacao
            Exit Function
        
        Case "btnExportarCadastroItens"
            Call FuncoesTXT.ExportarParaTxt(Tributacao, "CODIGO", "CFOP")
            Exit Function
        
        Case "btnImportarCadastroCorrelacoes"
            Call FuncoesExcel.ImportarCadastroCorrelacoes
            Exit Function
        
        Case "btnListarDivergencias"
            Call FuncoesFiltragem.ListarDivergencias
            Exit Function
        
        Case "btnEstruturarSPED"
            Call FuncoesSPEDFiscal.EstruturarSPED
            Exit Function
        
        Case "btnConsultarCEP0005", "btnConsultarCEP0100", "btnConsultarCEP0150"
            Call FuncoesAPI.ConsultarCEP(ActiveSheet)
            Exit Function
        
        Case "btnImportarProd0200"
            Call FuncoesExcel.ImportarCadastro0200
            Exit Function
        
        Case "btnExportarProd0200"
            Call FuncoesExcel.ExportarCadastro0200
            Exit Function
        
        Case "btnImportar0200Lote"
            Call FuncoesXML.ImportarCadastro0200XML("Lote")
            Exit Function
        
        Case "btnImportar0200Arquivo"
            Call FuncoesXML.ImportarCadastro0200XML("Arquivo")
            Exit Function
        
        Case "btnRemoverDuplicatas"
            If ActiveSheet.CodeName Like "reg*" Then Call FuncoesPlanilha.RemoverDuplicatas(ActiveSheet, "CHV_REG", "CHV_PAI_FISCAL")
            Exit Function
        
    End Select

End Function

