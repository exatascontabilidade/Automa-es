Attribute VB_Name = "FuncoesXML"
Option Explicit

Public Sub ImportarProtocolosCancelamento()

Dim dicEntradasCTes As New Dictionary
Dim dicEntradasNFe As New Dictionary
Dim dicSaidasNFCe As New Dictionary
Dim dicSaidasCTes As New Dictionary
Dim dicSaidasCFes As New Dictionary
Dim dicSaidasNFe As New Dictionary
Dim arrCanceladas As New ArrayList
Dim arrChaves As New ArrayList
Dim arrXMLs As New ArrayList
Dim Caminho As String
Dim Status As Boolean
Dim Comeco As Double
    
    Caminho = Util.SelecionarPasta("Selecione a pasta que contém os arquivos XML")
    If Caminho = "" Then Exit Sub
    
    Inicio = Now()
    Application.StatusBar = "Listando arquivos XML na estrutura de pastas..."
    
    Call Util.ListarArquivos(arrXMLs, Caminho)
    If arrXMLs.Count > 0 Then
        
        Status = CarregarDadosContribuinte
        If Not Status Then Exit Sub
        
        Call fnXML.CarregarProtocolosCancelamento(arrXMLs, arrCanceladas)
        
        Application.StatusBar = "Carregando dados das NFe de entrada, por favor aguarde..."
        Set dicEntradasNFe = Util.CriarDicionarioRegistro(EntNFe, "Chave de Acesso")
        
        Application.StatusBar = "Carregando dados dos CTe de entrada, por favor aguarde..."
        Set dicEntradasCTes = Util.CriarDicionarioRegistro(EntCTe, "Chave de Acesso")
        
        Application.StatusBar = "Carregando dados das NFe de saída, por favor aguarde..."
        Set dicSaidasNFe = Util.CriarDicionarioRegistro(SaiNFe, "Chave de Acesso")
        
        Application.StatusBar = "Carregando dados das NFCe, por favor aguarde..."
        Set dicSaidasNFCe = Util.CriarDicionarioRegistro(SaiNFCe, "Chave de Acesso")
        
        Application.StatusBar = "Carregando dados dos CTe de saída, por favor aguarde..."
        Set dicSaidasCTes = Util.CriarDicionarioRegistro(SaiCTe, "Chave de Acesso")
        
        Application.StatusBar = "Carregando dados dos CFe de saída, por favor aguarde..."
        Set dicSaidasCFes = Util.CriarDicionarioRegistro(SaiCFe, "Chave de Acesso")
        
        Application.StatusBar = "Atualizando situação dos documentos fiscais, por favor aguarde..."
        Call Util.AtualizarSituacaoDoce(EntNFe, dicEntradasNFe, arrCanceladas)
        Call Util.AtualizarSituacaoDoce(SaiNFe, dicSaidasNFe, arrCanceladas)
        Call Util.AtualizarSituacaoDoce(SaiNFCe, dicSaidasNFCe, arrCanceladas)
        Call Util.AtualizarSituacaoDoce(EntCTe, dicEntradasCTes, arrCanceladas)
        Call Util.AtualizarSituacaoDoce(SaiCTe, dicSaidasCTes, arrCanceladas)
        Call Util.AtualizarSituacaoDoce(SaiCFe, dicSaidasCFes, arrCanceladas)
        
        If dicEntradasNFe.Count > 0 Then
            Call Util.LimparDados(EntNFe, 4, False)
            Call Util.ExportarDadosDicionario(EntNFe, dicEntradasNFe)
        End If
        
        If dicEntradasCTes.Count > 0 Then
            Call Util.LimparDados(EntCTe, 4, False)
            Call Util.ExportarDadosDicionario(EntCTe, dicEntradasCTes)
        End If
        
        If dicSaidasNFe.Count > 0 Then
            Call Util.LimparDados(SaiNFe, 4, False)
            Call Util.ExportarDadosDicionario(SaiNFe, dicSaidasNFe)
        End If
        
        If dicSaidasNFCe.Count > 0 Then
            Call Util.LimparDados(SaiNFCe, 4, False)
            Call Util.ExportarDadosDicionario(SaiNFCe, dicSaidasNFCe)
        End If
        
        If dicSaidasCTes.Count > 0 Then
            Call Util.LimparDados(SaiCTe, 4, False)
            Call Util.ExportarDadosDicionario(SaiCTe, dicSaidasCTes)
        End If
        
        If dicSaidasCFes.Count > 0 Then
            Call Util.LimparDados(SaiCFe, 4, False)
            Call Util.ExportarDadosDicionario(SaiCFe, dicSaidasCFes)
        End If
        
        Call Util.MsgInformativa("Importação concluída com sucesso!", "Protocolos de Cancelamentos", Inicio)
        
    End If
    
End Sub

Public Sub IdentificarNotasReferenciadas()

Dim dicChavesReferenciadas As New Dictionary
Dim dicEntradasCTes As New Dictionary
Dim dicEntradasNFe As New Dictionary
Dim dicSaidasNFCe As New Dictionary
Dim dicSaidasCTes As New Dictionary
Dim dicSaidasCFes As New Dictionary
Dim dicSaidasNFe As New Dictionary
Dim arrChaves As New ArrayList
Dim arrXMLs As New ArrayList
Dim Caminho As String
Dim Status As Boolean
Dim Comeco As Double
    
    Caminho = Util.SelecionarPasta("Selecione a pasta que contém os arquivos XML")
    If Caminho = "" Then Exit Sub
    
    Inicio = Now()
    Application.StatusBar = "Listando arquivos XML na estrutura de pastas..."
    
    Call Util.ListarArquivos(arrXMLs, Caminho)
    If arrXMLs.Count > 0 Then
        
        Status = CarregarDadosContribuinte
        If Not Status Then Exit Sub
        
        Call fnXML.CarregarChavesReferenciadas(arrXMLs, dicChavesReferenciadas)
        
        Application.StatusBar = "Carregando dados das NFe de entrada, por favor aguarde..."
        Set dicEntradasNFe = Util.CriarDicionarioRegistro(EntNFe, "Chave de Acesso")
        
        Application.StatusBar = "Carregando dados dos CTe de entrada, por favor aguarde..."
        Set dicEntradasCTes = Util.CriarDicionarioRegistro(EntCTe, "Chave de Acesso")
        
        Application.StatusBar = "Atualizando situação dos documentos fiscais, por favor aguarde..."
        Call Util.AtualizarSituacaoDoce(EntNFe, dicEntradasNFe, , dicChavesReferenciadas)
        Call Util.AtualizarSituacaoDoce(EntCTe, dicEntradasCTes, , dicChavesReferenciadas)
        
        If dicEntradasNFe.Count > 0 Then
            Call Util.LimparDados(EntNFe, 4, False)
            Call Util.ExportarDadosDicionario(EntNFe, dicEntradasNFe)
        End If
        
        If dicEntradasCTes.Count > 0 Then
            Call Util.LimparDados(EntCTe, 4, False)
            Call Util.ExportarDadosDicionario(EntCTe, dicEntradasCTes)
        End If
        
        Call Util.MsgInformativa("Atualização concluída com sucesso!", "Identificação de Notas Devolvidas", Inicio)
        
    End If
    
End Sub

Public Sub ImportarDocumentosEletronicos()

Dim dicEntradasCTes As New Dictionary
Dim dicEntradasNFe As New Dictionary
Dim dicSaidasNFCe As New Dictionary
Dim dicSaidasCTes As New Dictionary
Dim dicSaidasCFes As New Dictionary
Dim dicSaidasNFe As New Dictionary
Dim arrChaves As New ArrayList
Dim Status As Boolean
Dim Comeco As Double
    
    Call ProcessarDocumentos.CarregarXMLS("Lote")
    Call ProcessarDocumentos.ListarTodosDocumentos
    
    If DocsFiscais.arrTodos.Count > 0 Then
        
        Status = CarregarDadosContribuinte
        If Not Status Then Exit Sub
        
        'Cria dicionários com as informações das NFe de entrada já importadas
        Application.StatusBar = "Carregando dados das NFe de entrada, por favor aguarde..."
        Set dicEntradasNFe = Util.CriarDicionarioRegistro(EntNFe, "Chave de Acesso")
        
        'Cria dicionários com as informações dos CTe de entrada já importadas
        Application.StatusBar = "Carregando dados dos CTe de entrada, por favor aguarde..."
        Set dicEntradasCTes = Util.CriarDicionarioRegistro(EntCTe, "Chave de Acesso")
        
        'Cria dicionários com as informações das NFe de saída já importadas
        Application.StatusBar = "Carregando dados das NFe de saída, por favor aguarde..."
        Set dicSaidasNFe = Util.CriarDicionarioRegistro(SaiNFe, "Chave de Acesso")
        
        'Cria dicionários com as informações das NFCe de saída já importadas
        Application.StatusBar = "Carregando dados das NFCe, por favor aguarde..."
        Set dicSaidasNFCe = Util.CriarDicionarioRegistro(SaiNFCe, "Chave de Acesso")
        
        'Cria dicionários com as informações dos CTe de saída já importadas
        Application.StatusBar = "Carregando dados dos CTe de saída, por favor aguarde..."
        Set dicSaidasCTes = Util.CriarDicionarioRegistro(SaiCTe, "Chave de Acesso")
        
        'Cria dicionários com as informações das CFe de saída já importadas
        Application.StatusBar = "Carregando dados dos CFe de saída, por favor aguarde..."
        Set dicSaidasCFes = Util.CriarDicionarioRegistro(SaiCFe, "Chave de Acesso")
        
        Application.StatusBar = "Atualizando situação dos documentos fiscais, por favor aguarde..."
        
        With DocsFiscais
            
            'Atualiza o status dos documentos caso estejam cancelados
            Call Util.AtualizarSituacaoDoce(EntNFe, dicEntradasNFe, .arrChavesCanceladas)
            Call Util.AtualizarSituacaoDoce(SaiNFe, dicSaidasNFe, .arrChavesCanceladas)
            Call Util.AtualizarSituacaoDoce(SaiNFCe, dicSaidasNFCe, .arrChavesCanceladas)
            Call Util.AtualizarSituacaoDoce(EntCTe, dicEntradasCTes, .arrChavesCanceladas)
            Call Util.AtualizarSituacaoDoce(SaiCTe, dicSaidasCTes, .arrChavesCanceladas)
            Call Util.AtualizarSituacaoDoce(SaiCFe, dicSaidasCFes, .arrChavesCanceladas)
            
        End With
        
        'Carrega chaves de acesso dos documentos já importados
        Call Util.CarregarChavesAcessoDoces(arrChaves)
        
        'Processa os documentos selecionados pelo usuário
        Call fnXML.ImportarDocumentosEletronicos(DocsFiscais.arrTodos, arrChaves, dicEntradasNFe, _
            dicSaidasNFe, dicEntradasCTes, dicSaidasCTes, dicSaidasNFCe, dicSaidasCFes, DocsFiscais.arrChavesCanceladas)
            
        'Exporta as informações de volta para as planilhas de origem já atualizadas
        If dicEntradasNFe.Count > 0 Then
            Call Util.LimparDados(EntNFe, 4, False)
            Call Util.ExportarDadosDicionario(EntNFe, dicEntradasNFe)
        End If
        
        If dicEntradasCTes.Count > 0 Then
            Call Util.LimparDados(EntCTe, 4, False)
            Call Util.ExportarDadosDicionario(EntCTe, dicEntradasCTes)
        End If
        
        If dicSaidasNFe.Count > 0 Then
            Call Util.LimparDados(SaiNFe, 4, False)
            Call Util.ExportarDadosDicionario(SaiNFe, dicSaidasNFe)
        End If
        
        If dicSaidasNFCe.Count > 0 Then
            Call Util.LimparDados(SaiNFCe, 4, False)
            Call Util.ExportarDadosDicionario(SaiNFCe, dicSaidasNFCe)
        End If
        
        If dicSaidasCTes.Count > 0 Then
            Call Util.LimparDados(SaiCTe, 4, False)
            Call Util.ExportarDadosDicionario(SaiCTe, dicSaidasCTes)
        End If
        
        If dicSaidasCFes.Count > 0 Then
            Call Util.LimparDados(SaiCFe, 4, False)
            Call Util.ExportarDadosDicionario(SaiCFe, dicSaidasCFes)
        End If
        
        Call Util.MsgInformativa("Importação concluída com sucesso!", "Importação de documentos eletrônicos", Inicio)
        
    End If
    
End Sub

Public Sub ImportarXMLsC100(ByVal TipoImportacao As String, Optional ByVal SPEDContrib As Boolean)

Dim ARQUIVO As String, Periodo$, Caminho$, CNPJEmit$, CNPJDest$, tpPart$, CHV_0000$, CHV_0001$, CHV_0140$, CHV_0150$
Dim VL_PIS_ST As String, VL_COFINS_ST$, Msg$, CHV_C001$, CHV_PAI$, CNPJ0140$, tpCont$, CNPJCont$
Dim arrCorrelacoesExistentes As New ArrayList
Dim dicCorrelacoes As New Dictionary
Dim dicTitulos0000 As New Dictionary
Dim dicTitulos0001 As New Dictionary
Dim dicTitulos0005 As New Dictionary
Dim dicTitulos0100 As New Dictionary
Dim dicTitulos0110 As New Dictionary
Dim dicTitulos0140 As New Dictionary
Dim dicTitulos0150 As New Dictionary
Dim dicTitulos0190 As New Dictionary
Dim dicTitulos0200 As New Dictionary
Dim dicTitulos0220 As New Dictionary
Dim dicTitulosC001 As New Dictionary
Dim dicTitulosC010 As New Dictionary
Dim dicTitulosC100 As New Dictionary
Dim dicTitulosC101 As New Dictionary
Dim dicTitulosC120 As New Dictionary
Dim dicTitulosC140 As New Dictionary
Dim dicTitulosC141 As New Dictionary
Dim dicTitulosC170 As New Dictionary
Dim dicTitulosC175Contr As New Dictionary
Dim dicTitulosC190 As New Dictionary
Dim dicTitulosC191 As New Dictionary
Dim dicTitulosCorrelacoes As New Dictionary
Dim dicDados0000 As New Dictionary
Dim dicDados0001 As New Dictionary
Dim dicDados0005 As New Dictionary
Dim dicDados0100 As New Dictionary
Dim dicDados0110 As New Dictionary
Dim dicDados0140 As New Dictionary
Dim dicDados0150 As New Dictionary
Dim dicDados0190 As New Dictionary
Dim dicDados0200 As New Dictionary
Dim dicDados0220 As New Dictionary
Dim dicDadosC001 As New Dictionary
Dim dicDadosC010 As New Dictionary
Dim dicDadosC100 As New Dictionary
Dim dicDadosC101 As New Dictionary
Dim dicDadosC120 As New Dictionary
Dim dicDadosC140 As New Dictionary
Dim dicDadosC141 As New Dictionary
Dim dicDadosC170 As New Dictionary
Dim dicDadosC175Contr As New Dictionary
Dim dicDadosC190 As New Dictionary
Dim dicDadosC191 As New Dictionary
Dim arrImportacoes As New ArrayList
Dim Produtos As IXMLDOMNodeList
Dim NFe As New DOMDocument60
Dim XML As Variant, Chave, Campos
Dim Comeco As Double
Dim i As Integer
Dim b As Long
    
    Call ProcessarDocumentos.CarregarXMLS(TipoImportacao)
    If TipoImportacao = "Arquivo" Then DocsFiscais.arrNFeNFCe.addRange DocsFiscais.arrTodos
    
    If DocsFiscais.arrNFeNFCe.Count > 0 Then
        
        Call Util.AtualizarBarraStatus("Iniciando importação dos XMLs...")
        
        'Reinicia a contagem de documentos sem validade jurídica
        DocsSemValidade = 0
        CNPJBase = VBA.Left(CNPJContribuinte, 8)
        
        'Inicia a importação dos protocolos de cancelamento
        Set dicCorrelacoes = Util.CriarDicionarioCorrelacoes(Correlacoes)
        Set dicTitulosCorrelacoes = Util.MapearTitulos(Correlacoes, 3)
        
        'Verifica tipo de SPED selecionado pelo usuário e carrega dados do cabeçalho
        If SPEDContrib Then
            
            Set dicDados0000 = Util.CriarDicionarioRegistro(reg0000_Contr, "ARQUIVO")
            Set dicTitulos0000 = Util.MapearTitulos(reg0000_Contr, 3)
            
            Set dicDados0110 = Util.CriarDicionarioRegistro(reg0110, "ARQUIVO")
            Set dicDados0140 = Util.CriarDicionarioRegistro(reg0140, "CNPJ")
            Set dicDadosC010 = Util.CriarDicionarioRegistro(regC010, "CNPJ")
            Set dicDadosC175Contr = Util.CriarDicionarioRegistro(regC175_Contr)
            
            Set dicTitulos0110 = Util.MapearTitulos(reg0110, 3)
            Set dicTitulos0140 = Util.MapearTitulos(reg0140, 3)
            Set dicTitulosC010 = Util.MapearTitulos(regC010, 3)
            Set dicTitulosC175Contr = Util.MapearTitulos(regC175_Contr, 3)
            
        Else
            
            Set dicDados0000 = Util.CriarDicionarioRegistro(reg0000, "ARQUIVO")
            Set dicTitulos0000 = Util.MapearTitulos(reg0000, 3)
            
            Set dicDados0005 = Util.CriarDicionarioRegistro(reg0005, "ARQUIVO")
            Set dicDadosC101 = Util.CriarDicionarioRegistro(regC101)
            
        End If
        
        'Carrega dados dos registros do bloco 0
        Set dicDados0001 = Util.CriarDicionarioRegistro(reg0001, "ARQUIVO")
        Set dicDados0150 = Util.CriarDicionarioRegistro(reg0150, "CNPJ", "CPF")
        Set dicDados0190 = Util.CriarDicionarioRegistro(reg0190)
        Set dicDados0200 = Util.CriarDicionarioRegistro(reg0200)
        Set dicDados0220 = Util.CriarDicionarioRegistro(reg0220)
        
        'Carrega dados dos registros do bloco C
        Set dicDadosC001 = Util.CriarDicionarioRegistro(regC001, "ARQUIVO")
        Set dicDadosC100 = Util.CriarDicionarioRegistro(regC100, "IND_OPER", "IND_EMIT", "CHV_NFE")
        Set dicDadosC120 = Util.CriarDicionarioRegistro(regC120)
        Set dicDadosC140 = Util.CriarDicionarioRegistro(regC140)
        Set dicDadosC141 = Util.CriarDicionarioRegistro(regC141)
        Set dicDadosC170 = Util.CriarDicionarioRegistro(regC170)
        
        'Carrega índices dos registros do bloco 0
        Set dicTitulos0140 = Util.MapearTitulos(reg0140, 3)
        Set dicTitulos0150 = Util.MapearTitulos(reg0150, 3)
        Set dicTitulos0190 = Util.MapearTitulos(reg0190, 3)
        Set dicTitulos0200 = Util.MapearTitulos(reg0200, 3)
        Set dicTitulos0220 = Util.MapearTitulos(reg0220, 3)
        
        'Carrega dados dos registros do bloco C
        Set dicTitulosC001 = Util.MapearTitulos(regC001, 3)
        Set dicTitulosC100 = Util.MapearTitulos(regC100, 3)
        Set dicTitulosC101 = Util.MapearTitulos(regC101, 3)
        Set dicTitulosC120 = Util.MapearTitulos(regC120, 3)
        Set dicTitulosC140 = Util.MapearTitulos(regC140, 3)
        Set dicTitulosC141 = Util.MapearTitulos(regC141, 3)
        Set dicTitulosC170 = Util.MapearTitulos(regC170, 3)
        Set dicTitulosC190 = Util.MapearTitulos(regC190, 3)
        Set dicTitulosC191 = Util.MapearTitulos(regC191, 3)
                
        b = 0
        Comeco = Timer
        For Each XML In DocsFiscais.arrNFeNFCe
            
            Call Util.AntiTravamento(b, 100, "Importando XML " & b + 1 & " de " & DocsFiscais.arrNFeNFCe.Count, DocsFiscais.arrNFeNFCe.Count, Comeco)
            Set NFe = fnXML.RemoverNamespaces(XML)
            
            If Not fnXML.ValidarXML(NFe) Then GoTo Prx:
            
            If fnXML.ValidarNFe(NFe) And fnXML.ValidarParticipante(NFe, SPEDContrib) Then
                
                With CamposC100
                    
                    .IND_OPER = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_OPER(fnXML.IdentificarTipoOperacao(NFe, ValidarTag(NFe, "//tpNF"), SPEDContrib))
                    .IND_EMIT = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_EMIT(fnXML.IdentificarTipoEmissao(NFe, SPEDContrib))
                    .COD_MOD = ValidarTag(NFe, "//mod")
                    .COD_PART = fnXML.IdentificarParticipante(NFe, VBA.Left(.IND_OPER, 1), VBA.Left(.IND_EMIT, 1))
                    .COD_SIT = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_COD_SIT(fnSPED.GerarCodigoSituacao(fnXML.ValidarSituacao(ValidarTag(NFe, "//cStat"))))
                    .SER = VBA.Format(ValidarTag(NFe, "//serie"), "000")
                    .NUM_DOC = ValidarTag(NFe, "//nNF")
                    .CHV_NFE = VBA.Right(ValidarTag(NFe, "//@Id"), 44)
                    
                    CNPJEmit = fnXML.ExtrairCNPJEmitente(NFe)
                    CNPJDest = fnXML.ExtrairCNPJDestinatario(NFe)
                    Periodo = Util.ExtrairPeriodo(fnXML.ExtrairDataDocumento(NFe))
                    
                    If UsarPeriodo And PeriodoEspecifico <> "" Then Periodo = VBA.Format(PeriodoEspecifico, "00/0000")
                    tpCont = fnXML.DefinirContribuinteNFe(NFe, SPEDContrib)
                    CNPJCont = ValidarTag(NFe, "//" & tpCont & "/CNPJ")
                    .ARQUIVO = Periodo & "-" & CNPJCont
                    
                    If SPEDContrib Then .ARQUIVO = Periodo & "-" & CNPJContribuinte
                    If Not dicDados0000.Exists(.ARQUIVO) Then
                        If SPEDContrib Then
                            Call fnXML.CriarRegistro0000_Contr(NFe, dicDados0000, dicDados0001, dicDados0110, Periodo, tpCont)
                        Else
                            Call fnXML.CriarRegistro0000(NFe, dicDados0000, dicDados0001, dicDados0005, dicDados0100, dicDadosC001, Periodo, tpCont)
                        End If
                    End If
                    
                    If dicDados0000.Exists(.ARQUIVO) Then
                        
                        Campos = dicDados0000(.ARQUIVO)
                        
                        If LBound(Campos) = 0 Then i = 1 Else i = 0
                        
                        CHV_0000 = Util.RemoverAspaSimples(Campos(dicTitulos0000("CHV_REG") - i))
                        CHV_0001 = fnSPED.GerarChaveRegistro(CHV_0000, "0001")
                        CHV_C001 = fnSPED.GerarChaveRegistro(CHV_0000, "C001")
                        
                    End If
                    
                    If .COD_MOD <> "65" Then
                        
                        tpPart = fnXML.DefinirParticipanteNFe(NFe, SPEDContrib)
                        If dicDados0150.Exists(.COD_PART) Then
                            CHV_0150 = VBA.Trim(.COD_PART)
                        End If
                        
                    Else
                        
                        .COD_PART = ""
                        
                    End If
                    
                    If SPEDContrib Then
                        Call fnXML.CriarRegistro0140(NFe, dicDados0140, dicTitulos0140, .ARQUIVO, CHV_0001, tpCont)
                        Call fnXML.CriarRegistroC010(NFe, dicDadosC010, dicTitulosC010, .ARQUIVO, CHV_C001, tpCont)
                    End If
                    
                    CHV_PAI = Util.SelecionarChaveSPED(CHV_C001, CNPJBase, CNPJEmit, CNPJDest, SPEDContrib)
                    Chave = VBA.Replace(VBA.Join(Array(.IND_OPER, .IND_EMIT, .CHV_NFE)), " ", "")
                    If dicDadosC100.Exists(Chave) Then
                        
                        .CHV_REG = dicDadosC100(Chave)(dicTitulosC100("CHV_REG"))
                        .ARQUIVO = dicDadosC100(Chave)(dicTitulosC100("ARQUIVO"))
                        .COD_PART = dicDadosC100(Chave)(dicTitulosC100("COD_PART"))
                        .IND_OPER = dicDadosC100(Chave)(dicTitulosC100("IND_OPER"))
                        .IND_EMIT = dicDadosC100(Chave)(dicTitulosC100("IND_EMIT"))
                        .COD_SIT = dicDadosC100(Chave)(dicTitulosC100("COD_SIT"))
                        
                    Else
                        
                        .COD_PART = fnXML.IdentificarParticipante(NFe, VBA.Left(.IND_OPER, 1), VBA.Left(.IND_EMIT, 1))
                        If dicDados0150.Exists(.COD_PART) Then
                            .COD_PART = dicDados0150(.COD_PART)(dicTitulos0150("COD_PART"))
                        End If
                        Call fnXML.CriarRegistroC100(NFe, dicDadosC100, DocsFiscais.arrChavesCanceladas, CHV_PAI, SPEDContrib)
                        
                    End If
                                        
                    If arrImportacoes.contains(Chave) Then GoTo Prx:
                    arrImportacoes.Add Chave
                    
                    Set Produtos = NFe.SelectNodes("//det")
                    
                    CHV_PAI = Util.SelecionarChaveSPED(CHV_0001, CNPJBase, CNPJEmit, CNPJDest, SPEDContrib)
                    If .COD_MOD <> "65" Then Call fnXML.CriarRegistro0150(NFe, dicDados0150, dicTitulos0150, .COD_PART, .ARQUIVO, CHV_PAI, tpPart)
                    
                    CHV_PAI = Util.SelecionarChaveSPED(CHV_C001, CNPJBase, CNPJEmit, CNPJDest, SPEDContrib)
                    Call fnXML.CriarRegistroC101(NFe, dicDadosC101, .ARQUIVO)
                    Call fnXML.CriarRegistroC120(Produtos, dicDadosC120, .ARQUIVO)
                    Call fnXML.CriarRegistroC140(NFe, dicDadosC140, dicDadosC141, .ARQUIVO)
                    Call fnXML.CriarRegistroC170(Produtos, VL_PIS_ST, VL_COFINS_ST, CNPJEmit, CNPJDest, dicDadosC170, .ARQUIVO, dicCorrelacoes, dicTitulosCorrelacoes, dicDados0000, dicTitulos0000, dicDados0190, dicDados0200, dicDados0220, dicTitulos0220, SPEDContrib)
                    Call fnXML.CriarRegistroC190(Produtos, dicTitulosC190, dicDadosC190, dicTitulosC191, dicDadosC191)
                    
                End With
                
            End If
Prx:
        Next XML
        
        Application.StatusBar = "Importando dados dos XMLS para o SPED, por favor aguarde..."
        
        'Exporta dados coletados dos registros
        If SPEDContrib Then
            Call Util.ExportarDadosDicionario(reg0000_Contr, dicDados0000, "A4")
        Else
            Call Util.ExportarDadosDicionario(reg0000, dicDados0000, "A4")
        End If
        
        Call Util.ExportarDadosDicionario(reg0001, dicDados0001, "A4")
        Call Util.ExportarDadosDicionario(reg0005, dicDados0005, "A4")
        Call Util.ExportarDadosDicionario(reg0100, dicDados0100, "A4")
        Call Util.ExportarDadosDicionario(reg0110, dicDados0110, "A4")
        Call Util.ExportarDadosDicionario(reg0140, dicDados0140, "A4")
        Call Util.ExportarDadosDicionario(reg0150, dicDados0150, "A4")
        Call Util.ExportarDadosDicionario(reg0190, dicDados0190, "A4")
        Call Util.ExportarDadosDicionario(reg0200, dicDados0200, "A4")
        Call Util.ExportarDadosDicionario(reg0220, dicDados0220, "A4")
        Call Util.ExportarDadosDicionario(regC001, dicDadosC001, "A4")
        Call Util.ExportarDadosDicionario(regC010, dicDadosC010, "A4")
        Call Util.ExportarDadosDicionario(regC100, dicDadosC100, "A4")
        Call Util.ExportarDadosDicionario(regC101, dicDadosC101, "A4")
        Call Util.ExportarDadosDicionario(regC120, dicDadosC120, "A4")
        Call Util.ExportarDadosDicionario(regC140, dicDadosC140, "A4")
        Call Util.ExportarDadosDicionario(regC141, dicDadosC141, "A4")
        Call Util.ExportarDadosDicionario(regC170, dicDadosC170, "A4")
        Call Util.ExportarDadosDicionario(regC190, dicDadosC190, "A4")
        Call Util.ExportarDadosDicionario(regC191, dicDadosC191, "A4")
        
        Call FuncoesSPEDFiscal.ZerarDicionariosEFD
        
        Call Util.AtualizarBarraStatus("Importação finalizada com sucesso!")
        If DocsSemValidade > 0 Then Call Util.MsgAlerta("Foram encontrados " & DocsSemValidade & " documento(s) sem validade jurídica", "Documentos sem validade jurídica")
        Call Util.MsgInformativa("Importação concluída com sucesso!", "Importação NF-e/NFC-e", Inicio)
        
    End If
    
End Sub

Public Sub ImportarXMLsA100(ByVal TipoImportacao As String)

Dim ARQUIVO As String, Periodo$, Caminho$, CNPJToma$, CNPJPrest$, CHV_PAI$, CHV_0000$, CHV_0001$, CHV_0110$, CHV_0140$, CHV_0150$, CHV_A001$, CHV_A010$, CHV_REG$, Msg$, tpCont$
Dim dicTitulos0000 As New Dictionary
Dim dicTitulos0001 As New Dictionary
Dim dicTitulos0100 As New Dictionary
Dim dicTitulos0110 As New Dictionary
Dim dicTitulos0140 As New Dictionary
Dim dicTitulos0150 As New Dictionary
Dim dicTitulos0200 As New Dictionary
Dim dicTitulosA001 As New Dictionary
Dim dicTitulosA010 As New Dictionary
Dim dicTitulosA100 As New Dictionary
Dim dicTitulosA101 As New Dictionary
Dim dicTitulosA170 As New Dictionary
Dim dicDados0000 As New Dictionary
Dim dicDados0001 As New Dictionary
Dim dicDados0100 As New Dictionary
Dim dicDados0110 As New Dictionary
Dim dicDados0140 As New Dictionary
Dim dicDados0150 As New Dictionary
Dim dicDados0200 As New Dictionary
Dim dicDadosA001 As New Dictionary
Dim dicDadosA010 As New Dictionary
Dim dicDadosA100 As New Dictionary
Dim dicDadosA101 As New Dictionary
Dim dicDadosA170 As New Dictionary
Dim Notas As IXMLDOMNodeList
Dim NFSe As New DOMDocument60
Dim Nota As IXMLDOMNode
Dim arrXMLs As New ArrayList
Dim XML As Variant, Campos
Dim Comeco As Double
Dim i As Integer
Dim b As Long
    
    'Verifica o tipo de importação selecionada pelo usuário
    If TipoImportacao = "Arquivo" Then Call Util.GuardarEnderecosArrayList("xml", arrXMLs)
    
    If TipoImportacao = "Lote" Then
        Caminho = Util.SelecionarPasta("Selecione a pasta que contém os arquivos XML")
        If Caminho = "" Then Exit Sub
        Inicio = Now()
        Call Util.ListarArquivos(arrXMLs, Caminho)
    End If
    
    'Verifica se a lista contém XMLs para processamento
    If arrXMLs.Count > 0 Then
        
        CNPJBase = VBA.Left(CNPJContribuinte, 8)
        
        'Carrega títulos dos registros do bloco 0
        Set dicTitulos0000 = Util.MapearTitulos(reg0000_Contr, 3)
        Set dicTitulos0001 = Util.MapearTitulos(reg0001, 3)
        Set dicTitulos0110 = Util.MapearTitulos(reg0110, 3)
        Set dicTitulos0140 = Util.MapearTitulos(reg0140, 3)
        Set dicTitulos0150 = Util.MapearTitulos(reg0150, 3)
        Set dicTitulos0200 = Util.MapearTitulos(reg0200, 3)
        
        'Carrega dados dos registros do bloco 0
        Set dicDados0000 = Util.CriarDicionarioRegistro(reg0000_Contr, "ARQUIVO")
        Set dicDados0001 = Util.CriarDicionarioRegistro(reg0001, "ARQUIVO")
        Set dicDados0110 = Util.CriarDicionarioRegistro(reg0110, "ARQUIVO")
        Set dicDados0140 = Util.CriarDicionarioRegistro(reg0140, "ARQUIVO", "CNPJ")
        Set dicDados0150 = Util.CriarDicionarioRegistro(reg0150, "CNPJ", "CPF")
        Set dicDados0200 = Util.CriarDicionarioRegistro(reg0200)
        
        'Carrega títulos dos registros do bloco A
        Set dicTitulosA001 = Util.MapearTitulos(regA001, 3)
        Set dicTitulosA010 = Util.MapearTitulos(regA010, 3)
        Set dicTitulosA100 = Util.MapearTitulos(regA100, 3)
        Set dicTitulosA170 = Util.MapearTitulos(regA170, 3)
        
        'Carrega dados dos registros do bloco A
        Set dicDadosA001 = Util.CriarDicionarioRegistro(regA001, "ARQUIVO")
        Set dicDadosA010 = Util.CriarDicionarioRegistro(regA010, "CNPJ")
        Set dicDadosA100 = Util.CriarDicionarioRegistro(regA100, "IND_OPER", "IND_EMIT", "CHV_NFSE")
        Set dicDadosA170 = Util.CriarDicionarioRegistro(regA170)
        
        b = 0
        Comeco = Timer
        For Each XML In arrXMLs
            
            Call Util.AntiTravamento(b, 100, "Etapa 3/3 [Importação da NFSe] - Importando XML " & b + 1 & " de " & arrXMLs.Count, arrXMLs.Count, Comeco)
            Set NFSe = fnXML.RemoverNamespaces(XML)
                                       
            If Not fnXML.ValidarXML(NFSe) Then GoTo Prx:
                                       
            If fnXML.ValidarNFSe(NFSe) Then
                
                Set Notas = NFSe.SelectNodes("//CompNfse")
                For Each Nota In Notas
                    
                    With CamposA100
                        
                        Periodo = VBA.Format(VBA.Left(Nota.SelectSingleNode("Nfse//Competencia").text, 10), "mm/yyyy")
                        If UsarPeriodo And PeriodoEspecifico <> "" Then Periodo = VBA.Format(PeriodoEspecifico, "00/0000")
                        
                        If Not fnNFSe.ExtrairCNPJContribunte(Nota) Like CNPJBase & "*" Then GoTo Prx:
                        
                        .ARQUIVO = Periodo & "-" & CNPJContribuinte
                        If Not dicDados0000.Exists(.ARQUIVO) Then Call fnNFSe.CriarRegistro0000(Nota, dicDados0000, Periodo)
                        
                        If dicDados0000.Exists(.ARQUIVO) Then
                            
                            Campos = dicDados0000(.ARQUIVO)
                            If LBound(Campos) = 0 Then i = 1 Else i = 0
                            
                            CHV_0000 = Util.RemoverAspaSimples(Campos(dicTitulos0000("CHV_REG") - i))
                            
                            'Cria registro 0001
                            CHV_0001 = fnSPED.GerarChaveRegistro(CHV_0000, "0001")
                            dicDados0001(.ARQUIVO) = Array("'0001", .ARQUIVO, CHV_0001, CHV_0000, 0)
                            
                            'Cria registro 0110
                            CHV_0110 = fnSPED.GerarChaveRegistro(CHV_0000, "0110")
                            dicDados0110(.ARQUIVO) = Array("'0110", .ARQUIVO, CHV_0110, CHV_0001, "", "", "", "")
                            
                            'Cria registro A001
                            CHV_A001 = fnSPED.GerarChaveRegistro(CHV_0000, "A001")
                            dicDadosA001(.ARQUIVO) = Array("A001", .ARQUIVO, CHV_A001, CHV_0000, 0)
                            
                            CHV_A010 = fnSPED.GerarChaveRegistro(CHV_0000, fnNFSe.ExtrairCNPJContribunte(Nota))
                            dicDadosA010(.ARQUIVO) = Array("A010", .ARQUIVO, CHV_A010, CHV_A001, "'" & fnNFSe.ExtrairCNPJContribunte(Nota))
                            
                        End If
                        
                        Call fnNFSe.CriarRegistro0140(Nota, dicDados0140, .ARQUIVO, CHV_0001)
                        
                        Call fnNFSe.DefinirDadosA100(Nota)
                        .CHV_NFSE = VBA.UCase(ValidarTag(Nota, "Nfse//CodigoVerificacao"))
                        
                        CHV_REG = VBA.Join(Array(.IND_OPER, .IND_EMIT, .CHV_NFSE))
                        If Not dicDadosA100.Exists(CHV_REG) Then _
                            Call fnNFSe.CriarRegistroA100(Nota, dicDadosA100, dicDadosA170, dicTitulosA170, CHV_A010)
                        
                        CHV_0140 = VBA.Join(Array(.ARQUIVO, CNPJContribuinte))
                        If dicDados0140.Exists(CHV_0140) Then
                            
                            Campos = dicDados0140(CHV_0140)
                            If LBound(Campos) = 0 Then i = 1 Else i = 0
                            CHV_0140 = Util.RemoverAspaSimples(Campos(dicTitulos0140("CHV_REG") - i))
                            
                        End If
                        
                        Call fnNFSe.CriarRegistro0150(Nota, dicDados0150, .ARQUIVO, CHV_0140)
                        Call fnNFSe.CriarRegistro0200(Nota, dicDados0200, .ARQUIVO, CHV_0140)
                        
                    End With
                
                Next Nota
                
            End If
Prx:
        Next XML
        
        Application.StatusBar = "Importando dados dos XMLS para o SPED, por favor aguarde..."
        
        'Exporta dados coletados dos registros
        Call Util.ExportarDadosDicionario(reg0000_Contr, dicDados0000, "A4")
        Call Util.ExportarDadosDicionario(reg0001, dicDados0001, "A4")
        Call Util.ExportarDadosDicionario(reg0100, dicDados0100, "A4")
        Call Util.ExportarDadosDicionario(reg0110, dicDados0110, "A4")
        Call Util.ExportarDadosDicionario(reg0140, dicDados0140, "A4")
        Call Util.ExportarDadosDicionario(reg0150, dicDados0150, "A4")
        Call Util.ExportarDadosDicionario(reg0200, dicDados0200, "A4")
        Call Util.ExportarDadosDicionario(regA001, dicDadosA001, "A4")
        Call Util.ExportarDadosDicionario(regA010, dicDadosA010, "A4")
        Call Util.ExportarDadosDicionario(regA100, dicDadosA100, "A4")
        Call Util.ExportarDadosDicionario(regA170, dicDadosA170, "A4")
        
        Call FuncoesSPEDFiscal.ZerarDicionariosEFD
        
        Application.StatusBar = "Importação finalizada com sucesso!"
        Call Util.MsgInformativa("Importação concluída com sucesso!", "Importação NFSe", Inicio)
        
    End If
    
End Sub

Public Sub ImportarRegistrosCTeXML(ByVal TipoImportacao As String, Optional ByVal SPEDContrib As Boolean)

Dim Caminho As String, Periodo$, CHV_0000$, CHV_0001$, CHV_0005$, CHV_0100$, CHV_0140$, CHV_0150$, CHV_D001$, CHV_D010$, tpPart$, tpCont$, CHV_0110$, CHV_PAI$, CNPJEmit$, CNPJToma$
Dim dicTitulos0000 As New Dictionary
Dim dicTitulos0001 As New Dictionary
Dim dicTitulos0005 As New Dictionary
Dim dicTitulos0100 As New Dictionary
Dim dicTitulos0110 As New Dictionary
Dim dicTitulos0140 As New Dictionary
Dim dicTitulos0150 As New Dictionary
Dim dicTitulosC001 As New Dictionary
Dim dicTitulosD001 As New Dictionary
Dim dicTitulosD010 As New Dictionary
Dim dicTitulosD100 As New Dictionary
Dim dicTitulosD101Contr As New Dictionary
Dim dicTitulosD105 As New Dictionary
Dim dicTitulosD190 As New Dictionary
Dim dicTitulosD200 As New Dictionary
Dim dicTitulosD201 As New Dictionary
Dim dicTitulosD205 As New Dictionary
Dim dicDados0000 As New Dictionary
Dim dicDados0001 As New Dictionary
Dim dicDados0005 As New Dictionary
Dim dicDados0100 As New Dictionary
Dim dicDados0110 As New Dictionary
Dim dicDados0140 As New Dictionary
Dim dicDados0150 As New Dictionary
Dim dicDadosC001 As New Dictionary
Dim dicDadosD001 As New Dictionary
Dim dicDadosD010 As New Dictionary
Dim dicDadosD100 As New Dictionary
Dim dicDadosD101Contr As New Dictionary
Dim dicDadosD105 As New Dictionary
Dim dicDadosD190 As New Dictionary
Dim dicDadosD200 As New Dictionary
Dim dicDadosD201 As New Dictionary
Dim dicDadosD205 As New Dictionary
Dim arrCanceladas As New ArrayList
Dim dicCorrelacoesCTeNFe As New Dictionary
Dim Produtos As IXMLDOMNodeList
Dim arrXMLs As New ArrayList
Dim CTe As New DOMDocument60
Dim XML As Variant, Chave, Campos
Dim Comeco As Double
Dim a As Long, i&

    Set dicDados0000 = Util.CriarDicionarioRegistro(reg0000, "ARQUIVO")
    
    If TipoImportacao = "Lote" Then
        Caminho = Util.SelecionarPasta("Selecione a pasta que contém os arquivos XML")
        If Caminho = "" Then Exit Sub
        Inicio = Now()
        Call Util.ListarArquivos(arrXMLs, Caminho)
    End If
    
    If TipoImportacao = "Arquivo" Then Call Util.GuardarEnderecosArrayList("xml", arrXMLs)
    
    If arrXMLs.Count > 0 Then
        
        Application.StatusBar = "Importando protocolos de cancelamento..."
        Call fnXML.CarregarProtocolosCancelamento(arrXMLs, arrCanceladas)
        
        If SPEDContrib Then
            
            Set dicDados0000 = Util.CriarDicionarioRegistro(reg0000_Contr, "ARQUIVO")
            Set dicDados0110 = Util.CriarDicionarioRegistro(reg0110, "ARQUIVO")
            Set dicDados0140 = Util.CriarDicionarioRegistro(reg0140, "CNPJ")
            Set dicDadosD010 = Util.CriarDicionarioRegistro(regD010, "CNPJ")
            Set dicDadosD101Contr = Util.CriarDicionarioRegistro(regD101_Contr)
            Set dicDadosD105 = Util.CriarDicionarioRegistro(regD105)
            Set dicDadosD200 = Util.CriarDicionarioRegistro(regD200, "COD_MOD", "SER", "DT_REF")
            Set dicDadosD201 = Util.CriarDicionarioRegistro(regD201)
            Set dicDadosD205 = Util.CriarDicionarioRegistro(regD205)
            
            Set dicTitulos0000 = Util.MapearTitulos(reg0000_Contr, 3)
            Set dicTitulos0110 = Util.MapearTitulos(reg0110, 3)
            Set dicTitulos0140 = Util.MapearTitulos(reg0140, 3)
            Set dicTitulosD010 = Util.MapearTitulos(regD010, 3)
            Set dicTitulosD101Contr = Util.MapearTitulos(regD101_Contr, 3)
            Set dicTitulosD105 = Util.MapearTitulos(regD105, 3)
            Set dicTitulosD200 = Util.MapearTitulos(regD200, 3)
            Set dicTitulosD201 = Util.MapearTitulos(regD201, 3)
            Set dicTitulosD205 = Util.MapearTitulos(regD205, 3)
            
        Else
            
            Set dicDados0000 = Util.CriarDicionarioRegistro(reg0000, "ARQUIVO")
            Set dicDados0005 = Util.CriarDicionarioRegistro(reg0005, "ARQUIVO")
            Set dicDadosD190 = Util.CriarDicionarioRegistro(regD190)
            
            Set dicTitulos0000 = Util.MapearTitulos(reg0000, 3)
            Set dicTitulos0005 = Util.MapearTitulos(reg0005, 3)
            Set dicTitulosD190 = Util.MapearTitulos(regD190, 3)
            
        End If
        
        Set dicDados0100 = Util.CriarDicionarioRegistro(reg0100, "ARQUIVO")
        Set dicDados0150 = Util.CriarDicionarioRegistro(reg0150)
        Set dicDadosC001 = Util.CriarDicionarioRegistro(regC001, "ARQUIVO")
        Set dicDadosD001 = Util.CriarDicionarioRegistro(regD001, "ARQUIVO")
        Set dicDadosD100 = Util.CriarDicionarioRegistro(regD100, "IND_OPER", "IND_EMIT", "CHV_CTE")
        
        Set dicTitulos0100 = Util.MapearTitulos(reg0100, 3)
        Set dicTitulos0150 = Util.MapearTitulos(reg0150, 3)
        Set dicTitulosC001 = Util.MapearTitulos(regC001, 3)
        Set dicTitulosD001 = Util.MapearTitulos(regD001, 3)
        Set dicTitulosD100 = Util.MapearTitulos(regD100, 3)
        
        a = 0
        Comeco = Timer
        For Each XML In arrXMLs
            
            Call Util.AntiTravamento(a, 100, "Importando XML " & a + 1 & " de " & arrXMLs.Count, arrXMLs.Count, Comeco)
            Set CTe = fnXML.RemoverNamespaces(XML)
            
            If Not fnXML.ValidarXML(CTe) Then GoTo Prx:
            Call CorrelacionarCTeNFe(CTe, dicCorrelacoesCTeNFe)
            
            If fnXML.ValidarXMLCTe(CTe) And fnXML.ValidarParticipanteCTe(CTe) Then
                
                
                
                CNPJContribuinte = Util.FormatarCNPJ(CadContrib.Range("CNPJContribuinte").value)
                CNPJBase = VBA.Left(CNPJContribuinte, 8)
                
                Periodo = VBA.Format(VBA.Left(CTe.SelectSingleNode("//dhEmi").text, 10), "mm/yyyy")
                tpCont = fnXML.DefinirContribuinteCTe(CTe, SPEDContrib)
                ARQUIVO = Periodo & "-" & CNPJContribuinte
                
                If Not dicDados0000.Exists(ARQUIVO) Then
                    
                    If SPEDContrib Then
                        
                        Call fnXML.CriarRegistro0000_Contr(CTe, dicDados0000, dicDados0001, dicDados0110, Periodo, tpCont)
                    
                    Else
                        
                        Call fnXML.CriarRegistro0000(CTe, dicDados0000, dicDados0001, dicDados0005, dicDados0100, dicDadosC001, Periodo, tpCont)
                    
                    End If
                    
                    
                End If
                
                If dicDados0000.Exists(ARQUIVO) Then
                    
                    Campos = dicDados0000(ARQUIVO)
                    If LBound(Campos) = 0 Then i = 1 Else i = 0
                    
                    CHV_0000 = Util.RemoverAspaSimples(Campos(dicTitulos0000("CHV_REG") - i))
                    
                    'Cria registro 0001
                    CHV_0001 = fnSPED.GerarChaveRegistro(CHV_0000, "0001")
                    dicDados0001(ARQUIVO) = Array("'0001", ARQUIVO, CHV_0001, CHV_0000, 0)
                    
                    'Cria registro 0100
                    CHV_0100 = fnSPED.GerarChaveRegistro(CHV_0001, "0100")
                    dicDados0100(ARQUIVO) = Array("'0100", ARQUIVO, CHV_0100, CHV_0001, "", "", "", "", "", "", "", "", "", "", "", "", "")
                    
                    'Cria registro 0001
                    CHV_0001 = fnSPED.GerarChaveRegistro(CHV_0000, "0001")
                    dicDados0001(ARQUIVO) = Array("'0001", ARQUIVO, CHV_0001, CHV_0000, 0)
                    
                    'Cria registro D001
                    CHV_D001 = fnSPED.GerarChaveRegistro(CHV_0000, "D001")
                    dicDadosD001(ARQUIVO) = Array("D001", ARQUIVO, CHV_D001, CHV_0000, 0)
                    
                    If SPEDContrib Then
                        
                        'Cria registro 0110
                        CHV_0110 = fnSPED.GerarChaveRegistro(CHV_0000, "0110")
                        dicDados0110(ARQUIVO) = Array("'0110", ARQUIVO, CHV_0110, CHV_0001, "", "", "", "")
                        
                        'Cria registro 0140
                        CHV_0140 = fnSPED.GerarChaveRegistro(CHV_0000, "0140")
                        Call fnXML.CriarRegistro0140(CTe, dicDados0140, dicTitulos0140, ARQUIVO, CHV_0001, tpCont)
                        
                        'Cria registro D010
                        CHV_D010 = fnSPED.GerarChaveRegistro(CHV_D001, CStr(fnXML.ExtrairCNPJContribuinte(CTe, tpCont)))
                        Call fnXML.CriarRegistroD010(CTe, dicDadosD010, dicTitulosD010, ARQUIVO, CHV_D001, tpCont)
                        
                    End If
                    
                End If
                
                With CamposD100
                    
                    .IND_OPER = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_D100_IND_OPER(fnXML.IdentificarTipoOperacaoCTe(CTe, ValidarTag(CTe, "//tpCTe")))
                    .IND_EMIT = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_EMIT(fnXML.IdentificarTipoEmissaoCTe(CTe))
                    .COD_PART = fnXML.IdentificarParticipanteCTe(CTe)
                    .COD_MOD = ValidarTag(CTe, "//mod")
                    .SER = VBA.Format(ValidarTag(CTe, "//serie"), "000")
                    .SUB = ValidarTag(CTe, "//subserie")
                    .NUM_DOC = ValidarTag(CTe, "//nCT")
                    .CHV_CTE = VBA.Right(ValidarTag(CTe, "//@Id"), 44)
                    
                    'Trata oerações de entrada de CTe para SPED Contribuições e lançamentos gerais para o SPED Fiscal
                    If (.IND_OPER Like "0*" And SPEDContrib) Or Not SPEDContrib Or ImportarCTeD100 Then
                        
                        Chave = VBA.Join(Array(.IND_OPER, .IND_EMIT, .CHV_CTE))
                        If dicDadosD100.Exists(Chave) Then
                            
                            .ARQUIVO = dicDadosD100(Chave)(dicTitulosD100("ARQUIVO"))
                            If dicDados0000.Exists(.ARQUIVO) Then Campos0000.UF = dicDados0000(.ARQUIVO)(dicTitulos0000("UF"))
                            .COD_PART = dicDadosD100(Chave)(dicTitulosD100("COD_PART"))
                            .IND_OPER = dicDadosD100(Chave)(dicTitulosD100("IND_OPER"))
                            .IND_EMIT = dicDadosD100(Chave)(dicTitulosD100("IND_EMIT"))
                            .CHV_REG = dicDadosD100(Chave)(dicTitulosD100("CHV_REG"))
                            
                        Else
                            
                            .COD_PART = fnXML.DefinirParticipanteCTe(CTe)
                            CNPJContribuinte = Util.FormatarCNPJ(CadContrib.Range("CNPJContribuinte").value)
                            CNPJBase = VBA.Left(CNPJContribuinte, 8)
                            Periodo = VBA.Format(VBA.Left(CTe.SelectSingleNode("//dhEmi").text, 10), "mm/yyyy")
                            If UsarPeriodo And PeriodoEspecifico <> "" Then Periodo = VBA.Format(PeriodoEspecifico, "00/0000")
                            .ARQUIVO = Periodo & "-" & CNPJContribuinte
                            If dicDados0000.Exists(.ARQUIVO) Then Campos0000.CHV_REG = dicDados0000(.ARQUIVO)(dicTitulos0000("CHV_REG"))
                            If dicDados0000.Exists(.ARQUIVO) Then Campos0000.UF = dicDados0000(.ARQUIVO)(dicTitulos0000("UF"))
                            
                            CNPJEmit = fnXML.ExtrairCNPJEmitente(CTe)
                            CNPJToma = fnXML.ExtrairCNPJTomador(CTe)
                            
                            CHV_PAI = Util.SelecionarChaveSPED(CHV_D001, CNPJBase, CNPJEmit, CNPJToma, SPEDContrib)
                            Call fnXML.CriarRegistroD100(CTe, dicDadosD100, dicDadosD101Contr, dicDadosD105, dicDadosD190, .ARQUIVO, CHV_PAI, SPEDContrib)
                            
                        End If
                    
                    'Trata oerações de saída de CTe para SPED Contribuições
                    ElseIf .IND_OPER Like "1*" And SPEDContrib Then
                            
                        With CamposD200
                            
                            .COD_MOD = ValidarTag(CTe, "//mod")
                            .COD_SIT = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_COD_SIT(fnSPED.GerarCodigoSituacao(fnXML.ValidarSituacao(ValidarTag(CTe, "//cStat"))))
                            .SER = VBA.Format(ValidarTag(CTe, "//serie"), "000")
                            .CFOP = ValidarTag(CTe, "//CFOP")
                            .DT_REF = VBA.Format(VBA.Left(ValidarTag(CTe, "//dhEmi"), 10), "yyyy-mm-dd")
                                                        
                            Call fnXML.CriarRegistroD200(CTe, dicDadosD200, dicTitulosD200, dicDadosD201, dicDadosD205, ARQUIVO, CHV_D010)
                            
                        End With
                        
                    End If
                    
                    tpPart = fnXML.DefinirParticipanteCTe(CTe)
                    If dicDados0150.Exists(.COD_PART) Then
                        CHV_0150 = dicDados0150(.COD_PART)(dicTitulos0150("CHV_REG"))
                    Else
                        CHV_0150 = VBA.Trim(.COD_PART)
                    End If
                    
                    CHV_PAI = Util.SelecionarChaveSPED(CHV_0001, CNPJBase, CNPJEmit, CNPJToma, SPEDContrib)
                    Call fnXML.CriarRegistro0150(CTe, dicDados0150, dicTitulos0150, CHV_0150, ARQUIVO, CHV_PAI, tpPart)
                    
                End With
                
            End If
Prx:
        Next XML
        
        Application.StatusBar = "Importando dados dos XMLS para o SPED, por favor aguarde..."
        
        If SPEDContrib Then
            
            Call Util.ExportarDadosDicionario(reg0000_Contr, dicDados0000, "A4")
            Call Util.ExportarDadosDicionario(reg0001, dicDados0001, "A4")
            Call Util.ExportarDadosDicionario(reg0110, dicDados0110, "A4")
            Call Util.ExportarDadosDicionario(reg0140, dicDados0140, "A4")
            Call Util.ExportarDadosDicionario(regD001, dicDadosD001, "A4")
            Call Util.ExportarDadosDicionario(regD010, dicDadosD010, "A4")
            Call Util.ExportarDadosDicionario(regD101_Contr, dicDadosD101Contr, "A4")
            Call Util.ExportarDadosDicionario(regD105, dicDadosD105, "A4")
            Call Util.ExportarDadosDicionario(regD200, dicDadosD200, "A4")
            Call Util.ExportarDadosDicionario(regD201, dicDadosD201, "A4")
            Call Util.ExportarDadosDicionario(regD205, dicDadosD205, "A4")
            
        Else
            
            Call Util.ExportarDadosDicionario(reg0000, dicDados0000, "A4")
            Call Util.ExportarDadosDicionario(reg0005, dicDados0005, "A4")
            
        End If
        
        Call Util.ExportarDadosDicionario(reg0100, dicDados0100, "A4")
        Call Util.ExportarDadosDicionario(reg0150, dicDados0150, "A4")
        Call Util.ExportarDadosDicionario(regD100, dicDadosD100, "A4")
        Call Util.ExportarDadosDicionario(regD190, dicDadosD190, "A4")
        
        Call Util.ExportarDadosDicionario(CorrelacoesCTeNFe, dicCorrelacoesCTeNFe)
        
        Call FuncoesSPEDFiscal.ZerarDicionariosEFD
        
        Application.StatusBar = "Importação finalizada com sucesso!"
        Call Util.MsgInformativa("Importação concluída com sucesso!", "Importação CT-e", Inicio)
        
    End If
    
End Sub

Public Sub ImportarRegistrosCFeXML(ByVal TipoImportacao As String)

Dim dicCorrelacoes As New Dictionary
Dim dicTitulos0000 As New Dictionary
Dim dicTitulos0190 As New Dictionary
Dim dicTitulos0200 As New Dictionary
Dim dicTitulosC800 As New Dictionary
Dim dicTitulosC810 As New Dictionary
Dim dicTitulosC850 As New Dictionary
Dim dicDados0000 As New Dictionary
Dim dicDados0190 As New Dictionary
Dim dicDados0200 As New Dictionary
Dim dicDadosC800 As New Dictionary
Dim dicDadosC810 As New Dictionary
Dim dicDadosC850 As New Dictionary
Dim Produtos As IXMLDOMNodeList
Dim Periodo As String, Caminho$, Chave$
Dim CFe As New DOMDocument60
Dim XML As Variant
Dim Comeco As Double

    Call ProcessarDocumentos.CarregarXMLS(TipoImportacao)
    
    If DocsFiscais.arrCFe.Count > 0 Then
        
        Application.StatusBar = "Carregando dados dos registros já importados..."
        Set dicDados0190 = Util.CriarDicionarioRegistro(reg0190)
        Set dicDados0200 = Util.CriarDicionarioRegistro(reg0200)
        Set dicDadosC800 = Util.CriarDicionarioRegistro(regC800, "CHV_CFE")
        Set dicDadosC810 = Util.CriarDicionarioRegistro(regC810)
        Set dicDadosC850 = Util.CriarDicionarioRegistro(regC850)
        
        Set dicTitulos0000 = Util.MapearTitulos(reg0000, 3)
        Set dicTitulos0190 = Util.MapearTitulos(reg0190, 3)
        Set dicTitulos0200 = Util.MapearTitulos(reg0200, 3)
        Set dicTitulosC800 = Util.MapearTitulos(regC800, 3)
        Set dicTitulosC810 = Util.MapearTitulos(regC810, 3)
        Set dicTitulosC850 = Util.MapearTitulos(regC850, 3)
        
        Application.StatusBar = "Importando dados da CF-e-SAT..."
        a = 0
        Comeco = Timer
        For Each XML In DocsFiscais.arrCFe
        
            Call Util.AntiTravamento(a, 100, "Importando CF-e " & a + 1 & " de " & DocsFiscais.arrCFe.Count, DocsFiscais.arrCFe.Count, Comeco)

            Set CFe = fnXML.RemoverNamespaces(XML)
            If Not fnXML.ValidarXML(CFe) Then GoTo Prx:
            
            If fnXML.ValidarXMLCFe(CFe) And fnXML.ValidarParticipante(CFe) Then
                
                With CamposC800
                    
                    .COD_SIT = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_COD_SIT(fnSPED.GerarCodigoSituacao(fnXML.ValidarSituacao("100")))
                    .CHV_CFE = Util.FormatarTexto(VBA.Right(ValidarTag(CFe, "//@Id"), 44))
                    Periodo = VBA.Format(VBA.Format(CFe.SelectSingleNode("//dEmi").text, "0000-00-00"), "mm/yyyy")
                    CNPJContribuinte = Util.FormatarCNPJ(CadContrib.Range("CNPJContribuinte").value)
                    CNPJBase = VBA.Left(CNPJContribuinte, 8)
                    .ARQUIVO = Periodo & "-" & CNPJContribuinte
                    .CHV_PAI = fnSPED.GerarChaveRegistro(.ARQUIVO, "C001")
                    
                    Chave = VBA.Join(Array(VBA.Replace(.CHV_CFE, "'", "")))
                    If dicDadosC800.Exists(Chave) Then
                        
                        .CHV_REG = dicDadosC800(Chave)(dicTitulosC800("CHV_REG"))
                        .ARQUIVO = dicDadosC800(Chave)(dicTitulosC800("ARQUIVO"))
                        .COD_SIT = dicDadosC800(Chave)(dicTitulosC800("COD_SIT"))
                        
                    Else
                        
                        Call fnXML.CriarRegistroC800(CFe, dicDadosC800, DocsFiscais.arrChavesCanceladas)
                        
                    End If
                    
                    Set Produtos = CFe.SelectNodes("//det")
                    'Call fnXML.CriarRegistro0190(Produtos, dicDados0190, dicTitulos0190, .ARQUIVO)
                    'Call fnXML.CriarRegistro0200(Produtos, dicDados0200, dicTitulos0200, .ARQUIVO)
                    Call fnXML.CriarRegistroC810(Produtos, dicDadosC810)
                    Call fnXML.CriarRegistroC850(Produtos, dicTitulosC850, dicDadosC850)
                    
                End With
                
            End If
Prx:
        Next XML
        
        Application.StatusBar = "Importando dados dos XMLS para o SPED, por favor aguarde..."
        
        Call Util.ExportarDadosDicionario(reg0190, dicDados0190, "A4")
        Call Util.ExportarDadosDicionario(reg0200, dicDados0200, "A4")
        Call Util.ExportarDadosDicionario(regC800, dicDadosC800, "A4")
        Call Util.ExportarDadosDicionario(regC810, dicDadosC810, "A4")
        Call Util.ExportarDadosDicionario(regC850, dicDadosC850, "A4")
        
        Call FuncoesSPEDFiscal.ZerarDicionariosEFD
        
        Call dicDadosC800.RemoveAll
        Call dicDadosC810.RemoveAll
        Call dicDadosC850.RemoveAll
        
        Application.StatusBar = "Importação finalizada com sucesso!"
        Call Util.MsgInformativa("Importação concluída com sucesso!", "Importação CF-e-SAT", Inicio)
        
    End If
    
End Sub

Public Sub ImportarRegistrosC140eFilhos(ByVal TipoImportacao As String)

Dim ARQUIVO As String, Caminho$, Chave$
Dim dicTitulosC100 As New Dictionary
Dim dicTitulosC140 As New Dictionary
Dim dicTitulosC141 As New Dictionary
Dim dicDadosC100 As New Dictionary
Dim dicDadosC140 As New Dictionary
Dim dicDadosC141 As New Dictionary
Dim Produtos As IXMLDOMNodeList
Dim arrXMLs As New ArrayList
Dim NFe As New DOMDocument60
Dim XML As Variant

    If TipoImportacao = "Lote" Then
        Caminho = Util.SelecionarPasta("Selecione a pasta que contém os arquivos XML")
        If Caminho = "" Then Exit Sub
        Call Util.ListarArquivos(arrXMLs, Caminho)
    End If
    
    If TipoImportacao = "Arquivo" Then Call Util.GuardarEnderecosArrayList("xml", arrXMLs)
    
    If arrXMLs.Count > 0 Then
        
        Set dicDadosC100 = Util.CriarDicionarioRegistro(regC100, "IND_OPER", "IND_EMIT", "CHV_NFE")
        Set dicDadosC140 = Util.CriarDicionarioRegistro(regC140)
        Set dicDadosC141 = Util.CriarDicionarioRegistro(regC141)

        Set dicTitulosC100 = Util.MapearTitulos(regC100, 3)
        Set dicTitulosC140 = Util.MapearTitulos(regC140, 3)
        Set dicTitulosC141 = Util.MapearTitulos(regC141, 3)
        
        For Each XML In arrXMLs
            
            Set NFe = fnXML.RemoverNamespaces(XML)
            If Not fnXML.ValidarXML(NFe) Then GoTo Prx:
            
            If fnXML.ValidarXMLNFe(NFe) And fnXML.ValidarParticipante(NFe) Then
                
                With CamposC100
                    
                    .IND_OPER = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_OPER(fnXML.IdentificarTipoOperacao(NFe, ValidarTag(NFe, "//tpNF")))
                    .IND_EMIT = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_EMIT(fnXML.IdentificarTipoEmissao(NFe))
                    .CHV_NFE = VBA.Right(ValidarTag(NFe, "//@Id"), 44)
                    
                    Chave = Util.UnirCampos(.IND_OPER, .IND_EMIT, .CHV_NFE)
                    If dicDadosC100.Exists(Chave) Then
                            
                        .CHV_REG = dicDadosC100(Chave)(dicTitulosC100("CHV_REG"))
                        .ARQUIVO = dicDadosC100(Chave)(dicTitulosC100("ARQUIVO"))
                        .COD_PART = dicDadosC100(Chave)(dicTitulosC100("COD_PART"))
                        .IND_OPER = dicDadosC100(Chave)(dicTitulosC100("IND_OPER"))
                        .IND_EMIT = dicDadosC100(Chave)(dicTitulosC100("IND_EMIT"))
                    
                        Call fnXML.CriarRegistroC140(NFe, dicDadosC140, dicDadosC141, .ARQUIVO)
                        
                    End If
                    
                End With
                
            End If
Prx:
        Next XML
        
        Call Util.ExportarDadosDicionario(regC140, dicDadosC140, "A4")
        Call Util.ExportarDadosDicionario(regC141, dicDadosC141, "A4")
        Call FuncoesAssistentesInteligentes.GerarRelatorioContasPagarReceber(True)

        Call Util.MsgAviso("Faturas importadas com sucesso!", "Importação de Faturas")
        
    End If
    
End Sub

Public Sub ImportarXMLSParaAnalise(ByVal tpImport As String)

Dim dicDivergencias As New Dictionary
Dim arrArqs As New ArrayList
Dim Caminho As String
Dim UltLin As Long
    
    If tpImport = "Lote" Then
        Caminho = Util.SelecionarPasta("Selecione a pasta que contém os arquivos XML")
        If Caminho = "" Then Exit Sub
        Call Util.ListarArquivos(arrArqs, Caminho)
    End If
    
    If tpImport = "Arquivo" Then Call Util.GuardarEnderecosArrayList("xml", arrArqs)
    'Call Util.InterromperProcesso
    
    If arrArqs.Count > 0 Then
        
        Inicio = Now()
        
        a = 0
        Comeco = Timer
        Application.StatusBar = "Carregando dados do SPED, por favor aguarde..."
        Set dicDivergencias = Util.CriarDicionarioRegistro(relDivergencias, "CHV_NFE")
        
        Call fnXML.ImportarXMLSParaAnalise(arrArqs, dicDivergencias)
        
        Call Util.LimparDados(relDivergencias, 4, False)
        Call Util.ExportarDadosDicionario(relDivergencias, dicDivergencias, "A4")
        Call FuncoesFormatacao.DeletarFormatacao
        Call FuncoesFormatacao.AplicarFormatacao(relDivergencias)
        Call FuncoesFormatacao.FormatarDivergencias(relDivergencias)
        
        Call dicDivergencias.RemoveAll
        relDivergencias.Activate
        
        Call Util.MsgInformativa("Análise concluída com sucesso!", "Importação de XMLS para Análise", Inicio)
        
    End If
    
End Sub

Public Sub ImportarCadastro0200XML(ByVal TipoImportacao As String)

Dim dicTitulos0000 As New Dictionary
Dim dicTitulos0190 As New Dictionary
Dim dicTitulos0200 As New Dictionary
Dim dicDados0000 As New Dictionary
Dim dicDados0001 As New Dictionary
Dim dicDados0005 As New Dictionary
Dim dicDados0100 As New Dictionary
Dim dicDados0190 As New Dictionary
Dim dicDados0200 As New Dictionary
Dim dicDadosC001 As New Dictionary
Dim Produtos As IXMLDOMNodeList
Dim arrXMLs As New ArrayList
Dim NFe As New DOMDocument60
Dim ARQUIVO As String, Periodo$, Caminho$, CNPJEmit$, tpPart$, CHV_0150$, CHV_0000$, CHV_0001$, Msg$
Dim XML As Variant, Chave
Dim i As Integer

    If PeriodoImportacao = "" Then
        Call Util.MsgAlerta("Informe o período ('MM/AAAA') que deseja inserir os itens para prosseguir com a importação.", "Período de importação não informado")
        Exit Sub
    Else
        Periodo = VBA.Format(PeriodoImportacao, "00/0000")
    End If
    
    If InscContribuinte = "" Then
        Call Util.MsgAlerta("Informe a Inscrição Estadual do Contribuinte.", "Inscrição Estadual não informada")
        CadContrib.Activate
        CadContrib.Range("InscContribuinte").Activate
        Exit Sub
    End If
    
    ProcessarDocumentos.LimparListaDocumentos
    Call ProcessarDocumentos.CarregarXMLS(TipoImportacao)
    
    If TipoImportacao = "Lote" Then
        Caminho = Util.SelecionarPasta("Selecione a pasta que contém os arquivos XML")
        If Caminho = "" Then Exit Sub
        Inicio = Now()
        Call Util.ListarArquivos(arrXMLs, Caminho)
    End If
    
    If TipoImportacao = "Arquivo" Then Call Util.GuardarEnderecosArrayList("xml", arrXMLs)
    
    If arrXMLs.Count > 0 Then
        
        Application.StatusBar = "Importando produtos dos XMLs..."
        
        Set dicDados0000 = Util.CriarDicionarioRegistro(reg0000, "ARQUIVO")
        Set dicDados0190 = Util.CriarDicionarioRegistro(reg0190)
        Set dicDados0200 = Util.CriarDicionarioRegistro(reg0200)
        
        Set dicTitulos0000 = Util.MapearTitulos(reg0000, 3)
        Set dicTitulos0190 = Util.MapearTitulos(reg0190, 3)
        Set dicTitulos0200 = Util.MapearTitulos(reg0200, 3)
        
        a = 0
        Comeco = Timer
        If Not Periodo Like "*/*" Then
            Periodo = VBA.Format(PeriodoImportacao, "00/0000")
        End If
        
        ARQUIVO = Periodo & "-" & CNPJContribuinte
        For Each XML In arrXMLs
            
            Call Util.AntiTravamento(a, 100, "Importando XML " & a + 1 & " de " & arrXMLs.Count, arrXMLs.Count, Comeco)
            Set NFe = fnXML.RemoverNamespaces(XML)
            
            If fnXML.ValidarXMLNFe(NFe) Then

                CNPJEmit = VBA.Mid(VBA.Right(ValidarTag(NFe, "//@Id"), 44), 7, 14)
                If CNPJEmit = CNPJContribuinte Then
                    
                    Set Produtos = NFe.SelectNodes("//det")
                    
                    If Not dicDados0000.Exists(ARQUIVO) Then
                        
                        Call fnXML.CriarRegistro0000(NFe, dicDados0000, _
                            dicDados0001, dicDados0005, dicDados0100, dicDadosC001, Periodo, "emit")
                        
                    End If
                    
                    If LBound(dicDados0000(ARQUIVO)) = 0 Then i = 1 Else i = 0
                    CHV_0000 = dicDados0000(ARQUIVO)(3 - i)
                    
                    CHV_0001 = fnSPED.GerarChaveRegistro(CHV_0000, "0001")
                    Call fnXML.CriarRegistro0190(Produtos, dicDados0190, dicTitulos0190, ARQUIVO, CHV_0001)
                    Call fnXML.CriarRegistro0200(Produtos, dicDados0200, dicTitulos0200, ARQUIVO, CHV_0001)
                    
                End If
                
            End If
            
        Next XML
        
        Application.StatusBar = "Importando dados dos XMLS para o SPED, por favor aguarde..."
        
        If dicDados0200.Count = 0 Then
            Msg = "Os XMLS selecionados não pertecem ao CNPJ do contribuinte." & vbCrLf & vbCrLf
            Msg = Msg & "Por favor selecione novos XMLS e tente novamente."
            Call Util.MsgAlerta(Msg, "XMLS inválidos")
            Exit Sub
        End If
        
        Call Util.ExportarDadosDicionario(reg0190, dicDados0190, "A4")
        Call Util.ExportarDadosDicionario(reg0200, dicDados0200, "A4")
        
        Application.StatusBar = "Importação finalizada com sucesso!"
        Call Util.MsgInformativa("Importação concluída com sucesso!", "Importação de Produtos da NF-e/NFC-e", Inicio)
        
    End If
    
End Sub

Public Function ImportarProdutosFornecedor(ByVal TipoImportacao As String)

Dim CNPJEmit As String, cProd$, uCom$, Chave$, Msg$
Dim arrCorrelacoesExistentes As New ArrayList
Dim dicCorrelacoesUnicas As New Dictionary
Dim dicCorrelacoes As New Dictionary
Dim Arqs As Variant, Correlacao
Dim arrSPED As New ArrayList
    
    'If TipoImportacao = "Lote" Or TipoImportacao = "Correlacionar" Then Call Rotinas.ListarArquivos(docsfiscais.arrnfenfce)
    'If TipoImportacao = "Arquivo" Then Call Util.GuardarEnderecosArrayList("xml", docsfiscais.arrnfenfce)
    If TipoImportacao = "Correlacionar" Then Call Util.GuardarEnderecosArrayList("txt", arrSPED)
    If TipoImportacao = "Lote" Or TipoImportacao = "Arquivo" Then Call ProcessarDocumentos.CarregarXMLS(TipoImportacao)
    
    Inicio = Now()
    
    If (DocsFiscais.arrNFeNFCe.Count + arrSPED.Count) > 0 Then
    
        Set dicCorrelacoes = Util.CriarDicionarioCorrelacoes(Correlacoes)
        If DocsFiscais.arrNFeNFCe.Count > 0 Then Call fnXML.CriarRegistroProdutoFornecedor(dicCorrelacoes, DocsFiscais.arrNFeNFCe)
        If arrSPED.Count > 0 Then Call fnSPED.ImportarDadosProdutoContribuinte(dicCorrelacoes, arrSPED)
    
        Call Util.LimparDados(Correlacoes, 4, False)
        Call Util.ExportarDadosDicionario(Correlacoes, dicCorrelacoes)
        
        If dicCorrelacoes.Count > 0 Then
        
            Call Util.MsgInformativa("Correlações geradas com sucesso!", "Correlação de produtos", Inicio)
            
        Else
            
            Msg = "O CNPJ do destinatário não apareceu em nenhuma das notas selecionadas." & vbCrLf & vbCrLf
            Msg = Msg & "Por favor verifique o CNPJ informado ou os XMLs selecionados e tente novamente."
            
            Call Util.MsgAlerta(Msg, "CNPJ do destinatário incompatível com XMLS selecionados")
            
        End If
    
    End If
    
End Function

Private Sub CorrelacionarCTeNFe(ByRef CTe As DOMDocument60, ByRef dicCorrelacoesCTeNFe As Dictionary)
    
Dim ChaveCTe As String, ChaveNFe$
Dim ChavesNFe As IXMLDOMNodeList
Dim vCTe As Double, vBCICMS#
Dim Chave As IXMLDOMNode
    
    Set ChavesNFe = CTe.SelectNodes("//chave")
    If ChavesNFe.Length = 0 Then Exit Sub
    
    ChaveCTe = VBA.Right(ValidarTag(CTe, "//@Id"), 44)
    vCTe = fnXML.ValidarValores(CTe, "//vRec")
    vBCICMS = fnXML.ValidarValores(CTe, "//imp/ICMS//vBC")
    
    For Each Chave In ChavesNFe
        
        ChaveNFe = Chave.text
        dicCorrelacoesCTeNFe(Util.UnirCampos(ChaveCTe, ChaveNFe)) = Array("'" & ChaveCTe, vCTe, vBCICMS, "'" & ChaveNFe, "", "", "", "", "")
        
    Next Chave
    
End Sub
