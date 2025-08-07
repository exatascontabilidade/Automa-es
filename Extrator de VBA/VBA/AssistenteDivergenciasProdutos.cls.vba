Attribute VB_Name = "AssistenteDivergenciasProdutos"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Const TitulosIgnorar As String = "REG, ARQUIVO, CHV_PAI_FISCAL, CHV_PAI_CONTRIBUICOES, CHV_REG, CHV_NFE, NUM_DOC, SER"
Private ValidacoesProdutos As New clsRegrasDivergenciasProdutos
Private DivQtdItensNota As DivergenciasQuantidadeItens
Private Divergencias As AssistenteDivergencias
Private GerenciadorSPED As clsRegistrosSPED
Private Check As VerificacoesCamposProdutos
Private dicTitulosRelatorio As Dictionary
Private dicTagsProdutos As New Dictionary
Private Apuracao As clsAssistenteApuracao
Private arrTitulosRelatorio As ArrayList
Private arrRelatorio As New ArrayList
Private CamposRelatorio As Variant
Private NomesCampos As Variant
Private CamposSPED As Variant
Private CamposXML As Variant

Public Sub GerarComparativoXMLSPED()

Dim Msg As String
Dim SPEDS As Variant
Dim CaminhoXMLS As String
Dim Result As VbMsgBoxResult
Dim arrChaves As New ArrayList
Dim RemovidaDuplicatas As Boolean
Dim TentavivaRemocaoDuplicidade As Boolean
    
    If Util.ChecarAusenciaDados(regC170, , "Dados ausentes no registro C170") Then Exit Sub
    
    Call Util.DesabilitarControles
    
    If Not ProcessarDocumentos.CarregarXMLS("Lote") Then Exit Sub
    Call InicializarObjetos
    
Reimportar:
    Call CarregarDadosC170
    Call CarregarDadosXML
    Call CorrelacionarNotas
    
    Application.StatusBar = "Processo concluído com sucesso!"
    If arrRelatorio.Count > 0 Then
        
        Call DadosValidacaoCFOP.CarregarTitulosRelatorio(ActiveSheet)
        Call ValidacoesProdutos.IdentificarDivergenciasProdutos(arrRelatorio)
        
        If relDivergenciasProdutos.AutoFilter.FilterMode Then relDivergenciasProdutos.ShowAllData
        Call Util.LimparDados(relDivergenciasProdutos, 4, False)
        
        Call Util.ExportarDadosArrayList(relDivergenciasProdutos, arrRelatorio)
        Call FuncoesFormatacao.DestacarInconsistencias(relDivergenciasProdutos)
        
        If Not TentavivaRemocaoDuplicidade Then
            
            RemovidaDuplicatas = AjustarDivergenciasQtdItensSPED_XML
            If RemovidaDuplicatas Then
                
                TentavivaRemocaoDuplicidade = True
                
                Msg = "Duplicidades removidas com sucesso!" & vbCrLf & vbCrLf
                Msg = Msg & "Deseja reimportar os XMLS para gerar uma nova análise?"
                
                Result = Util.MsgDecisao(Msg, "Regerar Análise dos XMLS")
                If Result = vbYes Then
                    
                    Call arrRelatorio.Clear
                    GoTo Reimportar:
                    
                End If
                
            End If
            
        End If
        
        Call Util.MsgInformativa("Relatório gerado com sucesso", "Relatório de Divergências de Produtos", Inicio)
        
    Else
        
        Msg = "Nenhum dado encontrado para geração do relatório." & vbCrLf & vbCrLf
        Msg = Msg & "Por favor verifique se o SPED e/ou XMLs foram importados e tente novamente."
        Call Util.MsgAlerta(Msg, "Relatório de Divergências de Produtos")
        
    End If
    
    Call ResetarObjetos
    
End Sub

Private Sub InicializarObjetos()
    
    Call arrRelatorio.Clear
    Call dicInconsistenciasIgnoradas.RemoveAll
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    
    Set Apuracao = New clsAssistenteApuracao
    Set GerenciadorSPED = New clsRegistrosSPED
    Set Divergencias = New AssistenteDivergencias
    
    Set dicTitulosRelatorio = Util.MapearTitulos(relDivergenciasProdutos, 3)
    
    With GerenciadorSPED
        
        Call .CarregarDadosRegistro0000("ARQUIVO")
        Call .CarregarDadosRegistro0000_Contr("ARQUIVO")
        Call .CarregarDadosRegistro0001("ARQUIVO")
        Call .CarregarDadosRegistro0140("CNPJ")
        Call .CarregarDadosRegistroC010
        Call .CarregarDadosRegistroC100
        Call .CarregarDadosRegistroC170
        
        If dtoRegSPED.r0000.Count = 0 Then _
            Call .CarregarDadosRegistro0190("CHV_PAI_CONTRIBUICOES", "UNID") _
                Else: Call .CarregarDadosRegistro0190("CHV_PAI_FISCAL", "UNID")
                
        If dtoRegSPED.r0000.Count = 0 Then _
            Call .CarregarDadosRegistro0200("CHV_PAI_CONTRIBUICOES", "COD_ITEM") _
                Else Call .CarregarDadosRegistro0200("CHV_PAI_FISCAL", "COD_ITEM")
            
    End With
    
End Sub

Private Sub ResetarObjetos()
    
    Set Apuracao = Nothing
    Set arrRelatorio = Nothing
    Set Divergencias = Nothing
    Set dicTitulosRelatorio = Nothing
    Set dicInconsistenciasIgnoradas = Nothing
        
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    Call Util.AtualizarBarraStatus(False)
    Call Util.DesabilitarControles
    
End Sub

Public Sub ReprocessarSugestoes()

Dim arrDivProdutos As New ArrayList
Dim arrRelatorio As New ArrayList
Dim dicTitulos As New Dictionary
Dim Campos As Variant
    
    If relDivergenciasProdutos.AutoFilterMode Then relDivergenciasProdutos.AutoFilter.ShowAllData
    
    If Util.ChecarAusenciaDados(relDivergenciasProdutos) Then Exit Sub
    Set arrDivProdutos = Util.CriarArrayListRegistro(relDivergenciasProdutos)
    Set dicTitulos = Util.MapearTitulos(relDivergenciasProdutos, 3)
    
    a = 0
    Comeco = Timer
    For Each Campos In arrDivProdutos
        
        Call Util.AntiTravamento(a, 10, "Reprocessando sugestões, por favor aguarde...", arrDivProdutos.Count, Comeco)
        
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            Campos(dicTitulos("INCONSISTENCIA")) = Empty
            Campos(dicTitulos("SUGESTAO")) = Empty
            
            arrRelatorio.Add Campos
            
        End If
        
    Next Campos
        
    Call ValidacoesProdutos.IdentificarDivergenciasProdutos(arrRelatorio)
    
    Call Util.LimparDados(relDivergenciasProdutos, 4, False)
    Call Util.ExportarDadosArrayList(relDivergenciasProdutos, arrRelatorio)
    Call FuncoesFormatacao.DestacarInconsistencias(relDivergenciasProdutos)
    
    Call Util.AtualizarBarraStatus("Processamento Concluído!")
    Call Util.AtualizarBarraStatus(False)
    
End Sub

Public Function AceitarSugestoesProdutos()

Dim Dados As Range, Linha As Range
Dim i As Long
    
    Set Dados = relDivergenciasProdutos.Range("A4").CurrentRegion
    Set arrRelatorio = New ArrayList
    
    If Dados Is Nothing Then
        Call Util.MsgAlerta("Não foi encontrada nenhuma inconsistência!", "Inconsistências Fiscais")
        Exit Function
    End If
    
    For Each Linha In Dados.Rows
        
        With CamposProduto
            
            CamposRelatorio = Application.index(Linha.Value2, 0, 0)
            
            Call DadosDivergenciasProdutos.ResetarCamposProduto
            If Util.ChecarCamposPreenchidos(CamposRelatorio) Then
                
                If Linha.Row > 3 Then Call DadosDivergenciasProdutos.CarregarDadosRegistroDivergenciaProdutos(CamposRelatorio)
                
                If Linha.EntireRow.Hidden = False And .SUGESTAO <> "" And Linha.Row > 3 Then
                    
                    Call AplicarSugestaoProduto
                    
                    .INCONSISTENCIA = Empty
                    .SUGESTAO = Empty
                    
                    CamposRelatorio = DadosDivergenciasProdutos.AtribuirCamposProduto(CamposRelatorio)
                    
                End If
                
                If Linha.Row > 3 Then arrRelatorio.Add CamposRelatorio
                
            End If
            
        End With
        
    Next Linha
    
    Call Util.AtualizarBarraStatus("Reanalisando inconsistências, por favor aguarde...")
    Call ValidacoesProdutos.IdentificarDivergenciasProdutos(arrRelatorio)
    
    If relDivergenciasProdutos.AutoFilter.FilterMode Then relDivergenciasProdutos.ShowAllData
    Call Util.LimparDados(relDivergenciasProdutos, 4, False)
    
    Call Util.ExportarDadosArrayList(relDivergenciasProdutos, arrRelatorio)
    Call FuncoesFormatacao.DestacarInconsistencias(relDivergenciasProdutos)
    
    Application.StatusBar = False
    
End Function

Private Function AjustarDivergenciasQtdItensSPED_XML() As Boolean

Dim RemovidasDuplicatas As Boolean
Dim arrChaves As New ArrayList
Dim Result As VbMsgBoxResult
Dim Msg As String
    
    Set arrChaves = ListarChavesDivergenciasQtdItensSPED_XML
    If arrChaves.Count > 0 Then
        
        Msg = "Foram identificadas notas com quantidades de itens lançados no SPED maior que a quantidade de itens do XML." & vbCrLf & vbCrLf
        Msg = Msg & "Deseja tentar unificar os itens em caso de lançamentos de produtos duplicados?"
        
        Result = Util.MsgDecisao(Msg, "Divergência de Itens SPED vs XML")
        
        If Result = vbYes Then
            
            RemovidasDuplicatas = rC170.UnificarProdutosDuplicadosEmNotasSelecionadas(arrChaves)
            If RemovidasDuplicatas Then AjustarDivergenciasQtdItensSPED_XML = True
            
        End If
        
    End If
    
End Function

Private Function ListarChavesDivergenciasQtdItensSPED_XML() As ArrayList

Dim SUGESTAO As String, CHV_PAI$
Dim arrChaves As New ArrayList
Dim arrDados As New ArrayList
Dim Campos As Variant
    
    Set arrDados = Util.CriarArrayListRegistro(relDivergenciasProdutos)
    
    For Each Campos In arrDados
        
        SUGESTAO = Campos(dicTitulosRelatorio("INCONSISTENCIA"))
        If SUGESTAO Like "A quantidade de Itens lançados no SPED*" Then
            
            CHV_PAI = Campos(dicTitulosRelatorio("CHV_PAI_FISCAL"))
            If Not arrChaves.contains(CHV_PAI) Then arrChaves.Add CHV_PAI
            
        End If
        
    Next Campos
    
    Set ListarChavesDivergenciasQtdItensSPED_XML = arrChaves
    
End Function

Private Function AplicarSugestaoProduto()
    
    With CamposProduto
        
        Select Case .SUGESTAO
            
            Case "Informar o mesmo código CEST do XML para o SPED"
                .CEST_SPED = .CEST_NF
                
            Case "Apagar valor do CEST informado no campo CEST_NF"
                .CEST_NF = ""
                
            Case "Informar o mesmo código de barras do XML para o SPED"
                .COD_BARRA_SPED = .COD_BARRA_NF
                
            Case "Informar o mesmo NCM do XML para o SPED"
                .COD_NCM_SPED = fnExcel.FormatarTexto(.COD_NCM_NF)
                
            Case "Apagar código de barras informado no SPED"
                .COD_BARRA_SPED = ""
                
            Case "Apagar CEST informado no SPED"
                .CEST_SPED = ""
                
            Case "Apagar NCM informado no SPED"
                .COD_NCM_SPED = ""
                
            Case "Apropiar crédito do ICMS"
                .VL_BC_ICMS_SPED = .VL_BC_ICMS_NF
                .ALIQ_ICMS_SPED = .ALIQ_ICMS_NF
                .VL_ICMS_SPED = .VL_ICMS_NF
                
            Case "Zerar campos do ICMS"
                .VL_BC_ICMS_SPED = 0
                .ALIQ_ICMS_SPED = 0
                .VL_ICMS_SPED = 0
                
            Case "Mudar o dígito de origem do CST_ICMS_SPED para 2"
                .CST_ICMS_SPED = "2" & VBA.Right(.CST_ICMS_SPED, VBA.Len(.CST_ICMS_SPED) - 1)
                
            Case "Mudar o dígito de origem do CST_ICMS_SPED para 7"
                .CST_ICMS_SPED = "7" & VBA.Right(.CST_ICMS_SPED, VBA.Len(.CST_ICMS_SPED) - 1)
                
            Case "Informar o CST 40 da tabela B para o campo CST_ICMS_SPED"
                .CST_ICMS_SPED = VBA.Left(.CST_ICMS_NF, 1) & "40"
                
            Case "Informar o CST 41 da tabela B para o campo CST_ICMS_SPED"
                .CST_ICMS_SPED = VBA.Left(.CST_ICMS_NF, 1) & "41"
                
            Case "Informar o CST 60 da tabela B para o campo CST_ICMS_SPED"
                .CST_ICMS_SPED = VBA.Left(.CST_ICMS_NF, 1) & "60"
                
            Case "Informar o CST 90 da tabela B para o campo CST_ICMS_SPED"
                .CST_ICMS_SPED = VBA.Left(.CST_ICMS_NF, 1) & "90"
                
            Case "Informar o mesmo valor do campo VL_ITEM_NF para o campo VL_ITEM_SPED"
                .VL_ITEM_SPED = .VL_ITEM_NF
                
            Case "Informar o mesmo valor do campo UNID_NF para o campo UNID_SPED"
                .UNID_SPED = .UNID_NF
                
            Case "Informar o mesmo valor do campo QTD_NF para o campo QTD_SPED"
                .QTD_SPED = .QTD_NF
                
            Case "Informar o mesmo valor do campo VL_DESC_NF para o campo VL_DESC_SPED"
                .VL_DESC_SPED = .VL_DESC_NF
                
            Case "Informar o mesmo valor do campo VL_BC_ICMS_NF para o campo VL_BC_ICMS_SPED", "Informar o mesmo valor do campo ALIQ_ICMS_NF para o campo ALIQ_ICMS_SPED", _
                 "Informar o mesmo valor do campo VL_ICMS_NF para o campo VL_ICMS_SPED"
                .VL_BC_ICMS_SPED = .VL_BC_ICMS_NF
                .ALIQ_ICMS_SPED = .ALIQ_ICMS_NF
                .VL_ICMS_SPED = .VL_ICMS_NF
                
            Case "Informar o mesmo valor do campo VL_BC_ICMS_ST_NF para o campo VL_BC_ICMS_ST_SPED", "Informar o mesmo valor do campo ALIQ_ST_NF para o campo ALIQ_ST_SPED", _
                 "Informar o mesmo valor do campo VL_ICMS_ST_NF para o campo VL_ICMS_ST_SPED"
                .VL_BC_ICMS_ST_SPED = .VL_BC_ICMS_ST_NF
                .ALIQ_ST_SPED = .ALIQ_ST_NF
                .VL_ICMS_ST_SPED = .VL_ICMS_ST_NF
                
            Case "Informar o mesmo valor do campo VL_BC_IPI_NF para o campo VL_BC_IPI_SPED", "Informar o mesmo valor do campo VL_IPI_NF para o campo VL_IPI_SPED", _
                 "Informar o mesmo valor do campo ALIQ_IPI_NF para o campo ALIQ_IPI_SPED"
                .VL_IPI_SPED = .VL_IPI_NF
                .VL_BC_IPI_SPED = .VL_BC_IPI_NF
                .ALIQ_IPI_SPED = .ALIQ_IPI_NF
                .VL_IPI_SPED = .VL_IPI_NF
                
            Case "Somar valor do campo VL_IPI_SPED ao campo VL_ITEM_SPED"
                .VL_ITEM_SPED = CDbl(.VL_ITEM_NF) + CDbl(.VL_IPI_SPED)
                .CST_IPI_SPED = ""
                .VL_BC_IPI_SPED = ""
                .ALIQ_IPI_SPED = ""
                .VL_IPI_SPED = ""
                
            Case "Somar valor do campo VL_ICMS_ST_SPED ao campo VL_ITEM_SPED"
                .VL_ITEM_SPED = CDbl(.VL_ITEM_SPED) + CDbl(.VL_ICMS_ST_SPED)
                .VL_BC_ICMS_ST_SPED = ""
                .ALIQ_ST_SPED = ""
                .VL_ICMS_ST_SPED = ""
                
            Case "Informar o mesmo valor do campo VL_ICMS_ST_NF para o campo VL_ICMS_ST_SPED"
                .VL_ICMS_ST_SPED = .VL_ICMS_ST_NF
                .VL_BC_ICMS_ST_SPED = .VL_BC_ICMS_ST_NF
                .ALIQ_ST_SPED = .ALIQ_ST_NF
                .VL_ICMS_ST_SPED = .VL_ICMS_ST_NF
                
            Case Else
            
        End Select
        
    End With
    
End Function

Public Function IgnorarInconsistencias()

Dim dicTitulos0000 As New Dictionary
Dim dicDados0000 As New Dictionary
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
    
    Set dicTitulos0000 = Util.MapearTitulos(reg0000, 3)
    Set dicDados0000 = Util.CriarDicionarioRegistro(reg0000, "ARQUIVO")
    
    Set dicTitulos = Util.MapearTitulos(relDivergenciasProdutos, 3)
    Set Dados = relDivergenciasProdutos.Range("A4").CurrentRegion
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

Dim CHV_REG As String, CHV_PAI$, CHV_0001$, CHV_0140$, CHV_0190$, CHV_0200$, CHV_C100$, COD_ITEM$, UNID$, ARQUIVO$
Dim Campos As Variant, Campos0190, Campos0200, CamposC170, dicCampos, regCampo, nCampo
Dim ExpReg As New ExportadorRegistros
Dim arrDivProdutos As New ArrayList
Dim dicTitulos As New Dictionary
Dim i As Integer, j As Integer
    
    If Util.ChecarAusenciaDados(relDivergenciasProdutos) Then Exit Sub
    
    Call Util.DesabilitarControles
    
    Inicio = Now()
    Application.StatusBar = "Preparando dados para atualização do SPED, por favor aguarde..."
    If relDivergenciasProdutos.AutoFilterMode Then relDivergenciasProdutos.AutoFilter.ShowAllData
    
    Campos0190 = Array("UNID")
    Campos0200 = Array("COD_BARRA", "COD_NCM", "EX_IPI", "CEST")
    CamposC170 = Array("NUM_ITEM", "COD_ITEM", "DESCR_COMPL", "QTD", "UNID", "VL_ITEM", "VL_DESC", "CST_ICMS", "CFOP", "VL_BC_ICMS", "ALIQ_ICMS", "VL_ICMS", "VL_BC_ICMS_ST", "ALIQ_ST", "VL_ICMS_ST", "CST_IPI", "VL_BC_IPI", "ALIQ_IPI", "VL_IPI", "CST_PIS", "VL_BC_PIS", "ALIQ_PIS", "QUANT_BC_PIS", "ALIQ_PIS_QUANT", "VL_PIS", "CST_COFINS", "VL_BC_COFINS", "ALIQ_COFINS", "QUANT_BC_COFINS", "ALIQ_COFINS_QUANT", "VL_COFINS")
    
    Call CarregarObjetos
    
    Set dicTitulos = Util.MapearTitulos(relDivergenciasProdutos, 3)
    If relDivergenciasProdutos.AutoFilterMode Then relDivergenciasProdutos.AutoFilter.ShowAllData
    
    Set arrDivProdutos = Util.CriarArrayListRegistro(relDivergenciasProdutos)
    
    a = 0
    Comeco = Timer
    For Each Campos In arrDivProdutos
        
        i = Util.VerificarPosicaoInicialArray(Campos)
        Call Util.AntiTravamento(a, 100, "Preparando dados para atualização do SPED, por favor aguarde...", arrDivProdutos.Count, Comeco)
        
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            ARQUIVO = Campos(dicTitulos("ARQUIVO"))
            COD_ITEM = Campos(dicTitulos("COD_ITEM_SPED"))
            UNID = Campos(dicTitulos("UNID_SPED"))
            
            CHV_PAI = Extrair_CHV_PAI_0190(dicTitulos, Campos, True)
            CHV_0190 = Util.UnirCampos(CHV_PAI, UNID)
            If Not dtoRegSPED.r0190.Exists(CHV_0190) Then
                
                CHV_0001 = dtoRegSPED.r0001(ARQUIVO)(dtoTitSPED.t0001("CHV_REG"))
                CHV_C100 = Campos(dicTitulos("CHV_PAI_CONTRIBUICOES"))
                CHV_0140 = Extrair_CHV_0140(CHV_C100)
                
                CHV_REG = fnSPED.GerarChaveRegistro(CHV_0001, UNID)
                dicCampos = Array("'0190", ARQUIVO, CHV_REG, CHV_0001, CHV_0140, UNID, UNID)
                
                dtoRegSPED.r0190(CHV_0190) = dicCampos
                
            End If
            
            CHV_PAI = Extrair_CHV_PAI_0200(dicTitulos, Campos, True)
            CHV_0200 = Util.UnirCampos(CHV_PAI, COD_ITEM)
            If dtoRegSPED.r0200.Exists(CHV_0200) Then
                
                dicCampos = dtoRegSPED.r0200(CHV_0200)
                j = Util.VerificarPosicaoInicialArray(dicCampos)
                For Each regCampo In Campos0200
                    
                    nCampo = regCampo & "_SPED"
                    dicCampos(dtoTitSPED.t0200(regCampo) - j) = Campos(dicTitulos(nCampo) - i)
                    
                Next regCampo
                
                dtoRegSPED.r0200(CHV_0200) = dicCampos
                
            End If
            
            CHV_REG = Campos(dicTitulos("CHV_REG"))
            If dtoRegSPED.rC170.Exists(CHV_REG) Then
                
                dicCampos = dtoRegSPED.rC170(CHV_REG)
                j = Util.VerificarPosicaoInicialArray(dicCampos)
                For Each regCampo In CamposC170
                    
                    If regCampo = "DESCR_COMPL" Then nCampo = "DESCR_ITEM_NF" Else nCampo = regCampo & "_SPED"
                    If regCampo Like "*UNID*" And Campos(dicTitulos(nCampo)) Like "* - *" Then Campos(dicTitulos(nCampo)) = VBA.Split(Campos(dicTitulos(nCampo)), " - ")(0)
                    
                    dicCampos(dtoTitSPED.tC170(regCampo) - j) = Campos(dicTitulos(nCampo) - i)
                    
                Next regCampo
                
                dtoRegSPED.rC170(CHV_REG) = dicCampos
                
            End If
            
        End If
        
    Next Campos
    
    Call ExpReg.ExportarRegistros("0190", "0200", "C170")
    
    Call Util.AtualizarBarraStatus("Atualizando dados do registro C190, por favor aguarde...")
    Call rC170.GerarC190(True)
    
    Call Util.AtualizarBarraStatus("Atualizando valores dos impostos no registro C100, por favor aguarde...")
    Call rC170.AtualizarImpostosC100(True)
    
    Call Util.AtualizarBarraStatus("Atualização concluída com sucesso!")
    Call Util.MsgInformativa("Registros atualizados com sucesso!", "Atualização de dados", Inicio)
    
    Call FuncoesFormatacao.AplicarFormatacao(ActiveSheet)
    'Call FuncoesFormatacao.DestacarInconsistencias(relDivergenciasProdutos)
    
    Call Util.AtualizarBarraStatus(False)
    Call Util.HabilitarControles
    Set ExpReg = Nothing
    
End Sub

Private Sub CarregarObjetos()
    
    Set GerenciadorSPED = New clsRegistrosSPED
    
    With GerenciadorSPED
        
        Call .CarregarDadosRegistro0000("ARQUIVO")
        Call .CarregarDadosRegistro0000_Contr("ARQUIVO")
        Call .CarregarDadosRegistro0001("ARQUIVO")
        Call .CarregarDadosRegistro0140("CNPJ")
        Call .CarregarDadosRegistroC010
        Call .CarregarDadosRegistroC100
        Call .CarregarDadosRegistroC170
        
        If dtoRegSPED.r0000.Count = 0 Then _
            Call .CarregarDadosRegistro0190("CHV_PAI_CONTRIBUICOES", "UNID") _
                Else: Call .CarregarDadosRegistro0190("CHV_PAI_FISCAL", "UNID")
                
        If dtoRegSPED.r0000.Count = 0 Then _
            Call .CarregarDadosRegistro0200("CHV_PAI_CONTRIBUICOES", "COD_ITEM") _
                Else Call .CarregarDadosRegistro0200("CHV_PAI_FISCAL", "COD_ITEM")
        
    End With
    
End Sub

Public Sub CarregarDadosC170()

Dim Dados As Range, Linha As Range
Dim Campos As Variant, Campo
Dim Comeco As Double
Dim Msg As String
Dim a As Long
    
    With Divergencias
        
        .TipoRelatorio = "SPED"
        Call .dicOperacoesSPED.RemoveAll
        Call .CarregarTitulosRelatorioProdutos
        Call .RedimensionarArray(.arrTitulosRelatorio.Count)
        
        a = 0
        Comeco = Timer
        For Each Campos In dtoRegSPED.rC170.Items()
            
            If Util.ChecarCamposPreenchidos(Campos) Then
                    
                Call Util.AntiTravamento(a, 10, Msg & "Carregando dados do registro C170", dtoRegSPED.rC170.Count, Comeco)
                
                For Each Campo In .arrTitulosRelatorio
                                    
                    Call TaratarCamposC170(Campos, Campo)
                    
                Next Campo
                
                Call .RegistrarOperacaoSPED
                
            End If
Prx:
        Next Campos
        
    End With
    
End Sub

Private Function TaratarCamposC170(ByVal Campos As Variant, ByVal Campo As String)

Const CamposIgnorar As String = "CHV_NFE, NUM_DOC, SER, DESCR_ITEM, COD_BARRA, COD_NCM, EX_IPI, CEST"
Dim ARQUIVO As String, CHV_REG$, CHV_PAI$, COD_ITEM$, UNID$
Dim VL_OPER As Double
    
    With Divergencias
        
        Select Case True
            
            Case CamposIgnorar Like "*" & Campo & "*"
            
            Case Campo Like "CHV_PAI_FISCAL"
                CHV_PAI = Campos(dtoTitSPED.tC170(Campo))
                .AtribuirValor Campo, CHV_PAI
                Call .ExtrairDadosC100(CHV_PAI)
                
            Case Campo Like "ARQUIVO"
                ARQUIVO = Campos(dtoTitSPED.tC170(Campo))
                .AtribuirValor Campo, ARQUIVO
                
            Case Campo Like "COD_ITEM"
                COD_ITEM = Campos(dtoTitSPED.tC170(Campo))
                CHV_PAI = Extrair_CHV_PAI_0200(dtoTitSPED.tC170, Campos, False)
                .AtribuirValor Campo, fnExcel.FormatarTexto(COD_ITEM)
                Call Divergencias.ExtrairDados0200(CHV_PAI, COD_ITEM)
                
            Case Campo Like "VL_OPER"
                VL_OPER = .Extrair_VL_OPER(dtoTitSPED.tC170, Campos)
                .AtribuirValor Campo, VL_OPER
                
            Case Campo Like "NUM_ITEM", Campo Like "CFOP"
                .AtribuirValor Campo, CInt(Campos(dtoTitSPED.tC170(Campo)))
                
            Case Campo Like "VL_*", Campo Like "ALIQ_*", Campo Like "QTD*"
                .AtribuirValor Campo, fnExcel.ConverterValores(Campos(dtoTitSPED.tC170(Campo)))
                
            Case Campo Like "REG", Campo Like "CHV_REG"
                .AtribuirValor Campo, Campos(dtoTitSPED.tC170(Campo))
                
            Case Campo Like "*UNID*"
                CHV_PAI = Extrair_CHV_PAI_0200(dtoTitSPED.tC170, Campos, False)
                UNID = Campos(dtoTitSPED.tC170(Campo))
                Call .ExtrairUnidade0190(CHV_PAI, UNID)
                
            Case Else
                .AtribuirValor Campo, fnExcel.FormatarTexto(Campos(dtoTitSPED.tC170(Campo)))
                
        End Select
        
    End With
    
End Function

Private Function Extrair_CHV_0140(ByRef CHV_C100 As String) As String

Dim i As Integer
Dim CHV_C010 As String, CNPJ$
Dim Campos0001 As Variant, Campos0140, CamposC010, CamposC100
    
    If dtoRegSPED.rC100.Exists(CHV_C100) Then
        
        CamposC100 = dtoRegSPED.rC100(CHV_C100)
        i = Util.VerificarPosicaoInicialArray(CamposC100)
        CHV_C010 = CamposC100(dtoTitSPED.tC100("CHV_PAI_CONTRIBUICOES") - i)
        
        If dtoRegSPED.rC010.Exists(CHV_C010) Then
            
            CamposC010 = dtoRegSPED.rC010(CHV_C010)
            i = Util.VerificarPosicaoInicialArray(CamposC010)
            CNPJ = CamposC010(dtoTitSPED.tC010("CNPJ") - i)
            
            If dtoRegSPED.r0140.Exists(CNPJ) Then
                
                Campos0140 = dtoRegSPED.r0140(CNPJ)
                i = Util.VerificarPosicaoInicialArray(Campos0140)
                Extrair_CHV_0140 = Campos0140(dtoTitSPED.t0140("CHV_REG") - i)
                Exit Function
                
            End If
            
        End If
        
    End If
    
End Function

Private Function Extrair_CHV_PAI_0190(ByRef dicTitulos As Dictionary, ByVal Campos As Variant, ByVal CamposSPED As Boolean) As String

Dim i As Integer
Dim Chave As String
Dim Campo As Variant, Valor
Dim Campos0001 As Variant, Campos0140, Campos0190, CamposC010, CamposC100
Dim ARQUIVO As String, CHV_REG$, CHV_PAI$, CHV_C010$, CHV_PAI_FISCAL$, CHV_PAI_CONTRIBUICOES$, CNPJ$, UNID$
    
    i = Util.VerificarPosicaoInicialArray(Campos)
    
    ARQUIVO = Campos(dicTitulos("ARQUIVO") - i)
    CHV_PAI_FISCAL = Campos(dicTitulos("CHV_PAI_FISCAL") - i)
    CHV_PAI_CONTRIBUICOES = Campos(dicTitulos("CHV_PAI_CONTRIBUICOES") - i)
    UNID = Campos(dicTitulos(IIf(CamposSPED = True, "UNID_SPED", "UNID")) - i)
    
    If CHV_PAI_FISCAL <> "" And dtoRegSPED.r0000.Count > 0 Then
        
        Campos0001 = dtoRegSPED.r0001(ARQUIVO)
        i = Util.VerificarPosicaoInicialArray(Campos0001)
        CHV_PAI = Campos0001(dtoTitSPED.t0001("CHV_REG") - i)
        
    ElseIf CHV_PAI_CONTRIBUICOES <> "" Then
        
        CamposC100 = dtoRegSPED.rC100(CHV_PAI_CONTRIBUICOES)
        i = Util.VerificarPosicaoInicialArray(CamposC100)
        CHV_C010 = CamposC100(dtoTitSPED.tC100("CHV_PAI_CONTRIBUICOES") - i)
        
        CamposC010 = dtoRegSPED.rC010(CHV_C010)
        i = Util.VerificarPosicaoInicialArray(CamposC010)
        CNPJ = CamposC010(dtoTitSPED.tC010("CNPJ") - i)
        
        Campos0140 = dtoRegSPED.r0140(CNPJ)
        i = Util.VerificarPosicaoInicialArray(Campos0140)
        CHV_PAI = Campos0140(dtoTitSPED.t0140("CHV_REG") - i)
        
    End If
    
    Extrair_CHV_PAI_0190 = CHV_PAI
    
End Function

Private Function Extrair_CHV_PAI_0200(ByRef dicTitulos As Dictionary, ByVal Campos As Variant, ByVal CamposSPED As Boolean) As String

Dim i As Integer
Dim Chave As String
Dim Campo As Variant, Valor
Dim Campos0001 As Variant, Campos0140, Campos0200, CamposC010, CamposC100
Dim ARQUIVO As String, CHV_REG$, CHV_PAI$, CHV_C010$, CHV_PAI_FISCAL$, CHV_PAI_CONTRIBUICOES$, CNPJ$, COD_ITEM$
    
    i = Util.VerificarPosicaoInicialArray(Campos)
    
    ARQUIVO = Campos(dicTitulos("ARQUIVO") - i)
    CHV_PAI_FISCAL = Campos(dicTitulos("CHV_PAI_FISCAL") - i)
    CHV_PAI_CONTRIBUICOES = Campos(dicTitulos("CHV_PAI_CONTRIBUICOES") - i)
    COD_ITEM = Campos(dicTitulos(IIf(CamposSPED = True, "COD_ITEM_SPED", "COD_ITEM")) - i)
    
    If CHV_PAI_FISCAL <> "" And dtoRegSPED.r0000.Count > 0 Then
        
        Campos0001 = dtoRegSPED.r0001(ARQUIVO)
        i = Util.VerificarPosicaoInicialArray(Campos0001)
        CHV_PAI = Campos0001(dtoTitSPED.t0001("CHV_REG") - i)
        
    ElseIf CHV_PAI_CONTRIBUICOES <> "" Then
        
        CamposC100 = dtoRegSPED.rC100(CHV_PAI_CONTRIBUICOES)
        i = Util.VerificarPosicaoInicialArray(CamposC100)
        CHV_C010 = CamposC100(dtoTitSPED.tC100("CHV_PAI_CONTRIBUICOES") - i)
        
        CamposC010 = dtoRegSPED.rC010(CHV_C010)
        i = Util.VerificarPosicaoInicialArray(CamposC010)
        CNPJ = CamposC010(dtoTitSPED.tC010("CNPJ") - i)
        
        Campos0140 = dtoRegSPED.r0140(CNPJ)
        i = Util.VerificarPosicaoInicialArray(Campos0140)
        CHV_PAI = Campos0140(dtoTitSPED.t0140("CHV_REG") - i)
        
    End If
    
    Extrair_CHV_PAI_0200 = CHV_PAI
    
End Function

Public Sub CarregarDadosXML()

Dim Produtos As IXMLDOMNodeList
Dim NFe As New DOMDocument60
Dim Produto As IXMLDOMNode
Dim NUM_ITEM As Integer
Dim COD_ITEM As String
Dim Comeco As Double
Dim XML As Variant
Dim Msg As String
    
    With Divergencias
        
        .TipoRelatorio = "XML"
        Call .dicOperacoesXML.RemoveAll
        Call .CarregarTitulosRelatorioProdutos
        Call .RedimensionarArray(.arrTitulosRelatorio.Count)
        
        a = 0
        Comeco = Timer
        Call ProcessarDocumentos.ListarTodosDocumentos
        
        For Each XML In DocsFiscais.arrTodos
            
            Call Util.AntiTravamento(a, 50, Msg & "Carregando dados dos XMLS", DocsFiscais.arrTodos.Count, Comeco)
            
            Set NFe = fnXML.RemoverNamespaces(XML)
            Call .ExtrairIdentificacaoNFe(NFe)
            
            Set Produtos = fnXML.ExtrairProdutosNFe(NFe)
            For Each Produto In Produtos
                
                NUM_ITEM = fnXML.ExtrairNumeroItemProduto(Produto)
                COD_ITEM = fnXML.ValidarTag(Produto, "prod/cProd")
                
                Call .ExtrairCadastroProduto(Produto)
                
                .AtribuirValor "NUM_ITEM", NUM_ITEM
                .AtribuirValor "COD_ITEM", "'" & COD_ITEM
                .AtribuirValor "QTD", fnXML.ValidarValores(Produto, "prod/qCom")
                .AtribuirValor "UNID", fnXML.ValidarTag(Produto, "prod/uCom")
                .AtribuirValor "CFOP", fnXML.ValidarTag(Produto, "prod/CFOP")
                .AtribuirValor "CST_ICMS", "'" & fnXML.ExtrairCST_CSOSN_ICMS(Produto)
                .AtribuirValor "VL_ITEM", fnXML.ValidarValores(Produto, "prod/vProd")
                .AtribuirValor "VL_DESC", fnXML.ValidarValores(Produto, "prod/vDesc")
                .AtribuirValor "VL_BC_ICMS", fnXML.ExtrairBaseICMS(Produto)
                .AtribuirValor "ALIQ_ICMS", fnXML.ExtrairAliquotaICMS(Produto)
                .AtribuirValor "VL_ICMS", fnXML.ExtrairValorICMS(Produto)
                .AtribuirValor "VL_BC_ICMS_ST", fnXML.ValidarValores(Produto, "imposto/ICMS//vBCST")
                .AtribuirValor "ALIQ_ST", fnXML.ExtrairAliquotaICMSST(Produto)
                .AtribuirValor "VL_ICMS_ST", fnXML.ExtrairValorICMSST(Produto)
                .AtribuirValor "CST_IPI", ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_IPI(fnXML.ValidarTag(Produto, "imposto/IPI//CST"))
                .AtribuirValor "VL_BC_IPI", fnXML.ValidarValores(Produto, "imposto/IPI//vBC")
                .AtribuirValor "ALIQ_IPI", fnXML.ValidarPercentual(Produto, "imposto/IPI//pIPI")
                .AtribuirValor "VL_IPI", fnXML.ValidarValores(Produto, "imposto/IPI//vIPI")
                .AtribuirValor "CST_PIS", ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(ValidarTag(Produto, "imposto/PIS//CST"))
                .AtribuirValor "VL_BC_PIS", fnXML.ValidarValores(Produto, "imposto/PIS//vBC")
                .AtribuirValor "ALIQ_PIS", fnXML.ValidarPercentual(Produto, "imposto/PIS//pPIS")
                .AtribuirValor "QUANT_BC_PIS", fnXML.ValidarValores(Produto, "imposto/PIS//qBCProd")
                .AtribuirValor "ALIQ_PIS_QUANT", fnXML.ValidarValores(Produto, "imposto/PIS//vAliqProd")
                .AtribuirValor "VL_PIS", fnXML.ValidarValores(Produto, "imposto/PIS//vPIS")
                .AtribuirValor "CST_COFINS", ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(ValidarTag(Produto, "imposto/COFINS//CST"))
                .AtribuirValor "VL_BC_COFINS", fnXML.ValidarValores(Produto, "imposto/COFINS//vBC")
                .AtribuirValor "ALIQ_COFINS", fnXML.ValidarPercentual(Produto, "imposto/COFINS//pCOFINS")
                .AtribuirValor "QUANT_BC_COFINS", fnXML.ValidarValores(Produto, "imposto/COFINS//qBCProd")
                .AtribuirValor "ALIQ_COFINS_QUANT", fnXML.ValidarValores(Produto, "imposto/COFINS//vAliqProd")
                .AtribuirValor "VL_COFINS", fnXML.ValidarValores(Produto, "imposto/COFINS//vCOFINS")
                .AtribuirValor "VL_OPER", .Extrair_VL_OPER_XML(Produto)
                
                If .VerificarExistenciaSPED Then Call .RegistrarOperacaoXML
                
            Next Produto
            
        Next XML
        
    End With
    
End Sub

Private Function CorrelacionarNotas()

Dim chNFe As Variant
    
    With Divergencias
        
        For Each chNFe In .dicOperacoesSPED.Keys()
            
            If .dicOperacoesXML.Exists(chNFe) Then
                
                Call VerificarDiferencaQuantidadeItens(chNFe)
                Call CorrelacionarProdutosNota(chNFe)
                
            Else
                
                Call RegistrarItensSemCorrelacao(chNFe)
                
            End If
            
            Call ResetarDivergenciaItensNota
            
        Next chNFe
        
    End With
    
End Function

Private Sub VerificarDiferencaQuantidadeItens(ByVal chNFe As String)
    
    With Divergencias
        
        DivQtdItensNota.QtdItensSPED = .dicOperacoesSPED(chNFe).Count
        DivQtdItensNota.QtdItensXML = .dicOperacoesXML(chNFe).Count
        
    End With
    
    With DivQtdItensNota
        
        If .QtdItensSPED <> .QtdItensXML Then .QtdDivergente = True Else .QtdDivergente = False
        
    End With
    
End Sub

Private Function CorrelacionarProdutosNota(ByVal chNFe As String)

Dim itemSPED As Variant
    
    With Divergencias
        
        For Each itemSPED In .dicOperacoesSPED(chNFe).Keys()
            
            Call CorrelacionarItensSPED_XML(chNFe, itemSPED)
            
        Next itemSPED
        
    End With
    
End Function

Private Function CorrelacionarItensSPED_XML(ByVal chNFe As String, ByVal itemSPED As Integer)

Const PONTUACAO_MINIMA As Double = 5
Dim MelhorCorrelacao As Integer
Dim MaiorPontuacao As Double
Dim Pontuacao As Double
Dim itemXML As Variant
    
    With Divergencias
        
        CamposSPED = Empty
        CamposSPED = .dicOperacoesSPED(chNFe)(itemSPED)
        For Each itemXML In .dicOperacoesXML(chNFe).Keys()
            
            CamposXML = Empty
            CamposXML = .dicOperacoesXML(chNFe)(itemXML)
            
            Call ResetarVerificacoesCampos
            Pontuacao = CalcularPontuacao
            
            If Pontuacao > MaiorPontuacao Then
                MaiorPontuacao = Pontuacao
                MelhorCorrelacao = itemXML
            End If
            
        Next itemXML
        
        If MelhorCorrelacao = 0 Then
            
            ReDim CamposXML(1 To UBound(CamposSPED))
            Call RegistrarMelhorCorrelacao(CamposSPED, CamposXML)
            
        Else
            
            CamposXML = .dicOperacoesXML(chNFe)(MelhorCorrelacao)
            If MaiorPontuacao >= PONTUACAO_MINIMA Then Call RegistrarMelhorCorrelacao(CamposSPED, CamposXML)
            
            Call .dicOperacoesXML(chNFe).Remove(MelhorCorrelacao)
            
        End If
        
    End With
    
End Function

Private Function CalcularPontuacao() As Double

Dim Pontuacao As Double
    
    Pontuacao = CalcularPontuacaoVL_OPER()
    Pontuacao = Pontuacao + CalcularPontuacaoVL_ITEM()
    Pontuacao = Pontuacao + CalcularPontuacaoDESCR_ITEM()
    Pontuacao = Pontuacao + CalcularPontuacaoNUM_ITEM()
    Pontuacao = Pontuacao + CalcularPontuacaoCOD_NCM()
    Pontuacao = Pontuacao + CalcularPontuacaoEX_IPI()
    Pontuacao = Pontuacao + CalcularPontuacaoCOD_BARRA()
    Pontuacao = Pontuacao + CalcularPontuacaoCEST()
    
    Pontuacao = Pontuacao + CalcularPontuacaoVL_IPI()
    Pontuacao = Pontuacao + CalcularPontuacaoVL_DESC()
    Pontuacao = Pontuacao + CalcularPontuacaoVL_ICMS_ST()
    Pontuacao = Pontuacao + CalcularPontuacaoUNID()
    CalcularPontuacao = Pontuacao + CalcularPontuacaoQTD()
    
End Function

Private Function CalcularPontuacaoVL_OPER() As Double

Dim Posicao As Byte
Dim VL_OPER_SPED As Double, VL_OPER_XML#
Dim Margem As Double
    
    Margem = 0.02
    With Divergencias
                
        .TipoRelatorio = "SPED"
        Posicao = .RetornarPosicaoTitulo("VL_OPER")
        
        VL_OPER_SPED = CamposSPED(Posicao)
        
        .TipoRelatorio = "XML"
        VL_OPER_XML = CamposXML(Posicao)
        
        Select Case True
            
            Case VL_OPER_SPED = VL_OPER_XML
                Check.VL_OPER = True
                CalcularPontuacaoVL_OPER = 5
                
            Case VBA.Abs(VBA.Round(VL_OPER_SPED - VL_OPER_XML, 2)) < Margem
                CalcularPontuacaoVL_OPER = 3
                
        End Select
        
    End With
    
End Function

Private Function CalcularPontuacaoVL_ITEM() As Double

Dim Posicao As Byte
Dim VL_ITEM_SPED As Double, VL_ITEM_XML#
Dim Margem As Double
    
    Margem = 0.02
    With Divergencias
        
        Posicao = .RetornarPosicaoTitulo("VL_ITEM")
        VL_ITEM_SPED = CamposSPED(Posicao)
        VL_ITEM_XML = CamposXML(Posicao)
        
        Select Case True
            
            Case VL_ITEM_SPED = VL_ITEM_XML And Check.VL_OPER
                Check.VL_ITEM = True
                CalcularPontuacaoVL_ITEM = 3
                
            Case VBA.Abs(VBA.Round(VL_ITEM_SPED - VL_ITEM_XML, 2)) < Margem
                CalcularPontuacaoVL_ITEM = 1
                
        End Select
        
    End With
    
End Function

Private Function CalcularPontuacaoDESCR_ITEM() As Double

Dim Posicao As Byte
Dim DESCR_ITEM_SPED As String, DESCR_ITEM_XML$
Dim Margem As Double
    
    Margem = 0.4
    With Divergencias
        
        Posicao = .RetornarPosicaoTitulo("DESCR_ITEM")
        DESCR_ITEM_SPED = LimparTexto(CamposSPED(Posicao))
        DESCR_ITEM_XML = LimparTexto(CamposXML(Posicao))
        
        Select Case True
            
            Case CompararDescricoes(DESCR_ITEM_XML, DESCR_ITEM_SPED) > Margem And Check.VL_OPER And Check.VL_ITEM
                Check.DESCR_ITEM = True
                CalcularPontuacaoDESCR_ITEM = 5
            
            Case CompararDescricoes(DESCR_ITEM_XML, DESCR_ITEM_SPED) > Margem And (Check.VL_OPER Or Check.VL_ITEM)
                Check.DESCR_ITEM = True
                CalcularPontuacaoDESCR_ITEM = 3
                
            Case CompararDescricoes(DESCR_ITEM_XML, DESCR_ITEM_SPED) > Margem
                Check.DESCR_ITEM = False
                CalcularPontuacaoDESCR_ITEM = 1
                
        End Select
        
    End With
    
End Function

Private Function CalcularPontuacaoNUM_ITEM() As Double

Dim Posicao As Byte
Dim NUM_ITEM_SPED As Double, NUM_ITEM_XML#
    
    With Divergencias
        
        Posicao = .RetornarPosicaoTitulo("NUM_ITEM")
        NUM_ITEM_SPED = CamposSPED(Posicao)
        NUM_ITEM_XML = CamposXML(Posicao)
        
        Select Case True
            
            Case NUM_ITEM_SPED = NUM_ITEM_XML And Check.DESCR_ITEM And Check.VL_OPER And Check.VL_ITEM
                Check.NUM_ITEM = True
                CalcularPontuacaoNUM_ITEM = 5
            
            Case NUM_ITEM_SPED = NUM_ITEM_XML And Check.DESCR_ITEM And (Check.VL_OPER Or Check.VL_ITEM)
                Check.NUM_ITEM = True
                CalcularPontuacaoNUM_ITEM = 3
                
            Case NUM_ITEM_SPED = NUM_ITEM_XML
                Check.NUM_ITEM = False
                CalcularPontuacaoNUM_ITEM = 1
                
        End Select
        
    End With
    
End Function

Private Function CalcularPontuacaoCOD_NCM() As Double

Dim Posicao As Byte
Dim COD_NCM_SPED As String, COD_NCM_XML$
    
    With Divergencias
        
        Posicao = .RetornarPosicaoTitulo("COD_NCM")
        COD_NCM_SPED = Util.RemoverAspaSimples(CamposSPED(Posicao))
        COD_NCM_XML = Util.RemoverAspaSimples(CamposXML(Posicao))
        
        Select Case True
            
            Case COD_NCM_SPED = COD_NCM_XML And Check.DESCR_ITEM And Check.VL_OPER And Check.VL_ITEM
                Check.COD_NCM = True
                CalcularPontuacaoCOD_NCM = 5
                
            Case COD_NCM_SPED = COD_NCM_XML And Check.DESCR_ITEM And (Check.VL_OPER Or Check.VL_ITEM)
                Check.COD_NCM = True
                CalcularPontuacaoCOD_NCM = 3
                
            Case COD_NCM_SPED = COD_NCM_XML
                Check.COD_NCM = False
                CalcularPontuacaoCOD_NCM = 1
                
        End Select
        
    End With
    
End Function

Private Function CalcularPontuacaoEX_IPI() As Double

Dim Posicao As Byte
Dim EX_IPI_SPED As String, EX_IPI_XML$
    
    With Divergencias
        
        Posicao = .RetornarPosicaoTitulo("EX_IPI")
        EX_IPI_SPED = Util.RemoverAspaSimples(CamposSPED(Posicao))
        EX_IPI_XML = Util.RemoverAspaSimples(CamposXML(Posicao))
        
        If Util.VerificarStringVazia(EX_IPI_SPED & EX_IPI_XML) Then Exit Function
        
        Select Case True
            
            Case EX_IPI_SPED = EX_IPI_XML And Check.DESCR_ITEM And Check.COD_NCM And Check.VL_OPER And Check.VL_ITEM
                Check.EX_IPI = True
                CalcularPontuacaoEX_IPI = 5
                
            Case EX_IPI_SPED = EX_IPI_XML And Check.DESCR_ITEM And Check.COD_NCM And (Check.VL_OPER Or Check.VL_ITEM)
                Check.EX_IPI = True
                CalcularPontuacaoEX_IPI = 3
                
            Case EX_IPI_SPED = EX_IPI_XML
                Check.EX_IPI = False
                CalcularPontuacaoEX_IPI = 1
                
        End Select
        
    End With
    
End Function

Private Function CalcularPontuacaoCOD_BARRA() As Double

Dim Posicao As Byte
Dim COD_BARRA_SPED As String, COD_BARRA_XML$
    
    With Divergencias
        
        Posicao = .RetornarPosicaoTitulo("COD_BARRA")
        COD_BARRA_SPED = Util.RemoverAspaSimples(CamposSPED(Posicao))
        COD_BARRA_XML = Util.RemoverAspaSimples(CamposXML(Posicao))
        
        If Util.VerificarStringVazia(COD_BARRA_SPED & COD_BARRA_XML) Then Exit Function
        
        Select Case True
            
            Case COD_BARRA_SPED = COD_BARRA_XML And Check.DESCR_ITEM And Check.VL_OPER And Check.VL_ITEM
                Check.COD_BARRA = True
                CalcularPontuacaoCOD_BARRA = 5
                
            Case COD_BARRA_SPED = COD_BARRA_XML And Check.DESCR_ITEM And (Check.VL_OPER Or Check.VL_ITEM)
                Check.COD_BARRA = True
                CalcularPontuacaoCOD_BARRA = 3
                
            Case COD_BARRA_SPED = COD_BARRA_XML
                Check.COD_BARRA = False
                CalcularPontuacaoCOD_BARRA = 1
                
        End Select
        
    End With
    
End Function

Private Function CalcularPontuacaoCEST() As Double

Dim Posicao As Byte
Dim CEST_SPED As String, CEST_XML$
    
    With Divergencias
        
        Posicao = .RetornarPosicaoTitulo("CEST")
        CEST_SPED = Util.RemoverAspaSimples(CamposSPED(Posicao))
        CEST_XML = Util.RemoverAspaSimples(CamposXML(Posicao))
        
        If Util.VerificarStringVazia(CEST_SPED & CEST_XML) Then Exit Function
        
        Select Case True
            
            Case CEST_SPED = CEST_XML And Check.DESCR_ITEM And Check.COD_NCM And Check.VL_OPER And Check.VL_ITEM
                Check.CEST = True
                CalcularPontuacaoCEST = 5
                
            Case CEST_SPED = CEST_XML And Check.DESCR_ITEM And Check.COD_NCM And (Check.VL_OPER Or Check.VL_ITEM)
                Check.CEST = True
                CalcularPontuacaoCEST = 3
                
            Case CEST_SPED = CEST_XML
                CalcularPontuacaoCEST = 1
                
        End Select
        
    End With
    
End Function

Private Function CalcularPontuacaoVL_DESC() As Double

Dim Posicao As Byte
Dim VL_DESC_SPED As Double, VL_DESC_XML#
Dim Margem As Double
    
    Margem = 0.02
    With Divergencias
        
        Posicao = .RetornarPosicaoTitulo("VL_DESC")
        VL_DESC_SPED = CamposSPED(Posicao)
        VL_DESC_XML = CamposXML(Posicao)
        
        Select Case True
            
            Case VL_DESC_SPED = VL_DESC_XML And Check.VL_OPER And Check.VL_ITEM
                Check.VL_DESC = True
                CalcularPontuacaoVL_DESC = 5
                
            Case VL_DESC_SPED = VL_DESC_XML And (Check.VL_OPER Or Check.VL_ITEM)
                Check.VL_DESC = True
                CalcularPontuacaoVL_DESC = 3
                
            Case VBA.Abs(VBA.Round(VL_DESC_SPED - VL_DESC_XML, 2)) < Margem
                Check.VL_DESC = False
                CalcularPontuacaoVL_DESC = 1
                
        End Select
        
    End With
    
End Function

Private Function CalcularPontuacaoVL_IPI() As Double

Dim Posicao As Byte
Dim VL_IPI_SPED As Double, VL_IPI_XML#
Dim Margem As Double
    
    Margem = 0.02
    With Divergencias
        
        Posicao = .RetornarPosicaoTitulo("VL_IPI")
        VL_IPI_SPED = CamposSPED(Posicao)
        VL_IPI_XML = CamposXML(Posicao)
        
        Select Case True
            
            Case VL_IPI_SPED = VL_IPI_XML And Check.VL_OPER And Check.VL_ITEM
                Check.VL_IPI = True
                CalcularPontuacaoVL_IPI = 5
                
            Case VL_IPI_SPED = VL_IPI_XML And (Check.VL_OPER Or Check.VL_ITEM)
                Check.VL_IPI = True
                CalcularPontuacaoVL_IPI = 3
                
            Case VBA.Abs(VBA.Round(VL_IPI_SPED - VL_IPI_XML, 2)) < Margem
                Check.VL_IPI = False
                CalcularPontuacaoVL_IPI = 1
                
        End Select
        
    End With
    
End Function

Private Function CalcularPontuacaoVL_ICMS_ST() As Double

Dim Posicao As Byte
Dim VL_ICMS_ST_SPED As Double, VL_ICMS_ST_XML#
Dim Margem As Double
    
    Margem = 0.02
    With Divergencias
        
        Posicao = .RetornarPosicaoTitulo("VL_ICMS_ST")
        VL_ICMS_ST_SPED = CamposSPED(Posicao)
        VL_ICMS_ST_XML = CamposXML(Posicao)
        
        Select Case True
            
            Case VL_ICMS_ST_SPED = VL_ICMS_ST_XML And Check.VL_OPER And Check.VL_ITEM
                Check.VL_ICMS_ST = True
                CalcularPontuacaoVL_ICMS_ST = 5
                
            Case VL_ICMS_ST_SPED = VL_ICMS_ST_XML And (Check.VL_OPER Or Check.VL_ITEM)
                Check.VL_ICMS_ST = True
                CalcularPontuacaoVL_ICMS_ST = 3
                
            Case VBA.Abs(VBA.Round(VL_ICMS_ST_SPED - VL_ICMS_ST_XML, 2)) < Margem
                Check.VL_ICMS_ST = False
                CalcularPontuacaoVL_ICMS_ST = 1
                
        End Select
        
    End With
    
End Function

Private Function CalcularPontuacaoUNID() As Double

Dim Posicao As Byte
Dim UNID_SPED As String, UNID_XML$
    
    With Divergencias
        
        Posicao = .RetornarPosicaoTitulo("UNID")
        UNID_SPED = CamposSPED(Posicao)
        UNID_XML = CamposXML(Posicao)
        
        Select Case True
            
            Case UNID_SPED = UNID_XML And Check.VL_OPER And Check.VL_ITEM
                Check.UNID = True
                CalcularPontuacaoUNID = 5
                
            Case UNID_SPED = UNID_XML And (Check.VL_OPER Or Check.VL_ITEM)
                Check.UNID = True
                CalcularPontuacaoUNID = 3
                
            Case UNID_SPED = UNID_XML
                Check.UNID = True
                CalcularPontuacaoUNID = 1
                
        End Select
        
    End With
    
End Function

Private Function CalcularPontuacaoQTD() As Double

Dim Posicao As Byte
Dim QTD_SPED As Double, QTD_XML#
    
    With Divergencias
        
        Posicao = .RetornarPosicaoTitulo("QTD")
        QTD_SPED = CamposSPED(Posicao)
        QTD_XML = CamposXML(Posicao)
        
        Select Case True
            
            Case QTD_SPED = QTD_XML And Check.UNID And Check.VL_OPER And Check.VL_ITEM
                Check.QTD = True
                CalcularPontuacaoQTD = 5
                
            Case QTD_SPED = QTD_XML And Check.UNID And (Check.VL_OPER Or Check.VL_ITEM)
                Check.QTD = True
                CalcularPontuacaoQTD = 3
                
            Case QTD_SPED = QTD_XML And Check.UNID
                Check.QTD = True
                CalcularPontuacaoQTD = 1
                
        End Select
        
    End With
    
End Function

Private Function CompararDescricoes(ByVal descXML As String, ByVal descSPED As String) As Double

Dim colXML As New Collection
Dim colSPED As New Collection
Dim colUniao As New Collection
Dim colIntersecao As New Collection
Dim Palavra As Variant
Dim Pontuacao As Double, totUniao#, totIntersecao#
            
    Set colXML = TokenizarPalavras(descXML)
    Set colSPED = TokenizarPalavras(descSPED)
    
    For Each Palavra In colXML
        On Error Resume Next
            colUniao.Add Palavra, Palavra
        On Error GoTo 0
    Next Palavra
    
    For Each Palavra In colSPED
        On Error Resume Next
            colUniao.Add Palavra, Palavra
            If Not IsError(colXML(Palavra)) Then
                colIntersecao.Add Palavra, Palavra
            End If
        On Error GoTo 0
    Next Palavra
    
    totIntersecao = colIntersecao.Count
    totUniao = colUniao.Count
    
    If totIntersecao <> 0 And totUniao <> 0 Then
        Pontuacao = totIntersecao / totUniao
    End If
    
    CompararDescricoes = Pontuacao
    
End Function

Private Function TokenizarPalavras(Texto As String) As Collection

Dim dicPalavras As New Collection
Dim Palavras() As String
Dim i As Integer
    
    Texto = LimparTexto(Texto)
    
    'Divide as palavras usando o espaço como delimitador
    Palavras = Split(Texto, " ")
    
    'Amarazena as palavras no Colletction
    For i = LBound(Palavras) To UBound(Palavras)
        On Error Resume Next
            dicPalavras.Add Palavras(i), Palavras(i)
        On Error GoTo 0
    Next i
    
    'Devolve as palavras tokenizadas como resultado da função
    Set TokenizarPalavras = dicPalavras
    
End Function

Private Function LimparTexto(ByVal Texto As String)

Dim i As Integer

    'Limpa o texto quando lote é informado na descrição do item
    If VBA.InStr(1, Texto, "lote ") > 0 Then
        i = VBA.InStr(1, Texto, "lote ") - 1
        Texto = VBA.Left(Texto, i)
    End If
    
    'Limpa o texto quando a data é informada na descrição do item
    If VBA.InStr(1, Texto, "dt.") > 0 Then
        i = VBA.InStr(1, Texto, "dt.") - 1
        Texto = VBA.Left(Texto, i)
    End If
    
    'Limpa o texto quando a data é informada na descrição do item
    If VBA.InStr(1, Texto, "DtFab ") > 0 Then
        i = VBA.InStr(1, Texto, "DtFab ") - 1
        Texto = VBA.Left(Texto, i)
    End If
    
    'Limpa o texto quando a data é informada na descrição do item
    If VBA.InStr(1, Texto, "vl:") > 0 Then
        i = VBA.InStr(1, Texto, "vl:") - 1
        Texto = VBA.Left(Texto, i)
    End If
    
    'Remove aspa simples
    Texto = VBA.Replace(Texto, "'", "")
    
    'Remove hífen
    Texto = VBA.Replace(Texto, "-", "")
        
    'Remove espaços duplos
    Texto = VBA.Replace(Texto, "  ", " ")
    
    'Remove espaços no início e no final do texto
    Texto = VBA.Trim(Texto)
    
    LimparTexto = Texto
    
End Function

Private Function RegistrarMelhorCorrelacao(ByVal CamposSPED As Variant, ByVal CamposXML As Variant)

Dim Titulo As Variant, TituloSPED, TituloXML
Dim Posicao As Byte
    
    With Divergencias
        
        Call .RedimensionarCampoRelatorioProdutos
        
        For Each Titulo In .arrTitulosRelatorio
            
            If Not TitulosIgnorar Like "*" & Titulo & "*" Then TituloSPED = Titulo & "_SPED" Else TituloSPED = Titulo
                        
            .TipoRelatorio = "SPED"
            Posicao = .RetornarPosicaoTitulo(Titulo)
            
            .TipoRelatorio = "Produtos"
            .AtribuirCorrelacao TituloSPED, CamposSPED(Posicao)
            
            If Not IgnorarCampoProduto(Titulo) Then
                
                TituloXML = Titulo & "_NF"
                
                .TipoRelatorio = "XML"
                Posicao = .RetornarPosicaoTitulo(Titulo)
                
                .TipoRelatorio = "Produtos"
                .AtribuirCorrelacao TituloXML, CamposXML(Posicao)
                
            End If
            
        Next Titulo
        
        If DivQtdItensNota.QtdDivergente Then Call InformarDivergenciaQuantidadeItens
        
        arrRelatorio.Add .CampoRelatorio
        
    End With
    
End Function

Private Function RegistrarItensSemCorrelacao(ByVal chNFe As String)

Dim Titulo As Variant, TituloSPED, CamposSPED
Dim Posicao As Byte
    
    With Divergencias
        
        For Each CamposSPED In .dicOperacoesSPED(chNFe).Items
            
            Call .RedimensionarCampoRelatorioProdutos
            
            For Each Titulo In .arrTitulosRelatorio
                
                If Not TitulosIgnorar Like "*" & Titulo & "*" Then TituloSPED = Titulo & "_SPED" Else TituloSPED = Titulo
                
                .TipoRelatorio = "SPED"
                Posicao = .RetornarPosicaoTitulo(Titulo)
                
                .TipoRelatorio = "Produtos"
                .AtribuirCorrelacao TituloSPED, CamposSPED(Posicao)
                Call InformarXMLNaoIdentificado
                
            Next Titulo
            
            arrRelatorio.Add .CampoRelatorio
            
        Next CamposSPED
        
    End With
    
End Function

Private Sub InformarXMLNaoIdentificado()

    With Divergencias
        
        .AtribuirCorrelacao "INCONSISTENCIA", "O XML dessa operação não foi importado"
        .AtribuirCorrelacao "SUGESTAO", "Inclua o XML dessa operação na pasta e gere o relatório novamente"
        
    End With
    
End Sub

Private Function IgnorarCampoProduto(ByVal Titulo As String) As Boolean
    
    If TitulosIgnorar Like "*" & Titulo & "*" Then IgnorarCampoProduto = True
    
End Function

Private Sub InformarDivergenciaQuantidadeItens()

Dim INCONSISTENCIA As String
    
    With DivQtdItensNota
        
        If .QtdItensSPED > .QtdItensXML Then INCONSISTENCIA = _
            "A quantidade de Itens lançados no SPED (" & .QtdItensSPED & " itens) é maior que a quantidade do XML (" & .QtdItensXML & " itens)"
            
        If .QtdItensSPED < .QtdItensXML Then INCONSISTENCIA = _
            "A quantidade de itens no XML (" & .QtdItensXML & " itens) é maior que os lançamentos no SPED (" & .QtdItensSPED & " itens)"
        
    End With
    
    With Divergencias
        
        .AtribuirCorrelacao "INCONSISTENCIA", INCONSISTENCIA
        .AtribuirCorrelacao "SUGESTAO", "Lançar no SPED a mesma quantidade de itens do XML"
        
    End With
    
End Sub

Public Function IncluirRegistro0150(ByVal NFe As IXMLDOMNode, ByRef CamposSPED As Variant, dicDados0150 As Dictionary, CHV_PAI As String)
    
Dim UltLin As Long
    
    With Campos0150
            
        .REG = "'0150"
        .ARQUIVO = CamposSPED(1)
        .CHV_REG = fnSPED.GerarChaveRegistro(CHV_PAI, Util.ApenasNumeros(CamposSPED(8)))
        .CHV_PAI = CHV_PAI
        .COD_PART = fnXML.ExtrairCNPJEmitente(NFe)
        .NOME = ValidarTag(NFe, "//emit/xNome")
        .COD_PAIS = fnExcel.FormatarTexto(ValidarTag(NFe, "//emit//cPais"))
        .CNPJ = fnExcel.FormatarTexto(ValidarTag(NFe, "//emit/CNPJ"))
        .CPF = fnExcel.FormatarTexto(ValidarTag(NFe, "//emit/CPF"))
        .IE = fnExcel.FormatarTexto(ValidarTag(NFe, "//emit/IE"))
        .COD_MUN = fnExcel.FormatarTexto(ValidarTag(NFe, "//emit//cMun"))
        .SUFRAMA = fnExcel.FormatarTexto(ValidarTag(NFe, "//emit/SUFRAMA"))
        .END = VBA.UCase(ValidarTag(NFe, "//emit//xLgr"))
        .NUM = fnExcel.FormatarTexto(ValidarTag(NFe, "//emit//nro"))
        .COMPL = VBA.UCase(ValidarTag(NFe, "//emit//xCpl"))
        .BAIRRO = VBA.UCase(ValidarTag(NFe, "//emit//xBairro"))
        If .COD_PAIS = "" Then .COD_PAIS = "1058"
                
        dicDados0150(.CHV_REG) = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI, "", "'" & .COD_PART, _
            .NOME, .COD_PAIS, .CNPJ, .CPF, .IE, .COD_MUN, .SUFRAMA, .END, .NUM, .COMPL, .BAIRRO)
        
    End With
            
End Function

Function CalcularPontuacaoRelativa(ByVal ValorNF As Double, ByVal ValorSPED As Double) As Double

Dim DIFERENCA As Double
    
    DIFERENCA = Abs(ValorNF - ValorSPED)
    
    ' Atribui uma pontuação baseada na diferença relativa
    If ValorNF = 0 And ValorSPED = 0 Then
        
        CalcularPontuacaoRelativa = 1
        
    ElseIf ValorNF = 0 Or ValorSPED = 0 Then
    
        CalcularPontuacaoRelativa = 0
        
    Else
    
        CalcularPontuacaoRelativa = 1 - (DIFERENCA / Application.WorksheetFunction.Max(Abs(ValorNF), Abs(ValorSPED)))
    
    End If
    
    ' Garante que a pontuação esteja entre 0 e 1
    If CalcularPontuacaoRelativa < 0 Then
    
        CalcularPontuacaoRelativa = 0
    
    ElseIf CalcularPontuacaoRelativa > 1 Then
        
        CalcularPontuacaoRelativa = 1
    
    End If
    
End Function

Private Function ValidarValores(ByVal ValorNF As Variant, ValorSPED As Variant) As Boolean
    
    If Not IsEmpty(ValorNF) And Not IsEmpty(ValorSPED) Then
        
        If ValorNF > 0 Or ValorSPED > 0 Then ValidarValores = True
        
    End If
    
End Function

Private Function CarregarTagsProdutos()
    
    Call dicTagsProdutos.RemoveAll
    
    'Tags Auxiliares
    dicTagsProdutos.Add "vProd", "prod/vProd"
    dicTagsProdutos.Add "vFrete", "prod/vFrete"
    dicTagsProdutos.Add "vSeg", "prod/vSeg"
    dicTagsProdutos.Add "vOutro", "prod/vOutro"
    dicTagsProdutos.Add "pCredSN", "imposto/ICMS//pCredSN"
    dicTagsProdutos.Add "vCredICMSSN", "imposto/ICMS//vCredICMSSN"
    dicTagsProdutos.Add "pFCP", "imposto/ICMS//pFCP"
    dicTagsProdutos.Add "vFCP", "imposto/ICMS//vFCP"
    dicTagsProdutos.Add "pFCPST", "imposto/ICMS//pFCPST"
    dicTagsProdutos.Add "vFCPST", "imposto/ICMS//vFCPST"
    
    'Tags do Produto
    dicTagsProdutos.Add "COD_ITEM", "prod/cProd"
    dicTagsProdutos.Add "DESCR_ITEM", "prod/xProd"
    dicTagsProdutos.Add "COD_BARRA", "FuncaoExtrair"
    dicTagsProdutos.Add "COD_NCM", "prod/NCM"
    dicTagsProdutos.Add "EX_IPI", "prod/EXTIPI"
    dicTagsProdutos.Add "CEST", "prod/CEST"
    dicTagsProdutos.Add "QTD", "prod/qCom"
    dicTagsProdutos.Add "UNID", "prod/uCom"
    
    'Tags da Operação
    dicTagsProdutos.Add "NUM_ITEM", "FuncaoExtrair"
    dicTagsProdutos.Add "CFOP", "prod/CFOP"
    dicTagsProdutos.Add "CST_ICMS", "FuncaoExtrair"
    dicTagsProdutos.Add "VL_ITEM", "FuncaoExtrair"
    dicTagsProdutos.Add "VL_DESC", "prod/vDesc"
    dicTagsProdutos.Add "VL_BC_ICMS", "FuncaoExtrair" '"imposto/ICMS//vBC"
    dicTagsProdutos.Add "ALIQ_ICMS", "FuncaoExtrair" '"imposto/ICMS//pICMS"
    dicTagsProdutos.Add "VL_ICMS", "FuncaoExtrair" '"imposto/ICMS//vICMS"
    dicTagsProdutos.Add "VL_BC_ICMS_ST", "FuncaoExtrair" '"imposto/ICMS//vBCST"
    dicTagsProdutos.Add "ALIQ_ICMS_ST", "FuncaoExtrair" '"imposto/ICMS//pICMSST"
    dicTagsProdutos.Add "VL_ICMS_ST", "FuncaoExtrair" '"imposto/ICMS//vICMSST"
    dicTagsProdutos.Add "CST_IPI", "imposto/IPI//CST"
    dicTagsProdutos.Add "VL_BC_IPI", "imposto/IPI//vBC"
    dicTagsProdutos.Add "ALIQ_IPI", "imposto/IPI//pIPI"
    dicTagsProdutos.Add "VL_IPI", "imposto/IPI//vIPI"
    dicTagsProdutos.Add "CST_PIS", "imposto/PIS//CST"
    dicTagsProdutos.Add "VL_BC_PIS", "imposto/PIS//vBC"
    dicTagsProdutos.Add "ALIQ_PIS", "imposto/PIS//pPIS"
    dicTagsProdutos.Add "QUANT_BC_PIS", "imposto/PIS/PISQtde/qBCProd"
    dicTagsProdutos.Add "ALIQ_PIS_QUANT", "imposto/PIS/PISQtde/vAliqProd"
    dicTagsProdutos.Add "VL_PIS", "imposto/PIS//vPIS"
    dicTagsProdutos.Add "CST_COFINS", "imposto/COFINS//CST"
    dicTagsProdutos.Add "VL_BC_COFINS", "imposto/COFINS//vBC"
    dicTagsProdutos.Add "ALIQ_COFINS", "imposto/COFINS//pCOFINS"
    dicTagsProdutos.Add "QUANT_BC_COFINS", "imposto/COFINS/COFINSQtde/qBCProd"
    dicTagsProdutos.Add "ALIQ_COFINS_QUANT", "imposto/COFINS/COFINSQtde/vAliqProd"
    dicTagsProdutos.Add "VL_COFINS", "imposto/COFINS//vCOFINS"
    dicTagsProdutos.Add "VL_OPER", "FuncaoExtrair"
    
End Function

Private Sub ResetarVerificacoesCampos()
    
    Dim CamposVazios As VerificacoesCamposProdutos
    LSet Check = CamposVazios
    
End Sub

Private Sub ResetarDivergenciaItensNota()
    
    Dim CamposVazios As DivergenciasQuantidadeItens
    LSet DivQtdItensNota = CamposVazios
    
End Sub
