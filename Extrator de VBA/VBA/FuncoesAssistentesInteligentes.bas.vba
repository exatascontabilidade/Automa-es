Attribute VB_Name = "FuncoesAssistentesInteligentes"
Option Explicit
Option Base 1

Private ValidacoesCFOP As New clsRegrasFiscaisCFOP

'Controles do Assistente de ICMS
Public Sub GerarRelatorioApuracaoICMS()

    Call Assistente.Fiscal.Apuracao.ICMS.GerarApuracaoAssistidaICMS
    'If Otimizacoes.OtimizacoesAtivas And Util.PossuiDadosInformados(ActiveSheet) Then Call Otimizacoes.SugerirCarregamentoTributacao(ActiveSheet)
    
End Sub

Public Function ReprocessarSugestoesApuracaoICMS()
    Call Assistente.Fiscal.Apuracao.ICMS.ReprocessarSugestoes
End Function

Public Function AceitarSugestoesApuracaoICMS()
    Call Assistente.Fiscal.Apuracao.ICMS.AceitarSugestoes
End Function

Public Function IgnorarInconsistenciasApuracaoICMS()
    Call Assistente.Fiscal.Apuracao.ICMS.IgnorarInconsistencias
End Function

Public Function AtualizarRegistrosApuracaoICMS()
    Call Assistente.Fiscal.Apuracao.ICMS.AtualizarRegistros
End Function

'Controles do Assistente de IPI
Public Sub GerarRelatorioApuracaoIPI()
    
    Call Assistente.Fiscal.Apuracao.IPI.GerarApuracaoAssistidaIPI
    'If Otimizacoes.OtimizacoesAtivas And Util.PossuiDadosInformados(ActiveSheet) Then Call Otimizacoes.SugerirCarregamentoTributacao(ActiveSheet)
    
End Sub

Public Function ReprocessarSugestoesApuracaoIPI()
    Call Assistente.Fiscal.Apuracao.IPI.ReprocessarSugestoes
End Function

Public Function AceitarSugestoesApuracaoIPI()
    Call Assistente.Fiscal.Apuracao.IPI.AceitarSugestoes
End Function

Public Function IgnorarInconsistenciasApuracaoIPI()
    Call Assistente.Fiscal.Apuracao.IPI.IgnorarInconsistencias
End Function

Public Function AtualizarRegistrosApuracaoIPI()
    Call Assistente.Fiscal.Apuracao.IPI.AtualizarRegistros
End Function


'Controles do Assistente de PIS/COFINS
Public Sub GerarRelatorioApuracaoPISCOFINS()

    Call Assistente.Fiscal.Apuracao.PISCOFINS.GerarApuracaoAssistidaPISCOFINS
    'If Otimizacoes.OtimizacoesAtivas And Util.PossuiDadosInformados(ActiveSheet) Then Call Otimizacoes.SugerirCarregamentoTributacao(ActiveSheet)
    
End Sub


'Controles do Assistente de Divergencias
Public Sub GerarRelatorioDivergencias()
    Call Assistente.Fiscal.Divergencias.GerarComparativoXMLSPED
End Sub

Public Function ReprocessarSugestoesDivergencias()
    Call Assistente.Fiscal.Divergencias.ReprocessarSugestoes
End Function

Public Function AceitarSugestoesDivergencias()
    Call Assistente.Fiscal.Divergencias.AceitarSugestoesProdutos
End Function

Public Function IgnorarInconsistenciasDivergencias()
    Call Assistente.Fiscal.Divergencias.IgnorarInconsistencias
End Function

Public Function ResetarInconsistenciasDivergencias()
    Call Assistente.ResetarInconsistencias
End Function

Public Sub AtualizarRegistrosDivergencias()
    Call Assistente.Fiscal.Divergencias.AtualizarRegistros
End Sub

'Controles do Assistente de Tributacao
Public Sub GerarRelatorioTributacao()
    Call Assistente.Fiscal.Tributacao.GerarAnaliseTributacao
End Sub

Public Function ReprocessarSugestoesTributacao()
    Call Assistente.Fiscal.Tributacao.ReprocessarSugestoes
End Function

Public Function AceitarSugestoesTributacao()
    Call Assistente.Fiscal.Tributacao.AceitarSugestoes
End Function

Public Function IgnorarInconsistenciasTributacao()
    Call Assistente.Fiscal.Tributacao.IgnorarInconsistencias
End Function

Public Sub AtualizarRegistrosTributacao()
    Call Assistente.Fiscal.Tributacao.AtualizarRegistros
End Sub

'Controles do Assistente de Estoque
Public Sub GerarRelatorioEstoque()
    Call Assistente.Estoque.GerarRelatorio
End Sub

Public Function ReprocessarSugestoesEstoque()
    Call Assistente.Estoque.ReprocessarSugestoes
End Function

Public Function AceitarSugestoesEstoque()
    Call Assistente.Estoque.AceitarSugestoes
End Function

Public Function IgnorarInconsistenciasEstoque()
    Call Assistente.Estoque.IgnorarInconsistencias
End Function

Public Sub AtualizarRegistrosEstoque()
    Call Assistente.Estoque.AtualizarRegistros
End Sub

'Controles do Assistente de PIS/COFINS
Public Function IgnorarInconsistenciasPISCOFINS()
    Call Assistente.Fiscal.Apuracao.PISCOFINS.IgnorarInconsistencias
End Function

'Controles do Assistente Tributário de PIS/COFINS
Public Function SalvarTributacaoPISCOFINS()
    Call Assistente.Tributario.PIS_COFINS.SalvarTributacaoPISCOFINS
End Function

Sub EnviarDados()

Dim Intervalo As Range
    
    Application.EnableEvents = False
        
        With ActiveSheet
            If .FilterMode Then .AutoFilter.ShowAllData
            .[A3].CurrentRegion.Copy
        End With
        
        Workbooks.Add
        ActiveSheet.Paste
        Application.CutCopyMode = False
        
    Application.EnableEvents = True
    
End Sub

Public Sub GerarRelatorioAnaliseProdutos()

Dim dicTitulosC100 As New Dictionary
Dim dicTitulosC170 As New Dictionary
Dim dicTitulosC177 As New Dictionary
Dim dicTitulos0200 As New Dictionary
Dim dicDados0200 As New Dictionary
Dim dicDadosC100 As New Dictionary
Dim dicDadosC177 As New Dictionary
Dim Dados As Range, Linha As Range
Dim arrRelatorio As New ArrayList
Dim dicTitulos As New Dictionary
Dim arrDados As New ArrayList
Dim Campos As Variant
Dim CHV_REG As String
Dim Comeco As Double
Dim a As Long
    
    Inicio = Now()
    Application.StatusBar = "Gerando relatório inteligente de produtos, por favor aguarde..."
    
    Set dicTitulos = Util.MapearTitulos(assApuracaoICMS, 3)
    
    Set dicTitulos0200 = Util.MapearTitulos(reg0200, 3)
    Set dicDados0200 = Util.CriarDicionarioRegistro(reg0200, "ARQUIVO", "COD_ITEM")
    
    Set dicTitulosC100 = Util.MapearTitulos(regC100, 3)
    Set dicDadosC100 = Util.CriarDicionarioRegistro(regC100)
    
    Set dicTitulosC177 = Util.MapearTitulos(regC177, 3)
    Set dicDadosC177 = Util.CriarDicionarioRegistro(regC177, "CHV_PAI_FISCAL")
    
    Set dicTitulosC170 = Util.MapearTitulos(regC170, 3)
    Set Dados = Util.DefinirIntervalo(regC170, 4, 3)
    
    If Dados Is Nothing Then
        Call Util.MsgAlerta("Não existem dados nos registros C170", "Dados indisponíveis")
        Exit Sub
    End If
    
    a = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 100, "Gerando relatório inteligente de produtos, por favor aguarde...", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            arrDados.Add Campos(dicTitulosC170("REG"))
            arrDados.Add Campos(dicTitulosC170("ARQUIVO"))
            arrDados.Add Campos(dicTitulosC170("CHV_PAI_FISCAL"))
            arrDados.Add Campos(dicTitulosC170("CHV_REG"))
            
            CHV_REG = Campos(dicTitulosC170("CHV_PAI_FISCAL"))
            If dicDadosC100.Exists(CHV_REG) Then
                
                arrDados.Add "'" & dicDadosC100(CHV_REG)(dicTitulosC100("CHV_NFE"))
                arrDados.Add "'" & dicDadosC100(CHV_REG)(dicTitulosC100("NUM_DOC"))
                arrDados.Add "'" & dicDadosC100(CHV_REG)(dicTitulosC100("SER"))
                
            Else
                
                arrDados.Add ""
                arrDados.Add ""
                arrDados.Add ""
                
            End If
            
            arrDados.Add "'" & Campos(dicTitulosC170("COD_ITEM"))
            
            'Coleta dados do registro 0200
            CHV_REG = VBA.Join(Array(Campos(dicTitulosC170("ARQUIVO")), Campos(dicTitulosC170("COD_ITEM"))))
            If dicDados0200.Exists(CHV_REG) Then
                
                arrDados.Add dicDados0200(CHV_REG)(dicTitulos0200("DESCR_ITEM"))
                arrDados.Add "'" & dicDados0200(CHV_REG)(dicTitulos0200("COD_BARRA"))
                arrDados.Add "'" & dicDados0200(CHV_REG)(dicTitulos0200("COD_NCM"))
                arrDados.Add "'" & dicDados0200(CHV_REG)(dicTitulos0200("EX_IPI"))
                arrDados.Add "'" & dicDados0200(CHV_REG)(dicTitulos0200("CEST"))
                arrDados.Add dicDados0200(CHV_REG)(dicTitulos0200("TIPO_ITEM"))
                
            Else
                
                arrDados.Add "ITEM NÃO IDENTIFICADO"
                arrDados.Add ""
                arrDados.Add ""
                arrDados.Add ""
                arrDados.Add ""
                arrDados.Add ""
                
            End If
            
            'Coleta dados do registro C177
            CHV_REG = Campos(dicTitulosC170("CHV_REG"))
            If dicDadosC177.Exists(CHV_REG) Then
                
                arrDados.Add dicDadosC177(CHV_REG)(dicTitulosC177("COD_INF_ITEM"))
                
            Else
                
                arrDados.Add ""
                
            End If
            
            arrDados.Add Campos(dicTitulosC170("IND_MOV"))
            arrDados.Add Campos(dicTitulosC170("CFOP"))
            arrDados.Add "'" & Campos(dicTitulosC170("CST_ICMS"))
            arrDados.Add Util.ValidarValores(Campos(dicTitulosC170("VL_ITEM")))
            arrDados.Add 0
            arrDados.Add Util.ValidarValores(Campos(dicTitulosC170("VL_DESC")))
            arrDados.Add Util.ValidarValores(Campos(dicTitulosC170("VL_BC_ICMS")))
            arrDados.Add Util.ValidarValores(Campos(dicTitulosC170("ALIQ_ICMS")))
            arrDados.Add Util.ValidarValores(Campos(dicTitulosC170("VL_ICMS")))
            arrDados.Add Util.ValidarValores(Campos(dicTitulosC170("VL_BC_ICMS_ST")))
            arrDados.Add Campos(dicTitulosC170("ALIQ_ST"))
            arrDados.Add Util.ValidarValores(Campos(dicTitulosC170("VL_ICMS_ST")))
            arrDados.Add Empty
            arrDados.Add Empty
            
        End If
        
        Campos = arrDados.toArray()
        Campos = GerarSugestoesProdutos(Campos, dicTitulos)
        arrRelatorio.Add Campos
        arrDados.Clear
        
    Next Linha
    
    
    Call Util.LimparDados(assApuracaoICMS, 4, False)
    Call Util.ExportarDadosArrayList(assApuracaoICMS, arrRelatorio)
    Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoICMS)

    Call Util.MsgInformativa("Relatório gerado com sucesso", "Relatório Inteligente de Produtos", Inicio)
    
    Application.StatusBar = False
    
End Sub

Public Sub GerarRelatorioPIS_COFINS()

Dim dicTitulos0200 As New Dictionary
Dim dicDados0200 As New Dictionary
Dim arrRelatorio As New ArrayList
Dim dicTitulos As New Dictionary
Dim Msg As String, Titulo$
    
    Msg = "Tem certeza que gostaria gerar o relatorio de apuração?"
    Titulo = "Confirmação de Operação"
    
    If Util.ConfirmarOperacao(Msg, Titulo) = vbNo Then Exit Sub
    
    Call Util.DesabilitarControles
        
        Inicio = Now()
        
        Call Util.AtualizarBarraStatus("Coletando dados dos registros, por favor aguarde...")
        
        Set dicTitulos = Util.MapearTitulos(assApuracaoPISCOFINS, 3)
        Set dicTitulos0200 = Util.MapearTitulos(reg0200, 3)
        Set dicDados0200 = Util.CriarDicionarioRegistro(reg0200)
        
        'Call Assistente.Fiscal.Apuracao.PISCOFINS.CarregarDadosA170(dicDados0200, dicTitulos0200, arrRelatorio, dicTitulos, "Etapa 2/3 [Carregando dados do registro 0200] - ")
        Call Assistente.Fiscal.Apuracao.PISCOFINS.CarregarDadosC170("Etapa 2/3 [Carregando dados do registro C170] - ")
        'Call Assistente.Fiscal.Apuracao.PISCOFINS.CarregarDadosC175("Etapa 2/3 [Carregando dados do registro C175] - ")
        Call Assistente.Fiscal.Apuracao.PISCOFINS.CarregarDadosC191("Etapa 2/3 [Carregando dados do registro C191] - ")
        'Call Assistente.Fiscal.Apuracao.PISCOFINS.CarregarDadosC195(dicDados0200, dicTitulos0200, arrRelatorio, dicTitulos, "Etapa 2/3 [Carregando dados do registro C175] - ")
        Call Assistente.Fiscal.Apuracao.PISCOFINS.CarregarDadosD201(arrRelatorio, dicTitulos, "Etapa 2/3 [Carregando dados do registro D201] - ")
        Call Assistente.Fiscal.Apuracao.PISCOFINS.CarregarDadosD205(arrRelatorio, dicTitulos, "Etapa 2/3 [Carregando dados do registro D205] - ")
        'Call Assistente.Fiscal.Apuracao.PISCOFINS.CarregarDadosF100(dicDados0200, dicTitulos0200, arrRelatorio, dicTitulos, "Etapa 2/3 [Carregando dados do registro F100] - ")
        Call Assistente.Fiscal.Apuracao.PISCOFINS.CarregarDadosF120(arrRelatorio, dicTitulos, "Etapa 2/3 [Carregando dados do registro F120] - ")
        
        Call Util.LimparDados(assApuracaoPISCOFINS, 4, False)
        
        If arrRelatorio.Count > 0 Then
        
            Call Util.AtualizarBarraStatus("Etapa 3/3 - Preparando relatório para exportação, por favor aguarde...")
            Call Util.ExportarDadosArrayList(assApuracaoPISCOFINS, arrRelatorio)
            Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoPISCOFINS)
            
            Call Util.AtualizarBarraStatus("Processo concluído com sucesso!")
            Call Util.MsgInformativa("Relatório gerado com sucesso", "Assistente de Apuração do PIS/COFINS", Inicio)
            
        Else
            
            Msg = "Não foi encontrado nenhum dado para geração do relatório." & vbCrLf & vbCrLf
            Msg = Msg & "Por favor verifique se o SPED e/ou XMLs foram importados e tente novamente."
            Call Util.MsgAlerta(Msg, "Assistente de Apuração do PIS/COFINS")
            
        End If
        
        Application.StatusBar = False
        
    Call Util.HabilitarControles
    
    If Otimizacoes.OtimizacoesAtivas And Util.PossuiDadosInformados(ActiveSheet) Then Call Otimizacoes.SugerirCarregamentoTributacao(ActiveSheet)
    
End Sub

Private Function GerarSugestoesProdutos(ByVal Campos As Variant, ByRef dicTitulos As Dictionary) As Variant

Dim VL_ICMS As Double, VL_BC_ICMS#, VL_ITEM#, VL_DESC#, VL_BC_ICMS_ST#, VL_ICMS_ST#
Dim CFOP As String, CST_ICMS$, TIPO_ITEM$, IND_MOV$, CEST$, COD_NCM$, ALIQ_ICMS$
Dim i As Integer
    
    If dicTabelaCFOP.Count = 0 Then Call ValidacoesCFOP.CarregarTabelaCFOP
    If dicTabelaCEST.Count = 0 Then Call Util.CarregarTabelaCEST(dicTabelaCEST)
    If dicTabelaNCM.Count = 0 Then Call Util.CarregarTabelaNCM(dicTabelaNCM)
        
    If LBound(Campos) = 0 Then i = 1
    
    IND_MOV = Campos(dicTitulos("IND_MOV") - i)
    TIPO_ITEM = Util.ApenasNumeros(Campos(dicTitulos("TIPO_ITEM") - i))
    COD_NCM = VBA.Format(Util.ApenasNumeros(Campos(dicTitulos("COD_NCM") - i)), "0000\.00\.00")
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    CEST = Util.ApenasNumeros(Campos(dicTitulos("CEST") - i))
    CST_ICMS = Util.ApenasNumeros(Campos(dicTitulos("CST_ICMS") - i))
    VL_ITEM = fnExcel.ConverterValores(Campos(dicTitulos("VL_ITEM") - i))
    VL_DESC = fnExcel.ConverterValores(Campos(dicTitulos("VL_DESC") - i))
    VL_BC_ICMS = fnExcel.ConverterValores(Campos(dicTitulos("VL_BC_ICMS") - i))
    VL_ICMS = fnExcel.ConverterValores(Campos(dicTitulos("VL_ICMS") - i))
    ALIQ_ICMS = fnExcel.FormatarPercentuais(Campos(dicTitulos("ALIQ_ICMS") - i))
    VL_BC_ICMS_ST = fnExcel.ConverterValores(Campos(dicTitulos("VL_BC_ICMS_ST") - i))
    VL_ICMS_ST = fnExcel.ConverterValores(Campos(dicTitulos("VL_ICMS_ST") - i))
    
    Select Case True
        
        Case Not dicTabelaNCM.Exists(COD_NCM) And COD_NCM <> ""
            Call Util.GravarSugestao(Campos, dicTitulos, _
                INCONSISTENCIA:="O NCM informado é inválido", _
                SUGESTAO:="Informar um NCM válido")
        
        Case CEST <> "" And VBA.Len(CEST) < 7
            Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA:="O CEST precisa ter 7 dígitos", SUGESTAO:="Adicionar zeros a esquerda do CEST")
                
        'CST_ICMS não informado
        Case CST_ICMS = ""
            Call Util.GravarSugestao(Campos, dicTitulos, "Campo 'CST_ICMS' não informado", "Informe um CST/ICMS para o campo")
                                            
        'CFOP não informado
        Case CFOP = ""
            Call Util.GravarSugestao(Campos, dicTitulos, "Campo 'CFOP' não informado", "Informe um CFOP para o campo")
            
        Case Not dicTabelaCEST.Exists(CEST) And CEST <> ""
            Call Util.GravarSugestao(Campos, dicTitulos, "O CEST informado não existe na tabela CEST", "Informar um CEST válido")

        Case Not dicTabelaCFOP.Exists(CFOP)
            Call Util.GravarSugestao(Campos, dicTitulos, "O CFOP informado não existe na tabela CFOP", "Informar um CFOP válido")
                        
        'CST_ICMS com 4 dígitos ou mais
        Case VBA.Len(CST_ICMS) <> 3
            Call Util.GravarSugestao(Campos, dicTitulos, "Campo 'CST_ICMS' informado com quantidade de dígitos diferente de 3", "Informe um CST/ICMS com 3 dígitos")
            
        'Operação com valor de ICMS e sem alíquota
        Case VL_BC_ICMS >= VL_ITEM - VL_DESC And VBA.Right(CST_ICMS, 2) = "20"
            Campos(dicTitulos("INCONSISTENCIA") - i) = "Operação tributada integralmente com CST_ICMS indicando operação com redução de base do ICMS"
            Campos(dicTitulos("SUGESTAO") - i) = "Alterar últimos 2 dígitos do campo CST_ICMS para 00"
            
        'Operação com valor de ICMS e sem alíquota
        Case VL_ICMS > 0 And ALIQ_ICMS = "0,00%"
            Campos(dicTitulos("INCONSISTENCIA") - i) = "Campo 'VL_ICMS' é maior que zero e campo 'ALIQ_ICMS' igual a zero"
            Campos(dicTitulos("SUGESTAO") - i) = "Definir uma alíquota válida para o campo ALIQ_ICMS"
            Exit Function
        'Operação sujeita a ST com CST_ICMS incorreto
        Case TIPO_ITEM = ""
            Campos(dicTitulos("INCONSISTENCIA") - i) = "O campo 'TIPO_ITEM' não foi informado"
            Campos(dicTitulos("SUGESTAO") - i) = "Definir um tipo de item para o produto"
            
        'Operação sujeita a ST com CST_ICMS incorreto
        Case IND_MOV = ""
            Campos(dicTitulos("INCONSISTENCIA") - i) = "O campo 'IND_MOV' não foi informado"
            Campos(dicTitulos("SUGESTAO") - i) = "Definir se há movimentação física para o produto: 0 - SIM / 1 - NÃO"
            
        'Operação sujeita a ST com CST_ICMS incorreto
        Case VBA.Left(CFOP, 1) < 4 And CFOP Like "#4##" And VL_ICMS > 0 And Not CFOP Like "*411"
            Campos(dicTitulos("INCONSISTENCIA") - i) = "Operação sujeita a ST com aproveitamento de crédito do ICMS"
            Campos(dicTitulos("SUGESTAO") - i) = "Zerar os campos VL_BC_ICMS, ALIQ_ICMS e VL_ICMS"
            
        'Operação com ativo imobilizado e aproveitamento de crédito do ICMS
        Case CFOP Like "*551" And VL_ICMS > 0
            Campos(dicTitulos("INCONSISTENCIA") - i) = "Operação com ativo imobilizado e aproveitamento de crédito do ICMS"
            Campos(dicTitulos("SUGESTAO") - i) = "Zerar os campos VL_BC_ICMS, ALIQ_ICMS e VL_ICMS"
            
        'Operação interna informada com dígito de origem indicando importação
        Case (CST_ICMS = "101" Or CST_ICMS = "102" Or CST_ICMS = "103" Or CST_ICMS = "201" Or CST_ICMS = "202" Or CST_ICMS = "203" Or CST_ICMS = "900") And VBA.Left(CFOP, 1) < 4
            Campos(dicTitulos("INCONSISTENCIA") - i) = "Informado CSOSN em operação de entrada"
            Campos(dicTitulos("SUGESTAO") - i) = "Alterar CST_ICMS para 090"
            
        'Operação interna informada com dígito de origem indicando importação
        Case Not CFOP Like "3*" And CST_ICMS Like "1*"
            Campos(dicTitulos("INCONSISTENCIA") - i) = "Informado dígito de origem de importação 1 para operação interna"
            Campos(dicTitulos("SUGESTAO") - i) = "Alterar dígito de origem do CST_ICMS para 2"
            
        'Operação interna informada com dígito de origem indicando importação
        Case Not CFOP Like "3*" And CST_ICMS Like "6*"
            Campos(dicTitulos("INCONSISTENCIA") - i) = "Informado dígito de origem de importação 6 para operação interna"
            Campos(dicTitulos("SUGESTAO") - i) = "Alterar dígito de origem do CST_ICMS para 7"
        
        'Operação sujeita a ST com CST_ICMS incorreto
        Case CFOP Like "#4##" And Not CFOP Like "*411" And Not CST_ICMS Like "*60"
            Campos(dicTitulos("INCONSISTENCIA") - i) = "Operação sujeita a ST com CST_ICMS inconsistente"
            Campos(dicTitulos("SUGESTAO") - i) = "Alterar últimos 2 dígitos do CST_ICMS para 60"
            
        'Operação com ativo imobilizado com CST_ICMS inconsistente
        Case CFOP Like "*551" And Not CST_ICMS Like "*90"
            Campos(dicTitulos("INCONSISTENCIA") - i) = "Operação com ativo imobilizado CST_ICMS inconsistente"
            Campos(dicTitulos("SUGESTAO") - i) = "Alterar últimos 2 dígitos do CST_ICMS para 90"
        
    End Select
    
    If CFOP <> "" Then
    
        Select Case True
        
            Case CFOP Like "*403" Or CFOP Like "*102"
            
                'Verifica VL_ICMS_ST
                If VL_ICMS_ST > 0 Then
                    Call Util.GravarSugestao(Campos, dicTitulos, _
                        INCONSISTENCIA:="O valor do campo 'VL_ICMS_ST' deve ser somado ao campo 'VL_ITEM' para operações de compra para revenda", _
                        SUGESTAO:="Somar valor do campo VL_ICMS_ST ao campo VL_ITEM")
                                            
                End If
            
            'Operações de uso e consumo com ST e combustíveis
            Case CFOP Like "*407" Or CFOP Like "*653"
                
                'Verifica VL_ICMS
                If VL_ICMS > 0 Then Call Util.GravarSugestao(Campos, dicTitulos, "Aproveitamento do VL_ICMS em operação de uso e consumo", "Zerar os campos VL_BC_ICMS, ALIQ_ICMS e VL_ICMS")
                
                'Verifica VL_ICMS_ST
                If VL_ICMS_ST > 0 Then
                    Call Util.GravarSugestao(Campos, dicTitulos, _
                        INCONSISTENCIA:="O valor do campo 'VL_ICMS_ST' deve ser somado ao campo 'VL_ITEM' para operações de uso e consumo", _
                        SUGESTAO:="Somar valor do campo 'VL_ICMS_ST' ao campo 'VL_ITEM'")
                                            
                End If
                
                'Verifica TIPO_ITEM
                If VBA.Left(TIPO_ITEM, 2) <> "07" Then Call Util.GravarSugestao(Campos, dicTitulos, "TIPO_ITEM informado não é de uso e consumo", "Alterar TIPO_ITEM para 07")
                
                'Verifica CST_ICMS
                If Not CST_ICMS Like "*60" Then Call Util.GravarSugestao(Campos, dicTitulos, "CST_ICMS inconsistente para operação de uso e consumo com ST", "Alterar últimos 2 dígitos do CST_ICMS para 60")
                
            'Operações de uso e consumo
            Case CFOP Like "*556"
            
                'Verifica VL_ICMS
                If VL_ICMS > 0 Then Call Util.GravarSugestao(Campos, dicTitulos, "aproveitamento do VL_ICMS em operação de uso e consumo", "Zerar os campos VL_BC_ICMS, ALIQ_ICMS e VL_ICMS")
                
                'Verifica VL_ICMS_ST
                If VL_ICMS_ST > 0 Then
                    Call Util.GravarSugestao(Campos, dicTitulos, _
                        INCONSISTENCIA:="O valor do campo 'VL_ICMS_ST' deve ser somado ao campo 'VL_ITEM' para operações de uso e consumo", _
                        SUGESTAO:="Somar valor do campo 'VL_ICMS_ST' ao campo 'VL_ITEM'")
                                            
                End If
                
                'Verifica TIPO_ITEM
                If VBA.Left(TIPO_ITEM, 2) <> "07" Then Call Util.GravarSugestao(Campos, dicTitulos, "TIPO_ITEM informado não é de uso e consumo", "Alterar TIPO_ITEM para 07")
                
                'Verifica CST_ICMS
                If Not CST_ICMS Like "*90" Then Call Util.GravarSugestao(Campos, dicTitulos, "CST_ICMS inconsistente para operação de uso e consumo", "Alterar últimos 2 dígitos do CST_ICMS para 90")
                                        
            Case CFOP < 4000 And (CFOP Like "#10#" Or CFOP Like "#11#" Or CFOP Like "#12#")
            
                If CST_ICMS Like "*00" And VL_ICMS = 0 Then
                    Call Util.GravarSugestao(Campos, dicTitulos, _
                            INCONSISTENCIA:="Operação de compra tributada integralmente (últimos 2 dígitos do campo 'CST_ICMS' = 00) sem aproveitamento de crédito do ICMS ('VL_ICMS' = 0)", _
                            SUGESTAO:="Alterar últimos 2 dígitos do campo 'CST_ICMS' para '90'")
                End If

            Case CFOP < 4000 And CFOP Like "#65#"
                
                If VL_ICMS > 0 Then
                    Call Util.GravarSugestao(Campos, dicTitulos, _
                        INCONSISTENCIA:="Aproveitamento do VL_ICMS em operação com combustíveis e lubrificantes", _
                        SUGESTAO:="Zerar os campos VL_BC_ICMS, ALIQ_ICMS e VL_ICMS")
                End If
                
                If Not CST_ICMS Like "*60" Then
                    Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="CST_ICMS inconsistente para operação de aquisição de combustíveis e lubrificantes", _
                    SUGESTAO:="Alterar últimos 2 dígitos do CST_ICMS para 60")
                    
                End If

        End Select
    
    End If
    
    If CST_ICMS <> "" Then
        
        Select Case True
            
            Case CST_ICMS Like "*10" And VL_ICMS > 0
                Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="Operação sujeita a ST (últimos 2 dígitos do campo 'CST_ICMS' = 10) com aproveitamento de crédito do ICMS ('VL_ICMS' > 0)", _
                    SUGESTAO:="Zerar os campos VL_BC_ICMS, ALIQ_ICMS e VL_ICMS")
                    
            Case CST_ICMS Like "*60" And VL_ICMS > 0
                Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="Operação sujeita a ST (últimos 2 dígitos do campo 'CST_ICMS' = 60) com aproveitamento de crédito do ICMS ('VL_ICMS' > 0)", _
                    SUGESTAO:="Zerar os campos VL_BC_ICMS, ALIQ_ICMS e VL_ICMS")
            
            Case CST_ICMS Like "*10" And Not CFOP Like "#4##" And Not CFOP Like "#6##" And Not CFOP Like "#9##"
                Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="Operação sujeita a ST (últimos 2 dígitos do campo 'CST_ICMS' = 10) com CFOP inconsistente", _
                    SUGESTAO:="Alterar CFOP para representar uma operação sujeita a ST")
                    
            Case CST_ICMS Like "*30" And Not CFOP Like "#4##" And Not CFOP Like "#6##" And Not CFOP Like "#9##"
                Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="Operação sujeita a ST (últimos 2 dígitos do campo 'CST_ICMS' = 30) com CFOP inconsistente", _
                    SUGESTAO:="Alterar CFOP para representar uma operação sujeita a ST")
                
            Case CST_ICMS Like "*60" And Not CFOP Like "#4##" And Not CFOP Like "#6##" And Not CFOP Like "#9##"
                Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="Operação sujeita a ST (últimos 2 dígitos do campo 'CST_ICMS' = 60) com CFOP inconsistente", _
                    SUGESTAO:="Alterar CFOP para representar uma operação sujeita a ST")
                
        End Select
        
    End If
    
    GerarSugestoesProdutos = Campos
    
End Function

Private Function GerarSugestoesContas(ByVal Campos As Variant, ByRef dicTitulos As Dictionary) As Variant

Dim VL_ICMS As Double, ALIQ_ICMS#, VL_BC_ICMS#, VL_ITEM#, VL_DESC#
Dim CHV_REG_C140 As String, IND_PGTO$
Dim i As Integer
    
    If LBound(Campos) = 0 Then i = 1
    
    CHV_REG_C140 = Campos(dicTitulos("CHV_REG_C140") - i)
    IND_PGTO = Campos(dicTitulos("IND_PGTO") - i)
        
    If CHV_REG_C140 = "" Then
    
        Select Case True
            
            Case IND_PGTO Like "1*"
                Call Util.GravarSugestao(Campos, dicTitulos, _
                        INCONSISTENCIA:="Operação a Prazo (IND_PGTO = 1) sem informações de fatura (C140) e vencimento (C141)", _
                        SUGESTAO:="Gerar informações de fatura e vencimento")
                                
        End Select
        
    End If
    
    GerarSugestoesContas = Campos
    
End Function

Private Function GerarSugestoesInventario(ByVal Campos As Variant, ByRef dicTitulos As Dictionary) As Variant

Dim DT_INV As String, MOT_INV$, DESCR_ITEM$, TIPO_ITEM$, IND_PROP$, COD_CTA$, UNID_INV$, COD_ITEM$, ARQUIVO$, UNID_0200$, CHV_0200$
Dim QTD As Double, VL_ITEM#, VL_UNIT#, VL_ITEM_CALC#, DIF_VL_ITEM#
Dim i As Integer
    
    If LBound(Campos) = 0 Then i = 1
    
    ARQUIVO = Campos(dicTitulos("ARQUIVO") - i)
    UNID_INV = Campos(dicTitulos("UNID") - i)
    DT_INV = Campos(dicTitulos("DT_INV") - i)
    MOT_INV = Campos(dicTitulos("MOT_INV") - i)
    COD_ITEM = Util.RemoverAspaSimples(Campos(dicTitulos("COD_ITEM") - i))
    DESCR_ITEM = Campos(dicTitulos("DESCR_ITEM") - i)
    TIPO_ITEM = Campos(dicTitulos("TIPO_ITEM") - i)
    QTD = fnExcel.ConverterValores(Campos(dicTitulos("QTD") - i), True, 3)
    VL_ITEM = fnExcel.ConverterValores(Campos(dicTitulos("VL_ITEM") - i), True, 2)
    VL_UNIT = fnExcel.ConverterValores(Campos(dicTitulos("VL_UNIT") - i), True, 6)
    IND_PROP = Campos(dicTitulos("IND_PROP") - i)
    COD_CTA = Campos(dicTitulos("COD_CTA") - i)
    CHV_0200 = Util.UnirCampos(ARQUIVO, COD_ITEM)
    VL_ITEM_CALC = fnExcel.ConverterValores(QTD * VL_UNIT, True, 2)
    DIF_VL_ITEM = VBA.Abs(fnExcel.ConverterValores(VL_ITEM - VL_ITEM_CALC, True, 2))
    
    If SPEDFiscal.dicDados0200.Exists(CHV_0200) Then UNID_0200 = SPEDFiscal.dicDados0200(CHV_0200)(SPEDFiscal.dicTitulos0200("UNID_INV"))
    
    Select Case True
        
        Case TIPO_ITEM Like "07*"
            Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="Itens de uso e consumo (TIPO_ITEM = 07) normalmente não são informados no inventário", _
                    SUGESTAO:="Excluir item do registro de inventário")
                    
        Case TIPO_ITEM Like "08*"
            Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="Itens de ativo imobilizado (TIPO_ITEM = 08) normalmente não são informados no inventário", _
                    SUGESTAO:="Excluir item do registro de inventário")
    
        Case TIPO_ITEM Like "09*"
            Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="Itens de serviço (TIPO_ITEM = 09) normalmente não são informados no inventário", _
                    SUGESTAO:="Excluir item do registro de inventário")
                    
        Case DESCR_ITEM = ""
            Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="Item sem cadastro no registro 0200", _
                    SUGESTAO:="Importar o cadastro de produtos no registro 0200")
                
        Case IND_PROP = ""
            Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="Campo 'IND_PROP' não informado", _
                    SUGESTAO:="Informar o código indicador de propriedade")

        Case TIPO_ITEM Like ""
            Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="Tipo do item não foi informado para o produto", _
                    SUGESTAO:="Informe um tipo de item para o produto")
                    
        Case QTD < 0
            Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="Itens com quantidades negativas não devem ser informados no inventário", _
                    SUGESTAO:="Excluir item do registro de inventário")
                    
        Case QTD = 0
            Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="Itens sem quantidade não devem ser informados no inventário", _
                    SUGESTAO:="Excluir item do registro de inventário")

        Case VL_UNIT < 0
            Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="O campo 'VL_UNIT' possui valor menor que 0 (zero)", _
                    SUGESTAO:="Informe um valor para o campo VL_UNIT")
                    
        Case VL_UNIT = 0
            Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="O campo 'VL_UNIT' possui valor 0 (zero)", _
                    SUGESTAO:="Informe um valor para o campo VL_UNIT")
                    
        Case VL_ITEM < VL_ITEM_CALC
            Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="Valor do campo VL_ITEM menor que o calculado [Diferença: R$ " & DIF_VL_ITEM & "]", _
                    SUGESTAO:="Recalcular o campo VL_ITEM")
                    
        Case VL_ITEM > VL_ITEM_CALC
            Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="Valor do campo VL_ITEM maior que o calculado [Diferença: R$ " & DIF_VL_ITEM & "]", _
                    SUGESTAO:="Recalcular o campo VL_ITEM")
                    
        Case VL_ITEM = 0
            Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="O campo 'VL_ITEM' possui valor 0 (zero)", _
                    SUGESTAO:="Recalcular o campo VL_ITEM")
                    
        Case UNID_INV <> UNID_0200
            Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="Unidade do H010 (" & UNID_INV & ") diferente do 0200 (" & UNID_0200 & ")", _
                    SUGESTAO:="Informar a mesma unidade do cadastro (0200) para o inventário (H010)")
                                        
        Case DT_INV = ""
            Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="Data do inventário não informada", _
                    SUGESTAO:="Informar a data do inventário, normalmente o último dia do ano anterior ao período do arquivo atual")
                            
        Case MOT_INV = ""
            Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="Motivo do inventário não informado", _
                    SUGESTAO:="Informar o motivo do inventário, normalmente '01' para apresentação do inventário anual")
                                                            
        Case COD_CTA = ""
            Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="Campo 'COD_CTA' não informado", _
                    SUGESTAO:="Informar o código da conta analítica do estoque")
                                        
    End Select

    GerarSugestoesInventario = Campos
    
End Function

Public Sub GerarRelatorioCustosPrecos()

Dim VL_MERC As Double, VL_FRT#, VL_SEG#, VL_OUT_DA#, VL_DESP#, VL_OPR#, Valor#, VL_MED#, VL_CUSTO_MED#, VL_PRECO_MED#, QTD_ENT#, QTD_SAI#, VL_CUSTO#, VL_PRECO#, VL_RESULTADO_MED#, VL_MARGEM#, VL_PRECO_MIN#
Dim CHV_REG As String, CFOP$, UF$, COD_ITEM$
Dim CUSTO As Integer, PRECO&, QTD&
Dim dicTitulos0150 As New Dictionary
Dim dicTitulos0200 As New Dictionary
Dim dicTitulosC100 As New Dictionary
Dim dicTitulosC170 As New Dictionary
Dim dicDados0150 As New Dictionary
Dim dicDados0200 As New Dictionary
Dim dicDadosC100 As New Dictionary
Dim dicDadosC170 As New Dictionary
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim dicDados As New Dictionary
Dim arrDados As New ArrayList
Dim Campos As Variant, MarkUp
    
    MarkUp = InputBox("Insira o markup: ", "Markup")
    If MarkUp = "" Then Exit Sub

    Inicio = Now()
    Application.StatusBar = "Gerando relatório inteligente de Assistente de Custos e Preços, por favor aguarde..."
    
    Set dicTitulos0150 = Util.MapearTitulos(reg0150, 3)
    Set dicDados0150 = Util.CriarDicionarioRegistro(reg0150, "ARQUIVO", "COD_PART")
    
    Set dicTitulos0200 = Util.MapearTitulos(reg0200, 3)
    Set dicDados0200 = Util.CriarDicionarioRegistro(reg0200, "ARQUIVO", "COD_ITEM")
    
    Set dicTitulosC100 = Util.MapearTitulos(regC100, 3)
    Set dicDadosC100 = Util.CriarDicionarioRegistro(regC100)
    
    Set dicTitulosC170 = Util.MapearTitulos(regC170, 3)
    Set dicDadosC170 = Util.CriarDicionarioRegistro(regC170)
    
    Set dicTitulos = Util.MapearTitulos(relCustosPrecos, 3)
    Set dicTitulosC170 = Util.MapearTitulos(regC170, 3)
    Set Dados = Util.DefinirIntervalo(regC170, 4, 3)
    If Dados Is Nothing Then Exit Sub
    
    a = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
    
        Call Util.AntiTravamento(a, 100, "Gerando relatório inteligente de Assistente de Custos e Preços, por favor aguarde...", Dados.Rows.Count, Comeco)
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            'Coleta dados do registro C100
            CHV_REG = Campos(dicTitulosC170("CHV_PAI_FISCAL"))
            
            CFOP = Campos(dicTitulosC170("CFOP"))
            If CFOP Like "#1##" Or CFOP Like "#4##" Then
                
                If dicDadosC100.Exists(CHV_REG) Then
                    VL_OPR = CDbl(Campos(dicTitulosC170("VL_ITEM")))
                    VL_MERC = dicDadosC100(CHV_REG)(dicTitulosC100("VL_MERC"))
                    VL_FRT = dicDadosC100(CHV_REG)(dicTitulosC100("VL_FRT"))
                    VL_SEG = dicDadosC100(CHV_REG)(dicTitulosC100("VL_SEG"))
                    VL_OUT_DA = dicDadosC100(CHV_REG)(dicTitulosC100("VL_OUT_DA"))
                    VL_DESP = VL_FRT + VL_SEG + VL_OUT_DA
                    If VL_MERC > 0 Then VL_OPR = VBA.Round((VL_OPR / VL_MERC) * VL_DESP + VL_OPR, 2)
                End If
                
                'Coleta dados do registro 0200
                CHV_REG = VBA.Join(Array(Campos(dicTitulosC170("ARQUIVO")), Campos(dicTitulosC170("COD_ITEM"))))
                COD_ITEM = Campos(dicTitulosC170("COD_ITEM"))
                arrDados.Add COD_ITEM
                
                If dicDados0200.Exists(CHV_REG) Then
                    
                    arrDados.Add dicDados0200(CHV_REG)(dicTitulos0200("DESCR_ITEM"))
                    arrDados.Add dicDados0200(CHV_REG)(dicTitulos0200("TIPO_ITEM"))
                    
                Else
                
                    arrDados.Add ""
                    arrDados.Add ""
                    
                End If
                
                QTD = Campos(dicTitulosC170("QTD"))
                
                If CFOP < 4000 Then
                    QTD_ENT = QTD
                    VL_CUSTO = VL_OPR
                    If QTD_ENT > 0 Then VL_CUSTO_MED = VBA.Round(VL_OPR / QTD_ENT, 2)
                Else
                    QTD_SAI = QTD
                    VL_PRECO = VL_OPR
                    If QTD_SAI > 0 Then VL_PRECO_MED = VBA.Round(VL_OPR / QTD_SAI, 2)
                End If
                
                VL_PRECO_MIN = VBA.Round(VL_CUSTO_MED / (1 - MarkUp / 100), 2)
                                                                                
                arrDados.Add QTD_ENT
                arrDados.Add VL_CUSTO
                arrDados.Add VL_CUSTO_MED
                arrDados.Add QTD_SAI
                arrDados.Add VL_PRECO
                arrDados.Add VL_PRECO_MED
                arrDados.Add VL_PRECO_MIN
                arrDados.Add VL_RESULTADO_MED
                arrDados.Add VL_MARGEM
                
                Campos = arrDados.toArray()
                
                CHV_REG = COD_ITEM
                If dicDados.Exists(CHV_REG) Then
                     
                    QTD_ENT = dicTitulos("QTD_ENT") - 1
                    QTD_SAI = dicTitulos("QTD_SAI") - 1
                    VL_CUSTO = dicTitulos("VL_CUSTO") - 1
                    VL_PRECO = dicTitulos("VL_PRECO") - 1
                    VL_CUSTO_MED = dicTitulos("VL_CUSTO_MED") - 1
                    VL_PRECO_MED = dicTitulos("VL_PRECO_MED") - 1
                    VL_PRECO_MIN = dicTitulos("VL_PRECO_MIN") - 1
                    VL_RESULTADO_MED = dicTitulos("VL_RESULTADO_MED") - 1
                    VL_MARGEM = dicTitulos("VL_MARGEM") - 1
                    
                    Campos(QTD_ENT) = dicDados(CHV_REG)(QTD_ENT) + CDbl(Campos(QTD_ENT))
                    Campos(QTD_SAI) = dicDados(CHV_REG)(QTD_SAI) + CDbl(Campos(QTD_SAI))
                    Campos(VL_CUSTO) = dicDados(CHV_REG)(VL_CUSTO) + CDbl(Campos(VL_CUSTO))
                    Campos(VL_PRECO) = dicDados(CHV_REG)(VL_PRECO) + CDbl(Campos(VL_PRECO))
                    If CDbl(Campos(QTD_ENT)) > 0 Then Campos(VL_CUSTO_MED) = VBA.Round(CDbl(Campos(VL_CUSTO)) / CDbl(Campos(QTD_ENT)), 2)
                    If CDbl(Campos(QTD_SAI)) > 0 Then Campos(VL_PRECO_MED) = VBA.Round(CDbl(Campos(VL_PRECO)) / CDbl(Campos(QTD_SAI)), 2)
                    Campos(VL_PRECO_MIN) = VBA.Round(CDbl(Campos(VL_CUSTO_MED)) / (1 - MarkUp / 100), 2)
                    
                    If Campos(VL_PRECO_MED) > 0 And Campos(VL_CUSTO_MED) > 0 Then
                        Campos(VL_RESULTADO_MED) = CDbl(Campos(VL_PRECO_MED)) - CDbl(Campos(VL_PRECO_MIN))
                        Campos(VL_MARGEM) = VBA.Round(Campos(VL_RESULTADO_MED) / Campos(VL_PRECO_MIN), 4)
                    End If
                    
                End If
                
                'Call fnExcel.DefinirTipoCampos(Campos, dicTitulos)
                dicDados(CHV_REG) = Campos
                
            End If
            
            arrDados.Clear
            QTD = 0: QTD_ENT = 0: QTD_SAI = 0: VL_CUSTO = 0: VL_PRECO = 0: VL_CUSTO_MED = 0: VL_PRECO_MED = 0: VL_MARGEM = 0: VL_RESULTADO_MED = 0: VL_PRECO_MIN = 0
            
        End If
        
    Next Linha
    
    Call Util.LimparDados(relCustosPrecos, 4, False)
    Call Util.ExportarDadosDicionario(relCustosPrecos, dicDados)
    Call Util.MsgInformativa("Relatório gerado com sucesso", "Relatório Inteligente de Assistente de Custos e Preços", Inicio)
    
    Application.StatusBar = False
    
End Sub

Public Sub AtualizarRegistrosProdutos()

Dim Campos As Variant, Campos0200, CamposC100, CamposC170, CamposC177, dicCampos, regCampo
Dim CHV_C170 As String, CHV_C177$, CHV_C100$, CHV_0200$
Dim dicTitulos0200 As New Dictionary
Dim dicTitulosC100 As New Dictionary
Dim dicTitulosC170 As New Dictionary
Dim dicTitulosC177 As New Dictionary
Dim dicDados0200 As New Dictionary
Dim dicDadosC100 As New Dictionary
Dim dicDadosC170 As New Dictionary
Dim dicDadosC177 As New Dictionary
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
    
    Inicio = Now()
    Application.StatusBar = "Preparando dados para atualização do SPED, por favor aguarde..."
    
    Campos0200 = Array("REG", "COD_BARRA", "COD_NCM", "EX_IPI", "CEST", "TIPO_ITEM")
    CamposC100 = Array("CHV_NFE", "NUM_DOC", "SER")
    CamposC170 = Array("IND_MOV", "CFOP", "VL_ITEM", "CST_ICMS", "VL_BC_ICMS", "ALIQ_ICMS", "VL_ICMS", "VL_BC_ICMS_ST", "ALIQ_ST", "VL_ICMS_ST")
    CamposC177 = Array("COD_INF_ITEM")
    
    Set dicTitulos0200 = Util.MapearTitulos(reg0200, 3)
    Set dicDados0200 = Util.CriarDicionarioRegistro(reg0200, "ARQUIVO", "COD_ITEM")
    
    Set dicTitulosC100 = Util.MapearTitulos(regC100, 3)
    Set dicDadosC100 = Util.CriarDicionarioRegistro(regC100)
    
    Set dicTitulosC170 = Util.MapearTitulos(regC170, 3)
    Set dicDadosC170 = Util.CriarDicionarioRegistro(regC170)
    
    Set dicTitulos = Util.MapearTitulos(assApuracaoICMS, 3)
    If assApuracaoICMS.AutoFilterMode Then assApuracaoICMS.AutoFilter.ShowAllData
    
    Set Dados = Util.DefinirIntervalo(assApuracaoICMS, 4, 3)
    If Dados Is Nothing Then Exit Sub
    
    a = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
    
        Call Util.AntiTravamento(a, 100, "Preparando dados para atualização do SPED, por favor aguarde...", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            CHV_0200 = VBA.Join(Array(Campos(dicTitulos("ARQUIVO")), Campos(dicTitulos("COD_ITEM"))))
            CHV_C100 = Campos(dicTitulos("CHV_PAI_FISCAL"))
            CHV_C170 = Campos(dicTitulos("CHV_REG"))
            
            'Atualizar dados do 0200
            If dicDados0200.Exists(CHV_0200) Then
                
                dicCampos = dicDados0200(CHV_0200)
                For Each regCampo In Campos0200
                    
                    If regCampo = "CEST" Or regCampo = "COD_BARRA" Or regCampo = "COD_NCM" Or regCampo = "EX_TIPI" Then
                        Campos(dicTitulos(regCampo)) = Util.FormatarTexto(Campos(dicTitulos(regCampo)))
                    End If
                    
                    If regCampo = "REG" Then Campos(dicTitulos(regCampo)) = "'0200"
                    dicCampos(dicTitulos0200(regCampo)) = Campos(dicTitulos(regCampo))
                    
                Next regCampo
                
                dicDados0200(CHV_0200) = dicCampos
                
            End If
            
            'Atualizar dados do C100
            If dicDadosC100.Exists(CHV_C100) Then
                
                dicCampos = dicDadosC100(CHV_C100)
                For Each regCampo In CamposC100
                    
                    dicCampos(dicTitulosC100(regCampo)) = Campos(dicTitulos(regCampo))
                    
                Next regCampo
                
                dicDadosC100(CHV_C100) = dicCampos
                
            End If
            
            'Atualizar dados do C170
            If dicDadosC170.Exists(CHV_C170) Then
                
                dicCampos = dicDadosC170(CHV_C170)
                For Each regCampo In CamposC170
                    
                    If regCampo = "CST_ICMS" Or regCampo = "COD_BARRA" Then
                        Campos(dicTitulos(regCampo)) = fnExcel.FormatarTexto(Campos(dicTitulos(regCampo)))
                    End If
                    
                    If regCampo Like "VL_*" Then Campos(dicTitulos(regCampo)) = VBA.Round(Campos(dicTitulos(regCampo)), 2)
                    dicCampos(dicTitulosC170(regCampo)) = Campos(dicTitulos(regCampo))
                
                Next regCampo
                
                dicDadosC170(CHV_C170) = dicCampos
                
            End If
                        
            'Atualizar dados do C177
            CHV_C177 = fnSPED.GerarChaveRegistro(CHV_C170, "C177")
            If dicDadosC177.Exists(CHV_C177) Then
                
                dicCampos = dicDadosC177(CHV_C177)
                For Each regCampo In CamposC177

                    dicCampos(dicTitulosC177(regCampo)) = Campos(dicTitulos(regCampo))
                    
                Next regCampo
                
                dicDadosC177(CHV_C177) = dicCampos
                
            ElseIf Campos(dicTitulos("COD_INF_ITEM")) <> "" Then
                
                dicCampos = Array("C177", Campos(dicTitulos("ARQUIVO")), CHV_C177, CHV_C170, Campos(dicTitulos("COD_INF_ITEM")))
                dicDadosC177(CHV_C177) = dicCampos
                
            End If
            
        End If

    Next Linha
    
    Application.StatusBar = "Atualizando dados do registro 0200, por favor aguarde..."
    Call Util.LimparDados(reg0200, 4, False)
    Call Util.ExportarDadosDicionario(reg0200, dicDados0200)
    
    Application.StatusBar = "Atualizando dados do registro C100, por favor aguarde..."
    Call Util.LimparDados(regC100, 4, False)
    Call Util.ExportarDadosDicionario(regC100, dicDadosC100)
    
    Application.StatusBar = "Atualizando dados do registro C170, por favor aguarde..."
    Call Util.LimparDados(regC170, 4, False)
    Call Util.ExportarDadosDicionario(regC170, dicDadosC170)
    
    Application.StatusBar = "Atualizando dados do registro C177, por favor aguarde..."
    Call Util.LimparDados(regC177, 4, False)
    Call Util.ExportarDadosDicionario(regC177, dicDadosC177)
    
    Application.StatusBar = "Atualizando dados do registro C190, por favor aguarde..."
    Call rC170.GerarC190(True)
    
    Application.StatusBar = "Atualizando valores dos impostos no registro C100, por favor aguarde..."
    Call rC170.AtualizarImpostosC100(True)
    
    Application.StatusBar = "Atualização concluída com sucesso!"
    Call Util.MsgInformativa("Registros atualizados com sucesso!", "Atualização de dados", Inicio)
    Application.StatusBar = False
    
End Sub

Public Function AceitarSugestoes()

Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim Campos As Variant, Campos0200, CamposC170, dicCampos, regCampo
Dim dicDados As New Dictionary
    
    Inicio = Now()
    Application.StatusBar = "Implementando sugestões selecionadas, por favor aguarde..."
    
    Set dicTitulos = Util.MapearTitulos(assApuracaoICMS, 3)
    Set Dados = assApuracaoICMS.Range("A4").CurrentRegion
    If Dados Is Nothing Then
        Call Util.MsgAlerta("Não foi encontrada nenhuma inconsistência!", "Inconsistências Fiscais")
        Exit Function
    End If
    
    For Each Linha In Dados.Rows
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
                        
            If Linha.EntireRow.Hidden = False And Campos(dicTitulos("SUGESTAO")) <> "" And Linha.Row > 3 Then
VERIFICAR:
                Select Case Campos(dicTitulos("SUGESTAO"))
                    
                    Case "Zerar os campos VL_BC_ICMS, ALIQ_ICMS e VL_ICMS"
                        Campos(dicTitulos("VL_BC_ICMS")) = 0
                        Campos(dicTitulos("ALIQ_ICMS")) = 0
                        Campos(dicTitulos("VL_ICMS")) = 0
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = GerarSugestoesProdutos(Campos, dicTitulos)
                        GoTo VERIFICAR:
                        
                    Case "Alterar últimos 2 dígitos do CST_ICMS para 90"
                        Campos(dicTitulos("CST_ICMS")) = VBA.Left(Campos(dicTitulos("CST_ICMS")), 1) & "90"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = GerarSugestoesProdutos(Campos, dicTitulos)
                        GoTo VERIFICAR:
                        
                    Case "Alterar últimos 2 dígitos do CST_ICMS para 60"
                        Campos(dicTitulos("CST_ICMS")) = VBA.Left(Campos(dicTitulos("CST_ICMS")), 1) & "60"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = GerarSugestoesProdutos(Campos, dicTitulos)
                        GoTo VERIFICAR:
                        
                    Case "Alterar CST_ICMS para 090"
                        Campos(dicTitulos("CST_ICMS")) = "090"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = GerarSugestoesProdutos(Campos, dicTitulos)
                        GoTo VERIFICAR:
                        
                    Case "Alterar dígito de origem do CST_ICMS para 2"
                        Campos(dicTitulos("CST_ICMS")) = "2" & VBA.Right(Campos(dicTitulos("CST_ICMS")), 2)
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = GerarSugestoesProdutos(Campos, dicTitulos)
                        GoTo VERIFICAR:
                        
                    Case "Alterar dígito de origem do CST_ICMS para 7"
                        Campos(dicTitulos("CST_ICMS")) = "7" & VBA.Right(Campos(dicTitulos("CST_ICMS")), 2)
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = GerarSugestoesProdutos(Campos, dicTitulos)
                        GoTo VERIFICAR:
                        
                    Case "Alterar TIPO_ITEM para 07"
                        Campos(dicTitulos("TIPO_ITEM")) = "07 - Material de Uso e Consumo"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = GerarSugestoesProdutos(Campos, dicTitulos)
                        GoTo VERIFICAR:
                        
                    Case "Alterar últimos 2 dígitos do campo CST_ICMS para 90"
                        Campos(dicTitulos("CST_ICMS")) = VBA.Left(Campos(dicTitulos("CST_ICMS")), 1) & "90"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = GerarSugestoesProdutos(Campos, dicTitulos)
                        GoTo VERIFICAR:
                        
                    Case "Alterar últimos 2 dígitos do campo CST_ICMS para 00"
                        Campos(dicTitulos("CST_ICMS")) = VBA.Left(Campos(dicTitulos("CST_ICMS")), 1) & "00"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = GerarSugestoesProdutos(Campos, dicTitulos)
                        GoTo VERIFICAR:
                        
                    Case "Gerar informações de fatura e vencimento"
                        Campos = GerarRegistrosC140eC141(Campos, dicTitulos)
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = GerarSugestoesContas(Campos, dicTitulos)
                        GoTo VERIFICAR:
                        
                    Case "Somar valor do campo VL_ICMS_ST ao campo VL_ITEM"
                        Campos(dicTitulos("VL_ITEM")) = CDbl(Campos(dicTitulos("VL_ITEM"))) + CDbl(Campos(dicTitulos("VL_ICMS_ST")))
                        Campos(dicTitulos("VL_BC_ICMS_ST")) = 0
                        Campos(dicTitulos("ALIQ_ST")) = 0
                        Campos(dicTitulos("VL_ICMS_ST")) = 0
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = GerarSugestoesProdutos(Campos, dicTitulos)
                        GoTo VERIFICAR:
                        
                    Case "Adicionar zeros a esquerda do CEST"
                        Campos(dicTitulos("CEST")) = "'" & VBA.Format(Util.ApenasNumeros(Campos(dicTitulos("CEST"))), VBA.String(7, "0"))
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = GerarSugestoesProdutos(Campos, dicTitulos)
                        GoTo VERIFICAR:
                        
                    Case Else
                        
                End Select
                
            End If
            
            If Linha.Row > 3 Then dicDados(Campos(dicTitulos("CHV_REG"))) = Campos
            
        End If
        
    Next Linha

    If assApuracaoICMS.AutoFilterMode Then assApuracaoICMS.AutoFilter.ShowAllData
    Call Util.LimparDados(assApuracaoICMS, 4, False)
    Call Util.ExportarDadosDicionario(assApuracaoICMS, dicDados)
    
    Call FuncoesFormatacao.AplicarFormatacao(ActiveSheet)
    Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoICMS)
    
    Call Util.MsgInformativa("Sugestões aplicadas com sucesso!", "Inclusão de Sugestões", Inicio)
    Application.StatusBar = False
    
End Function



Public Function AceitarSugestoesContas()

Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim Campos As Variant, Campos0200, CamposC170, dicCampos, regCampo
Dim dicDados As New Dictionary
    
    If dias = 0 Then
        Call Util.MsgAlerta("Selecione uma opção no controle 'vencimento para:'", "Vencimento não informado")
        Exit Function
    End If
    
    Inicio = Now()
    Application.StatusBar = "Implementando sugestões selecionadas, por favor aguarde..."
    
    Set dicTitulos = Util.MapearTitulos(relInteligenteContas, 3)
    Set Dados = relInteligenteContas.Range("A4").CurrentRegion
    If Dados Is Nothing Then
        Call Util.MsgAlerta("Não foi encontrada nenhuma inconsistência!", "Inconsistências Fiscais")
        Exit Function
    End If
    
    For Each Linha In Dados.Rows
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            If Linha.EntireRow.Hidden = False And Campos(dicTitulos("SUGESTAO")) <> "" And Linha.Row > 3 Then
VERIFICAR:
                Select Case Campos(dicTitulos("SUGESTAO"))
                    
                    Case "Gerar informações de fatura e vencimento"
                        Campos = GerarRegistrosC140eC141(Campos, dicTitulos)
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = GerarSugestoesContas(Campos, dicTitulos)
                        GoTo VERIFICAR:
                        
                    Case Else
                    
                End Select
                
            End If
            
            If Linha.Row > 3 Then dicDados(dicDados.Count) = Campos
            
        End If
        
    Next Linha
    
    If relInteligenteContas.AutoFilterMode Then relInteligenteContas.AutoFilter.ShowAllData
    Call Util.LimparDados(relInteligenteContas, 4, False)
    Call Util.ExportarDadosDicionario(relInteligenteContas, dicDados)
    Call FuncoesFormatacao.AplicarFormatacao(ActiveSheet)
    
    Call Util.MsgInformativa("Sugestões aplicadas com sucesso!", "Inclusão de Sugestões", Inicio)
    Application.StatusBar = False
    
End Function

Public Function AceitarSugestoesInventario()

Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim Campos As Variant, Campos0200, CamposC170, dicCampos, regCampo
Dim dicDados As New Dictionary
Dim ARQUIVO As String, UNID_INV$, COD_ITEM$, CHV_0200$
Dim QTD As Double, VL_UNIT#
    
    Inicio = Now()
    Application.StatusBar = "Implementando sugestões selecionadas, por favor aguarde..."
    
    Set dicTitulos = Util.MapearTitulos(relInteligenteInventario, 3)
    Set Dados = relInteligenteInventario.Range("A4").CurrentRegion
    If Dados Is Nothing Then
        Call Util.MsgAlerta("Não foi encontrada nenhuma inconsistência!", "Inconsistências Fiscais")
        Exit Function
    End If
    
    For Each Linha In Dados.Rows
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            If Linha.EntireRow.Hidden = False And Campos(dicTitulos("SUGESTAO")) <> "" And Linha.Row > 3 Then
VERIFICAR:
                Select Case Campos(dicTitulos("SUGESTAO"))
                    
                    Case "Excluir item do registro de inventário"
                        GoTo Prx:
                        
                    Case "Informar a mesma unidade do cadastro (0200) para o inventário (H010)"
                        ARQUIVO = Campos(dicTitulos("ARQUIVO"))
                        COD_ITEM = Campos(dicTitulos("COD_ITEM"))
                        CHV_0200 = Util.UnirCampos(ARQUIVO, COD_ITEM)
                        If SPEDFiscal.dicDados0200.Exists(CHV_0200) Then
                            
                            Campos(dicTitulos("UNID")) = SPEDFiscal.dicDados0200(CHV_0200)(SPEDFiscal.dicTitulos0200("UNID_INV"))
                            Campos(dicTitulos("SUGESTAO")) = Empty
                            Campos(dicTitulos("INCONSISTENCIA")) = Empty
                            Campos = GerarSugestoesInventario(Campos, dicTitulos)
                            
                        End If
                        
                    Case "Recalcular o campo VL_ITEM"
                        QTD = fnExcel.ConverterValores(Campos(dicTitulos("QTD")), True, 3)
                        VL_UNIT = fnExcel.ConverterValores(Campos(dicTitulos("VL_UNIT")), True, 6)
                        Campos(dicTitulos("VL_ITEM")) = fnExcel.ConverterValores(QTD * VL_UNIT, True, 2)
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos = GerarSugestoesInventario(Campos, dicTitulos)
                        
                    Case Else
                    
                End Select
                
            End If
            
            If Linha.Row > 3 Then dicDados(dicDados.Count) = Campos
            
        End If
Prx:
    Next Linha
    
    'Call ReprocessarSugestoesInventario
    
    If relInteligenteInventario.AutoFilterMode Then relInteligenteInventario.AutoFilter.ShowAllData
    Call Util.LimparDados(relInteligenteInventario, 4, False)
    Call Util.ExportarDadosDicionario(relInteligenteInventario, dicDados)
    
    Call FuncoesFormatacao.AplicarFormatacao(relInteligenteInventario)
    Call FuncoesFormatacao.DestacarInconsistencias(relInteligenteInventario)
    
    Call Util.MsgInformativa("Sugestões aplicadas com sucesso!", "Inclusão de Sugestões", Inicio)
    Application.StatusBar = False
    
End Function

Public Sub GerarRelatorioContasPagarReceber(Optional ByVal OmitirMsg As Boolean)

Dim REG As String, ARQUIVO$, CHV_REG_C100$, CHV_REG_C140$, CHV_REG_C141$, IND_OPER$, IND_EMIT$, COD_SIT$, NUM_DOC$, CHV_NFE$, DT_DOC$, DT_E_S$, IND_PGTO$, IND_TIT$, DESC_TIT$, NUM_TIT$, NUM_PARC$, DT_VCTO$, i$
Dim VL_DOC As Double, QTD_PARC#, VL_TIT#, VL_PARC#
Dim dicTitulosC140 As New Dictionary
Dim dicTitulosC141 As New Dictionary
Dim dicTitulosC100 As New Dictionary
Dim dicDadosC100 As New Dictionary
Dim dicDadosC140 As New Dictionary
Dim dicDadosC141 As New Dictionary
Dim Dados As Range, Linha As Range
Dim arrRelatorio As New ArrayList
Dim dicTitulos As New Dictionary
Dim arrDados As New ArrayList
Dim Campos As Variant
    
    Inicio = Now()
    Application.StatusBar = "Gerando relatório inteligente de contas a pagar e a receber, por favor aguarde..."
    
    Set dicTitulos = Util.MapearTitulos(relInteligenteContas, 3)
    
    Set dicTitulosC140 = Util.MapearTitulos(regC140, 3)
    Set dicDadosC140 = Util.CriarDicionarioRegistro(regC140)
    
    Set dicTitulosC141 = Util.MapearTitulos(regC141, 3)
    Set dicDadosC141 = Util.CriarDicionarioRegistro(regC141)
    
    Set dicTitulosC100 = Util.MapearTitulos(regC100, 3)
    Set Dados = Util.DefinirIntervalo(regC100, 4, 3)
    If Dados Is Nothing Then
        Call Util.MsgAlerta("Não existem dados nos registros C100", "Dados indisponíveis")
        Exit Sub
    End If
    
    a = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 100, "Gerando relatório inteligente de contas a pagar e receber, por favor aguarde...", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            COD_SIT = Campos(dicTitulosC100("COD_SIT"))
            If Not COD_SIT Like "00*" And Not COD_SIT Like "01*" And Not COD_SIT Like "08*" Then GoTo Prx:
            
            REG = Campos(dicTitulosC100("REG"))
            ARQUIVO = Campos(dicTitulosC100("ARQUIVO"))
            CHV_REG_C100 = Campos(dicTitulosC100("CHV_REG"))
            IND_OPER = Campos(dicTitulosC100("IND_OPER"))
            IND_EMIT = Campos(dicTitulosC100("IND_EMIT"))
            NUM_DOC = Campos(dicTitulosC100("NUM_DOC"))
            CHV_NFE = Campos(dicTitulosC100("CHV_NFE"))
            DT_DOC = Campos(dicTitulosC100("DT_DOC"))
            VL_DOC = Campos(dicTitulosC100("VL_DOC"))
            IND_PGTO = Campos(dicTitulosC100("IND_PGTO"))
            
            CHV_REG_C140 = fnSPED.GerarChaveRegistro(CStr(CHV_REG_C100), CStr("C140"))
            
            'Coleta dados do registro C140
            If dicDadosC140.Exists(CHV_REG_C140) Then
                
                IND_TIT = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_TIT(dicDadosC140(CHV_REG_C140)(dicTitulosC140("IND_TIT")))
                DESC_TIT = dicDadosC140(CHV_REG_C140)(dicTitulosC140("DESC_TIT"))
                NUM_TIT = dicDadosC140(CHV_REG_C140)(dicTitulosC140("NUM_TIT"))
                QTD_PARC = dicDadosC140(CHV_REG_C140)(dicTitulosC140("QTD_PARC"))
                VL_TIT = dicDadosC140(CHV_REG_C140)(dicTitulosC140("VL_TIT"))
                
            Else
                
                CHV_REG_C140 = ""
                IND_TIT = ""
                DESC_TIT = ""
                NUM_TIT = ""
                QTD_PARC = 0
                VL_TIT = 0
                
            End If
            
            
            i = 1
            CHV_REG_C141 = fnSPED.GerarChaveRegistro(CStr(CHV_REG_C140), CStr(VBA.Format(i, "00")))
            'Coleta dados do registro C141

                Do
                            
                    If dicDadosC141.Exists(CHV_REG_C141) Then
                        
                        NUM_PARC = VBA.Format(dicDadosC141(CHV_REG_C141)(dicTitulosC141("NUM_PARC")), "00")
                        DT_VCTO = dicDadosC141(CHV_REG_C141)(dicTitulosC141("DT_VCTO"))
                        VL_PARC = dicDadosC141(CHV_REG_C141)(dicTitulosC141("VL_PARC"))
                        
                    Else
                        
                        CHV_REG_C141 = ""
                        NUM_PARC = ""
                        DT_VCTO = ""
                        VL_PARC = 0
                        
                    End If
                    
                    Campos = Array(REG, ARQUIVO, CHV_REG_C100, CHV_REG_C140, CHV_REG_C141, IND_OPER, IND_EMIT, COD_SIT, NUM_DOC, "'" & CHV_NFE, DT_DOC, VL_DOC, IND_PGTO, IND_TIT, DESC_TIT, NUM_TIT, QTD_PARC, VL_TIT, NUM_PARC, DT_VCTO, VL_PARC, Empty, Empty)
                    Campos = GerarSugestoesContas(Campos, dicTitulos)
                    'Call fnExcel.DefinirTipoCampos(Campos, dicTitulos)
                    arrRelatorio.Add Campos
                    
                    i = CInt(i) + 1
                    CHV_REG_C141 = fnSPED.GerarChaveRegistro(CStr(CHV_REG_C140), CStr(VBA.Format(i, "00")))
                    
                Loop Until Not dicDadosC141.Exists(CHV_REG_C141)
                
        End If
        
Prx:
        arrDados.Clear
        
    Next Linha
    
    Call Util.LimparDados(relInteligenteContas, 4, False)
    Call Util.ExportarDadosArrayList(relInteligenteContas, arrRelatorio)
    Call FuncoesFormatacao.DestacarInconsistencias(relInteligenteContas)

    If Not OmitirMsg Then Call Util.MsgInformativa("Relatório de contas gerado com sucesso!", "Relatório de Contas", Inicio)
    
    Application.StatusBar = False
    
End Sub

Public Function ReprocessarSugestoesProdutos()

Dim Dados As Range, Linha As Range
Dim arrRelatorio As New ArrayList
Dim dicTitulos As New Dictionary
Dim Campos As Variant
    
    Set dicTitulos = Util.MapearTitulos(assApuracaoICMS, 3)
    If assApuracaoICMS.AutoFilterMode Then assApuracaoICMS.AutoFilter.ShowAllData
    
    Set Dados = Util.DefinirIntervalo(assApuracaoICMS, 4, 3)
    If Dados Is Nothing Then Exit Function
    
    a = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 100, "Reprocessando sugestões, por favor aguarde...", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            Campos(dicTitulos("INCONSISTENCIA")) = Empty
            Campos(dicTitulos("SUGESTAO")) = Empty
            Campos = FuncoesAssistentesInteligentes.GerarSugestoesProdutos(Campos, dicTitulos)
            
            arrRelatorio.Add Campos
            
        End If
        
    Next Linha
    
    Call Util.LimparDados(assApuracaoICMS, 4, False)
    Call Util.ExportarDadosArrayList(assApuracaoICMS, arrRelatorio)
    Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoICMS)
    
End Function

Public Function ReprocessarSugestoesContas()

Dim Dados As Range, Linha As Range
Dim arrRelatorio As New ArrayList
Dim dicTitulos As New Dictionary
Dim Campos As Variant
    
    Set dicTitulos = Util.MapearTitulos(relInteligenteContas, 3)
    If relInteligenteContas.AutoFilterMode Then relInteligenteContas.AutoFilter.ShowAllData
    
    Set Dados = Util.DefinirIntervalo(relInteligenteContas, 4, 3)
    If Dados Is Nothing Then Exit Function
    
    a = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 100, "Reprocessando sugestões, por favor aguarde...", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            Campos(dicTitulos("INCONSISTENCIA")) = Empty
            Campos(dicTitulos("SUGESTAO")) = Empty
            Campos = GerarSugestoesContas(Campos, dicTitulos)
            Call fnExcel.DefinirTipoCampos(Campos, dicTitulos)
            arrRelatorio.Add Campos
            
        End If
        
    Next Linha
    
    Call Util.LimparDados(relInteligenteContas, 4, False)
    Call Util.ExportarDadosArrayList(relInteligenteContas, arrRelatorio)
    
End Function

Public Sub AtualizarContasPagarReceber()

Dim dicTitulosC100 As New Dictionary
Dim dicTitulosC140 As New Dictionary
Dim dicTitulosC141 As New Dictionary
Dim dicDadosC100 As New Dictionary
Dim dicDadosC140 As New Dictionary
Dim dicDadosC141 As New Dictionary
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim Campos As Variant, CamposC100, CamposC140, CamposC141, dicCampos, regCampo, nCampo
Dim CHV_C141 As String, CHV_C100$, CHV_C140$
    
    Inicio = Now()
    Application.StatusBar = "Preparando dados para atualização do SPED, por favor aguarde..."
    
    CamposC100 = Array("IND_OPER", "IND_EMIT", "COD_SIT", "NUM_DOC", "CHV_NFE", "DT_DOC", "VL_DOC", "IND_PGTO")
    CamposC140 = Array("IND_EMIT", "IND_TIT", "DESC_TIT", "NUM_TIT", "QTD_PARC", "VL_TIT")
    CamposC141 = Array("NUM_PARC", "DT_VCTO", "VL_PARC")
    
    Set dicTitulosC100 = Util.MapearTitulos(regC100, 3)
    Set dicDadosC100 = Util.CriarDicionarioRegistro(regC100)
    
    Set dicTitulosC140 = Util.MapearTitulos(regC140, 3)
    Set dicDadosC140 = Util.CriarDicionarioRegistro(regC140)
    
    Set dicTitulosC141 = Util.MapearTitulos(regC141, 3)
    Set dicDadosC141 = Util.CriarDicionarioRegistro(regC141)
    
    Set dicTitulos = Util.MapearTitulos(relInteligenteContas, 3)
    If relInteligenteContas.AutoFilterMode Then relInteligenteContas.AutoFilter.ShowAllData
    
    Set Dados = Util.DefinirIntervalo(relInteligenteContas, 4, 3)
    If Dados Is Nothing Then Exit Sub
    
    a = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 100, "Preparando dados para atualização do SPED, por favor aguarde...", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            'Atualizar dados do C100
            CHV_C100 = CStr(Campos(dicTitulos("CHV_REG_C100")))
            If dicDadosC100.Exists(CHV_C100) Then
                
                dicCampos = dicDadosC100(CHV_C100)
                For Each regCampo In CamposC100
                    
                    dicCampos(dicTitulosC100(regCampo)) = Campos(dicTitulos(regCampo))
                    
                Next regCampo
                
                dicDadosC100(CHV_C100) = dicCampos
            
            End If
            
            'Atualizar dados do C140
            CHV_C140 = CStr(Campos(dicTitulos("CHV_REG_C140")))
            If CHV_C140 <> "" Then
            
                If dicDadosC140.Exists(CHV_C140) Then
                    
                    dicCampos = dicDadosC140(CHV_C140)
                    For Each regCampo In CamposC140
                        
                        dicCampos(dicTitulosC140(regCampo)) = Campos(dicTitulos(regCampo))
                        
                    Next regCampo
                    
                    dicDadosC140(CHV_C140) = dicCampos
                    
                Else
                    
                    dicCampos = Array("C140", Campos(dicTitulos("ARQUIVO")), CHV_C140, CHV_C100, "", Campos(dicTitulos("IND_EMIT")), _
                                      Campos(dicTitulos("IND_TIT")), Campos(dicTitulos("DESC_TIT")), Campos(dicTitulos("NUM_TIT")), _
                                      Campos(dicTitulos("QTD_PARC")), Campos(dicTitulos("VL_TIT")))
                                      
                    dicDadosC140(CHV_C140) = dicCampos
                    
                End If
            
            End If
            
            'Atualizar dados do C141
            CHV_C141 = CStr(Campos(dicTitulos("CHV_REG_C141")))
            If CHV_C141 <> "" Then
                
                If dicDadosC141.Exists(CHV_C141) Then
                    
                    dicCampos = dicDadosC141(CHV_C141)
                    For Each regCampo In CamposC141
                        
                        dicCampos(dicTitulosC141(regCampo)) = Campos(dicTitulos(regCampo))
                        
                    Next regCampo
                    
                    dicDadosC141(CHV_C141) = dicCampos
                
                Else
                    
                    dicCampos = Array("C141", Campos(dicTitulos("ARQUIVO")), CHV_C141, CHV_C140, "", _
                        VBA.Format(Campos(dicTitulos("NUM_PARC")), "00"), Campos(dicTitulos("DT_VCTO")), Campos(dicTitulos("VL_PARC")))
    
                    dicDadosC141(CHV_C141) = dicCampos
                
                End If
            
            End If
            
        End If
        
    Next Linha
    
    Application.StatusBar = "Atualizando dados do registro C100, por favor aguarde..."
    Call Util.LimparDados(regC100, 4, False)
    Call Util.ExportarDadosDicionario(regC100, dicDadosC100)
    
    Application.StatusBar = "Atualizando dados do registro C140, por favor aguarde..."
    Call Util.LimparDados(regC140, 4, False)
    Call Util.ExportarDadosDicionario(regC140, dicDadosC140)
    
    Application.StatusBar = "Atualizando dados do registro C141, por favor aguarde..."
    Call Util.LimparDados(regC141, 4, False)
    Call Util.ExportarDadosDicionario(regC141, dicDadosC141)
    
    Application.StatusBar = "Geração concluída com sucesso!"
    Call Util.MsgInformativa("Registros atualizados com sucesso!", "Atualização de registros do SPED", Inicio)
    Application.StatusBar = False
    
End Sub

Public Function GerarRegistrosC140eC141(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim Dados As Range, Linha As Range
Dim CHV_REG_C100 As String
        
    If Util.ChecarCamposPreenchidos(Campos) Then
        
        CHV_REG_C100 = Campos(dicTitulos("CHV_REG_C100"))
        Campos(dicTitulos("CHV_REG_C140")) = fnSPED.GerarChaveRegistro(CHV_REG_C100, "C140")
        Campos(dicTitulos("CHV_REG_C141")) = fnSPED.GerarChaveRegistro(CStr(Campos(dicTitulos("CHV_REG_C140"))), "01")
        Campos(dicTitulos("IND_TIT")) = "00 - Duplicata"
        Campos(dicTitulos("DESC_TIT")) = "Duplicata ref NF " & Campos(dicTitulos("NUM_DOC"))
        Campos(dicTitulos("NUM_TIT")) = Campos(dicTitulos("NUM_DOC"))
        Campos(dicTitulos("QTD_PARC")) = 1
        Campos(dicTitulos("VL_TIT")) = Campos(dicTitulos("VL_DOC"))
        Campos(dicTitulos("NUM_PARC")) = "01"
        Campos(dicTitulos("DT_VCTO")) = Campos(dicTitulos("DT_DOC")) + dias
        Campos(dicTitulos("VL_PARC")) = Campos(dicTitulos("VL_DOC"))
        
    End If
    
    GerarRegistrosC140eC141 = Campos
    
End Function

Public Function ImportarInventarioFisico()

Dim dtIni As String, dtFim$, ARQUIVO$, CHV_REG_0000$, CHV_REG_0001$, CHV_REG_0150$, CHV_REG_0190$, CHV_REG_0200$, CHV_REG_H001$, COD_ITEM$, UNID$, COD_PART$
Dim dicTitulosInventario As New Dictionary
Dim dicTitulos0150 As New Dictionary
Dim dicTitulos0190 As New Dictionary
Dim dicTitulos0200 As New Dictionary
Dim Caminho As Variant, Campos, Titulo, Campo
Dim Dados As Range, Linha As Range
Dim dicDados0150 As New Dictionary
Dim dicDados0190 As New Dictionary
Dim dicDados0200 As New Dictionary
Dim dicTitulos As New Dictionary
Dim PastaDeTrabalho As Workbook
Dim arrCampos As New ArrayList
Dim arrDados As New ArrayList
Dim Plan As Worksheet
    
    If PeriodoInventario = "" Then
        Call Util.MsgAlerta("Informe o período ('MM/AAAA') que deseja inserir os itens para prosseguir com a importação.", "Período de importação não informado")
        Exit Function
    End If
    
    Caminho = Util.SelecionarArquivo("xlsx")
    If VarType(Caminho) = 11 Then Exit Function
    
    Inicio = Now()
    With CadContrib
        
        CNPJContribuinte = Util.FormatarCNPJ(CadContrib.Range("CNPJContribuinte").value)
        CNPJBase = VBA.Left(CNPJContribuinte, 8)
        If VBA.IsNumeric(CadContrib.Range("InscContribuinte").value) Then InscContribuinte = CadContrib.Range("InscContribuinte").value * 1
        dtIni = Util.FormatarData("01" & PeriodoInventario)
        dtFim = VBA.Format(Util.FimMes(dtIni), "ddmmyyyy")
        dtIni = VBA.Format(dtIni, "ddmmyyyy")
        CHV_REG_0000 = Cripto.MD5(fnSPED.MontarChaveRegistro(Array("", dtIni, dtFim, CNPJContribuinte, "", InscContribuinte)))
        CHV_REG_0001 = fnSPED.GerarChaveRegistro(CHV_REG_0000, "0001")
        CHV_REG_H001 = fnSPED.GerarChaveRegistro(CHV_REG_0000, "H001")
        
        ARQUIVO = VBA.Format(PeriodoInventario, "00/0000") & "-" & CNPJContribuinte
        If InscContribuinte = "" Then
            Call Util.MsgAlerta("Informe a Inscrição Estadual do Contribuinte.", "Inscrição Estadual não informada")
            CadContrib.Activate
            CadContrib.Range("InscContribuinte").Activate
            Exit Function
        End If
        
    End With
    
    Set PastaDeTrabalho = Workbooks.Open(Caminho)
    ActiveWindow.visible = False
    
    Set dicTitulos0150 = Util.MapearTitulos(reg0150, 3)
    Set dicTitulos0190 = Util.MapearTitulos(reg0190, 3)
    Set dicTitulos0200 = Util.MapearTitulos(reg0200, 3)
    
    Set dicDados0150 = Util.CriarDicionarioRegistro(reg0150)
    Set dicDados0190 = Util.CriarDicionarioRegistro(reg0190)
    Set dicDados0200 = Util.CriarDicionarioRegistro(reg0200)
    
    For Each Plan In PastaDeTrabalho.Worksheets
        
        With Plan
            
            If .AutoFilterMode Then .AutoFilter.ShowAllData
            Set dicTitulos = Util.MapearTitulos(Plan, 1)
            Set dicTitulosInventario = Util.MapearTitulos(relInteligenteInventario, 3)
            
            Set Dados = Util.DefinirIntervalo(Plan, 2, 1)
            If Dados Is Nothing Then GoTo Prx:
            
            For Each Linha In Dados.Rows
                
                Campos = Application.index(Linha.Value2, 0, 0)
                If Util.ChecarCamposPreenchidos(Campos) Then
                    
                    For Each Titulo In dicTitulosInventario
                        
                        Select Case Titulo
                            
                            Case "ARQUIVO"
                                arrCampos.Add ARQUIVO
                                GoTo PrxTit:
                                
                            Case "COD_ITEM"
                                COD_ITEM = Campos(dicTitulos(Titulo))
                                CHV_REG_0200 = fnSPED.GerarChaveRegistro(CHV_REG_0001, COD_ITEM)
                                
                            Case "DESCR_ITEM", "TIPO_ITEM", "COD_NCM", "CEST"
                                If dicTitulos.Exists(Titulo) Then Campo = Campos(dicTitulos(Titulo)) Else Campo = ""
                                If dicDados0200.Exists(CHV_REG_0200) Then Campo = dicDados0200(CHV_REG_0200)(dicTitulos0200(Titulo))
                                If Titulo = "DESCR_ITEM" And Campo = "" Then Campo = "ITEM NÃO IDENTIFICADO NO 0200"
                                arrCampos.Add Campo
                                GoTo PrxTit:
                                
                            Case "UNID"
                                 UNID = Campos(dicTitulos(Titulo))
                                 CHV_REG_0190 = fnSPED.GerarChaveRegistro(CHV_REG_0001, UNID)
                                 
                            Case "COD_PART"
                                COD_PART = Campos(dicTitulos(Titulo))
                                CHV_REG_0150 = fnSPED.GerarChaveRegistro(CHV_REG_0001, COD_PART)
                                If dicTitulos.Exists(Titulo) Then Campo = Campos(dicTitulos(Titulo)) Else Campo = ""
                                If dicDados0150.Exists(CHV_REG_0150) Then Campo = dicDados0150(CHV_REG_0150)(dicTitulos0150(Titulo))
                                arrCampos.Add Campo
                                GoTo PrxTit:
                                
                            Case "QTD", "VL_UNIT", "VL_ITEM", "VL_ITEM_IR", "VL_BC_ICMS", "VL_ICMS"
                                If dicTitulos.Exists(Titulo) Then Campo = Campos(dicTitulos(Titulo)) Else Campo = 0
                                If Not IsNumeric(Campo) Or Campo = "" Then Campo = 0
                                arrCampos.Add Campo
                                GoTo PrxTit:
                                
                            Case "INCONSISTENCIA", "SUGESTAO"
                                arrCampos.Add Empty
                                GoTo PrxTit:
                                
                        End Select
                        
                        If dicTitulos.Exists(Titulo) Then arrCampos.Add Campos(dicTitulos(Titulo)) Else arrCampos.Add ""
PrxTit:
                    Next Titulo
                    
                    CHV_REG_0190 = fnSPED.GerarChaveRegistro(CHV_REG_0001, UNID)
                    If Not dicDados0190.Exists(CHV_REG_0190) Then Call IncluirUnid0190(dicDados0190, ARQUIVO, CHV_REG_0001, CHV_REG_0190, UNID)
                    Campos = GerarSugestoesInventario(arrCampos.toArray(), dicTitulosInventario)
                    Call fnExcel.DefinirTipoCampos(Campos, dicTitulosInventario)
                    arrDados.Add Campos
                    arrCampos.Clear
                    
                End If

            Next Linha
            
            Exit For
            
        End With
        
Prx:
    
    Next Plan
    
    Application.DisplayAlerts = False
        PastaDeTrabalho.Close
    Application.DisplayAlerts = True
    
    If arrDados.Count = 0 Then
        Call Util.MsgAlerta("O arquivo selecionado não possui dados de produtos.", "Arquivo sem dados informados")
        Exit Function
    End If
    
    
    Call Util.LimparDados(reg0190, 4, False)
    Call Util.ExportarDadosDicionario(reg0190, dicDados0190)
    
    Call Util.LimparDados(relInteligenteInventario, 4, False)
    Call Util.ExportarDadosArrayList(relInteligenteInventario, arrDados)
    
    Call Util.MsgInformativa("Registros de Inventário importados com sucesso!", "Importação de Inventário (Bloco H)", Inicio)
    
End Function

Public Function ReprocessarSugestoesInventario()

Dim Dados As Range, Linha As Range
Dim arrRelatorio As New ArrayList
Dim dicTitulos As New Dictionary
Dim Campos As Variant
    
    Set dicTitulos = Util.MapearTitulos(relInteligenteInventario, 3)
    If relInteligenteInventario.AutoFilterMode Then relInteligenteInventario.AutoFilter.ShowAllData
    
    Set Dados = Util.DefinirIntervalo(relInteligenteInventario, 4, 3)
    If Dados Is Nothing Then Exit Function
    
    Call DadosSPEDFiscal.CarregarDadosRegistro0200
        
    a = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 100, "Reprocessando sugestões, por favor aguarde...", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            Campos(dicTitulos("INCONSISTENCIA")) = Empty
            Campos(dicTitulos("SUGESTAO")) = Empty
            Campos = GerarSugestoesInventario(Campos, dicTitulos)
            Call fnExcel.DefinirTipoCampos(Campos, dicTitulos)
            arrRelatorio.Add Campos
            
        End If
        
    Next Linha
    
    Call Util.LimparDados(relInteligenteInventario, 4, False)
    Call Util.ExportarDadosArrayList(relInteligenteInventario, arrRelatorio)
    
    Call FuncoesFormatacao.AplicarFormatacao(relInteligenteInventario)
    Call FuncoesFormatacao.DestacarInconsistencias(relInteligenteInventario)
        
End Function

Public Sub AtualizarInventario()

Dim Campos As Variant, CamposH005, CamposH010, CamposH020, dicCampos, regCampo, nCampo, Campo
Dim CHV_0000 As String, CHV_H001$, CHV_H005$, CHV_H010$, CHV_H020$, DT_INV$, MOT_INV$, COD_ITEM$, IND_PROP$, COD_PART$, ARQUIVO$, Periodo$
Dim dicTitulosH005 As New Dictionary
Dim dicTitulosH010 As New Dictionary
Dim dicTitulosH020 As New Dictionary
Dim dicDadosH005 As New Dictionary
Dim dicDadosH010 As New Dictionary
Dim dicDadosH020 As New Dictionary
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim Reiniciado As Boolean

    Inicio = Now()
    Application.StatusBar = "Preparando dados para atualização do SPED, por favor aguarde..."
    
    CamposH005 = Array("VL_INV")
    CamposH010 = Array("COD_ITEM", "UNID", "QTD", "VL_UNIT", "VL_ITEM", "IND_PROP", "COD_PART", "TXT_COMPL", "COD_CTA", "VL_ITEM_IR")
    CamposH020 = Array("CST_ICMS", "VL_BC_ICMS", "VL_ICMS")
    
    Set dicTitulosH005 = Util.MapearTitulos(regH005, 3)
    Set dicDadosH005 = Util.CriarDicionarioRegistro(regH005)
    
    Set dicTitulosH010 = Util.MapearTitulos(regH010, 3)
    'Set dicDadosH010 = Util.CriarDicionarioRegistro(regH010)
    
    Set dicTitulosH020 = Util.MapearTitulos(regH020, 3)
    Set dicDadosH020 = Util.CriarDicionarioRegistro(regH020)
    
    Set dicTitulos = Util.MapearTitulos(relInteligenteInventario, 3)
    If relInteligenteInventario.AutoFilterMode Then relInteligenteInventario.AutoFilter.ShowAllData
    
    Set Dados = Util.DefinirIntervalo(relInteligenteInventario, 4, 3)
    If Dados Is Nothing Then Exit Sub
    
    Call Util.LimparDados(regH005, 4, False)
    Call Util.LimparDados(regH010, 4, False)
    Call Util.LimparDados(regH020, 4, False)
    
    a = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 100, "Preparando dados para atualização do SPED, por favor aguarde...", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            ARQUIVO = Campos(dicTitulos("ARQUIVO"))
            Periodo = VBA.Left(ARQUIVO, 7)
            CHV_0000 = fnSPED.GerarChvReg0000(Periodo)
            CHV_H001 = fnSPED.GerarChaveRegistro(CHV_0000, "H001")
            
            'Atualizar dados do H005
            DT_INV = VBA.Format(Campos(dicTitulos("DT_INV")), "ddmmyyyy")
            MOT_INV = VBA.Left(Campos(dicTitulos("MOT_INV")), 2)
            
            CHV_H005 = fnSPED.GerarChaveRegistro(CHV_H001, DT_INV, MOT_INV)
            If dicDadosH005.Exists(CHV_H005) And Not Reiniciado Then
                Call dicDadosH005.Remove(CHV_H005)
                Reiniciado = True
            End If
            
            If dicDadosH005.Exists(CHV_H005) Then
                
                dicCampos = dicDadosH005(CHV_H005)
                For Each regCampo In CamposH005
                    
                    If regCampo = "VL_INV" Then
                        Campo = "VL_ITEM"
                        dicCampos(dicTitulosH005(regCampo)) = VBA.Round(CDbl(Campos(dicTitulos(Campo))) + CDbl(dicCampos(dicTitulosH005(regCampo))), 2)
                        
                    End If
                    
                Next regCampo
                
            Else
                
                Reiniciado = True
                dicCampos = Array("H005", ARQUIVO, CHV_H005, CHV_H001, Campos(dicTitulos("DT_INV")), Campos(dicTitulos("VL_ITEM")), "'" & MOT_INV)
                
            End If

            dicDadosH005(CHV_H005) = dicCampos
            
            'Atualizar dados do H010
            COD_ITEM = Campos(dicTitulos("COD_ITEM"))
            IND_PROP = VBA.Left(Campos(dicTitulos("IND_PROP")), 1)
            COD_PART = Campos(dicTitulos("COD_PART"))
            
            CHV_H010 = fnSPED.GerarChaveRegistro(CHV_H005, COD_ITEM, IND_PROP, COD_PART)
            If dicDadosH010.Exists(CHV_H010) Then
                
                dicCampos = dicDadosH010(CHV_H010)
                For Each regCampo In CamposH010
                    
                    dicCampos(dicTitulosH010(regCampo)) = Campos(dicTitulos(regCampo))
                    
                Next regCampo
                
            Else
                
                dicCampos = Array("H010", ARQUIVO, CHV_H010, CHV_H005, "'" & Campos(dicTitulos("COD_ITEM")), _
                                  Campos(dicTitulos("UNID")), Campos(dicTitulos("QTD")), Campos(dicTitulos("VL_UNIT")), _
                                  Campos(dicTitulos("VL_ITEM")), Campos(dicTitulos("IND_PROP")), Campos(dicTitulos("COD_PART")), _
                                  Campos(dicTitulos("TXT_COMPL")), Campos(dicTitulos("COD_CTA")), Campos(dicTitulos("VL_ITEM_IR")))
                
            End If
            
            dicDadosH010(CHV_H010) = dicCampos
            
            'Atualizar dados do H020
            CHV_H020 = fnSPED.GerarChaveRegistro(CHV_H010, "H020")
            If dicDadosH020.Exists(CHV_H020) Then
                
                dicCampos = dicDadosH020(CHV_H020)
                For Each regCampo In CamposH020
                    
                    dicCampos(dicTitulosH020(regCampo)) = Campos(dicTitulos(regCampo))
                    
                Next regCampo
                
            Else
                
                If Campos(dicTitulos("CST_ICMS")) <> "" Then _
                    dicCampos = Array("H020", ARQUIVO, CHV_H020, CHV_H010, "'" & Campos(dicTitulos("CST_ICMS")), _
                                  VBA.Round(Campos(dicTitulos("VL_BC_ICMS")), 2), VBA.Round(Campos(dicTitulos("VL_ICMS")), 2))
                                
            End If
            
            If VarType(dicCampos) = 8204 Then dicDadosH020(CHV_H020) = dicCampos
            
        End If
        
    Next Linha
    
    Application.StatusBar = "Atualizando dados do registro H005, por favor aguarde..."
    Call Util.ExportarDadosDicionario(regH005, dicDadosH005)
    
    Application.StatusBar = "Atualizando dados do registro H010, por favor aguarde..."
    Call Util.ExportarDadosDicionario(regH010, dicDadosH010)
    
    Application.StatusBar = "Atualizando dados do registro H020, por favor aguarde..."
    Call Util.ExportarDadosDicionario(regH020, dicDadosH020)
    
    Application.StatusBar = "Geração concluída com sucesso!"
    Call Util.MsgInformativa("Registros de inventário incluídos no SPED com sucesso!", "Inclusão de inventário no SPED Fiscal", Inicio)
    Application.StatusBar = False
    
End Sub

Public Function AtualizarRegistrosApuracaoPIS_COFINS()
    Call Assistente.Fiscal.Apuracao.PISCOFINS.AtualizarRegistrosPIS_COFINS
End Function

Public Function GerarRelatorioInventario()

Dim Comeco As Double, VL_MERC#, VL_ITEM#, VL_DESP#, VL_FRT#, VL_SEG#, VL_OUT#, VL_ADIC#
Dim CHV_REG As String, CST_PIS$, CST_COFINS$, ARQUIVO$, CHV_0001$, COD_ITEM$, Periodo$, Msg$
Dim dicTitulos As New Dictionary
Dim dicTitulos0000 As New Dictionary
Dim dicTitulos0001 As New Dictionary
Dim dicTitulos0200 As New Dictionary
Dim dicTitulosH005 As New Dictionary
Dim dicTitulosH010 As New Dictionary
Dim dicTitulosH020 As New Dictionary
Dim dicTitulosC177 As New Dictionary
Dim Dados As Range, Linha As Range
Dim dicDados0000 As New Dictionary
Dim dicDados0001 As New Dictionary
Dim dicDados0200 As New Dictionary
Dim dicDadosH005 As New Dictionary
Dim dicDadosH020 As New Dictionary
Dim dicDadosC177 As New Dictionary
Dim arrDados As New ArrayList
Dim arrRelatorio As New ArrayList
Dim Campos As Variant
Dim b As Long
    
    Inicio = Now()
    Set Dados = Util.DefinirIntervalo(regH010, 4, 3)
        
    If Dados Is Nothing Then
        
        Msg = "Nenhum dado encontrado para geração do relatório." & vbCrLf & vbCrLf
        Msg = Msg & "Por favor verifique se o SPED foi importado e tente novamente."
        Call Util.MsgAlerta(Msg, "Relatório de Inventário")
        Exit Function
    
    End If
    
    Set dicTitulos = Util.MapearTitulos(relInteligenteInventario, 3)
    
    Set dicTitulos0000 = Util.MapearTitulos(reg0000, 3)
    Set dicDados0000 = Util.CriarDicionarioRegistro(reg0000, "ARQUIVO")
    
    Set dicTitulos0001 = Util.MapearTitulos(reg0001, 3)
    Set dicDados0001 = Util.CriarDicionarioRegistro(reg0001, "ARQUIVO")
    
    Set dicTitulos0200 = Util.MapearTitulos(reg0200, 3)
    Set dicDados0200 = Util.CriarDicionarioRegistro(reg0200)
    
    Set dicTitulosH005 = Util.MapearTitulos(regH005, 3)
    Set dicDadosH005 = Util.CriarDicionarioRegistro(regH005)
    
    Set dicTitulosH020 = Util.MapearTitulos(regH020, 3)
    Set dicDadosH020 = Util.CriarDicionarioRegistro(regH020)
    
    Set dicTitulosH010 = Util.MapearTitulos(regH010, 3)
    
    Call DadosSPEDFiscal.CarregarDadosRegistro0200
        
    b = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(b, 100, "Carregando dados do registro H010", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            ARQUIVO = Campos(dicTitulosH010("ARQUIVO"))
            If dicDados0000.Exists(ARQUIVO) And Periodo <> ARQUIVO Then
                Call fnSPED.AtuailzarCadastroContribuinte(dicDados0000, dicTitulos0000, ARQUIVO, Periodo)
            End If
            
            arrDados.Add ARQUIVO
            
            'Carrega dados do registro H005
            CHV_REG = Campos(dicTitulosH010("CHV_PAI_FISCAL"))
            If dicDadosH005.Exists(CHV_REG) Then
                
                arrDados.Add dicDadosH005(CHV_REG)(dicTitulosH005("DT_INV"))
                arrDados.Add ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_MOT_INV(dicDadosH005(CHV_REG)(dicTitulosH005("MOT_INV")))

            Else
                
                arrDados.Add ""
                arrDados.Add ""
                
            End If
            
            
            arrDados.Add fnExcel.FormatarTexto(Campos(dicTitulosH010("COD_ITEM")))
            
            COD_ITEM = Campos(dicTitulosH010("COD_ITEM"))
            If dicDados0001.Exists(ARQUIVO) Then CHV_0001 = dicDados0001(ARQUIVO)(dicTitulos0001("CHV_REG")) Else CHV_0001 = ""
            
            'Coleta dados do registro 0200
            CHV_REG = fnSPED.GerarChaveRegistro(CHV_0001, COD_ITEM)
            If dicDados0200.Exists(CHV_REG) Then
                
                arrDados.Add dicDados0200(CHV_REG)(dicTitulos0200("DESCR_ITEM"))
                arrDados.Add dicDados0200(CHV_REG)(dicTitulos0200("TIPO_ITEM"))
                arrDados.Add fnExcel.FormatarTexto(dicDados0200(CHV_REG)(dicTitulos0200("COD_NCM")))
                arrDados.Add fnExcel.FormatarTexto(dicDados0200(CHV_REG)(dicTitulos0200("CEST")))
                
            Else
                
                arrDados.Add "ITEM NÃO IDENTIFICADO"
                arrDados.Add ""
                arrDados.Add ""
                arrDados.Add ""
                
            End If
            
            arrDados.Add Campos(dicTitulosH010("UNID"))
            arrDados.Add Campos(dicTitulosH010("QTD"))
            arrDados.Add fnExcel.FormatarValores(Campos(dicTitulosH010("VL_UNIT")))
            arrDados.Add fnExcel.FormatarValores(Campos(dicTitulosH010("VL_ITEM")))
            arrDados.Add Campos(dicTitulosH010("IND_PROP"))
            arrDados.Add fnExcel.FormatarTexto(Campos(dicTitulosH010("COD_PART")))
            arrDados.Add fnExcel.FormatarTexto(Campos(dicTitulosH010("TXT_COMPL")))
            arrDados.Add fnExcel.FormatarTexto(Campos(dicTitulosH010("COD_CTA")))
            arrDados.Add fnExcel.FormatarValores(Campos(dicTitulosH010("VL_ITEM_IR")))
            
            'Coleta dados do registro H020
            CHV_REG = fnSPED.GerarChaveRegistro(CStr(Campos(dicTitulosH010("CHV_REG"))), "H020")
            If dicDadosH020.Exists(CHV_REG) Then
                
                arrDados.Add fnExcel.FormatarTexto(dicDadosH020(CHV_REG)(dicTitulosH020("CST_ICMS")))
                arrDados.Add fnExcel.FormatarValores(dicDadosH020(CHV_REG)(dicTitulosH020("VL_BC_ICMS")))
                arrDados.Add fnExcel.FormatarValores(dicDadosH020(CHV_REG)(dicTitulosH020("VL_ICMS")))
                
            Else
                
                arrDados.Add ""
                arrDados.Add 0
                arrDados.Add 0
                
            End If
            
            arrDados.Add Empty
            arrDados.Add Empty
            
            Campos = GerarSugestoesInventario(arrDados.toArray(), dicTitulos)
            arrRelatorio.Add Campos
            arrDados.Clear
                        
        End If
        
    Next Linha
    
    Call Util.LimparDados(relInteligenteInventario, 4, False)
    Call Util.ExportarDadosArrayList(relInteligenteInventario, arrRelatorio)
    
    Call FuncoesFormatacao.AplicarFormatacao(relInteligenteInventario)
    Call FuncoesFormatacao.DestacarInconsistencias(relInteligenteInventario)
    
    Call Util.MsgInformativa("Relatório gerado com sucesso!", "Relatório Inteligente de Inventário", Inicio)
    
    Application.StatusBar = False
    
End Function
