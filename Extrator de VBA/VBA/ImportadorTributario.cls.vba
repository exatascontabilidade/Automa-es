Attribute VB_Name = "ImportadorTributario"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' --- Dependências Externas
Private aTributario As New AssistenteTributario

Private Type ConfiguracaoImportacao
    
    DicTitulosDestino As Dictionary
    DicDadosExistentes As Dictionary
    DicTitulosOrigem As Dictionary
    NomeTributoDestino As String
    CamposChaveTributo As Variant
    planilhaOrigem As Worksheet
    planilhaDestino As Worksheet
    
End Type

'=======================================================================================
' ORQUESTRADOR PRINCIPAL
'=======================================================================================

Public Sub ImportarTributacao(ByVal planilhaDestino As Worksheet)
    Dim arquivoImportacao As Workbook
    Dim planilhaOrigem As Worksheet
    Dim configuracoes As ConfiguracaoImportacao ' Usando um Type para agrupar dados
    Dim novasLinhasTributarias As ArrayList
    Dim mensagemValidacao As String
    Dim tempoInicio As Date

    tempoInicio = Now()
    Application.ScreenUpdating = False
    AtualizarStatusBar "Iniciando processo de importação..."

    Set arquivoImportacao = SelecionarAbrirArquivoImportacao()
    If arquivoImportacao Is Nothing Then GoTo FinalizarExecucao ' Usuário cancelou ou erro na abertura

    Set planilhaOrigem = arquivoImportacao.Worksheets(1)

    Call CarregarConfiguracoesImportacao(planilhaDestino, planilhaOrigem, configuracoes)

    mensagemValidacao = ValidarLayoutArquivoImportacao(configuracoes)

    If Len(mensagemValidacao) > 0 Then
        Call ExibirMensagemAlerta("Erro de Validação", mensagemValidacao)
        GoTo FinalizarExecucao ' Pula para fechar arquivo e limpar
    End If

    AtualizarStatusBar "Validação concluída. Processando linhas do arquivo..."
    Set novasLinhasTributarias = ProcessarLinhasArquivoImportacao(planilhaOrigem, configuracoes)

    If novasLinhasTributarias Is Nothing Or novasLinhasTributarias.Count = 0 Then
        Call ExibirMensagemAlerta("Importação Concluída", "Nenhuma tributação nova encontrada no arquivo importado.")
    Else
        AtualizarStatusBar "Exportando " & novasLinhasTributarias.Count & " novas tributações..."
        Call ExportarNovasTributacoes(planilhaDestino, novasLinhasTributarias)
        Call ExibirMensagemSucesso("Importação Concluída", "Tributação importada com sucesso!", tempoInicio)
    End If

FinalizarExecucao:
    Call FecharArquivoImportacao(arquivoImportacao) ' Fecha o workbook mesmo se ocorreu erro antes
    Call LimparObjetosConfiguracao(configuracoes)
    Set novasLinhasTributarias = Nothing
    Set planilhaOrigem = Nothing
    Application.ScreenUpdating = True
    AtualizarStatusBar False ' Limpa a barra de status
End Sub

'=======================================================================================
' MANIPULAÇÃO DE ARQUIVO E CONFIGURAÇÃO
'=======================================================================================

Private Function SelecionarAbrirArquivoImportacao() As Workbook
    Dim CaminhoArquivo As Variant
    Dim wb As Workbook

    CaminhoArquivo = Util.SelecionarArquivo("xlsx") ' Assume Util.SelecionarArquivo existe
    If VarType(CaminhoArquivo) = vbBoolean And CaminhoArquivo = False Then Exit Function ' Cancelado

    On Error GoTo ErroAbrir
    Set wb = Workbooks.Open(CaminhoArquivo, ReadOnly:=True) ' Abrir como somente leitura é mais seguro
    wb.Windows(1).visible = False ' Ocultar janela
    Set SelecionarAbrirArquivoImportacao = wb
    On Error GoTo 0
    Exit Function

ErroAbrir:
    Call ExibirMensagemAlerta("Erro ao Abrir Arquivo", "Não foi possível abrir o arquivo:" & vbCrLf & CaminhoArquivo & vbCrLf & "Erro: " & Err.Description)
    Set SelecionarAbrirArquivoImportacao = Nothing
End Function

Private Sub CarregarConfiguracoesImportacao(ByVal PlanDestino As Worksheet, ByVal planOrigem As Worksheet, ByRef config As ConfiguracaoImportacao)
    AtualizarStatusBar "Carregando configurações e mapeamentos..."
    Set config.DicTitulosDestino = Util.MapearTitulos(PlanDestino, 3)
    Set config.DicDadosExistentes = aTributario.CarregarTributacoesSalvas(PlanDestino)
    If config.DicDadosExistentes Is Nothing Then Set config.DicDadosExistentes = New Dictionary
    Set config.DicTitulosOrigem = Util.MapearTitulos(planOrigem, 1)
    config.NomeTributoDestino = aTributario.ExtrairNomeTributo(PlanDestino)
    config.CamposChaveTributo = aTributario.ObterNomesCamposChave(PlanDestino, True)
    Set config.planilhaOrigem = planOrigem
    Set config.planilhaDestino = PlanDestino
End Sub

Private Sub LimparObjetosConfiguracao(ByRef config As ConfiguracaoImportacao)
    Set config.DicTitulosDestino = Nothing
    Set config.DicDadosExistentes = Nothing
    Set config.DicTitulosOrigem = Nothing
    Set config.planilhaOrigem = Nothing
    Set config.planilhaDestino = Nothing
    ' CamposChaveTributo e NomeTributoDestino são tipos simples/variant, não precisam Set = Nothing
End Sub

Private Sub FecharArquivoImportacao(ByRef wb As Workbook)
    If Not wb Is Nothing Then
        If wb.name <> ThisWorkbook.name Then ' Segurança extra
            Application.DisplayAlerts = False
            wb.Close SaveChanges:=False
            Application.DisplayAlerts = True
        End If
        Set wb = Nothing
    End If
End Sub

'=======================================================================================
' LÓGICA DE VALIDAÇÃO
'=======================================================================================

Private Function ValidarLayoutArquivoImportacao(ByRef config As ConfiguracaoImportacao) As String
    AtualizarStatusBar "Validando layout do arquivo..."

    If Not TitulosArquivoContemCamposChave(config.DicTitulosOrigem, config.CamposChaveTributo) Then
        Dim msgChave As String
        msgChave = "O arquivo não contém as colunas chave obrigatórias." & vbCrLf
        If IsArray(config.CamposChaveTributo) Then
             msgChave = msgChave & "Esperado: " & VBA.Join(config.CamposChaveTributo, ", ")
        ElseIf Not IsEmpty(config.CamposChaveTributo) Then
             msgChave = msgChave & "Esperado: " & config.CamposChaveTributo
        End If
        ValidarLayoutArquivoImportacao = msgChave
        Exit Function
    End If

    If Not TitulosArquivoCorrespondemAoTipoTributo(config.DicTitulosOrigem, config.NomeTributoDestino) Then
        Dim msgTipo As String
        Dim colsEsperadas As Variant
        msgTipo = "O arquivo não parece ser do tipo (" & config.NomeTributoDestino & ")." & vbCrLf
        msgTipo = msgTipo & "Não contém as colunas específicas esperadas."
        colsEsperadas = ObterNomesColunasObrigatoriasParaTributo(config.NomeTributoDestino)
        If IsArray(colsEsperadas) Then
             msgTipo = msgTipo & vbCrLf & "Esperadas: " & VBA.Join(colsEsperadas, ", ")
        End If
        ValidarLayoutArquivoImportacao = msgTipo
        Exit Function
    End If

    ValidarLayoutArquivoImportacao = "" ' String vazia indica sucesso
End Function

Private Function TitulosArquivoContemCamposChave(ByVal titulosArquivo As Dictionary, ByVal camposChaveNecessarios As Variant) As Boolean
    Dim Mapeamento As Long
    Dim TotalCamposChave As Long
    Dim Campo As Variant

    If titulosArquivo Is Nothing Then Exit Function ' Falso se não há títulos

    TotalCamposChave = ContarElementos(camposChaveNecessarios)
    If TotalCamposChave = 0 Then
        TitulosArquivoContemCamposChave = True ' Válido se não há chaves a verificar
        Exit Function
    End If

    Mapeamento = 0
    If IsArray(camposChaveNecessarios) Then
        For Each Campo In camposChaveNecessarios
            If titulosArquivo.Exists(Campo) Then Mapeamento = Mapeamento + 1
        Next Campo
    Else
        If titulosArquivo.Exists(camposChaveNecessarios) Then Mapeamento = 1
    End If

    TitulosArquivoContemCamposChave = (Mapeamento >= TotalCamposChave)
End Function

Private Function TitulosArquivoCorrespondemAoTipoTributo(ByVal titulosArquivo As Dictionary, ByVal nomeTributoEsperado As String) As Boolean
    Dim colunasObrigatorias As Variant
    Dim Coluna As Variant

    If titulosArquivo Is Nothing Then Exit Function ' Falso se não há títulos

    colunasObrigatorias = ObterNomesColunasObrigatoriasParaTributo(nomeTributoEsperado)

    If Not IsArray(colunasObrigatorias) Then
        TitulosArquivoCorrespondemAoTipoTributo = True ' Válido se não há colunas específicas a checar
        Exit Function
    End If

    For Each Coluna In colunasObrigatorias
        If Not titulosArquivo.Exists(Coluna) Then
            TitulosArquivoCorrespondemAoTipoTributo = False ' Falha se uma coluna obrigatória falta
            Exit Function
        End If
    Next Coluna

    TitulosArquivoCorrespondemAoTipoTributo = True ' Todas as colunas obrigatórias encontradas
End Function

Private Function ObterNomesColunasObrigatoriasParaTributo(ByVal NomeTributo As String) As Variant
    Select Case UCase(NomeTributo)
        Case "ICMS": ObterNomesColunasObrigatoriasParaTributo = Array("CST_ICMS", "ALIQ_ICMS")
        Case "IPI": ObterNomesColunasObrigatoriasParaTributo = Array("CST_IPI", "ALIQ_IPI")
        Case "PIS", "COFINS", "PIS_COFINS", "PIS E COFINS": ObterNomesColunasObrigatoriasParaTributo = Array("CST_PIS", "ALIQ_PIS", "CST_COFINS", "ALIQ_COFINS")
        Case Else: ObterNomesColunasObrigatoriasParaTributo = Empty
    End Select
End Function

'=======================================================================================
' PROCESSAMENTO DE DADOS
'=======================================================================================

Private Function ProcessarLinhasArquivoImportacao(ByVal planOrigem As Worksheet, ByRef config As ConfiguracaoImportacao) As ArrayList
    Dim intervaloDados As Range
    Dim linhaAtual As Range
    Dim dadosLinhaArquivo As Variant ' Receberá o array 1D
    Dim chaveTributacao As String
    Dim linhasParaAdicionar As New ArrayList
    Dim tempoInicioLoop As Double
    Dim linhaAtualIndex As Long

    On Error Resume Next
    If planOrigem.AutoFilterMode Then planOrigem.AutoFilter.ShowAllData
    On Error GoTo 0

    Set intervaloDados = Util.DefinirIntervalo(planOrigem, 2, 1)
    If intervaloDados Is Nothing Then
        Set ProcessarLinhasArquivoImportacao = linhasParaAdicionar
        Exit Function
    End If

    tempoInicioLoop = Timer()
    linhaAtualIndex = 0

    For Each linhaAtual In intervaloDados.Rows
        linhaAtualIndex = linhaAtualIndex + 1
        Call Util.AntiTravamento(linhaAtualIndex, 100, "Processando linha " & linhaAtualIndex & " de " & intervaloDados.Rows.Count, intervaloDados.Rows.Count, tempoInicioLoop)

        ' *** Alteração: Usando Application.Index para obter array 1D ***
        dadosLinhaArquivo = Application.index(linhaAtual.Value2, 0, 0)

        If Util.ChecarCamposPreenchidos(dadosLinhaArquivo) Then ' Assume que Util.ChecarCamposPreenchidos lida com array 1D
            chaveTributacao = GerarChaveTributacaoUnica(config.CamposChaveTributo, dadosLinhaArquivo, config.DicTitulosOrigem)

            If Not config.DicDadosExistentes.Exists(chaveTributacao) Then
                Dim dadosLinhaDestino As Variant
                dadosLinhaDestino = MapearDadosLinhaParaDestino(dadosLinhaArquivo, config.DicTitulosOrigem, config.DicTitulosDestino)

                If Util.ChecarCamposPreenchidos(dadosLinhaDestino) Then
                     linhasParaAdicionar.Add dadosLinhaDestino
                End If
            End If
        End If
    Next linhaAtual

    Set ProcessarLinhasArquivoImportacao = linhasParaAdicionar
End Function

Private Function GerarChaveTributacaoUnica(ByVal nomesCamposChave As Variant, ByVal dadosLinha As Variant, ByVal titulosOrigem As Dictionary) As String
    Dim partesChave As New ArrayList
    Dim separador As String: separador = Chr(7)
    Dim campoNome As Variant
    Dim indiceColuna As Long
    Dim valorCampo As String

    ' Verifica se dadosLinha é um array válido antes de prosseguir
    If Not IsArray(dadosLinha) Then Exit Function

    If IsArray(nomesCamposChave) Then
        For Each campoNome In nomesCamposChave
            If titulosOrigem.Exists(campoNome) Then
                indiceColuna = titulosOrigem(campoNome)
                ' *** Alteração: Acesso 1D e verificação de limites ***
                If indiceColuna >= LBound(dadosLinha) And indiceColuna <= UBound(dadosLinha) Then
                    valorCampo = Util.RemoverAspaSimples(CStr(dadosLinha(indiceColuna)))
                    partesChave.Add valorCampo
                Else
                    partesChave.Add "" ' Índice fora dos limites
                End If
            Else
                partesChave.Add "" ' Campo chave não encontrado nos títulos
            End If
        Next campoNome
    ElseIf Not IsEmpty(nomesCamposChave) Then
        If titulosOrigem.Exists(nomesCamposChave) Then
             indiceColuna = titulosOrigem(nomesCamposChave)
             ' *** Alteração: Acesso 1D e verificação de limites ***
             If indiceColuna >= LBound(dadosLinha) And indiceColuna <= UBound(dadosLinha) Then
                 valorCampo = Util.RemoverAspaSimples(CStr(dadosLinha(indiceColuna)))
                 partesChave.Add valorCampo
             Else
                 partesChave.Add ""
             End If
        Else
             partesChave.Add ""
        End If
    End If

    GerarChaveTributacaoUnica = VBA.Join(partesChave.toArray(), separador)
    Set partesChave = Nothing
End Function

Private Function MapearDadosLinhaParaDestino(ByVal dadosOrigem As Variant, ByVal titulosOrigem As Dictionary, ByVal titulosDestino As Dictionary) As Variant
    Dim dadosDestino() As Variant
    Dim tituloDestino As Variant
    Dim indiceDestino As Long
    Dim indiceOrigem As Long
    Dim colunasIgnoradas As String: colunasIgnoradas = ",INCONSISTENCIA,SUGESTAO,"

    If titulosDestino Is Nothing Or titulosDestino.Count = 0 Then Exit Function
    ' Verifica se dadosOrigem é um array válido
    If Not IsArray(dadosOrigem) Then Exit Function

    ReDim dadosDestino(1 To titulosDestino.Count)

    For Each tituloDestino In titulosDestino.Keys
        indiceDestino = titulosDestino(tituloDestino)

        If InStr(1, colunasIgnoradas, "," & tituloDestino & ",", vbTextCompare) = 0 Then
            If titulosOrigem.Exists(tituloDestino) Then
                indiceOrigem = titulosOrigem(tituloDestino)
                 ' *** Alteração: Acesso 1D e verificação de limites ***
                If indiceOrigem >= LBound(dadosOrigem) And indiceOrigem <= UBound(dadosOrigem) Then
                     
                     If tituloDestino Like "ALIQ_*" Then dadosDestino(indiceDestino) = dadosOrigem(indiceOrigem) _
                        Else dadosDestino(indiceDestino) = fnExcel.FormatarTipoDado(tituloDestino, dadosOrigem(indiceOrigem))
                        
                Else
                     dadosDestino(indiceDestino) = vbNullString ' Índice origem inválido
                End If
            Else
                dadosDestino(indiceDestino) = vbNullString ' Coluna destino não existe na origem
            End If
        Else
             dadosDestino(indiceDestino) = vbNullString ' Coluna ignorada
        End If
    Next tituloDestino

    MapearDadosLinhaParaDestino = dadosDestino
End Function

Private Sub ExportarNovasTributacoes(ByVal PlanDestino As Worksheet, ByVal novasLinhas As ArrayList)
    Call Util.ExportarDadosArrayList(PlanDestino, novasLinhas) ' Assume Util.ExportarDadosArrayList existe
End Sub

'=======================================================================================
' FUNÇÕES UTILITÁRIAS INTERNAS E FEEDBACK
'=======================================================================================

Private Function ContarElementos(ByVal item As Variant) As Long
    If IsArray(item) Then
        If UBound(item) < LBound(item) Then ContarElementos = 0 Else ContarElementos = UBound(item) - LBound(item) + 1
    ElseIf Not IsEmpty(item) Then
        ContarElementos = 1
    Else
        ContarElementos = 0
    End If
End Function

Private Sub AtualizarStatusBar(ByVal Texto As Variant)
    Application.StatusBar = Texto
End Sub

Private Sub ExibirMensagemAlerta(ByVal Titulo As String, ByVal Mensagem As String)
    Call Util.MsgAlerta(Mensagem, Titulo)
End Sub

Private Sub ExibirMensagemSucesso(ByVal Titulo As String, ByVal Mensagem As String, ByVal tempoInicio As Date)
    Call Util.MsgInformativa(Mensagem, Titulo, tempoInicio)
End Sub
