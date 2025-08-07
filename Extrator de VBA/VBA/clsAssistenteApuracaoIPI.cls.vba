Attribute VB_Name = "clsAssistenteApuracaoIPI"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private dicTitulos As New Dictionary
Private arrRelatorio As New ArrayList
Private Apuracao As clsAssistenteApuracao
Private ValidacoesNCM As New clsRegrasFiscaisNCM
Private ValidacoesCFOP As New clsRegrasFiscaisCFOP

Public Function GerarApuracaoAssistidaIPI()

Dim arrDocsC100 As New ArrayList
Dim Msg As String
    
    Inicio = Now()
    
    Call arrRelatorio.Clear
    Call dicInconsistenciasIgnoradas.RemoveAll
    Call DadosValidacaoCFOP.CarregarTitulosRelatorio(ActiveSheet)
    
    Set dicTitulos = Util.MapearTitulos(assApuracaoIPI, 3)
    Set Apuracao = New clsAssistenteApuracao
    
    Call CarregarDadosC170(arrDocsC100, Msg)
    Call CarregarDadosC190(arrDocsC100, Msg)
    
    Application.StatusBar = "Processo concluído com sucesso!"
    If arrRelatorio.Count > 0 Then
        
        On Error Resume Next
            If assApuracaoIPI.AutoFilter.FilterMode Then assApuracaoIPI.ShowAllData
        On Error GoTo 0
        Call Util.LimparDados(assApuracaoIPI, 4, False)
        
        Call Util.ExportarDadosArrayList(assApuracaoIPI, arrRelatorio)
        Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoIPI)
        Call Util.MsgInformativa("Relatório gerado com sucesso", "Assistente de Apuração do IPI", Inicio)
        
    Else
        
        Msg = "Nenhum dado encontrado para geração do relatório." & vbCrLf & vbCrLf
        Msg = Msg & "Por favor verifique se o SPED foi importado e tente novamente."
        Call Util.MsgAlerta(Msg, "Assistente de Apuração do IPI")
        
    End If
    
    Application.StatusBar = False
    
    Set Apuracao = Nothing
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    
End Function

Public Sub CarregarDadosC170(ByRef arrDocsC100 As ArrayList, ByVal Msg As String)

Dim ARQUIVO As String, CHV_REG$, CHV_0001$, CHV_0150$, CHV_C100$, COD_ITEM$, COD_PART$, COD_NAT$, UF_CONTRIB$, CONTRIBUINTE$, CNPJ_ESTABELECIMENTO$
Dim dicTitulosC170 As New Dictionary
Dim Dados As Range, Linha As Range
Dim Comeco As Double, VL_ITEM#
Dim Campos As Variant
Dim b As Long
        
    Set Dados = Util.DefinirIntervalo(regC170, 4, 3)
    If Dados Is Nothing Then Exit Sub
    
    Set Apuracao.dicTitulos = dicTitulos
    Set dicTitulosC170 = Util.MapearTitulos(regC170, 3)
    
    b = 0
    Comeco = Timer
    With Apuracao
        
        Call .CarregarDadosRegistro0000
        
        For Each Linha In Dados.Rows
            
            .RedimensionarArray (dicTitulos.Count)
            Call Util.AntiTravamento(b, 100, Msg & "Carregando dados do registro C170", Dados.Rows.Count, Comeco)
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                'Carrega as variáveis necessárias
                ARQUIVO = Campos(dicTitulosC170("ARQUIVO"))
                UF_CONTRIB = .ExtrairUFContribuinte(ARQUIVO)
                CNPJ_ESTABELECIMENTO = fnExcel.FormatarTexto(VBA.Split(ARQUIVO, "-")(1))
                
                CHV_C100 = Campos(dicTitulosC170("CHV_PAI_FISCAL"))
                If CHV_C100 <> "" Then arrDocsC100.Add CHV_C100
                
                COD_PART = .ExtrairCOD_PART_C100(CHV_C100)
                CHV_0001 = .ExtrairCHV_0001(ARQUIVO)
                COD_ITEM = Campos(dicTitulosC170("COD_ITEM"))
                COD_NAT = Campos(dicTitulosC170("COD_NAT"))
                VL_ITEM = fnExcel.FormatarValores(Campos(dicTitulosC170("VL_ITEM")))
                
                'Extrai dados do registro C100
                Call .ExtrairDadosC100(CHV_C100)
                
                'Extrai dados do registro 0150
                Call .ExtrairDados0150(CHV_0001, COD_PART, False, True)
                
                'Extrai dados do registro 0200
                Call .ExtrairDados0200(CHV_0001, COD_ITEM, False, True)
                
                'Extrai dados do registro 0400
                'Call .ExtrairDados0400(ARQUIVO, COD_NAT)
                
                'Atribui valores aos campos do relatório
                .AtribuirValor "REG", Campos(dicTitulosC170("REG"))
                .AtribuirValor "ARQUIVO", ARQUIVO
                .AtribuirValor "CHV_PAI_FISCAL", CHV_C100
                .AtribuirValor "CHV_REG", Campos(dicTitulosC170("CHV_REG"))
                .AtribuirValor "CNPJ_ESTABELECIMENTO", CNPJ_ESTABELECIMENTO
                .AtribuirValor "UF_CONTRIB", UF_CONTRIB
                .AtribuirValor "COD_ITEM", fnExcel.FormatarTexto(COD_ITEM)
                .AtribuirValor "IND_MOV", Campos(dicTitulosC170("IND_MOV"))
                .AtribuirValor "IND_APUR", Campos(dicTitulosC170("IND_APUR"))
                .AtribuirValor "COD_ENQ", Campos(dicTitulosC170("COD_ENQ"))
                .AtribuirValor "CFOP", Campos(dicTitulosC170("CFOP"))
                .AtribuirValor "VL_ITEM", VL_ITEM
                .AtribuirValor "VL_DESP", .ExtrairVL_DESP_C100(CHV_C100, VL_ITEM)
                .AtribuirValor "VL_DESC", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_DESC")))
                .AtribuirValor "CST_IPI", fnExcel.FormatarTexto(Campos(dicTitulosC170("CST_IPI")))
                .AtribuirValor "VL_BC_IPI", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_BC_IPI")))
                .AtribuirValor "ALIQ_IPI", fnExcel.ConverterValores(Campos(dicTitulosC170("ALIQ_IPI")))
                .AtribuirValor "VL_IPI", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_IPI")))
                
            End If
            
            Campos = ValidarRegrasFiscais(.Campo)
            arrRelatorio.Add Campos
            
        Next Linha
    
    End With
    
End Sub

Public Sub CarregarDadosC190(ByRef arrDocsC100 As ArrayList, ByVal Msg As String)

Dim ARQUIVO As String, CHV_REG$, CHV_0001$, CHV_0150$, CHV_C100$, COD_PART$, COD_ITEM$, COD_NAT$, UF_CONTRIB$, CONTRIBUINTE$, CNPJ_ESTABELECIMENTO$
Dim dicTitulosC190 As New Dictionary
Dim Dados As Range, Linha As Range
Dim Comeco As Double, VL_OPR#
Dim Campos As Variant
Dim b As Long
        
    Set Dados = Util.DefinirIntervalo(regC190, 4, 3)
    If Dados Is Nothing Then Exit Sub
    
    Set Apuracao.dicTitulos = dicTitulos
    Set dicTitulosC190 = Util.MapearTitulos(regC190, 3)
    
    b = 0
    Comeco = Timer
    With Apuracao
        
        Call .CarregarDadosRegistro0000
        
        For Each Linha In Dados.Rows
            
            .RedimensionarArray (dicTitulos.Count)
            Call Util.AntiTravamento(b, 100, Msg & "Carregando dados do registro C190", Dados.Rows.Count, Comeco)
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                'Carrega as variáveis necessárias
                ARQUIVO = Campos(dicTitulosC190("ARQUIVO"))
                UF_CONTRIB = .ExtrairUFContribuinte(ARQUIVO)
                CNPJ_ESTABELECIMENTO = fnExcel.FormatarTexto(VBA.Split(ARQUIVO, "-")(1))
                
                CHV_C100 = Campos(dicTitulosC190("CHV_PAI_FISCAL"))
                If arrDocsC100.contains(CHV_C100) Then GoTo Prx:
                
                COD_PART = .ExtrairCOD_PART_C100(CHV_C100)
                CHV_0001 = .ExtrairCHV_0001(ARQUIVO)
                VL_OPR = fnExcel.FormatarValores(Campos(dicTitulosC190("VL_OPR")))
                
                'Extrai dados do registro C100
                Call .ExtrairDadosC100(CHV_C100)
                
                'Extrai dados do registro 0150
                Call .ExtrairDados0150(CHV_0001, COD_PART, False, True)
                
                'Extrai dados do registro 0200
                Call .ExtrairDados0200(CHV_0001, COD_ITEM, False, True)
                
                'Atribui valores aos campos do relatório
                .AtribuirValor "UF_CONTRIB", UF_CONTRIB
                .AtribuirValor "DESCR_ITEM", "O REGISTRO C190 NÃO POSSUI DADOS DE PRODUTOS"
                .AtribuirValor "REG", Campos(dicTitulosC190("REG"))
                .AtribuirValor "ARQUIVO", ARQUIVO
                .AtribuirValor "CHV_PAI_FISCAL", CHV_C100
                .AtribuirValor "CHV_REG", Campos(dicTitulosC190("CHV_REG"))
                .AtribuirValor "CNPJ_ESTABELECIMENTO", CNPJ_ESTABELECIMENTO
                .AtribuirValor "CFOP", Campos(dicTitulosC190("CFOP"))
                .AtribuirValor "VL_ITEM", VL_OPR
                .AtribuirValor "VL_DESP", .ExtrairVL_DESP_C100(CHV_C100, VL_OPR)
                .AtribuirValor "VL_IPI", fnExcel.ConverterValores(Campos(dicTitulosC190("VL_IPI")))
                
            End If
            
            Campos = ValidarRegrasFiscais(.Campo)
            arrRelatorio.Add Campos
Prx:
        Next Linha
        
    End With
    
End Sub

Public Function ValidarRegrasFiscais(ByRef Campos As Variant) As Variant

Dim REG As String, UF$
Dim AjustePosicaoArray As Integer
    
    If UBound(Campos) = -1 Then
        ValidarRegrasFiscais = Campos
        Exit Function
    End If
    
    If LBound(Campos) = 0 Then AjustePosicaoArray = 1 Else AjustePosicaoArray = 0
    REG = Campos(dicTitulos("REG") - AjustePosicaoArray)
    
    Select Case REG
        
        Case "C170"
            Call ValidarRegrasFiscaisC170(Campos)
            
        Case "C190"
            Call ValidarRegrasFiscaisC190(Campos)
            
    End Select
    
    ValidarRegrasFiscais = Campos
    
End Function

Public Function ValidarRegrasFiscaisC170(ByRef Campos As Variant)

Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call ValidacoesCFOP.ValidarCampo_CFOP(Campos, "IPI")
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoIPI.ValidarCampo_DT_ENT_SAI(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call ValidacoesNCM.ValidarCampo_COD_NCM(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_IND_MOV(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_TIPO_ITEM(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoIPI.ValidarCampo_IND_APUR(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoIPI.ValidarCampo_VL_IPI(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoIPI.ValidarCampo_ALIQ_IPI(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoIPI.ValidarCampo_CST_IPI(Campos, dicTitulos)
    
End Function

Public Function ValidarRegrasFiscaisC190(ByRef Campos As Variant)

Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call ValidacoesCFOP.ValidarCampo_CFOP(Campos, "IPI")
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoIPI.ValidarCampo_IND_APUR(Campos, dicTitulos)
    
End Function

Public Function ReprocessarSugestoes()

Dim Dados As Range, Linha As Range
Dim arrRelatorio As New ArrayList
Dim Campos As Variant
    
    Set dicTitulos = Util.MapearTitulos(assApuracaoIPI, 3)
    If assApuracaoIPI.AutoFilterMode Then assApuracaoIPI.AutoFilter.ShowAllData
    
    Call DadosValidacaoCFOP.CarregarTitulosRelatorio(ActiveSheet)
    Set Dados = Util.DefinirIntervalo(assApuracaoIPI, 4, 3)
    If Dados Is Nothing Then Exit Function
    
    a = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 100, "Reprocessando sugestões, por favor aguarde...", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            Campos(dicTitulos("INCONSISTENCIA")) = Empty
            Campos(dicTitulos("SUGESTAO")) = Empty
            Call ValidarRegrasFiscais(Campos)
            
            arrRelatorio.Add Campos
            
        End If
        
    Next Linha
    
    Call Util.LimparDados(assApuracaoIPI, 4, False)
    Call Util.ExportarDadosArrayList(assApuracaoIPI, arrRelatorio)
    Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoIPI)
    
    Call Util.AtualizarBarraStatus("Processamento Concluído!")
        
End Function

Public Function AceitarSugestoes()

Dim Dados As Range, Linha As Range
Dim arrDados As New ArrayList
Dim UltimaSugestao As String
Dim Campos As Variant
    
    Inicio = Now()
    Application.StatusBar = "Implementando sugestões selecionadas, por favor aguarde..."
    
    Set dicTitulos = Util.MapearTitulos(assApuracaoIPI, 3)
    Set Dados = assApuracaoIPI.Range("A4").CurrentRegion
    If Dados Is Nothing Then
        Call Util.MsgAlerta("Não foi encontrada nenhuma inconsistência!", "Inconsistências Fiscais")
        Exit Function
    End If
    
    For Each Linha In Dados.Rows
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
                        
            If Linha.EntireRow.Hidden = False And Campos(dicTitulos("SUGESTAO")) <> "" And Linha.Row > 3 Then
VERIFICAR:
                Call Util.AntiTravamento(a, 10, "Aplicando sugestões sugeridas...", Dados.Rows.Count, Comeco)
                Select Case Campos(dicTitulos("SUGESTAO"))
                    
                   Case "Alterar valor do campo COD_SIT para: 08 - Regime Especial ou Norma Específica"
                        Campos(dicTitulos("COD_SIT")) = "08 - Regime Especial ou Norma Específica"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Call ValidarRegrasFiscais(Campos)
                        
                    Case "Informar CST_IPI 00 - Entrada com recuperação de crédito"
                        Campos(dicTitulos("CST_IPI")) = "00 - Entrada com Recuperação de Crédito"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Call ValidarRegrasFiscais(Campos)
                        GoTo VERIFICAR:
                    
                    Case "Informar CST_IPI 49 - Outras Entradas"
                        Campos(dicTitulos("CST_IPI")) = "49 - Outras Entradas"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Call ValidarRegrasFiscais(Campos)
                        GoTo VERIFICAR:
                        
                    Case "Zerar campos do IPI"
                        Campos(dicTitulos("VL_BC_IPI")) = 0
                        Campos(dicTitulos("ALIQ_IPI")) = 0
                        Campos(dicTitulos("VL_IPI")) = 0
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = ValidarRegrasFiscais(Campos)
                        GoTo VERIFICAR:
                        
                    Case "Alterar o valor do campo TIPO_ITEM para 00"
                        Campos(dicTitulos("TIPO_ITEM")) = "00 - Mercadoria para Revenda"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = ValidarRegrasFiscais(Campos)
                        GoTo VERIFICAR:
                        
                    Case "Alterar o valor do campo TIPO_ITEM para 07"
                        Campos(dicTitulos("TIPO_ITEM")) = "07 - Material de Uso e Consumo"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = ValidarRegrasFiscais(Campos)
                        GoTo VERIFICAR:
                        
                    Case "Alterar o valor do campo TIPO_ITEM para 08"
                        Campos(dicTitulos("TIPO_ITEM")) = "08 - Ativo Imobilizado"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = ValidarRegrasFiscais(Campos)
                        GoTo VERIFICAR:
                        
                    Case "Apagar valor informado no campo COD_NCM"
                        Campos(dicTitulos("COD_NCM")) = ""
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = ValidarRegrasFiscais(Campos)
                        GoTo VERIFICAR:
                        
                    Case "Adicionar zeros a esquerda do campo COD_NCM"
                        Campos(dicTitulos("CEST")) = "'" & VBA.Format(Util.ApenasNumeros(Campos(dicTitulos("CEST"))), VBA.String(7, "0"))
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = ValidarRegrasFiscais(Campos)
                        GoTo VERIFICAR:
                        
                    Case "Zerar Alíquota do IPI"
                        Campos(dicTitulos("ALIQ_IPI")) = 0
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = ValidarRegrasFiscais(Campos)
                        GoTo VERIFICAR:
                    
                    Case "Recalcular o campos VL_IPI"
                        Campos(dicTitulos("VL_IPI")) = fnExcel.ConverterValores(Campos(dicTitulos("VL_BC_IPI")) * Campos(dicTitulos("ALIQ_IPI")), True, 2)
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = ValidarRegrasFiscais(Campos)
                        GoTo VERIFICAR:
                        
                End Select
                
            End If
            
            If Linha.Row > 3 Then arrDados.Add Campos
            
        End If
        
    Next Linha

    If arrDados.Count = 0 Then
        Call Util.MsgAlerta("Não existem sugestões para processar!", "Sugestões Fiscais")
        Exit Function
    End If
    
    
    If assApuracaoIPI.AutoFilterMode Then assApuracaoIPI.AutoFilter.ShowAllData
    Call Util.LimparDados(assApuracaoIPI, 4, False)
    Call Util.ExportarDadosArrayList(assApuracaoIPI, arrDados)
    
    Call FuncoesFormatacao.AplicarFormatacao(ActiveSheet)
    Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoIPI)
    
    Call Util.MsgInformativa("Sugestões aplicadas com sucesso!", "Inclusão de Sugestões", Inicio)
    Application.StatusBar = False
    
End Function

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
    
    Set dicTitulos = Util.MapearTitulos(assApuracaoIPI, 3)
    Set Dados = assApuracaoIPI.Range("A4").CurrentRegion
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
                Call ValidarRegrasFiscais(Campos)
                
            End If
            
            If Linha.Row > 3 Then dicDados(Campos(dicTitulos("CHV_REG"))) = Campos
            
        End If
        
    Next Linha

    If dicInconsistenciasIgnoradas.Count = 0 Then
        Call Util.MsgAlerta("Não existem Inconsistêncais a ignorar!", "Ignorar Inconsistências")
        Exit Function
    End If
    
    If assApuracaoIPI.AutoFilterMode Then assApuracaoIPI.AutoFilter.ShowAllData
    Call Util.LimparDados(assApuracaoIPI, 4, False)
    Call Util.ExportarDadosDicionario(assApuracaoIPI, dicDados)
    
    Call FuncoesFormatacao.AplicarFormatacao(ActiveSheet)
    Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoIPI)
    
    Call Util.MsgInformativa("Inconsistências ignoradas com sucesso!", "Ignorar Inconsistências", Inicio)
    Application.StatusBar = False
    
End Function

Public Sub AtualizarRegistros()

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
Dim Status As Boolean

'    If Otimizacoes.OtimizacoesAtivas Then
'
'        Status = Otimizacoes.OtimizarAtualizacaoRegistros(assApuracaoIPI)
'        If Not Status Then Exit Sub
'
'    End If
    
    Inicio = Now()
    Application.StatusBar = "Preparando dados para atualização do SPED, por favor aguarde..."
    
    Campos0200 = Array("REG", "COD_BARRA", "COD_NCM", "EX_IPI", "TIPO_ITEM")
    CamposC100 = Array("CHV_NFE", "NUM_DOC", "SER")
    CamposC170 = Array("IND_MOV", "IND_APUR", "COD_ENQ", "CFOP", "CST_IPI", "VL_ITEM", "VL_DESC", "VL_BC_IPI", "ALIQ_IPI", "VL_IPI")
    
    Set dicTitulos0200 = Util.MapearTitulos(reg0200, 3)
    Set dicDados0200 = Util.CriarDicionarioRegistro(reg0200, "ARQUIVO", "COD_ITEM")
    
    Set dicTitulosC100 = Util.MapearTitulos(regC100, 3)
    Set dicDadosC100 = Util.CriarDicionarioRegistro(regC100)
    
    Set dicTitulosC170 = Util.MapearTitulos(regC170, 3)
    Set dicDadosC170 = Util.CriarDicionarioRegistro(regC170)
    
    Set dicTitulos = Util.MapearTitulos(assApuracaoIPI, 3)
    If assApuracaoIPI.AutoFilterMode Then assApuracaoIPI.AutoFilter.ShowAllData
    
    Set Dados = Util.DefinirIntervalo(assApuracaoIPI, 4, 3)
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
                    
                    If regCampo = "CST_IPI" Or regCampo = "COD_BARRA" Then
                        Campos(dicTitulos(regCampo)) = fnExcel.FormatarTexto(Campos(dicTitulos(regCampo)))
                    End If
                    
                    If regCampo Like "VL_*" Then Campos(dicTitulos(regCampo)) = VBA.Round(Campos(dicTitulos(regCampo)), 2)
                    dicCampos(dicTitulosC170(regCampo)) = Campos(dicTitulos(regCampo))
                
                Next regCampo
                
                dicDadosC170(CHV_C170) = dicCampos
                
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
    
    Application.StatusBar = "Atualizando dados do registro C190, por favor aguarde..."
    Call rC170.GerarC190(True)
    
    Application.StatusBar = "Atualizando valores dos impostos no registro C100, por favor aguarde..."
    Call rC170.AtualizarImpostosC100(True)
    Call r0200.AtualizarCodigoGenero(True)
    
    'Atualizando Saldos de Apuração do IPI
    Call AtualizarApuracaoIPI

    Call FuncoesFormatacao.AplicarFormatacao(ActiveSheet)
    Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoIPI)
    
    Application.StatusBar = "Atualização concluída com sucesso!"
    Call Util.MsgInformativa("Registros atualizados com sucesso!", "Atualização de dados", Inicio)
    Application.StatusBar = False
    
End Sub

Public Function ListarTributacoesIPI()

Dim Tributacao As New AssistenteTributario
        
    Call Tributacao.ListarTributacoes(assApuracaoIPI, assTributacaoIPI)
    
End Function

Public Sub AtualizarApuracaoIPI()

Dim dicTitulos As New Dictionary

Dim Dados As Range, Linha As Range
Dim dicDados0000 As New Dictionary
Dim dicDadosE001 As New Dictionary
Dim dicDadosE500 As New Dictionary
Dim dicDadosE510 As New Dictionary
Dim dicDadosE520 As New Dictionary
Dim dicDadosE530 As New Dictionary
Dim dicSaidasE510 As New Dictionary
Dim dicEntradasE510 As New Dictionary
Dim dicTitulos0000 As New Dictionary
Dim dicTitulosE001 As New Dictionary
Dim dicTitulosE500 As New Dictionary
Dim dicTitulosE510 As New Dictionary
Dim dicTitulosE520 As New Dictionary
Dim dicTitulosE530 As New Dictionary
Dim dicAuxiliarE500 As New Dictionary
Dim dicAuxiliarE510 As New Dictionary
Dim dicAuxiliarE520 As New Dictionary
Dim dicAuxiliarE530 As New Dictionary

Dim VL_BC_IPI As Double, VL_CONT_IPI#, VL_IPI#, VL_ITEM#, VL_DESP#, VL_DESC#
Dim Campos As Variant, Campos0000, CamposE001, CamposE510, CamposIPI, dicCampos, regCampo
Dim CHV_0000 As String, CHV_E001$, CHV_E500$, ARQUIVO$, IND_APUR$, DT_INI$, DT_FIN$, DT_ENT_SAI$, DT_DOC$, CFOP$, CST_IPI$
    
    Inicio = Now()
    Application.StatusBar = "Gerando apuração do IPI, por favor aguarde..."
    
    CamposIPI = Array("IND_APUR", "CFOP", "CST_IPI", "VL_ITEM", "VL_DESC", "VL_BC_IPI", "VL_IPI")
    
    'carregando títulos dos registros do IPI
    Set dicTitulos = Util.MapearTitulos(assApuracaoIPI, 3)
    Set dicTitulos0000 = Util.MapearTitulos(reg0000, 3)
    Set dicTitulosE001 = Util.MapearTitulos(regE001, 3)
    Set dicTitulosE500 = Util.MapearTitulos(regE500, 3)
    Set dicTitulosE510 = Util.MapearTitulos(regE510, 3)
    Set dicTitulosE520 = Util.MapearTitulos(regE520, 3)
    Set dicTitulosE530 = Util.MapearTitulos(regE530, 3)
    
    'Carregando dados dos registros do IPI
    Set dicDados0000 = Util.CriarDicionarioRegistro(reg0000, "ARQUIVO")
    Set dicDadosE001 = Util.CriarDicionarioRegistro(regE001, "ARQUIVO")
    Set dicDadosE500 = Util.CriarDicionarioRegistro(regE500)
    Set dicDadosE510 = Util.CriarDicionarioRegistro(regE510, "CHV_PAI_FISCAL", "CFOP", "CST_IPI")
    Set dicDadosE520 = Util.CriarDicionarioRegistro(regE520)
    Set dicDadosE530 = Util.CriarDicionarioRegistro(regE530, "CHV_PAI_FISCAL", "IND_AJ", "COD_AJ", "IND_DOC", "DESCR_AJ")
    
    Call Util.SegregarEntradasSaidas(dicDadosE510, dicTitulosE510, dicEntradasE510, dicSaidasE510)
    
    If regC170.AutoFilterMode Then regC170.AutoFilter.ShowAllData
    
    Set Dados = Util.DefinirIntervalo(assApuracaoIPI, 4, 3)
    If Dados Is Nothing Then Exit Sub
    
    a = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 100, "Atualizar dados da apuração do IPI, por favor aguarde...", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            ARQUIVO = Campos(dicTitulos("ARQUIVO"))
            
            'Extrai informações do registro 0000 do SPED FISCAL
            CHV_0000 = fnSPED.ExtrairCampoDicionario(dicDados0000, dicTitulos0000, ARQUIVO, "CHV_REG")
            DT_INI = fnExcel.FormatarData(fnSPED.ExtrairCampoDicionario(dicDados0000, dicTitulos0000, ARQUIVO, "DT_INI"))
            DT_FIN = fnExcel.FormatarData(fnSPED.ExtrairCampoDicionario(dicDados0000, dicTitulos0000, ARQUIVO, "DT_FIN"))
            
            'Extrai chave do registro E001
            CHV_E001 = fnSPED.ExtrairCampoDicionario(dicDadosE001, dicTitulosE001, ARQUIVO, "CHV_REG")
            
            'Extrair campos do assistente de tributação do IPI
            CFOP = fnSPED.ExtrairCampoArray(Campos, dicTitulos, "CFOP")
            VL_ITEM = fnExcel.ConverterValores(fnSPED.ExtrairCampoArray(Campos, dicTitulos, "VL_ITEM"))
            VL_DESP = fnExcel.ConverterValores(fnSPED.ExtrairCampoArray(Campos, dicTitulos, "VL_DESP"))
            VL_DESC = fnExcel.ConverterValores(fnSPED.ExtrairCampoArray(Campos, dicTitulos, "VL_DESC"))
            VL_CONT_IPI = VL_ITEM + VL_DESP - VL_DESC
            
            CST_IPI = fnSPED.ExtrairCampoArray(Campos, dicTitulos, "CST_IPI")
            VL_BC_IPI = fnExcel.ConverterValores(fnSPED.ExtrairCampoArray(Campos, dicTitulos, "VL_BC_IPI"))
            VL_IPI = fnExcel.ConverterValores(fnSPED.ExtrairCampoArray(Campos, dicTitulos, "VL_IPI"))
            DT_DOC = fnExcel.FormatarData(fnSPED.ExtrairCampoArray(Campos, dicTitulos, "DT_DOC"))
            DT_ENT_SAI = fnExcel.FormatarData(fnSPED.ExtrairCampoArray(Campos, dicTitulos, "DT_ENT_SAI"))
            IND_APUR = fnSPED.ExtrairCampoArray(Campos, dicTitulos, "IND_APUR")
            
            CHV_E500 = GerarRegistroE500(dicAuxiliarE500, dicTitulosE500, ARQUIVO, CHV_E001, IND_APUR, CFOP, DT_INI, DT_FIN, DT_DOC, DT_ENT_SAI)
            Call GerarRegistroE510(dicAuxiliarE510, dicTitulosE510, ARQUIVO, CHV_E500, CFOP, CST_IPI, VL_CONT_IPI, VL_BC_IPI, VL_IPI)
            
        End If
        
    Next Linha
    
    Call LimparDicionarioE510(dicDadosE510, dicAuxiliarE510, dicTitulosE510)
    Call Util.AtualizarDicionario(dicDadosE510, dicAuxiliarE510)
    
    Call Util.AtualizarDicionario(dicDadosE500, dicAuxiliarE500)
    
    Call GerarRegistroE520(dicDadosE520, dicAuxiliarE520, dicTitulosE520, dicDadosE510, dicTitulosE510, dicDadosE530, dicTitulosE530)
    
    Application.StatusBar = "Atualizando dados de apuração no E500, por favor aguarde..."
    Call Util.LimparDados(regE500, 4, False)
    Call Util.ExportarDadosDicionario(regE500, dicDadosE500)
    
    Application.StatusBar = "Atualizando dados de apuração no E510, por favor aguarde..."
    Call Util.LimparDados(regE510, 4, False)
    Call Util.ExportarDadosDicionario(regE510, dicDadosE510)
    
    Application.StatusBar = "Atualizando dados de apuração no E520, por favor aguarde..."
    Call Util.LimparDados(regE520, 4, False)
    Call Util.ExportarDadosDicionario(regE520, dicDadosE520)
    
    Application.StatusBar = "Atualizando dados de apuração no E530, por favor aguarde..."
    Call Util.LimparDados(regE530, 4, False)
    Call Util.ExportarDadosDicionario(regE530, dicDadosE530)
    
    Application.StatusBar = "Atualização concluída com sucesso!"
    Application.StatusBar = False
    
End Sub

Private Function GerarRegistroE500(ByRef dicDadosE500 As Dictionary, ByRef dicTitulosE500 As Dictionary, _
    ByVal ARQUIVO As String, ByVal CHV_PAI As String, ByVal IND_APUR As String, ByVal CFOP As String, _
    ByVal DT_INI As String, ByVal DT_FIN As String, ByVal DT_EMI As String, ByVal DT_ENT As String) As String

Dim CHV_REG As String, DT_INI_SPED$, DT_FIN_SPED$, IND_APUR_SPED$
Dim Campos As Variant
    
    If IND_APUR Like "1*" Then Call GerarPeriodoDecendial(CFOP, DT_EMI, DT_ENT, DT_INI, DT_FIN)
    
    DT_INI_SPED = VBA.Format(DT_INI, "ddmmyyyy")
    DT_FIN_SPED = VBA.Format(DT_FIN, "ddmmyyyy")
    IND_APUR_SPED = Util.ApenasNumeros(IND_APUR)
    
    CHV_REG = fnSPED.GerarChaveRegistro(CHV_PAI, IND_APUR_SPED, DT_INI_SPED, DT_FIN_SPED)
    Campos = Array("E500", ARQUIVO, CHV_REG, CHV_PAI, "", IND_APUR, DT_INI, DT_FIN)
    
    If Not dicDadosE500.Exists(CHV_REG) Then dicDadosE500(CHV_REG) = Campos
    
    GerarRegistroE500 = CHV_REG
    
End Function

Private Function GerarPeriodoDecendial(ByVal CFOP As Long, ByVal DT_EMI As String, _
    ByVal DT_ENT As String, ByRef DT_INI As String, ByRef DT_FIN As String) As Variant
    
Dim dia As Integer, DECENDIO As Integer
Dim Data As String
    
    ' Determina a data base com base no CFOP
    If CFOP > 4000 Then dia = Day(DT_EMI) Else dia = Day(DT_ENT)
    If CFOP > 4000 Then Data = DT_EMI Else Data = DT_ENT
    
    ' Calcula o decêndio
    DECENDIO = Int((dia - 1) / 10) + 1
    
    ' Define as datas de início e fim do decêndio
    Select Case DECENDIO
        
        Case 1
            DT_INI = Format(Data, "yyyy-mm") & "-01"
            DT_FIN = Format(Data, "yyyy-mm") & "-10"
            
        Case 2
            DT_INI = Format(Data, "yyyy-mm") & "-11"
            DT_FIN = Format(Data, "yyyy-mm") & "-20"
            
        Case 3
            DT_INI = Format(Data, "yyyy-mm") & "-21"
            DT_FIN = Format(Data, "yyyy-mm") & "-" & Right(Format(DateSerial(Year(Data), Month(Data) + 1, 0), "yyyy-mm-dd"), 2)  ' Último dia do mês
            
    End Select
    
    ' Retorna as datas de início e fim do decêndio
    GerarPeriodoDecendial = Array(DT_INI, DT_FIN)
    
End Function

Private Function GerarRegistroE510(ByRef dicDadosE510 As Dictionary, ByRef dicTitulosE510 As Dictionary, _
    ByVal ARQUIVO As String, ByVal CHV_PAI As String, ByVal CFOP As String, ByVal CST_IPI As String, _
    ByVal VL_CONT_IPI As Double, ByVal VL_BC_IPI As Double, ByVal VL_IPI As Double)
    
Dim CHV_REG As String, Chave$, CST_IPI_SPED$
Dim Campos As Variant
    
    CST_IPI_SPED = Util.ApenasNumeros(CST_IPI)
    Chave = VBA.Join(Array(CHV_PAI, CFOP, CST_IPI_SPED))
    CHV_REG = fnSPED.GerarChaveRegistro(CHV_PAI, CFOP, CST_IPI_SPED)
    Campos = Array("E510", ARQUIVO, CHV_REG, CHV_PAI, "", CFOP, CST_IPI, VL_CONT_IPI, VL_BC_IPI, VL_IPI)
    
    Call Util.SomarValoresDicionario(dicDadosE510, Campos, CHV_REG)
    
End Function

Private Function GerarRegistroE520(ByRef dicDadosE520 As Dictionary, ByRef dicAuxiliarE520 As Dictionary, _
    ByRef dicTitulosE520 As Dictionary, ByRef dicDadosE510 As Dictionary, ByRef dicTitulosE510 As Dictionary, _
    ByRef dicDadosE530 As Dictionary, ByRef dicTitulosE530 As Dictionary)
    
Dim Campos As Variant, Chave
Dim arrChavesE520 As New ArrayList
Dim CHV_REG As String, CHV_PAI$, CFOP$
Dim VL_SD_ANT_IPI As Double, VL_DEB_IPI#, VL_CRED_IPI#, VL_OD_IPI#, VL_OC_IPI#, VL_SC_IPI#, VL_SD_IPI#, VL_IPI#
    
    For Each Chave In dicDadosE510.Keys()
        
        CHV_PAI = fnSPED.ExtrairCampoDicionario(dicDadosE510, dicTitulosE510, Chave, "CHV_PAI_FISCAL")
        CHV_REG = fnSPED.GerarChaveRegistro(CHV_PAI, "E520")
        VL_IPI = fnSPED.ExtrairCampoDicionario(dicDadosE510, dicTitulosE510, Chave, "VL_IPI")
        CFOP = fnSPED.ExtrairCampoDicionario(dicDadosE510, dicTitulosE510, Chave, "CFOP")

        Call ContabilizarValoresIPI(CFOP, VL_DEB_IPI, VL_CRED_IPI, VL_IPI)
        Campos = Array("E520", ARQUIVO, CHV_REG, CHV_PAI, "", VL_SD_ANT_IPI, VL_DEB_IPI, VL_CRED_IPI, VL_OD_IPI, VL_OC_IPI, VL_SC_IPI, VL_SD_IPI)
        
        Call Util.SomarValoresDicionario(dicAuxiliarE520, Campos, CHV_REG)
        
    Next Chave
    
    Set arrChavesE520 = Util.ListarValoresUnicos(regE520, 4, 3, "CHV_REG")
    For Each Chave In arrChavesE520
        
        Call AtualizarSaldoCredorIPI(dicDadosE520, dicAuxiliarE520, dicTitulosE520, Chave)
        Call ContabilizarAjustesIPI(dicDadosE530, dicTitulosE530, dicAuxiliarE520, dicTitulosE520, Chave)
        Call AtualizarSaldoIPI(dicAuxiliarE520, dicTitulosE520, Chave)
        
    Next Chave
    
    Set dicDadosE520 = dicAuxiliarE520
    
End Function

Private Function AtualizarSaldoCredorIPI(ByRef dicDadosE520 As Dictionary, ByRef dicAuxiliarE520 As Dictionary, ByRef dicTitulosE520 As Dictionary, ByVal CHV_REG As String) As Double
    
Dim VL_SD_ANT_IPI As Double
    
    If dicDadosE520.Exists(CHV_REG) Then
        VL_SD_ANT_IPI = Util.ExtrairDadoDicionario(dicDadosE520, dicTitulosE520, CHV_REG, "VL_SD_ANT_IPI")
    End If
    
    If dicAuxiliarE520.Exists(CHV_REG) Then
        Call Util.AlterarCampoDicionario(dicAuxiliarE520, dicTitulosE520, CHV_REG, "VL_SD_ANT_IPI", VL_SD_ANT_IPI)
    End If
    
End Function

Private Function ContabilizarValoresIPI(ByVal CFOP As String, ByRef VL_DEB_IPI As Double, _
    ByRef VL_CRED_IPI As Double, ByVal VL_IPI As Double) As Double
    
    If CFOP > 4000 Then VL_DEB_IPI = VL_IPI Else VL_CRED_IPI = VL_IPI

End Function

Private Sub AtualizarSaldoIPI(ByRef dicDadosE520 As Dictionary, ByRef dicTitulosE520 As Dictionary, ByVal CHV_REG As String)

Dim VL_DEB_IPI As Double, VL_OD_IPI#, VL_SD_ANT_IPI#, VL_CRED_IPI#, VL_OC_IPI#, VL_SC_IPI#, VL_SD_IPI#, Saldo#
    
    VL_DEB_IPI = Util.ExtrairDadoDicionario(dicDadosE520, dicTitulosE520, CHV_REG, "VL_DEB_IPI")
    VL_OD_IPI = Util.ExtrairDadoDicionario(dicDadosE520, dicTitulosE520, CHV_REG, "VL_OD_IPI")
    VL_SD_ANT_IPI = Util.ExtrairDadoDicionario(dicDadosE520, dicTitulosE520, CHV_REG, "VL_SD_ANT_IPI")
    VL_CRED_IPI = Util.ExtrairDadoDicionario(dicDadosE520, dicTitulosE520, CHV_REG, "VL_CRED_IPI")
    VL_OC_IPI = Util.ExtrairDadoDicionario(dicDadosE520, dicTitulosE520, CHV_REG, "VL_OC_IPI")
    VL_SC_IPI = Util.ExtrairDadoDicionario(dicDadosE520, dicTitulosE520, CHV_REG, "VL_SC_IPI")
    VL_SD_IPI = Util.ExtrairDadoDicionario(dicDadosE520, dicTitulosE520, CHV_REG, "VL_SD_IPI")
    
    Saldo = (VL_DEB_IPI + VL_OD_IPI) - (VL_SD_ANT_IPI + VL_CRED_IPI + VL_OC_IPI)
    
    If Saldo < 0 Then
        
        VL_SC_IPI = Abs(Saldo)
        VL_SD_IPI = 0
        
    Else
        
        VL_SC_IPI = 0
        VL_SD_IPI = Saldo
        
    End If
    
    Call Util.AlterarCampoDicionario(dicDadosE520, dicTitulosE520, CHV_REG, "VL_SC_IPI", VL_SC_IPI)
    Call Util.AlterarCampoDicionario(dicDadosE520, dicTitulosE520, CHV_REG, "VL_SD_IPI", VL_SD_IPI)
    
End Sub

Private Function ContabilizarAjustesIPI(ByRef dicDadosE530 As Dictionary, ByRef dicTitulosE530 As Dictionary, _
    ByRef dicDadosE520 As Dictionary, ByRef dicTitulosE520 As Dictionary, ByVal CHV_PAI As String) As Double
    
Dim VL_OD_IPI As Double, VL_OC_IPI#, VL_AJ#
Dim CHV_PAI_E530 As String
Dim CamposE530 As Variant
Dim IND_AJ As Byte
    
    For Each CamposE530 In dicDadosE530.Items()
        
        CHV_PAI_E530 = fnSPED.ExtrairCampoArray(CamposE530, dicTitulosE530, "CHV_PAI_FISCAL")
        If CHV_PAI_E530 = CHV_PAI Then
            
            VL_AJ = fnSPED.ExtrairCampoArray(CamposE530, dicTitulosE530, "VL_AJ")
            IND_AJ = fnSPED.ExtrairCampoArray(CamposE530, dicTitulosE530, "IND_AJ")
            Select Case True
                
                Case IND_AJ Like "0*"
                    VL_OD_IPI = VL_OD_IPI + VL_AJ
                    
                Case IND_AJ Like "*"
                    VL_OC_IPI = VL_OC_IPI + VL_AJ
                    
            End Select
            
        End If
        
    Next CamposE530
    
    Call Util.AlterarCampoDicionario(dicDadosE520, dicTitulosE520, CHV_PAI, "VL_OD_IPI", VL_OD_IPI)
    Call Util.AlterarCampoDicionario(dicDadosE520, dicTitulosE520, CHV_PAI, "VL_OC_IPI", VL_OC_IPI)
    
End Function

Private Function LimparDicionarioE510(ByRef dicDadosE510 As Dictionary, ByRef dicAuxiliarE510 As Dictionary, ByRef dicTitulos As Dictionary)

Dim arrCFOPs As New ArrayList
Dim Campos As Variant, Chave
Dim CFOP As String
Dim i As Byte
    
    Set arrCFOPs = Util.ListarValoresUnicosDicionario(dicAuxiliarE510, dicTitulos, "CFOP")
    
    For Each Chave In dicDadosE510.Keys()
        
        Campos = dicDadosE510(Chave)
        If LBound(Campos) = 0 Then i = 1
        
        CFOP = Campos(dicTitulos("CFOP"))
        If arrCFOPs.contains(CFOP) Then Call dicDadosE510.Remove(Chave)
        
    Next Chave
    
End Function
