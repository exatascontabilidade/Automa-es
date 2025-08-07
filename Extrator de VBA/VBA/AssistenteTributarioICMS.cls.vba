Attribute VB_Name = "AssistenteTributarioICMS"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private dicTitulos As New Dictionary
Private Const RegistrosIgnorados As String = "C190"

Public Function AceitarSugestoes()

Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim arrDados As New ArrayList
Dim Campos As Variant
    
    Inicio = Now()
    Application.StatusBar = "Implementando sugestões selecionadas, por favor aguarde..."
    
    Set dicTitulos = Util.MapearTitulos(assTributacaoICMS, 3)
    Set Dados = assTributacaoICMS.Range("A4").CurrentRegion
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
                    
                    Case "Informar alíquota de 1,65% para o PIS"
                        Campos(dicTitulos("ALIQ_PIS")) = 0.0165
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Call Assistente.Tributario.PIS_COFINS.VerificarInconsistenciasCadastrais(Campos, dicTitulos)
                        GoTo VERIFICAR:
                        
                    Case "Informar alíquota de 0,65% para o PIS"
                        Campos(dicTitulos("ALIQ_PIS")) = 0.0065
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Call Assistente.Tributario.PIS_COFINS.VerificarInconsistenciasCadastrais(Campos, dicTitulos)
                        GoTo VERIFICAR:
                        
                    Case "Informar alíquota de 7,60% para a COFINS"
                        Campos(dicTitulos("ALIQ_COFINS")) = 0.076
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Call Assistente.Tributario.PIS_COFINS.VerificarInconsistenciasCadastrais(Campos, dicTitulos)
                        
                    Case "Informar alíquota de 3,00% para a COFINS"
                        Campos(dicTitulos("ALIQ_COFINS")) = 0.03
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Call Assistente.Tributario.PIS_COFINS.VerificarInconsistenciasCadastrais(Campos, dicTitulos)
                        
                    Case "Zerar alíquota do PIS"
                        Campos(dicTitulos("ALIQ_PIS")) = 0
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Call Assistente.Tributario.PIS_COFINS.VerificarInconsistenciasCadastrais(Campos, dicTitulos)
                        GoTo VERIFICAR:
                        
                    Case "Zerar alíquota da COFINS"
                        Campos(dicTitulos("ALIQ_COFINS")) = 0
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Call Assistente.Tributario.PIS_COFINS.VerificarInconsistenciasCadastrais(Campos, dicTitulos)
                        
                    Case "Alterar CST_PIS para 49", "Informar CST_PIS 49 - Outras Operações de Saída"
                        Campos(dicTitulos("CST_PIS")) = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(49)
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Call Assistente.Tributario.PIS_COFINS.VerificarInconsistenciasCadastrais(Campos, dicTitulos)
                        GoTo VERIFICAR:
                        
                    Case "Informar CST_PIS igual a 70 - Operação de Aquisição sem Direito a Crédito"
                        Campos(dicTitulos("CST_PIS")) = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(70)
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Call Assistente.Tributario.PIS_COFINS.VerificarInconsistenciasCadastrais(Campos, dicTitulos)
                        GoTo VERIFICAR:
                        
                    Case "Informar CST_COFINS igual a 70 - Operação de Aquisição sem Direito a Crédito"
                        Campos(dicTitulos("CST_COFINS")) = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(70)
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Call Assistente.Tributario.PIS_COFINS.VerificarInconsistenciasCadastrais(Campos, dicTitulos)
                        
                    Case "Informar CST_PIS 98 - Outras Operações de Entrada"
                        Campos(dicTitulos("CST_PIS")) = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(98)
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Call Assistente.Tributario.PIS_COFINS.VerificarInconsistenciasCadastrais(Campos, dicTitulos)
                        GoTo VERIFICAR:
                        
                    Case "Alterar CST_COFINS para 49", "Informar CST_COFINS 49 - Outras Operações de Saída"
                        Campos(dicTitulos("CST_COFINS")) = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(49)
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Call Assistente.Tributario.PIS_COFINS.VerificarInconsistenciasCadastrais(Campos, dicTitulos)
                        
                    Case "Informar CST_COFINS 98 - Outras Operações de Entrada"
                        Campos(dicTitulos("CST_COFINS")) = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(98)
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Call Assistente.Tributario.PIS_COFINS.VerificarInconsistenciasCadastrais(Campos, dicTitulos)
                        
                    Case "Alterar o valor do campo TIPO_ITEM para 00"
                        Campos(dicTitulos("TIPO_ITEM")) = "00 - Mercadoria para Revenda"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Call Assistente.Tributario.PIS_COFINS.VerificarInconsistenciasCadastrais(Campos, dicTitulos)
                        
                End Select
                
            End If
            
            If Linha.Row > 3 Then arrDados.Add Campos
            
        End If
        
    Next Linha
    
    If assTributacaoICMS.AutoFilterMode Then assTributacaoICMS.AutoFilter.ShowAllData
    Call Util.LimparDados(assTributacaoICMS, 4, False)
    Call Util.ExportarDadosArrayList(assTributacaoICMS, arrDados)
    
    Call FuncoesFormatacao.AplicarFormatacao(ActiveSheet)
    Call FuncoesFormatacao.DestacarInconsistencias(assTributacaoICMS)
    
    Call Util.MsgInformativa("Sugestões aplicadas com sucesso!", "Assistente de Tributação", Inicio)
    Application.StatusBar = False
    
End Function

Public Function ImportarTributacaoICMS()
        
    Call impTributario.ImportarTributacao(assTributacaoICMS)
    
End Function

Public Sub VerificarTributacaoICMS()

Dim ChaveTrib As String, REG$, CFOP_CORRETO$, CFOP$, DESCR_ITEM$, NCM$, DT_REF$, Msg$
Dim dicEstruturaTributaria As New Dictionary
Dim Tributacao As New AssistenteTributario
Dim dicDadosTributarios As New Dictionary
Dim arrDadosApuracao As New ArrayList
Dim arrRelatorio As New ArrayList
Dim Campos As Variant, CamposTrib
    
    Inicio = Now()
    
    Set arrDadosApuracao = Util.CriarArrayListRegistro(assApuracaoICMS)
    If arrDadosApuracao.Count = 0 Then
    
        Msg = "Precisa haver dados no assistente de apuração para usar esse recurso."
        Call Util.MsgAlerta(Msg, "Assistente Tributário de ICMS")
    
    End If
    
    With Tributacao
        
        Set .dicDadosTributarios = .CarregarTributacoesSalvas(assTributacaoICMS)
        
        If .dicDadosTributarios.Count = 0 Then
            
            Msg = "Não existem dados tributários de ICMS cadastrados para realizar a análise." & vbCrLf & vbCrLf
            Msg = Msg & "Por favor realize a importação do cadastro de tributação do ICMS para utilizar esse recurso."
            
            Call Util.MsgAlerta(Msg, "Assistente Tributário de ICMS")
            Exit Sub
            
        End If
        
        Set .dicTitulosTributacao = Util.MapearTitulos(assTributacaoICMS, 3)
        Set .dicTitulosApuracao = Util.MapearTitulos(assApuracaoICMS, 3)
        Call .CarregarEstruturaTributaria(assTributacaoICMS)
        
        For Each Campos In arrDadosApuracao
            
            REG = Campos(.dicTitulosApuracao("REG"))
            DT_REF = .ExtrairDataReferencia(Campos)
            
            'Apaga registro de inconsistências e sugestões
            Campos(.dicTitulosApuracao("INCONSISTENCIA")) = Empty
            Campos(.dicTitulosApuracao("SUGESTAO")) = Empty
            
            If RegistrosIgnorados Like "*" & REG & "*" Then GoTo Prx:
            
            ChaveTrib = .GerarChaveTributacao(assApuracaoICMS, Campos)
            If .dicDadosTributarios.Exists(ChaveTrib) Then
                
                CamposTrib = .ExtrairCamposTributarios(ChaveTrib, DT_REF)
                If Not VBA.IsEmpty(CamposTrib) Then Call ValidarRegrasTributariasICMS(Campos, .dicTitulosApuracao, CamposTrib, .dicTitulosTributacao)
                
            Else
                
                CFOP_CORRETO = IdentificarOperacoesIncorretas(Tributacao, assApuracaoICMS, Campos)
                
                If CFOP_CORRETO <> "" Then
                    
                    CFOP = Campos(.dicTitulosApuracao("CFOP"))
                    NCM = Campos(.dicTitulosApuracao("COD_NCM"))
                    DESCR_ITEM = Campos(.dicTitulosApuracao("DESCR_ITEM"))
                    
                    Campos(.dicTitulosApuracao("INCONSISTENCIA")) = "CFOP (" & CFOP & ") incorreto para tributação do item: " & DESCR_ITEM & " (NCM: " & NCM & ")"
                    Campos(.dicTitulosApuracao("SUGESTAO")) = "Aplicar CFOP " & CFOP_CORRETO & " para a operação."
                    
                Else
                    
                    Campos(.dicTitulosApuracao("INCONSISTENCIA")) = "Operação não cadastrada no Assistente de Tributação do ICMS"
                    Campos(.dicTitulosApuracao("SUGESTAO")) = "Cadastrar Operação no Assistente de tributação do ICMS"
                    
                End If
                
            End If
Prx:
            arrRelatorio.Add Campos
            
        Next Campos
        
    End With
    
    If arrRelatorio.Count > 0 Then
        
        Application.StatusBar = "Exportando resultado da análise!"
        Call Util.LimparDados(assApuracaoICMS, 4, False)
        
        Call Util.ExportarDadosArrayList(assApuracaoICMS, arrRelatorio)
        Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoICMS)
        
        Application.StatusBar = "Verificação concluída com sucesso!"
        Call Util.MsgInformativa("Verificação concluída com sucesso!", "Assistente de Tributação ICMS", Inicio)
        
    Else
        
        Msg = "Não foi encontrado nenhum dado para geração do relatório." & vbCrLf & vbCrLf
        Msg = Msg & "Por favor verifique se o SPED e/ou XMLs foram importados e tente novamente."
        Call Util.MsgAlerta(Msg, "Assistente de Tributação ICMS")
        
    End If
        
    Application.StatusBar = False
    
End Sub

Public Sub AplicarTributacaoICMS()

Dim REG As String, CFOP$, SUGESTAO$, DT_REF$
Dim CamposTrib As Variant, Campos, ChaveTrib
Dim Tributacao As New AssistenteTributario
Dim Apuracao As New clsAssistenteApuracao
Dim Dados As Range, Linha As Range
Dim arrDados As New ArrayList
    
    Inicio = Now()
    Application.StatusBar = "Aplicando tributações selecionadas, por favor aguarde."
    
    With Tributacao
        
        Set .dicTitulosApuracao = Util.MapearTitulos(assApuracaoICMS, 3)
        Set .dicTitulosTributacao = Util.MapearTitulos(assTributacaoICMS, 3)
        Set .dicDadosTributarios = .CarregarTributacoesSalvas(assTributacaoICMS)
        Set .dicTitulos = Util.MapearTitulos(assApuracaoICMS, 3)
        
        Set Dados = assApuracaoICMS.Range("A4").CurrentRegion
        If Dados Is Nothing Then
            
            Call Util.MsgAlerta("Não foi encontrada nenhuma inconsistência tributária!", "Inconsistências Tributárias")
            Exit Sub
            
        End If
        
        For Each Linha In Dados.Rows
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                If Linha.EntireRow.Hidden = False And Campos(.dicTitulosApuracao("SUGESTAO")) <> "" And Linha.Row > 3 Then
                    
                    REG = Util.RemoverAspaSimples(Campos(.dicTitulosApuracao("REG")))
                    If RegistrosIgnorados Like "*" & REG & "*" Then GoTo Prx:
                    
                    ARQUIVO = (Campos(.dicTitulosApuracao("ARQUIVO")))
                    DT_REF = .ExtrairDataReferencia(Campos)
                    
                    ChaveTrib = .GerarChaveTributacao(assApuracaoICMS, Campos)
                    CamposTrib = .ExtrairCamposTributarios(ChaveTrib, DT_REF)
                    
Reprocessar:
                    SUGESTAO = Campos(.dicTitulosApuracao("SUGESTAO"))
                    Select Case True
                        
                        Case SUGESTAO = "Aplicar o TIPO_ITEM cadastrado na Tributação"
                            Campos(.dicTitulosApuracao("TIPO_ITEM")) = CamposTrib(.dicTitulosTributacao("TIPO_ITEM"))
                            Call .LimparInconsistenciasSugestoes(Campos)
                            Call ValidarRegrasTributariasICMS(Campos, .dicTitulosApuracao, CamposTrib, .dicTitulosTributacao)
                            
                        Case SUGESTAO = "Aplicar o COD_BARRA cadastrado na Tributação"
                            Campos(.dicTitulosApuracao("COD_BARRA")) = CamposTrib(.dicTitulosTributacao("COD_BARRA"))
                            Call .LimparInconsistenciasSugestoes(Campos)
                            Call ValidarRegrasTributariasICMS(Campos, .dicTitulosApuracao, CamposTrib, .dicTitulosTributacao)
                            
                        Case SUGESTAO = "Aplicar o NCM cadastrado na Tributação"
                            Campos(.dicTitulosApuracao("COD_NCM")) = CamposTrib(.dicTitulosTributacao("COD_NCM"))
                            Call .LimparInconsistenciasSugestoes(Campos)
                            Call ValidarRegrasTributariasICMS(Campos, .dicTitulosApuracao, CamposTrib, .dicTitulosTributacao)
                            
                        Case SUGESTAO = "Aplicar a EX_IPI cadastrada na Tributação"
                            Campos(.dicTitulosApuracao("EX_IPI")) = CamposTrib(.dicTitulosTributacao("EX_IPI"))
                            Call .LimparInconsistenciasSugestoes(Campos)
                            Call ValidarRegrasTributariasICMS(Campos, .dicTitulosApuracao, CamposTrib, .dicTitulosTributacao)
                            
                        Case SUGESTAO = "Aplicar o CEST cadastrado na Tributação"
                            Campos(.dicTitulosApuracao("CEST")) = CamposTrib(.dicTitulosTributacao("CEST"))
                            Call .LimparInconsistenciasSugestoes(Campos)
                            Call ValidarRegrasTributariasICMS(Campos, .dicTitulosApuracao, CamposTrib, .dicTitulosTributacao)
                            
                        Case SUGESTAO = "Aplicar a alíquota do ICMS-ST cadastrada na Tributação"
                            Campos(.dicTitulosApuracao("ALIQ_ST")) = CamposTrib(.dicTitulosTributacao("ALIQ_ST"))
                            Call .LimparInconsistenciasSugestoes(Campos)
                            Call ValidarRegrasTributariasICMS(Campos, .dicTitulosApuracao, CamposTrib, .dicTitulosTributacao)
                            GoTo Reprocessar:
                            
                        Case SUGESTAO = "Aplicar a alíquota do ICMS cadastrada na Tributação"
                            Campos(.dicTitulosApuracao("ALIQ_ICMS")) = CamposTrib(.dicTitulosTributacao("ALIQ_ICMS"))
                            Call .LimparInconsistenciasSugestoes(Campos)
                            Call ValidarRegrasTributariasICMS(Campos, .dicTitulosApuracao, CamposTrib, .dicTitulosTributacao)
                            GoTo Reprocessar:
                            
                        Case SUGESTAO = "Aplicar o CST_ICMS cadastrado na Tributação"
                            Campos(.dicTitulosApuracao("CST_ICMS")) = CamposTrib(.dicTitulosTributacao("CST_ICMS"))
                            Call .LimparInconsistenciasSugestoes(Campos)
                            Call ValidarRegrasTributariasICMS(Campos, .dicTitulosApuracao, CamposTrib, .dicTitulosTributacao)
                            GoTo Reprocessar:
                            
                        Case SUGESTAO Like "Aplicar CFOP * para a operação."
                            CFOP = Util.ApenasNumeros(SUGESTAO)
                            Campos(.dicTitulosApuracao("CFOP")) = CFOP
                            Call .LimparInconsistenciasSugestoes(Campos)
                            ChaveTrib = .GerarChaveTributacao(assTributacaoICMS, Campos, True)
                            DT_REF = .ExtrairDataReferencia(Campos)
                            CamposTrib = .ExtrairCamposTributarios(ChaveTrib, DT_REF)
                            
                            Call ValidarRegrasTributariasICMS(Campos, .dicTitulosApuracao, CamposTrib, .dicTitulosTributacao)
                            
                        Case SUGESTAO Like "Aplicar o indicador de movimento cadastrado na Tributação"
                            Campos(.dicTitulosApuracao("IND_MOV")) = CamposTrib(.dicTitulosTributacao("IND_MOV"))
                            Call .LimparInconsistenciasSugestoes(Campos)
                            Call ValidarRegrasTributariasICMS(Campos, .dicTitulosApuracao, CamposTrib, .dicTitulosTributacao)
                            
                    End Select
                
                End If
Prx:
                If Linha.Row > 3 Then arrDados.Add Campos
                
            End If
            
        Next Linha
        
    End With
    
    If assApuracaoICMS.AutoFilterMode Then assApuracaoICMS.AutoFilter.ShowAllData
    Call Util.LimparDados(assApuracaoICMS, 4, False)
    Call Util.ExportarDadosArrayList(assApuracaoICMS, arrDados)
    
    Call FuncoesFormatacao.AplicarFormatacao(ActiveSheet)
    Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoICMS)
    
    Call Util.MsgInformativa("Tributações aplicadas com sucesso!", "Aplicação de Tributações PIS/COFINS", Inicio)
    Application.StatusBar = False
    
End Sub

Private Function ValidarRegrasTributariasICMS(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
    If Campos(dicTitulosApuracao("INCONSISTENCIA")) = "" Then _
        Call RegrasTributarias.ValidarCampo_COD_NCM(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
        
    If Campos(dicTitulosApuracao("INCONSISTENCIA")) = "" Then _
        Call RegrasTributarias.ValidarCampo_CEST(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
        
    If Campos(dicTitulosApuracao("INCONSISTENCIA")) = "" Then _
        Call RegrasTributarias.ValidarCampo_COD_BARRA(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
        
    If Campos(dicTitulosApuracao("INCONSISTENCIA")) = "" Then _
        Call RegrasTributarias.ValidarCampo_TIPO_ITEM(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
        
    If Campos(dicTitulosApuracao("INCONSISTENCIA")) = "" Then _
        Call RegrasTributarias.ValidarCampo_EX_IPI(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
        
    If Campos(dicTitulosApuracao("INCONSISTENCIA")) = "" Then _
        Call RegrasTributarias.ValidarCampo_IND_MOV(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
        
    If Campos(dicTitulosApuracao("INCONSISTENCIA")) = "" Then _
        Call RegrasTributarias.ICMS.ValidarCampo_CST_ICMS(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
        
    If Campos(dicTitulosApuracao("INCONSISTENCIA")) = "" Then _
        Call RegrasTributarias.ICMS.ValidarCampo_ALIQ_ICMS(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
        
    If Campos(dicTitulosApuracao("INCONSISTENCIA")) = "" Then _
        Call RegrasTributarias.ICMS.ValidarCampo_ALIQ_ST(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
        
End Function

Public Function VerificarInconsistenciasCadastrais(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
    
    If Campos(dicTitulos("INCONSISTENCIA")) = "" Then Call RegrasCadastrais.VerificarCFOP(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA")) = "" Then Call RegrasCadastrais.ValidarCST_ICMS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA")) = "" Then Call RegrasCadastrais.ICMS.ValidarCampo_CFOP(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA")) = "" Then Call RegrasCadastrais.ICMS.ValidarCampo_CST_ICMS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA")) = "" Then Call RegrasCadastrais.ICMS.ValidarCampo_ALIQ_ICMS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA")) = "" Then Call RegrasCadastrais.ICMS.ValidarCampo_ALIQ_ST(Campos, dicTitulos)
    
End Function

Private Function IdentificarOperacoesIncorretas(ByRef Tributacao As AssistenteTributario, _
    ByRef Plan As Worksheet, ByRef CamposApuracao As Variant) As String

Dim CamposChave As Variant, Chave, Campos, Valor
Dim CFOP As String, CFOPS_INCORRETOS$
Dim dicAtual As Dictionary
    
    With Tributacao
        
        Set dicAtual = .dicEstruturaTributaria
        CamposChave = Tributacao.ObterNomesCamposChave(Plan)
        
        'Navega pela estrutura do dicionário usando o array CamposChave
        For Each Chave In CamposChave
            
            If Chave = "CFOP" Then
                
                IdentificarOperacoesIncorretas = IdentificarCFOPCorreto(Tributacao, dicAtual, CamposApuracao)
                Exit Function
                
            End If
            
            Valor = CamposApuracao(.dicTitulosApuracao(Chave))
            If Not dicAtual.Exists(Valor) Then Exit Function
            Set dicAtual = dicAtual(Valor)
            
        Next Chave
        
    End With
    
End Function

Private Function IdentificarCFOPCorreto(ByRef Tributacao As AssistenteTributario, _
    ByRef dicOperacoes As Dictionary, ByRef CamposApuracao As Variant)

Dim Operacao As Variant, Campos
Dim CFOP_APURACAO As String, CFOP_CORRETO$, CFOPS_INCORRETOS$
    
    With Tributacao
        
        'Itera pelas operações no último nível do dicionário
        For Each Operacao In dicOperacoes.Keys
            
            If Not IsEmpty(Operacao) Then
            
                Campos = dicOperacoes(Operacao)
                CFOP_APURACAO = CamposApuracao(.dicTitulosApuracao("CFOP"))
                CFOP_CORRETO = Campos(.dicTitulosTributacao("CFOP"))
                CFOPS_INCORRETOS = Campos(.dicTitulosTributacao("CFOPS_INCORRETOS"))
                
                If CFOPS_INCORRETOS Like "*" & CFOP_APURACAO & "*" Then
                  
                  IdentificarCFOPCorreto = CFOP_CORRETO
                  Exit Function
                  
                End If
            
            End If
            
        Next Operacao
        
    End With
    
End Function

Private Function CarregarEstruturaTributariaICMS(ByRef dicDados As Dictionary)

Dim CNPJ_ESTABELECIMENTO As String, COD_ITEM$, CFOP$
Dim arrCamposChave As New ArrayList
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim Campos As Variant
    
    With assTributacaoICMS
        
        If .AutoFilterMode Then .AutoFilter.ShowAllData
        Set Dados = Util.DefinirIntervalo(assTributacaoICMS, 4, 3)
        
        If Dados Is Nothing Then Exit Function
        Set dicTitulos = Util.MapearTitulos(assTributacaoICMS, 3)
        
        For Each Linha In Dados.Rows
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                                
                CNPJ_ESTABELECIMENTO = Util.FormatarCNPJ(Campos(dicTitulos("CNPJ")))
                COD_ITEM = Campos(dicTitulos("COD_ITEM"))
                CFOP = Campos(dicTitulos("CFOP"))
                                    
                'Cria estrutura de Dicionários
                If Not dicDados.Exists(CNPJ_ESTABELECIMENTO) Then Set dicDados(CNPJ_ESTABELECIMENTO) = New Dictionary
                If Not dicDados(CNPJ_ESTABELECIMENTO).Exists(COD_ITEM) Then Set dicDados(CNPJ_ESTABELECIMENTO)(COD_ITEM) = New Dictionary
                If Not dicDados(CNPJ_ESTABELECIMENTO)(COD_ITEM).Exists(CFOP) Then Set dicDados(CNPJ_ESTABELECIMENTO)(COD_ITEM)(CFOP) = New Dictionary
                
                'Armazenas informações tributárias na estrutura de dicionários
                dicDados(CNPJ_ESTABELECIMENTO)(COD_ITEM)(CFOP) = Campos
                
             End If
             
        Next Linha
        
    End With
    
End Function

Public Function ReprocessarSugestoes()

Dim Dados As Range, Linha As Range
Dim arrRelatorio As New ArrayList
Dim dicTitulos As New Dictionary
Dim COD_INC_TRIB As String
Dim Campos As Variant
        
    Set dicTitulos = Util.MapearTitulos(assTributacaoICMS, 3)
    If assTributacaoICMS.AutoFilterMode Then assTributacaoICMS.AutoFilter.ShowAllData
    
    Set Dados = Util.DefinirIntervalo(assTributacaoICMS, 4, 3)
    If Dados Is Nothing Then Exit Function
    
    a = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 100, "Reprocessando inconsistências, por favor aguarde...", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
                        
            Campos(dicTitulos("INCONSISTENCIA")) = Empty
            Campos(dicTitulos("SUGESTAO")) = Empty
            Call VerificarInconsistenciasCadastrais(Campos, dicTitulos)
            arrRelatorio.Add Campos
            
        End If
        
    Next Linha
    
    Application.StatusBar = "Atualizando relatório de inconsistências, isso pode levar alguns segundos! Por favor aguarde..."
    Call Util.LimparDados(assTributacaoICMS, 4, False)
    Call Util.ExportarDadosArrayList(assTributacaoICMS, arrRelatorio)
    Call FuncoesFormatacao.DestacarInconsistencias(assTributacaoICMS)
        
    Call Util.AtualizarBarraStatus("Processamento Concluído!")
    
End Function

Public Sub SalvarTributacaoICMS()

Dim Tributacao As New AssistenteTributario
Dim dicDadosTributarios As New Dictionary
Dim dicTitulosApuracao As New Dictionary
Dim Dados As Range, Linha As Range
Dim REG As String, Msg$
Dim Campos As Variant
Dim Comeco As Double
Dim b As Long
    
    Inicio = Now()
    
    Set dicTitulosApuracao = Util.MapearTitulos(assApuracaoICMS, 3)
    Set Dados = Util.DefinirIntervalo(assApuracaoICMS, 4, 3)
    
    If Dados Is Nothing Then
        
        Msg = "Sem dados a processar!" & vbCrLf & vbCrLf
        Msg = Msg & "O relatório precisa de dados para esa função funcionar."
        
        Call Util.MsgAlerta(Msg, "Assistente Tributário do PIS/COFINS")
        Exit Sub
        
    End If
    
    b = 0
    Comeco = Timer
    With Tributacao
        
        Set .dicTitulos = Util.MapearTitulos(assTributacaoICMS, 3)
        Set .dicDadosTributarios = .CarregarTributacoesSalvas(assTributacaoICMS)
        
        For Each Linha In Dados.Rows
            
            .RedimensionarArray (.dicTitulos.Count)
            Call Util.AntiTravamento(b, 100, Msg & "Carregando dados da apuração", Dados.Rows.Count, Comeco)
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                REG = Campos(dicTitulosApuracao("REG"))
                If RegistrosIgnorados Like "*" & REG & "*" Then GoTo Prx:
                
                'Atribui valores aos campos do relatório
                .AtribuirValor "CNPJ_ESTABELECIMENTO", fnExcel.FormatarTexto(Campos(dicTitulosApuracao("CNPJ_ESTABELECIMENTO")))
                .AtribuirValor "UF_CONTRIB", Campos(dicTitulosApuracao("UF_CONTRIB"))
                .AtribuirValor "TIPO_PART", Campos(dicTitulosApuracao("TIPO_PART"))
                .AtribuirValor "CONTRIBUINTE", Campos(dicTitulosApuracao("CONTRIBUINTE"))
                .AtribuirValor "COD_ITEM", fnExcel.FormatarTexto(Campos(dicTitulosApuracao("COD_ITEM")))
                .AtribuirValor "DESCR_ITEM", Campos(dicTitulosApuracao("DESCR_ITEM"))
                .AtribuirValor "TIPO_ITEM", Campos(dicTitulosApuracao("TIPO_ITEM"))
                .AtribuirValor "COD_BARRA", fnExcel.FormatarTexto(Campos(dicTitulosApuracao("COD_BARRA")))
                .AtribuirValor "COD_NCM", fnExcel.FormatarTexto(Campos(dicTitulosApuracao("COD_NCM")))
                .AtribuirValor "EX_IPI", fnExcel.FormatarTexto(Campos(dicTitulosApuracao("EX_IPI")))
                .AtribuirValor "CEST", fnExcel.FormatarTexto(Campos(dicTitulosApuracao("CEST")))
                .AtribuirValor "IND_MOV", Campos(dicTitulosApuracao("IND_MOV"))
                .AtribuirValor "UF_PART", Campos(dicTitulosApuracao("UF_PART"))
                .AtribuirValor "CFOP", Campos(dicTitulosApuracao("CFOP"))
                .AtribuirValor "CST_ICMS", fnExcel.FormatarTexto(Campos(dicTitulosApuracao("CST_ICMS")))
                .AtribuirValor "ALIQ_ICMS", Campos(dicTitulosApuracao("ALIQ_ICMS"))
                .AtribuirValor "ALIQ_ST", Campos(dicTitulosApuracao("ALIQ_ST"))
                
            End If
            
            Call .RegistrarTributacao(assTributacaoICMS, REG)
Prx:
        Next Linha
    
        Application.StatusBar = Msg & "Atualizando dados tributários..."
        Call Util.LimparDados(assTributacaoICMS, 4, False)
        
        Call .PrepararTributacoesParaExportacao
        Call Util.ExportarDadosArrayList(assTributacaoICMS, .arrTributacoes)
        
        Call Util.MsgInformativa("Tributação do ICMS atualizada com sucesso!", "Assistente de Tributação do ICMS", Inicio)
    
    End With
        
End Sub
