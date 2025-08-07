Attribute VB_Name = "AssistenteTributarioIPI"
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
    
    Set dicTitulos = Util.MapearTitulos(assTributacaoIPI, 3)
    Set Dados = assTributacaoIPI.Range("A4").CurrentRegion
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
    
    If assTributacaoIPI.AutoFilterMode Then assTributacaoIPI.AutoFilter.ShowAllData
    Call Util.LimparDados(assTributacaoIPI, 4, False)
    Call Util.ExportarDadosArrayList(assTributacaoIPI, arrDados)
    
    Call FuncoesFormatacao.AplicarFormatacao(ActiveSheet)
    Call FuncoesFormatacao.DestacarInconsistencias(assTributacaoIPI)
    
    Call Util.MsgInformativa("Sugestões aplicadas com sucesso!", "Assistente de Tributação", Inicio)
    Application.StatusBar = False
    
End Function

Public Function ImportarTributacaoIPI()

    Call impTributario.ImportarTributacao(assTributacaoIPI)
    
End Function

Public Sub VerificarTributacaoIPI()

Dim ChaveTrib As String, REG$, CFOP_CORRETO$, CFOP$, DESCR_ITEM$, NCM$, DT_REF$, Msg$
Dim dicEstruturaTributaria As New Dictionary
Dim Tributacao As New AssistenteTributario
Dim dicDadosTributarios As New Dictionary
Dim arrDadosApuracao As New ArrayList
Dim arrRelatorio As New ArrayList
Dim Campos As Variant, CamposTrib
    
    Inicio = Now()
    
    Set arrDadosApuracao = Util.CriarArrayListRegistro(assApuracaoIPI)
    If arrDadosApuracao.Count = 0 Then
    
        Msg = "Precisa haver dados no assistente de apuração para usar esse recurso."
        Call Util.MsgAlerta(Msg, "Assistente Tributário de IPI")
    
    End If
    
    With Tributacao
        
        Set .dicDadosTributarios = .CarregarTributacoesSalvas(assTributacaoIPI)
        
        If .dicDadosTributarios.Count = 0 Then
            
            Msg = "Não existem dados tributários de IPI cadastrados para realizar a análise." & vbCrLf & vbCrLf
            Msg = Msg & "Por favor realize a importação do cadastro de tributação do IPI para utilizar esse recurso."
            
            Call Util.MsgAlerta(Msg, "Assistente Tributário de IPI")
            Exit Sub
            
        End If
        
        Set .dicTitulosTributacao = Util.MapearTitulos(assTributacaoIPI, 3)
        Set .dicTitulosApuracao = Util.MapearTitulos(assApuracaoIPI, 3)
        Call .CarregarEstruturaTributaria(assTributacaoIPI)
        
        For Each Campos In arrDadosApuracao
            
            REG = Campos(.dicTitulosApuracao("REG"))
            DT_REF = .ExtrairDataReferencia(Campos)
            
            'Apaga rgistro de inconsistências e sugestões
            Campos(.dicTitulosApuracao("INCONSISTENCIA")) = Empty
            Campos(.dicTitulosApuracao("SUGESTAO")) = Empty
            
            If RegistrosIgnorados Like "*" & REG & "*" Then GoTo Prx:
                
            ChaveTrib = .GerarChaveTributacao(assApuracaoIPI, Campos)
            If .dicDadosTributarios.Exists(ChaveTrib) Then
                
                CamposTrib = .ExtrairCamposTributarios(ChaveTrib, DT_REF)
                Call ValidarRegrasTributariasIPI(Campos, .dicTitulosApuracao, CamposTrib, .dicTitulosTributacao)
                
            Else
                
                CFOP_CORRETO = IdentificarOperacoesIncorretas(Tributacao, assApuracaoIPI, Campos)
                
                If CFOP_CORRETO <> "" Then
                    
                    CFOP = Campos(.dicTitulosApuracao("CFOP"))
                    NCM = Campos(.dicTitulosApuracao("COD_NCM"))
                    DESCR_ITEM = Campos(.dicTitulosApuracao("DESCR_ITEM"))
                    
                    Campos(.dicTitulosApuracao("INCONSISTENCIA")) = "CFOP (" & CFOP & ") incorreto para tributação do item: " & DESCR_ITEM & " (NCM: " & NCM & ")"
                    Campos(.dicTitulosApuracao("SUGESTAO")) = "Aplicar CFOP " & CFOP_CORRETO & " para a operação."
                    
                Else
                    
                    Campos(.dicTitulosApuracao("INCONSISTENCIA")) = "Operação não cadastrada no Assistente de Tributação do IPI"
                    Campos(.dicTitulosApuracao("SUGESTAO")) = "Cadastrar Operação no Assistente de tributação do IPI"
                    
                End If
                
            End If
Prx:
            arrRelatorio.Add Campos
            
        Next Campos
        
    End With
    
    If arrRelatorio.Count > 0 Then
        
        Application.StatusBar = "Exportando resultado da análise!"
        Call Util.LimparDados(assApuracaoIPI, 4, False)
        
        Call Util.ExportarDadosArrayList(assApuracaoIPI, arrRelatorio)
        Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoIPI)
        
        Application.StatusBar = "Verificação concluída com sucesso!"
        Call Util.MsgInformativa("Verificação concluída com sucesso!", "Assistente de Tributação IPI", Inicio)
        
    Else
        
        Msg = "Não foi encontrado nenhum dado para geração do relatório." & vbCrLf & vbCrLf
        Msg = Msg & "Por favor verifique se o SPED e/ou XMLs foram importados e tente novamente."
        Call Util.MsgAlerta(Msg, "Assistente de Tributação IPI")
        
    End If
        
    Application.StatusBar = False
    
End Sub

Public Sub SalvarTributacaoIPI()

Dim Tributacao As New AssistenteTributario
Dim dicDadosTributarios As New Dictionary
Dim dicTitulosApuracao As New Dictionary
Dim Dados As Range, Linha As Range
Dim REG As String, Msg$
Dim Campos As Variant
Dim Comeco As Double
Dim b As Long
    
    Inicio = Now()
    
    Set dicTitulosApuracao = Util.MapearTitulos(assApuracaoIPI, 3)
    Set Dados = Util.DefinirIntervalo(assApuracaoIPI, 4, 3)
    
    If Dados Is Nothing Then
        
        Msg = "Sem dados a processar!" & vbCrLf & vbCrLf
        Msg = Msg & "O relatório precisa de dados para esa função funcionar."
        
        Call Util.MsgAlerta(Msg, "Assistente Tributário do PIS/COFINS")
        Exit Sub
        
    End If
    
    b = 0
    Comeco = Timer
    With Tributacao
        
        Set .dicTitulos = Util.MapearTitulos(assTributacaoIPI, 3)
        Set .dicDadosTributarios = .CarregarTributacoesSalvas(assTributacaoIPI)
        
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
                .AtribuirValor "COD_ITEM", fnExcel.FormatarTexto(Campos(dicTitulosApuracao("COD_ITEM")))
                .AtribuirValor "DESCR_ITEM", Campos(dicTitulosApuracao("DESCR_ITEM"))
                .AtribuirValor "TIPO_ITEM", Campos(dicTitulosApuracao("TIPO_ITEM"))
                .AtribuirValor "COD_BARRA", fnExcel.FormatarTexto(Campos(dicTitulosApuracao("COD_BARRA")))
                .AtribuirValor "COD_NCM", fnExcel.FormatarTexto(Campos(dicTitulosApuracao("COD_NCM")))
                .AtribuirValor "EX_IPI", fnExcel.FormatarTexto(Campos(dicTitulosApuracao("EX_IPI")))
                .AtribuirValor "IND_MOV", Campos(dicTitulosApuracao("IND_MOV"))
                .AtribuirValor "UF_PART", Campos(dicTitulosApuracao("UF_PART"))
                .AtribuirValor "CFOP", Campos(dicTitulosApuracao("CFOP"))
                .AtribuirValor "IND_APUR", Campos(dicTitulosApuracao("IND_APUR"))
                .AtribuirValor "COD_ENQ", fnExcel.FormatarTexto(Campos(dicTitulosApuracao("COD_ENQ")))
                .AtribuirValor "CST_IPI", fnExcel.FormatarTexto(Campos(dicTitulosApuracao("CST_IPI")))
                .AtribuirValor "ALIQ_IPI", Campos(dicTitulosApuracao("ALIQ_IPI"))
                
            End If
            
            Call .RegistrarTributacao(assTributacaoIPI, REG)
Prx:
        Next Linha
            
        Application.StatusBar = Msg & "Atualizando dados tributários..."
        Call Util.LimparDados(assTributacaoIPI, 4, False)
        
        Call .PrepararTributacoesParaExportacao
        Call Util.ExportarDadosArrayList(assTributacaoIPI, .arrTributacoes)
        
        Call Util.MsgInformativa("Tributação do IPI atualizada com sucesso!", "Assistente de Tributação do IPI", Inicio)
        
    End With
    
End Sub

Public Sub AplicarTributacaoIPI()

Dim REG As String, CFOP$, SUGESTAO$, DT_REF$
Dim CamposTrib As Variant, Campos, ChaveTrib
Dim Tributacao As New AssistenteTributario
Dim Apuracao As New clsAssistenteApuracao
Dim Dados As Range, Linha As Range
Dim arrDados As New ArrayList
    
    Inicio = Now()
    Application.StatusBar = "Aplicando tributações selecionadas, por favor aguarde."
    
    With Tributacao
        
        Set .dicTitulosApuracao = Util.MapearTitulos(assApuracaoIPI, 3)
        Set .dicTitulosTributacao = Util.MapearTitulos(assTributacaoIPI, 3)
        Set .dicDadosTributarios = .CarregarTributacoesSalvas(assTributacaoIPI)
        Set .dicTitulos = Util.MapearTitulos(assApuracaoIPI, 3)
        
        Set Dados = assApuracaoIPI.Range("A4").CurrentRegion
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
                    
                    ChaveTrib = .GerarChaveTributacao(assApuracaoIPI, Campos)
                    CamposTrib = .ExtrairCamposTributarios(ChaveTrib, DT_REF)
                    
Reprocessar:
                    SUGESTAO = Campos(.dicTitulosApuracao("SUGESTAO"))
                    Select Case True
                        
                        Case SUGESTAO = "Aplicar o TIPO_ITEM cadastrado na Tributação"
                            Campos(.dicTitulosApuracao("TIPO_ITEM")) = CamposTrib(.dicTitulosTributacao("TIPO_ITEM"))
                            Call .LimparInconsistenciasSugestoes(Campos)
                            Call ValidarRegrasTributariasIPI(Campos, .dicTitulosApuracao, CamposTrib, .dicTitulosTributacao)
                            
                        Case SUGESTAO = "Aplicar o COD_BARRA cadastrado na Tributação"
                            Campos(.dicTitulosApuracao("COD_BARRA")) = CamposTrib(.dicTitulosTributacao("COD_BARRA"))
                            Call .LimparInconsistenciasSugestoes(Campos)
                            Call ValidarRegrasTributariasIPI(Campos, .dicTitulosApuracao, CamposTrib, .dicTitulosTributacao)
                            
                        Case SUGESTAO = "Aplicar o NCM cadastrado na Tributação"
                            Campos(.dicTitulosApuracao("COD_NCM")) = CamposTrib(.dicTitulosTributacao("COD_NCM"))
                            Call .LimparInconsistenciasSugestoes(Campos)
                            Call ValidarRegrasTributariasIPI(Campos, .dicTitulosApuracao, CamposTrib, .dicTitulosTributacao)
                            
                        Case SUGESTAO = "Aplicar a EX_IPI cadastrada na Tributação"
                            Campos(.dicTitulosApuracao("EX_IPI")) = CamposTrib(.dicTitulosTributacao("EX_IPI"))
                            Call .LimparInconsistenciasSugestoes(Campos)
                            Call ValidarRegrasTributariasIPI(Campos, .dicTitulosApuracao, CamposTrib, .dicTitulosTributacao)
                            
                        Case SUGESTAO = "Aplicar a alíquota do IPI cadastrada na Tributação"
                            Campos(.dicTitulosApuracao("ALIQ_IPI")) = CamposTrib(.dicTitulosTributacao("ALIQ_IPI"))
                            Call .LimparInconsistenciasSugestoes(Campos)
                            Call ValidarRegrasTributariasIPI(Campos, .dicTitulosApuracao, CamposTrib, .dicTitulosTributacao)
                            GoTo Reprocessar:
                            
                        Case SUGESTAO = "Aplicar o CST_IPI cadastrado na Tributação"
                            Campos(.dicTitulosApuracao("CST_IPI")) = CamposTrib(.dicTitulosTributacao("CST_IPI"))
                            Call .LimparInconsistenciasSugestoes(Campos)
                            Call ValidarRegrasTributariasIPI(Campos, .dicTitulosApuracao, CamposTrib, .dicTitulosTributacao)
                            GoTo Reprocessar:
                            
                        Case SUGESTAO = "Aplicar o COD_ENQ cadastrado na Tributação"
                            Campos(.dicTitulosApuracao("COD_ENQ")) = CamposTrib(.dicTitulosTributacao("COD_ENQ"))
                            Call .LimparInconsistenciasSugestoes(Campos)
                            Call ValidarRegrasTributariasIPI(Campos, .dicTitulosApuracao, CamposTrib, .dicTitulosTributacao)
                            GoTo Reprocessar:
                            
                        Case SUGESTAO = "Aplicar o IND_APUR cadastrado na Tributação"
                            Campos(.dicTitulosApuracao("IND_APUR")) = CamposTrib(.dicTitulosTributacao("IND_APUR"))
                            Call .LimparInconsistenciasSugestoes(Campos)
                            Call ValidarRegrasTributariasIPI(Campos, .dicTitulosApuracao, CamposTrib, .dicTitulosTributacao)
                            GoTo Reprocessar:
                            
                        Case SUGESTAO Like "Aplicar CFOP * para a operação."
                            CFOP = Util.ApenasNumeros(SUGESTAO)
                            Campos(.dicTitulosApuracao("CFOP")) = CFOP
                            Call .LimparInconsistenciasSugestoes(Campos)
                            ChaveTrib = .GerarChaveTributacao(assTributacaoIPI, Campos, True)
                            DT_REF = .ExtrairDataReferencia(Campos)
                            CamposTrib = .ExtrairCamposTributarios(ChaveTrib, DT_REF)
                            
                            Call ValidarRegrasTributariasIPI(Campos, .dicTitulosApuracao, CamposTrib, .dicTitulosTributacao)
                        
                        Case SUGESTAO Like "Aplicar o indicador de movimento cadastrado na Tributação"
                            Campos(.dicTitulosApuracao("IND_MOV")) = CamposTrib(.dicTitulosTributacao("IND_MOV"))
                            Call .LimparInconsistenciasSugestoes(Campos)
                            Call ValidarRegrasTributariasIPI(Campos, .dicTitulosApuracao, CamposTrib, .dicTitulosTributacao)
                        
                    End Select
                
                End If
                
            End If
Prx:
            If Linha.Row > 3 Then arrDados.Add Campos
            
        Next Linha
        
    End With
    
    If assApuracaoIPI.AutoFilterMode Then assApuracaoIPI.AutoFilter.ShowAllData
    Call Util.LimparDados(assApuracaoIPI, 4, False)
    Call Util.ExportarDadosArrayList(assApuracaoIPI, arrDados)
    
    Call FuncoesFormatacao.AplicarFormatacao(ActiveSheet)
    Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoIPI)
    
    Call Util.MsgInformativa("Tributações aplicadas com sucesso!", "Aplicação de Tributações PIS/COFINS", Inicio)
    Application.StatusBar = False
    
End Sub

Private Function ValidarRegrasTributariasIPI(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
    'Regras Gerais
    If Campos(dicTitulosApuracao("INCONSISTENCIA")) = "" Then _
        Call RegrasTributarias.ValidarCampo_COD_NCM(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
        
    If Campos(dicTitulosApuracao("INCONSISTENCIA")) = "" Then _
        Call RegrasTributarias.ValidarCampo_COD_BARRA(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
        
    If Campos(dicTitulosApuracao("INCONSISTENCIA")) = "" Then _
        Call RegrasTributarias.ValidarCampo_TIPO_ITEM(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
        
    If Campos(dicTitulosApuracao("INCONSISTENCIA")) = "" Then _
        Call RegrasTributarias.ValidarCampo_EX_IPI(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
        
    If Campos(dicTitulosApuracao("INCONSISTENCIA")) = "" Then _
        Call RegrasTributarias.ValidarCampo_IND_MOV(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
        
    'Regras Específicas
    If Campos(dicTitulosApuracao("INCONSISTENCIA")) = "" Then _
        Call RegrasTributarias.IPI.ValidarCampo_CST_IPI(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
        
    If Campos(dicTitulosApuracao("INCONSISTENCIA")) = "" Then _
        Call RegrasTributarias.IPI.ValidarCampo_ALIQ_IPI(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
        
    If Campos(dicTitulosApuracao("INCONSISTENCIA")) = "" Then _
        Call RegrasTributarias.IPI.ValidarCampo_COD_ENQ(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
        
    If Campos(dicTitulosApuracao("INCONSISTENCIA")) = "" Then _
        Call RegrasTributarias.IPI.ValidarCampo_IND_APUR(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
        
End Function

Public Function VerificarInconsistenciasCadastrais(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
    
    If Campos(dicTitulos("INCONSISTENCIA")) = "" Then Call RegrasCadastrais.VerificarCFOP(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA")) = "" Then Call RegrasCadastrais.ValidarCST_IPI(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA")) = "" Then Call RegrasCadastrais.IPI.ValidarCampo_CST_IPI(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA")) = "" Then Call RegrasCadastrais.IPI.ValidarCampo_ALIQ_IPI(Campos, dicTitulos)
    
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

Public Function ReprocessarSugestoes()

Dim Dados As Range, Linha As Range
Dim arrRelatorio As New ArrayList
Dim dicTitulos As New Dictionary
Dim COD_INC_TRIB As String
Dim Campos As Variant
        
    Set dicTitulos = Util.MapearTitulos(assTributacaoIPI, 3)
    If assTributacaoIPI.AutoFilterMode Then assTributacaoIPI.AutoFilter.ShowAllData
    
    Set Dados = Util.DefinirIntervalo(assTributacaoIPI, 4, 3)
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
    Call Util.LimparDados(assTributacaoIPI, 4, False)
    Call Util.ExportarDadosArrayList(assTributacaoIPI, arrRelatorio)
    Call FuncoesFormatacao.DestacarInconsistencias(assTributacaoIPI)
    
    Call Util.AtualizarBarraStatus("Processamento Concluído!")
    
End Function
