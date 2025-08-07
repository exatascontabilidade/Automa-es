Attribute VB_Name = "ControlDocs_v2"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    
    On Error Resume Next
    Call FuncoesControlDocs.EnviarAcionamentos("RELATORIO_ACIONAMENTOS")
    Call ConfigControlDocs.ResetarConfiguracoes
    If FuncoesControlDocs.ObterUuidComputador = "8A1DD300-CCB5-11EC-B9D4-478D28DC6B00" Then Call FuncoesControlDocs.VersionarProjeto
    Call fnSeguranca.VerificarDadosTributarios("fechar")
    
End Sub

Private Sub Workbook_Open()
    
    VerificacoesControlDocs.VerificarAtualizacao
    Call VerificacoesControlDocs.VerificarConfiguracoesControlDocs
    Call ConsultarStatusAssinatura("CONSULTAR_ASSINATURA")
    
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    
    Call AplicarFormatacao(Sh)
    
    If Application.StatusBar <> False Then Application.StatusBar = False
    If Sh.CodeName = "relDivergencias" Then Call FuncoesFormatacao.FormatarDivergencias(Sh)
    If Sh.CodeName Like "relInteligente*" Then Call FuncoesFormatacao.DestacarInconsistencias(Sh)
    If Sh.CodeName Like "res*" Then Call FuncoesFormatacao.DestacarInconsistencias(Sh)
    If Sh.CodeName Like "relDiv*" Then Call FuncoesFormatacao.DestacarInconsistencias(Sh)
    If Sh.CodeName Like "ass*" Then Call FuncoesFormatacao.DestacarInconsistencias(Sh)
    If Sh.CodeName = "relICMS" Then Call FuncoesFormatacao.FormatarInconsistencias(relICMS)
    If Sh.CodeName = "relCorrelacoes" Then Call FuncoesFormatacao.DestacarMelhorCorrelacao(relCorrelacoes)
    
    Call AtualizarRibbon(ActiveSheet.name)
    
    If Not Rib Is Nothing Then
        
        On Error Resume Next
        Select Case True
            
            Case (Sh.CodeName = "Autenticacao") Or (Sh.CodeName = "CadContrib")
                Rib.ActivateTab "tbControlDocs"
                
            Case (VBA.Left(Sh.CodeName, 3) = "Ent") Or (VBA.Left(Sh.CodeName, 3) = "Sai")
                Rib.ActivateTab "tbDocumentos"
                
            Case (VBA.Left(Sh.CodeName, 3) = "reg")
                Rib.ActivateTab "tbRegEFD"
                
            Case (VBA.Left(Sh.CodeName, 5) = "Livro") Or (VBA.Left(Sh.CodeName, 5) = "Corre") _
                Or (VBA.Left(Sh.CodeName, 5) = "Diver" Or (VBA.Left(Sh.CodeName, 4) = "Trib"))
                Rib.ActivateTab "tbAnalises"
                
            Case Sh.CodeName = "assApuracaoICMS"
                Rib.ActivateTab "tbAssistentesFiscais"
                Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoICMS)
                
            Case Sh.CodeName = "assApuracaoIPI"
                Rib.ActivateTab "tbAssistentesFiscais"
                Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoIPI)
                
            Case Sh.CodeName = "assApuracaoPISCOFINS"
                Rib.ActivateTab "tbAssistentesFiscais"
                Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoPISCOFINS)
                
            Case Sh.CodeName = "assTributacaoPISCOFINS"
                Call AssTributario.DestacarNovosCadastros(assTributacaoPISCOFINS)
                
            Case Sh.CodeName = "assTributacaoICMS"
                Call AssTributario.DestacarNovosCadastros(assTributacaoICMS)
                
            Case Sh.CodeName = "assTributacaoIPI"
                Call AssTributario.DestacarNovosCadastros(assTributacaoIPI)
                
        End Select
        
    End If
    
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

Dim Valor As String
Dim Intervalo As Range
Dim Valores As Variant
Dim dicTitulos As New Dictionary

    On Error GoTo Tratar:
    
    If Application.StatusBar <> False Then Application.StatusBar = False
    Select Case True
        
        'Valida se a planilha é válida para os recursos
        Case (VBA.Left(Sh.CodeName, 3) <> "rel" And VBA.Left(Sh.CodeName, 3) <> "ass") And VBA.Left(Sh.CodeName, 3) <> "reg"
            Exit Sub
            
    End Select
    
    Call Util.DesabilitarControles
        
        'Filtros de Dados
        If (Target.Count = 1) And (Target.Row = 1) And (Target.Column <> 1) Then
            
            Valor = Target.value
            If Valor = "" Then
                If Sh.AutoFilterMode Then Sh.AutoFilter.ShowAllData
                Call Util.HabilitarControles
                Exit Sub
                
            ElseIf VBA.InStr(1, Valor, ",") > 0 Then
                Valores = VBA.Split(Valor, ",")
                
            ElseIf VBA.InStr(1, Valor, ";") > 0 Then
                Valores = VBA.Split(Valor, ";")
                                                            
            End If
            
            Set Intervalo = Util.DefinirIntervalo(Sh, 4, 3)
            If Not Intervalo Is Nothing Then
                
                If IsEmpty(Valores) Then
                    Intervalo.AutoFilter Field:=Target.Column, Criteria1:=Valor
                Else
                    Intervalo.AutoFilter Field:=Target.Column, Criteria1:=Valores, Operator:=xlFilterValues
                End If
                
            End If
    
        End If
        
        'Tratamento de Enumerações
        If Target.Row > 3 And Target.Count = 1 And Target.value <> "" Then
            
            Application.DisplayFormulaBar = True
            Set dicTitulos = Util.MapearTitulos(Sh, 3)
            
            If dicTitulos.Exists("ARQUIVO") Then ARQUIVO = Sh.Cells(Target.Row, dicTitulos("ARQUIVO")).value
            
            If Target.Column = dicTitulos("REGIME_TRIBUTARIO") Then Target.value = RegrasCadastrais.PIS_COFINS.ValidarEnumeracao_REGIME_TRIBUTARIO(Target.value)
            If Target.Column = dicTitulos("COD_CONS") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_COD_CONS(Target.value)
            If Target.Column = dicTitulos("COD_FIN") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_COD_FIN(Target.value)
            If Target.Column = dicTitulos("COD_GRUPO_TENSAO") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_COD_GRUPO_TENSAO(Target.value)
            If Target.Column = dicTitulos("COD_INC_TRIB") Then Target.value = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_COD_INC_TRIB(Target.value)
            If Target.Column = dicTitulos("COD_SIT") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_COD_SIT(Target.value)
            If Target.Column = dicTitulos("COD_TIPO_CONT") Then Target.value = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_COD_TIPO_CONT(Target.value)
            If Target.Column = dicTitulos("COD_NAT_CC") Then Target.value = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_COD_NAT_CC(Target.value)
            If Target.Column = dicTitulos("CST_COFINS") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(Target.value)
            If Target.Column = dicTitulos("CST_PIS") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(Target.value)
            If Target.Column = dicTitulos("CST_IPI") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_IPI(Target.value)
            If Target.Column = dicTitulos("FIN_DOCE") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_FIN_DOCE(Target.value)
            If Target.Column = dicTitulos("IND_APRO_CRED") Then Target.value = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_IND_APRO_CRED(Target.value)
            If Target.Column = dicTitulos("IND_CTA") Then Target.value = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_IND_CTA(Target.value)
            If Target.Column = dicTitulos("IND_APUR") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_APUR(Target.value)
            If Target.Column = dicTitulos("IND_ATIV") And Sh.name = "0000_Contr" Then Target.value = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_IND_ATIV(Target.value)
            If Target.Column = dicTitulos("IND_ATIV") And Sh.name = "0000" Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_ATIV(Target.value)
            If Target.Column = dicTitulos("IND_DEST") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_DEST(Target.value)
            If Target.Column = dicTitulos("IND_EMIT") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_EMIT(Target.value)
            If Target.Column = dicTitulos("IND_FRT") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_FRT(Target.value)
            If Target.Column = dicTitulos("IND_MOV") And Sh.name = "C170" Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_C170_IND_MOV(Target.value)
            If Target.Column = dicTitulos("IND_MOV") And Sh.name = "Assistente de Apuração ICMS" Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_C170_IND_MOV(Target.value)
            If Target.Column = dicTitulos("IND_MOV") And Sh.name = "Assist. Tributação PIS e COFINS" Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_C170_IND_MOV(Target.value)
            If Target.Column = dicTitulos("IND_MOV") And Sh.name Like "*001" Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_MOV(Target.value)
            If Target.Column = dicTitulos("IND_NAT_PJ") Then Target.value = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_IND_NAT_PJ(Target.value)
            If Target.Column = dicTitulos("IND_OPER") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_OPER(Target.value)
            If Target.Column = dicTitulos("IND_PGTO") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_PGTO(Target.value)
            If Target.Column = dicTitulos("IND_PROP") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_PROP(Target.value)
            If Target.Column = dicTitulos("IND_REG_CUM") Then Target.value = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_IND_REG_CUM(Target.value)
            If Target.Column = dicTitulos("IND_TIT") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_TIT(Target.value)
            If Target.Column = dicTitulos("MOT_INV") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_MOT_INV(Target.value)
            If Target.Column = dicTitulos("TIPO_ESCRIT") Then Target.value = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_TIPO_ESCRIT(Target.value)
            If Target.Column = dicTitulos("IND_NAT_PJ") Then Target.value = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_IND_NAT_PJ(Target.value)
            If Target.Column = dicTitulos("IND_SIT_ESP") Then Target.value = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_IND_SIT_ESP(Target.value)
            If Target.Column = dicTitulos("TIPO_ITEM") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_TIPO_ITEM(Target.value)
            If Target.Column = dicTitulos("TP_ASSINANTE") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_TP_ASSINANTE(Target.value)
            If Target.Column = dicTitulos("TP_CT_E") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_TP_CT_E(Target.value)
            If Target.Column = dicTitulos("TP_LIGACAO") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_TP_LIGACAO(Target.value)
            If Target.Column = dicTitulos("COD_DOC_IMP") Then Target.value = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_COD_DOC_IMP(Target.value)
            If Target.Column = dicTitulos("NAT_BC_CRED") Then Target.value = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_NAT_BC_CRED(Target.value)
            If Target.Column = dicTitulos("CST_PIS_NF") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(Target.value)
            If Target.Column = dicTitulos("CST_PIS_SPED") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(Target.value)
            If Target.Column = dicTitulos("CST_COFINS_NF") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(Target.value)
            If Target.Column = dicTitulos("CST_COFINS_SPED") Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(Target.value)
            If Target.Column = dicTitulos("COD_NAT") And Not Sh.name Like "*0400*" Then Target.value = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_COD_NAT(ARQUIVO, Target.value)
            
        End If
    
    Call Util.HabilitarControles
    
    Exit Sub
Tratar:

    Call Util.HabilitarControles
    If Err.Number = 6 Then Exit Sub

End Sub

