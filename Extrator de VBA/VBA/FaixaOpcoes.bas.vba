Attribute VB_Name = "FaixaOpcoes"
Option Explicit

Sub InicializarFaixaPersonalizada(ribbon As IRibbonUI)
    
    On Error Resume Next
    Set Rib = ribbon
    Application.DisplayAlerts = False
        Rib.ActivateTab "tbControlDocs"
        If CNPJContribuinte = "" Then CadContrib.Activate
        If EmailAssinante = "" Then relGestaoAssinatura.Activate
    Application.DisplayAlerts = True
    
    Call RecarregarRibbon
    
End Sub

Sub ForcarErro()
    ' Força um erro personalizado com código de erro 9999
    Err.Raise 9999, "ForcarErro", "Este é um erro forçado para testar a rotina"
End Sub

Sub RecarregarRibbon()
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
End Sub

Sub getClique(control As IRibbonControl, ByRef Check)
        
    If ControlPressionado Then
        Call DocumentacaoControlDocs.AcessarDocumentacao(control)
        Rib.InvalidateControl control.id
        Exit Sub
    End If
        
    Select Case True
        
        Case control.id = "chItensNotasProprias"
            ExportarC170Proprios = Check

        Case control.id = "chItensNotasProprias"
            ExportarC170Proprios = Check
            
        Case control.id = "chPISCOFINS"
            ExportarPISCOFINS = Check
            
        Case control.id = "chC170Contrib"
            ExportarC170Contribuicoes = Check
            
        Case control.id = "chC140"
            ExportarC140Filhos = Check
            
        Case control.id = "chC175Contrib"
            ExportarC175Contruicoes = Check
            
        Case control.id = "chDescPISCOFINS"
            DesconsiderarPISCOFINS = Check

        Case control.id = "chAbatimento"
            DesconsiderarAbatimento = Check
        
        Case control.id = "chIPI"
            SomarIPIProdutos = Check
            
        Case control.id = "chNFeSemValidade"
            chNFSemValidade = Check
        
        Case control.id = "chICMSST"
            SomarICMSSTProdutos = Check
            
        Case control.id = "chIgnorarEmissoesProprias"
            IgnorarEmissoesProprias = Check
                
        Case control.id = "chCadItensTercProprios"
            CadItensFornecProprios = Check
            
        Case control.id = "chLinhasGrade"
            ConfiguracoesControlDocs.Range("LinhasGrade").value = Check
            Call ConfigControlDocs.RemoverLinhasGrade(Not Check)
        
        Case control.id = "chApropriarCreditosICMS"
            ApropriarCreditosICMS = Check
        
        Case control.id = "chImportarCTeD100"
            ImportarCTeD100 = Check
        
        Case (control.id = "chImportPeriodo")
            StatusPeriodo = Check
            If Not Check Then PeriodoEspecifico = ""
            Rib.InvalidateControl "lbInstrucoesPeriodo"
            Rib.InvalidateControl "btnDefinirPeriodo"
            Rib.InvalidateControl "ebPeriodo"
            Rib.InvalidateControl "lbXML"
        
        Case (control.id = "chIgnoreQtdUnidXML")
            ConfiguracoesControlDocs.Range("IgnorarQtdUnidXML").value = Check
        
        'TODO: CRIAR CHECKBOX PARA IMPORTAR APENAS NOTAS DE EMISSÃO PRÓPRIA NOS XMLS
        
    End Select
    
End Sub

Sub getText(control As IRibbonControl, Optional ByRef text)
    
    If control.id = "ebPeriodo" Then text = PeriodoEspecifico
    If control.id = "edPeriodo" Then text = PeriodoImportacao
    If control.id = "edInventario" Then text = PeriodoInventario
    
End Sub

Sub LimparTexto(control As IRibbonControl, Optional ByRef text)
    text = ""
End Sub

Sub AcessarRegistro(control As IRibbonControl, Optional text As String)

On Error GoTo Tratar:
        
    Select Case True
        
        Case (VBA.Len(text) = 5) And (VBA.UCase(VBA.Right(text, 1)) = "C")
            text = VBA.Left(text, 4) & "_Contr"
            
    End Select
    
    If text <> "" Then
        text = VBA.UCase(text)
        Worksheets(text).Activate
        Call GetVisible(control, True)
        Call LimparTexto(control, "")
        Call AtualizarRibbon(text)
    End If
    
Exit Sub

Tratar:

MsgBox "O registro informado está inválido!" & vbCrLf & _
       "Por favor verifique o registro digitado e tente novamente.", _
       vbExclamation, "Registro Inválido"
    
End Sub

Sub GetVisible(control As IRibbonControl, Optional ByRef visible)
    
Dim nReg As String
    
    On Error GoTo TratarErro:
    
    With ActiveSheet
        
        Select Case True
            
            Case control.id = "lbInstrucoesPeriodo" Or control.id = "ebPeriodo" Or control.id = "lbXML" Or control.id = "btnDefinirPeriodo"
                If StatusPeriodo Then visible = True Else visible = False
            
            Case control.id = "grEnt"
                If control.Tag = "fEntNFe" Then visible = True Else visible = False
                
            Case control.id = "grSai"
                If control.Tag = "Quebra" Then visible = True Else visible = False
                
            Case (VBA.InStr(1, .CodeName, "NFe")) Or (VBA.InStr(1, .CodeName, "CTe")) Or (VBA.InStr(1, .CodeName, "NFCe")) Or (VBA.InStr(1, .CodeName, "CFe"))
                If control.Tag = "DocsSemLancar" Then visible = True
                
            Case (.CodeName = "CadContrib")
                If control.Tag = "CadContrib" Then visible = True
                
            Case (.CodeName = "Divergencias")
                If control.Tag = "Divergências Fiscais" Then visible = True
                
            Case (.CodeName = "relICMS")
                If control.Tag = "Livro ICMS" Then visible = True
                
            Case (.CodeName = "LivroIPI")
                If control.Tag = "Livro IPI" Then visible = True
                
            Case (.CodeName = "LivroPISCOFINS")
                If control.Tag = "Livro PIS-COFINS" Then visible = True
                
            Case (.CodeName = "Autenticacao")
                If control.Tag = "Autenticação" Then visible = True
            
            Case (.CodeName = "Correlacoes")
                If control.Tag = "Correlação Produtos" Then visible = True
                
            Case (.CodeName = "Tributacao")
                If control.Tag = "btnImportarItensSPED" Then visible = True
                
            Case (.CodeName = "assApuracaoICMS")
                If control.Tag = "Análise Produtos" Then visible = True
                
            Case (.CodeName = "assApuracaoPISCOFINS")
                If control.Tag = "Assistente de PIS e COFINS" Then visible = True
                
            Case (.CodeName = "relInteligenteDivergencias")
                If control.Tag = "Assistente de Divergencias" Then visible = True
                
            Case (.CodeName = "assTributacaoICMS")
                If control.Tag = "Assistente de Tributação ICMS" Then visible = True
                
            Case (.CodeName = "relInteligenteTribIPI")
                If control.Tag = "Assistente de Tributação ICMS" Then visible = True
                
            Case (.CodeName = "relInteligenteTribPISCOFINS")
                If control.Tag = "Assistente de Tributação ICMS" Then visible = True
                
            Case (.CodeName = "relCustosPrecos")
                If control.Tag = "Assistente de Custos e Preços" Then visible = True
                
            Case (.CodeName = "relInteligenteEstoque")
                If control.Tag = "Assistente de Estoque" Then visible = True
                
            Case (.CodeName = "relInteligenteContas")
                If control.Tag = "Assistente de Contas" Then visible = True
                
            Case (.CodeName = "relInventário")
                If control.Tag = "Assistente de Inventário" Then visible = True
                
            Case .CodeName Like "*_Contr"
                nReg = VBA.Mid(.CodeName, 4, 4)
                If control.Tag Like "*f" & nReg & "_Contr*" Then visible = True
                
            Case .CodeName Like "reg*"
                nReg = VBA.Right(.CodeName, 4)
                If control.Tag Like "f*" & nReg Then visible = True
                
        End Select
        
    End With
    
    If control.Tag Like MyTag Then visible = True
    
Exit Sub
TratarErro:
    
     MsgBox "Erro!", vbCritical, "Erro"
    
End Sub

Sub AtualizarRibbon(Tag As String)
     MyTag = Tag
     If Rib Is Nothing Then MsgBox "Reinicie a planilha." Else Rib.Invalidate
End Sub

Sub IrPara(control As IRibbonControl)
    Worksheets(control.Tag).Activate
    Call AtualizarRibbon(control.Tag)
End Sub

Sub MostrarGrupos(control As IRibbonControl, Optional Plan As Worksheet)
    
    Call AtualizarRibbon(control.Tag)
    On Error Resume Next
    Plan.Activate
    
End Sub

Public Sub getEnabled(ByRef control As IRibbonControl, ByRef Valor)
    
    If control.id = "ebPeriodo" Then Valor = StatusPeriodo
    
End Sub

Function getPressed(control As IRibbonControl, ByRef Check)
    
    Select Case True
        
        Case control.id = "chImportPeriodo"
            Check = StatusPeriodo
        
        Case control.id = "btnDefinirPeriodo"
            Check = UsarPeriodo
            
        Case control.id = "chItensNotasProprias"
            Check = ExportarC170Proprios

        Case control.id = "chItensNotasProprias"
            Check = ExportarC170Proprios
            
        Case control.id = "chPISCOFINS"
            Check = ExportarPISCOFINS
            
        Case control.id = "chC170Contrib"
            Check = ExportarC170Contribuicoes
            
        Case control.id = "chC140"
            Check = ExportarC140Filhos
            
        Case control.id = "chC175Contrib"
            Check = ExportarC175Contruicoes
            
        Case control.id = "chDescPISCOFINS"
            Check = DesconsiderarPISCOFINS

        Case control.id = "chAbatimento"
            Check = DesconsiderarAbatimento
            
        Case control.id = "chIPI"
            Check = SomarIPIProdutos
            
        Case control.id = "chICMSST"
            Check = SomarICMSSTProdutos
            
        Case control.id = "chNFeSemValidade"
            Check = chNFSemValidade
            
        Case control.id = "chCadItensTercProprios"
            Check = CadItensFornecProprios
            
        Case control.id = "chApropriarCreditosICMS"
            Check = ApropriarCreditosICMS
        
        Case control.id = "chIgnorarEmissoesProprias"
            Check = IgnorarEmissoesProprias
            
        Case control.id = "chLinhasGrade"
            Check = CBool(ConfiguracoesControlDocs.Range("LinhasGrade").value)
            
        Case (control.id = "chIgnoreQtdUnidXML")
            Check = CBool(ConfiguracoesControlDocs.Range("IgnorarQtdUnidXML").value)
            
        Case control.id = "chImportarCTeD100"
            Check = ImportarCTeD100
            
    End Select
    
End Function

Public Function DefinirPeriodo(control As IRibbonControl, Optional text As String)
    
    If control.id = "ebPeriodo" Then PeriodoEspecifico = text
    If control.id = "edPeriodo" Then PeriodoImportacao = text
    If control.id = "edInventario" Then PeriodoInventario = text
    
End Function

Public Sub getOption(control As IRibbonControl, id As String, index As Integer)
    
    If ControlPressionado Then
        Call DocumentacaoControlDocs.AcessarDocumentacao(control)
        Exit Sub
    End If
    
    Select Case id
        
        Case "D00"
            dias = 0
            
        Case "D07"
            dias = 7
            
        Case "D15"
            dias = 15
            
        Case "D21"
            dias = 21
            
        Case "D28"
            dias = 28
            
        Case "D30"
            dias = 30
            
    End Select
    
End Sub

Public Sub getSelectedItemID(control As IRibbonControl, ByRef retorno)
    
    If ControlPressionado Then
        Call DocumentacaoControlDocs.AcessarDocumentacao(control)
        Rib.InvalidateControl control.id
        Exit Sub
    End If
    
    retorno = "D" & VBA.Format(dias, "00")
    
End Sub

Sub VerificarUUID(control As IRibbonControl, ByRef value)

Dim Uuid As String

    Uuid = FuncoesControlDocs.ObterUuidComputador()
    If Uuid = "8A1DD300-CCB5-11EC-B9D4-478D28DC6B00" Then value = True
    
End Sub
