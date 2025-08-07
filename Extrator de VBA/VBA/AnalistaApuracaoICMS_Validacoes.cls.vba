Attribute VB_Name = "AnalistaApuracaoICMS_Validacoes"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private dtoValidacoes As DTOsClasseValidacoesResumoICMS

Private Type CamposResumoICMS
    
    CFOP As String
    CST_ICMS As String
    VL_ITEM As Double
    VL_BC_ICMS As Double
    ALIQ_ICMS As Double
    VL_ICMS As Double
    VL_BC_ICMS_ST As Double
    VL_ICMS_ST As Double
    INCONSISTENCIA As String
    SUGESTAO As String
    
End Type

Private Type DTOsClasseValidacoesResumoICMS
    
    dicTitulosResumoICMS As Dictionary
    dtoResumoICMS As CamposResumoICMS
    dicRegexCST_ICMS As Dictionary
    arrVerificacoes As ArrayList
    dicRegexCFOP As Dictionary
    dicTitulos As Dictionary
    Campos As Variant
    
End Type

Public Sub InicializarObjetos()

Dim CustomPart As New clsCustomPartXML
    
    Call CarregarFuncoesVerificacao
    
    Set dtoValidacoes.dicTitulosResumoICMS = Util.MapearTitulos(resICMS, 3)
    Set dtoValidacoes.dicRegexCFOP = JsonConverter.ParseJson(CustomPart.ExtrairJsonXmlPart("RegexCFOP"))
    Set dtoValidacoes.dicRegexCST_ICMS = JsonConverter.ParseJson(CustomPart.ExtrairJsonXmlPart("RegexCST_ICMS"))
    
End Sub

Public Sub ReprocessarResumoICMS(ByVal arrCampos As Variant)

Dim Campos As Variant
    
    Call CarregarFuncoesVerificacao
    
    For Each Campos In arrCampos
        
        Call ValidarRegrasResumoICMS(Campos)
        
    Next Campos
    
End Sub

Public Function ValidarRegrasResumoICMS(ByRef CamposResumoICMS As Variant) As Variant

Dim Verificacao As Variant
    
    With dtoValidacoes
        
        .Campos = CamposResumoICMS
        Call CarregarDadosDTO
        
        For Each Verificacao In dtoValidacoes.arrVerificacoes
            
            CallByName Me, CStr(Verificacao), VbMethod
            
            If .dtoResumoICMS.INCONSISTENCIA <> "" Then
                
                Call RegistrarInconsistencia
                Exit For
                
            End If
            
        Next Verificacao
        
        ValidarRegrasResumoICMS = .Campos
        
    End With
    
End Function

Public Sub CarregarFuncoesVerificacao()
    
    With dtoValidacoes
        
        Set .arrVerificacoes = New ArrayList
                
        .arrVerificacoes.Add "VerificarValoresOperacao"
        .arrVerificacoes.Add "VerificarOperacoesFiscais"
        
    End With
    
End Sub

Public Sub VerificarOperacoesFiscais()
    
    With dtoValidacoes.dtoResumoICMS
        
        Select Case True
            
            Case ValidarPadraoCFOP(.CFOP, "CompraComercializacao") And ValidarPadraoCST_ICMS(.CST_ICMS, "STGeral")
                .INCONSISTENCIA = "CST_ICMS (" & .CST_ICMS & ") inconsistente com CFOP " & .CFOP
                .SUGESTAO = "Corrigir CST_ICMS informado na operação"
                
            Case .CFOP Like "[1-2]403" And Not .CST_ICMS Like "#6[0-1]"
                .INCONSISTENCIA = "CST_ICMS (" & .CST_ICMS & ") inconsistente com CFOP " & .CFOP
                .SUGESTAO = "Informar CST_ICMS " & VBA.Left(.CST_ICMS, 1) & "60 para a operação"
                
            Case ValidarPadraoCFOP(.CFOP, "EntradaST") And Not ValidarPadraoCST_ICMS(.CST_ICMS, "STGeral")
                .INCONSISTENCIA = "CST_ICMS (" & .CST_ICMS & ") inconsistente com CFOP " & .CFOP
                .SUGESTAO = "Corrigir CST_ICMS informado na operação"
                
        End Select
        
    End With
    
End Sub

Public Sub VerificarValoresOperacao()
    
    With dtoValidacoes.dtoResumoICMS
        
        Select Case True
            
            Case VerificarOperacaoTributadaIntegralmente
                
            Case VerificarOperacaoBaseReduzida
                
            Case VerificarOperacaoIsentaNaoTributada
                
            Case VerificarOperacaoSubstituida
                
        End Select
        
    End With
    
End Sub

Private Function VerificarOperacaoTributadaIntegralmente() As Boolean
    
    With dtoValidacoes.dtoResumoICMS
        
        Select Case True
            
            Case .CST_ICMS Like "[0-8]00" And .VL_ICMS = 0
                .INCONSISTENCIA = "CST_ICMS (" & .CST_ICMS & ") com campo VL_ICMS = 0"
                .SUGESTAO = "Corrigir CST_ICMS informado na operação"
                VerificarOperacaoTributadaIntegralmente = True
                
        End Select
        
    End With
    
End Function

Private Function VerificarOperacaoBaseReduzida() As Boolean
    
    With dtoValidacoes.dtoResumoICMS
        
        Select Case True
            
            Case .CST_ICMS Like "[0-8]20" And .VL_ITEM <= .VL_BC_ICMS
                .INCONSISTENCIA = "CST_ICMS (" & .CST_ICMS & ") com VL_ITEM < VL_BC_ICMS"
                .SUGESTAO = "Corrigir CST_ICMS informado na operação"
                VerificarOperacaoBaseReduzida = True
                
            Case .CST_ICMS Like "[0-8]20" And .VL_ICMS = 0
                .INCONSISTENCIA = "CST_ICMS (" & .CST_ICMS & ") com campo VL_ICMS = 0"
                .SUGESTAO = "Corrigir CST_ICMS informado na operação"
                VerificarOperacaoBaseReduzida = True
                
        End Select
        
    End With
    
End Function

Private Function VerificarOperacaoIsentaNaoTributada() As Boolean
    
    With dtoValidacoes.dtoResumoICMS
        
        Select Case True
            
            Case .CST_ICMS Like "[0-8]4[0-1]" And .VL_ICMS > 0
                .INCONSISTENCIA = "CST_ICMS (" & .CST_ICMS & ") com campo VL_ICMS > 0"
                .SUGESTAO = "Corrigir CST_ICMS informado na operação"
                VerificarOperacaoIsentaNaoTributada = True
                
            Case .CST_ICMS Like "[0-8]4[0-1]" And .VL_BC_ICMS > 0 And .VL_ICMS = 0
                .INCONSISTENCIA = "CST_ICMS (" & .CST_ICMS & ") com campo VL_BC_ICMS > 0 e campo VL_ICMS = 0"
                .SUGESTAO = "Zerar valores do campo VL_BC_ICMS"
                VerificarOperacaoIsentaNaoTributada = True
                
        End Select
        
    End With
    
End Function

Private Function VerificarOperacaoSubstituida() As Boolean
    
    With dtoValidacoes.dtoResumoICMS
        
        Select Case True
            
            Case .CST_ICMS Like "[0-8]60" And .VL_ICMS > 0
                .INCONSISTENCIA = "CST_ICMS (" & .CST_ICMS & ") com campo VL_ICMS > 0"
                .SUGESTAO = "Corrigir CST_ICMS informado na operação"
                VerificarOperacaoSubstituida = True
                
            Case .CST_ICMS Like "[0-8]60" And .VL_BC_ICMS > 0 And .VL_ICMS = 0
                .INCONSISTENCIA = "CST_ICMS (" & .CST_ICMS & ") com campo VL_BC_ICMS > 0 e campo VL_ICMS = 0"
                .SUGESTAO = "Zerar valores do campo VL_BC_ICMS"
                VerificarOperacaoSubstituida = True
                
        End Select
        
    End With
    
End Function

Private Sub CarregarDadosDTO()
    
    With dtoValidacoes
        
        .dtoResumoICMS.CFOP = .Campos(.dicTitulosResumoICMS("CFOP"))
        .dtoResumoICMS.CST_ICMS = Util.RemoverAspaSimples(.Campos(.dicTitulosResumoICMS("CST_ICMS")))
        .dtoResumoICMS.VL_ITEM = .Campos(.dicTitulosResumoICMS("VL_ITEM"))
        .dtoResumoICMS.VL_BC_ICMS = .Campos(.dicTitulosResumoICMS("VL_BC_ICMS"))
        .dtoResumoICMS.ALIQ_ICMS = .Campos(.dicTitulosResumoICMS("ALIQ_ICMS"))
        .dtoResumoICMS.VL_ICMS = .Campos(.dicTitulosResumoICMS("VL_ICMS"))
        .dtoResumoICMS.VL_BC_ICMS_ST = .Campos(.dicTitulosResumoICMS("VL_BC_ICMS_ST"))
        .dtoResumoICMS.VL_ICMS_ST = .Campos(.dicTitulosResumoICMS("VL_ICMS_ST"))
        .dtoResumoICMS.INCONSISTENCIA = .Campos(.dicTitulosResumoICMS("INCONSISTENCIA"))
        .dtoResumoICMS.SUGESTAO = .Campos(.dicTitulosResumoICMS("SUGESTAO"))
        
    End With
    
End Sub

Private Sub ExtrairDadosDTO()
    
    With dtoValidacoes
        
        .Campos(.dicTitulosResumoICMS("CFOP")) = .dtoResumoICMS.CFOP
        .Campos(.dicTitulosResumoICMS("CST_ICMS")) = fnExcel.FormatarTexto(.dtoResumoICMS.CST_ICMS)
        .Campos(.dicTitulosResumoICMS("VL_ITEM")) = .dtoResumoICMS.VL_ITEM
        .Campos(.dicTitulosResumoICMS("VL_BC_ICMS")) = .dtoResumoICMS.VL_BC_ICMS
        .Campos(.dicTitulosResumoICMS("ALIQ_ICMS")) = .dtoResumoICMS.ALIQ_ICMS
        .Campos(.dicTitulosResumoICMS("VL_ICMS")) = .dtoResumoICMS.VL_ICMS
        .Campos(.dicTitulosResumoICMS("VL_BC_ICMS_ST")) = .dtoResumoICMS.VL_BC_ICMS_ST
        .Campos(.dicTitulosResumoICMS("VL_ICMS_ST")) = .dtoResumoICMS.VL_ICMS_ST
        .Campos(.dicTitulosResumoICMS("INCONSISTENCIA")) = .dtoResumoICMS.INCONSISTENCIA
        .Campos(.dicTitulosResumoICMS("SUGESTAO")) = .dtoResumoICMS.SUGESTAO
        
    End With
    
End Sub

Private Sub RegistrarInconsistencia()
    
    With dtoValidacoes
        
        .Campos(.dicTitulosResumoICMS("INCONSISTENCIA")) = .dtoResumoICMS.INCONSISTENCIA
        .Campos(.dicTitulosResumoICMS("SUGESTAO")) = .dtoResumoICMS.SUGESTAO
        
    End With
    
End Sub

Private Function ValidarPadraoCFOP(ByVal Texto As String, ByVal Padrao As String) As Boolean

Dim regex As New RegExp
    
    With dtoValidacoes
        
        If .dicRegexCFOP.Exists(Padrao) Then
            
            regex.Pattern = .dicRegexCFOP(Padrao)
            regex.Global = False
            
        End If
        
        ValidarPadraoCFOP = regex.Test(Texto)
        
    End With
    
End Function

Private Function ValidarPadraoCST_ICMS(ByVal Texto As String, ByVal Padrao As String) As Boolean

Dim regex As New RegExp
    
    With dtoValidacoes
        
        If .dicRegexCST_ICMS.Exists(Padrao) Then
            
            regex.Pattern = .dicRegexCST_ICMS(Padrao)
            regex.Global = False
            
        End If
        
        ValidarPadraoCST_ICMS = regex.Test(Texto)
        
    End With
    
End Function

Public Function ResetarDTOs()

Dim dtoVazio As DTOsClasseValidacoesResumoICMS
    
    LSet dtoValidacoes = dtoVazio
    
End Function
