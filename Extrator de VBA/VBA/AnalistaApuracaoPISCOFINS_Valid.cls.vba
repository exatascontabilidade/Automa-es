Attribute VB_Name = "AnalistaApuracaoPISCOFINS_Valid"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private dtoValidacoes As DTOsClasseValidacoesResumoPISCOFINS

Private Type CamposResumoPISCOFINS
    
    CFOP As String
    CST_PISCOFINS As String
    VL_ITEM As Double
    VL_BC_PISCOFINS As Double
    ALIQ_PISCOFINS As Double
    VL_PISCOFINS As Double
    VL_BC_PISCOFINS_ST As Double
    VL_PISCOFINS_ST As Double
    INCONSISTENCIA As String
    SUGESTAO As String
    
End Type

Private Type DTOsClasseValidacoesResumoPISCOFINS
    
    dicTitulosResumoPISCOFINS As Dictionary
    dtoResumoPISCOFINS As CamposResumoPISCOFINS
    dicRegexCST_PISCOFINS As Dictionary
    arrVerificacoes As ArrayList
    dicRegexCFOP As Dictionary
    dicTitulos As Dictionary
    Campos As Variant
    
End Type

Public Sub InicializarObjetos()

Dim CustomPart As New clsCustomPartXML
    
    Call CarregarFuncoesVerificacao
    
    Set dtoValidacoes.dicTitulosResumoPISCOFINS = Util.MapearTitulos(resPISCOFINS, 3)
    Set dtoValidacoes.dicRegexCFOP = JsonConverter.ParseJson(CustomPart.ExtrairJsonXmlPart("RegexCFOP"))
    Set dtoValidacoes.dicRegexCST_PISCOFINS = JsonConverter.ParseJson(CustomPart.ExtrairJsonXmlPart("RegexCST_PISCOFINS"))
    
End Sub

Public Sub ReprocessarResumoPISCOFINS(ByVal arrCampos As Variant)

Dim Campos As Variant
    
    Call CarregarFuncoesVerificacao
    
    For Each Campos In arrCampos
        
        Call ValidarResumoPISCOFINS(Campos)
        
    Next Campos
    
End Sub

Public Function ValidarResumoPISCOFINS(ByRef CamposResumoPISCOFINS As Variant) As Variant

Dim Verificacao As Variant
    
    With dtoValidacoes
        
        'Remover essas partes depois de implementar as regras de validação.
        ValidarResumoPISCOFINS = CamposResumoPISCOFINS  '<--
        Exit Function                                   '<--
        
        .Campos = CamposResumoPISCOFINS
        Call CarregarDadosDTO
        
        For Each Verificacao In dtoValidacoes.arrVerificacoes
            
            CallByName Me, CStr(Verificacao), VbMethod
            
            If .dtoResumoPISCOFINS.INCONSISTENCIA <> "" Then
                
                Call RegistrarInconsistencia
                Exit For
                
            End If
            
        Next Verificacao
        
        ValidarResumoPISCOFINS = .Campos
        
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
    
    With dtoValidacoes.dtoResumoPISCOFINS
        
        Select Case True
            
            Case ValidarPadraoCFOP(.CFOP, "CompraComercializacao") And ValidarPadraoCST_PISCOFINS(.CST_PISCOFINS, "STGeral")
                .INCONSISTENCIA = "CST_PISCOFINS (" & .CST_PISCOFINS & ") inconsistente com CFOP " & .CFOP
                .SUGESTAO = "Corrigir CST_PISCOFINS informado na operação"
                
            Case .CFOP Like "[1-2]403" And Not .CST_PISCOFINS Like "#6[0-1]"
                .INCONSISTENCIA = "CST_PISCOFINS (" & .CST_PISCOFINS & ") inconsistente com CFOP " & .CFOP
                .SUGESTAO = "Informar CST_PISCOFINS " & VBA.Left(.CST_PISCOFINS, 1) & "60 para a operação"
                
            Case ValidarPadraoCFOP(.CFOP, "EntradaST") And Not ValidarPadraoCST_PISCOFINS(.CST_PISCOFINS, "STGeral")
                .INCONSISTENCIA = "CST_PISCOFINS (" & .CST_PISCOFINS & ") inconsistente com CFOP " & .CFOP
                .SUGESTAO = "Corrigir CST_PISCOFINS informado na operação"
                
        End Select
        
    End With
    
End Sub

Public Sub VerificarValoresOperacao()
    
    With dtoValidacoes.dtoResumoPISCOFINS
        
        Select Case True
            
            Case .CST_PISCOFINS Like "[0-8]00" And .VL_PISCOFINS = 0
                .INCONSISTENCIA = "CST_PISCOFINS (" & .CST_PISCOFINS & ") com campo VL_PISCOFINS = 0"
                .SUGESTAO = "Corrigir CST_PISCOFINS informado na operação"
                
            Case .CST_PISCOFINS Like "[0-8]20" And .VL_ITEM <= .VL_BC_PISCOFINS
                .INCONSISTENCIA = "CST_PISCOFINS (" & .CST_PISCOFINS & ") com VL_ITEM < VL_BC_PISCOFINS"
                .SUGESTAO = "Corrigir CST_PISCOFINS informado na operação"
                
            Case .CST_PISCOFINS Like "[0-8]20" And .VL_PISCOFINS = 0
                .INCONSISTENCIA = "CST_PISCOFINS (" & .CST_PISCOFINS & ") com campo VL_PISCOFINS = 0"
                .SUGESTAO = "Corrigir CST_PISCOFINS informado na operação"
                
            Case .CST_PISCOFINS Like "[0-8]4[0-1]" And .VL_PISCOFINS > 0
                .INCONSISTENCIA = "CST_PISCOFINS (" & .CST_PISCOFINS & ") com campo VL_PISCOFINS > 0"
                .SUGESTAO = "Corrigir CST_PISCOFINS informado na operação"
                
            Case .CST_PISCOFINS Like "[0-8]60" And .VL_PISCOFINS > 0
                .INCONSISTENCIA = "CST_PISCOFINS (" & .CST_PISCOFINS & ") com campo VL_PISCOFINS > 0"
                .SUGESTAO = "Corrigir CST_PISCOFINS informado na operação"
                    
        End Select
        
    End With
    
End Sub

Private Sub CarregarDadosDTO()
    
    With dtoValidacoes
        
        .dtoResumoPISCOFINS.CFOP = .Campos(.dicTitulosResumoPISCOFINS("CFOP"))
        .dtoResumoPISCOFINS.CST_PISCOFINS = .Campos(.dicTitulosResumoPISCOFINS("CST_PISCOFINS"))
        .dtoResumoPISCOFINS.VL_ITEM = .Campos(.dicTitulosResumoPISCOFINS("VL_ITEM"))
        .dtoResumoPISCOFINS.VL_BC_PISCOFINS = .Campos(.dicTitulosResumoPISCOFINS("VL_BC_PISCOFINS"))
        .dtoResumoPISCOFINS.ALIQ_PISCOFINS = .Campos(.dicTitulosResumoPISCOFINS("ALIQ_PISCOFINS"))
        .dtoResumoPISCOFINS.VL_PISCOFINS = .Campos(.dicTitulosResumoPISCOFINS("VL_PISCOFINS"))
        .dtoResumoPISCOFINS.VL_BC_PISCOFINS_ST = .Campos(.dicTitulosResumoPISCOFINS("VL_BC_PISCOFINS_ST"))
        .dtoResumoPISCOFINS.VL_PISCOFINS_ST = .Campos(.dicTitulosResumoPISCOFINS("VL_PISCOFINS_ST"))
        .dtoResumoPISCOFINS.INCONSISTENCIA = .Campos(.dicTitulosResumoPISCOFINS("INCONSISTENCIA"))
        .dtoResumoPISCOFINS.SUGESTAO = .Campos(.dicTitulosResumoPISCOFINS("SUGESTAO"))
        
    End With
    
End Sub

Private Sub ExtrairDadosDTO()
    
    With dtoValidacoes
        
        .Campos(.dicTitulosResumoPISCOFINS("CFOP")) = .dtoResumoPISCOFINS.CFOP
        .Campos(.dicTitulosResumoPISCOFINS("CST_PISCOFINS")) = .dtoResumoPISCOFINS.CST_PISCOFINS
        .Campos(.dicTitulosResumoPISCOFINS("VL_ITEM")) = .dtoResumoPISCOFINS.VL_ITEM
        .Campos(.dicTitulosResumoPISCOFINS("VL_BC_PISCOFINS")) = .dtoResumoPISCOFINS.VL_BC_PISCOFINS
        .Campos(.dicTitulosResumoPISCOFINS("ALIQ_PISCOFINS")) = .dtoResumoPISCOFINS.ALIQ_PISCOFINS
        .Campos(.dicTitulosResumoPISCOFINS("VL_PISCOFINS")) = .dtoResumoPISCOFINS.VL_PISCOFINS
        .Campos(.dicTitulosResumoPISCOFINS("VL_BC_PISCOFINS_ST")) = .dtoResumoPISCOFINS.VL_BC_PISCOFINS_ST
        .Campos(.dicTitulosResumoPISCOFINS("VL_PISCOFINS_ST")) = .dtoResumoPISCOFINS.VL_PISCOFINS_ST
        .Campos(.dicTitulosResumoPISCOFINS("INCONSISTENCIA")) = .dtoResumoPISCOFINS.INCONSISTENCIA
        .Campos(.dicTitulosResumoPISCOFINS("SUGESTAO")) = .dtoResumoPISCOFINS.SUGESTAO
        
    End With
    
End Sub

Private Sub RegistrarInconsistencia()
    
    With dtoValidacoes
        
        .Campos(.dicTitulosResumoPISCOFINS("INCONSISTENCIA")) = .dtoResumoPISCOFINS.INCONSISTENCIA
        .Campos(.dicTitulosResumoPISCOFINS("SUGESTAO")) = .dtoResumoPISCOFINS.SUGESTAO
        
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

Private Function ValidarPadraoCST_PISCOFINS(ByVal Texto As String, ByVal Padrao As String) As Boolean

Dim regex As New RegExp
    
    With dtoValidacoes
        
        If .dicRegexCST_PISCOFINS.Exists(Padrao) Then
            
            regex.Pattern = .dicRegexCST_PISCOFINS(Padrao)
            regex.Global = False
            
        End If
        
        ValidarPadraoCST_PISCOFINS = regex.Test(Texto)
        
    End With
    
End Function

Public Function ResetarDTOs()

Dim dtoVazio As DTOsClasseValidacoesResumoPISCOFINS
    
    LSet dtoValidacoes = dtoVazio
    
End Function

