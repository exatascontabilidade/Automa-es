Attribute VB_Name = "clsRegrasTributariasICMS"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function ValidarCampo_ALIQ_ICMS(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
Dim ALIQ_ICMS As Double, ALIQ_ICMS_TRIB#
Dim INCONSISTENCIA As String, SUGESTAO$, CFOP$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração do ICMS
    ALIQ_ICMS = fnExcel.FormatarPercentuais(Campos(dicTitulosApuracao("ALIQ_ICMS") - i))
    CFOP = Campos(dicTitulosApuracao("CFOP") - i)
    
    If LBound(Campos) = 0 Then i = 1
    'Carregar campos do assistente de Tributação do ICMS
    
    ALIQ_ICMS_TRIB = fnExcel.FormatarPercentuais(CamposTrib(dicTitulosTributacao("ALIQ_ICMS") - i))
    
    Select Case True
        
        Case ALIQ_ICMS <> ALIQ_ICMS_TRIB
            INCONSISTENCIA = "ALIQ_ICMS divergente: " & fnExcel.FormatarPercentuais(ALIQ_ICMS) & " (informado) vs " & fnExcel.FormatarPercentuais(ALIQ_ICMS_TRIB) & " (cadastrado) para a operação com CFOP " & CFOP
            SUGESTAO = "Aplicar a alíquota do ICMS cadastrada na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_CST_ICMS(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)

Dim i As Byte, t As Byte
Dim arrCST_ICMS As New ArrayList
Dim CST_ICMS As String, CST_ICMS_TRIB$
Dim INCONSISTENCIA As String, SUGESTAO$, CFOP$
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração do ICMS
    CST_ICMS = Campos(dicTitulosApuracao("CST_ICMS") - i)
    CFOP = Campos(dicTitulosApuracao("CFOP") - i)
    
    If LBound(CamposTrib) = 0 Then t = 1
    
    'Carregar campos do assistente de Tributação do ICMS
    CST_ICMS_TRIB = CamposTrib(dicTitulosTributacao("CST_ICMS") - t)
    
    Select Case True
        
        Case CST_ICMS <> CST_ICMS_TRIB
            INCONSISTENCIA = "CST_ICMS divergente: " & CST_ICMS & " (informado) vs " & CST_ICMS_TRIB & " (cadastrado) para a operação com CFOP " & CFOP
            SUGESTAO = "Aplicar o CST_ICMS cadastrado na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_ALIQ_ST(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
Dim ALIQ_ST As Double, ALIQ_ST_TRIB#
Dim INCONSISTENCIA As String, SUGESTAO$, CFOP$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração do ICMS
    ALIQ_ST = fnExcel.FormatarPercentuais(Campos(dicTitulosApuracao("ALIQ_ST") - i))
    CFOP = Campos(dicTitulosApuracao("CFOP") - i)

    If LBound(Campos) = 0 Then i = 1
    'Carregar campos do assistente de Tributação do ICMS
    
    ALIQ_ST_TRIB = fnExcel.FormatarPercentuais(CamposTrib(dicTitulosTributacao("ALIQ_ST") - i))
    
    Select Case True
        
        Case ALIQ_ST <> ALIQ_ST_TRIB
            INCONSISTENCIA = "ALIQ_ST divergente: " & fnExcel.FormatarPercentuais(fnExcel.FormatarPercentuais(ALIQ_ST)) & " (informado) vs " & fnExcel.FormatarPercentuais(fnExcel.FormatarPercentuais(ALIQ_ST_TRIB)) & " (cadastrado) para a operação com CFOP " & CFOP
            SUGESTAO = "Aplicar a alíquota do ICMS-ST cadastrada na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function
