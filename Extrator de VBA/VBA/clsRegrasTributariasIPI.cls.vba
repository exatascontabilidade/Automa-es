Attribute VB_Name = "clsRegrasTributariasIPI"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function ValidarCampo_ALIQ_IPI(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
Dim ALIQ_IPI As Double, ALIQ_IPI_TRIB#
Dim INCONSISTENCIA As String, SUGESTAO$, CFOP$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração do IPI
    ALIQ_IPI = fnExcel.FormatarPercentuais(Campos(dicTitulosApuracao("ALIQ_IPI") - i))
    CFOP = Campos(dicTitulosApuracao("CFOP") - i)
    
    If LBound(Campos) = 0 Then i = 1
    'Carregar campos do assistente de Tributação do IPI
    
    ALIQ_IPI_TRIB = fnExcel.FormatarPercentuais(CamposTrib(dicTitulosTributacao("ALIQ_IPI") - i))
    
    Select Case True
        
        Case ALIQ_IPI <> ALIQ_IPI_TRIB
            INCONSISTENCIA = "ALIQ_IPI divergente: " & fnExcel.FormatarPercentuais(ALIQ_IPI) & " (informado) vs " & fnExcel.FormatarPercentuais(ALIQ_IPI_TRIB) & " (cadastrado) para a operação com CFOP " & CFOP
            SUGESTAO = "Aplicar a alíquota do IPI cadastrada na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_CST_IPI(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)

Dim i As Byte, t As Byte
Dim arrCST_IPI As New ArrayList
Dim CST_IPI As String, CST_IPI_TRIB$
Dim INCONSISTENCIA As String, SUGESTAO$, CFOP$
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração do IPI
    CST_IPI = Util.ApenasNumeros(Campos(dicTitulosApuracao("CST_IPI") - i))
    CFOP = Campos(dicTitulosApuracao("CFOP") - i)
    
    If LBound(CamposTrib) = 0 Then t = 1
    
    'Carregar campos do assistente de Tributação do IPI
    CST_IPI_TRIB = Util.ApenasNumeros(CamposTrib(dicTitulosTributacao("CST_IPI") - t))
    
    Select Case True
        
        Case CST_IPI <> CST_IPI_TRIB
            INCONSISTENCIA = "CST_IPI divergente: " & CST_IPI & " (informado) vs " & CST_IPI_TRIB & " (cadastrado) para a operação com CFOP " & CFOP
            SUGESTAO = "Aplicar o CST_IPI cadastrado na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_COD_ENQ(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)

Dim i As Byte, t As Byte
Dim arrCOD_ENQ As New ArrayList
Dim COD_ENQ As String, COD_ENQ_TRIB$
Dim INCONSISTENCIA As String, SUGESTAO$, CFOP$
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração do IPI
    COD_ENQ = Util.ApenasNumeros(Campos(dicTitulosApuracao("COD_ENQ") - i))
    CFOP = Campos(dicTitulosApuracao("CFOP") - i)
    
    If LBound(CamposTrib) = 0 Then t = 1
    
    'Carregar campos do assistente de Tributação do IPI
    COD_ENQ_TRIB = Util.ApenasNumeros(CamposTrib(dicTitulosTributacao("COD_ENQ") - t))
    
    Select Case True
        
        Case COD_ENQ <> COD_ENQ_TRIB
            INCONSISTENCIA = "COD_ENQ divergente: " & COD_ENQ & " (informado) vs " & COD_ENQ_TRIB & " (cadastrado) para a operação com CFOP " & CFOP
            SUGESTAO = "Aplicar o COD_ENQ cadastrado na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_IND_APUR(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)

Dim i As Byte, t As Byte
Dim arrIND_APUR As New ArrayList
Dim IND_APUR As String, IND_APUR_TRIB$
Dim INCONSISTENCIA As String, SUGESTAO$, CFOP$
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração do IPI
    IND_APUR = Util.ApenasNumeros(Campos(dicTitulosApuracao("IND_APUR") - i))
    CFOP = Campos(dicTitulosApuracao("CFOP") - i)
    
    If LBound(CamposTrib) = 0 Then t = 1
    
    'Carregar campos do assistente de Tributação do IPI
    IND_APUR_TRIB = Util.ApenasNumeros(CamposTrib(dicTitulosTributacao("IND_APUR") - t))
    
    Select Case True
        
        Case IND_APUR <> IND_APUR_TRIB
            INCONSISTENCIA = "IND_APUR divergente: " & IND_APUR & " (informado) vs " & IND_APUR_TRIB & " (cadastrado) para a operação com CFOP " & CFOP
            SUGESTAO = "Aplicar o IND_APUR cadastrado na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function


