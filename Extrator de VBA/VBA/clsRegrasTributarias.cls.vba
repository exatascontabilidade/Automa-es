Attribute VB_Name = "clsRegrasTributarias"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public PIS_COFINS As New clsRegrasTributariasPISCOFINS
Public ICMS As New clsRegrasTributariasICMS
Public IPI As New clsRegrasTributariasIPI

Public Function ValidarCampo_COD_NCM(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
Dim COD_NCM As String, COD_NCM_TRIB$, DESCR_ITEM$
Dim INCONSISTENCIA As String, SUGESTAO$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração PIS/COFINS
    COD_NCM = Campos(dicTitulosApuracao("COD_NCM") - i)
    DESCR_ITEM = Campos(dicTitulosApuracao("DESCR_ITEM") - i)
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Tributação PIS/COFINS
    COD_NCM_TRIB = CamposTrib(dicTitulosTributacao("COD_NCM") - i)
    
    Select Case True
        
        Case COD_NCM <> COD_NCM_TRIB
            INCONSISTENCIA = "COD_NCM divergente: " & COD_NCM & " (informado) vs " & COD_NCM_TRIB & " (cadastrado) para o item: " & DESCR_ITEM
            SUGESTAO = "Aplicar o NCM cadastrado na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_CEST(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
Dim CEST As String, CEST_TRIB$, DESCR_ITEM$
Dim INCONSISTENCIA As String, SUGESTAO$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração PIS/COFINS
    CEST = Campos(dicTitulosApuracao("CEST") - i)
    DESCR_ITEM = Campos(dicTitulosApuracao("DESCR_ITEM") - i)
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Tributação PIS/COFINS
    CEST_TRIB = CamposTrib(dicTitulosTributacao("CEST") - i)
    
    Select Case True
        
        Case CEST <> CEST_TRIB
            INCONSISTENCIA = "CEST divergente: " & CEST & " (informado) vs " & CEST_TRIB & " (cadastrado) para o item: " & DESCR_ITEM
            SUGESTAO = "Aplicar o CEST cadastrado na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_EX_IPI(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
Dim EX_IPI As String, EX_IPI_TRIB$
Dim INCONSISTENCIA As String, SUGESTAO$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração PIS/COFINS
    EX_IPI = Campos(dicTitulosApuracao("EX_IPI") - i)
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Tributação PIS/COFINS
    EX_IPI_TRIB = CamposTrib(dicTitulosTributacao("EX_IPI") - i)
    
    Select Case True
        
        Case EX_IPI <> EX_IPI_TRIB
            INCONSISTENCIA = "Campo EX_IPI (" & EX_IPI & ") divergente do cadastrado na tributação (" & EX_IPI_TRIB & ")"
            SUGESTAO = "Aplicar a EX_IPI cadastrada na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_COD_BARRA(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
Dim COD_BARRA As String, COD_BARRA_TRIB$
Dim INCONSISTENCIA As String, SUGESTAO$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração PIS/COFINS
    COD_BARRA = Campos(dicTitulosApuracao("COD_BARRA") - i)
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Tributação PIS/COFINS
    COD_BARRA_TRIB = CamposTrib(dicTitulosTributacao("COD_BARRA") - i)
    
    Select Case True
        
        Case COD_BARRA <> COD_BARRA_TRIB
            INCONSISTENCIA = "Campo COD_BARRA (" & COD_BARRA & ") divergente do cadastrado na tributação (" & COD_BARRA_TRIB & ")"
            SUGESTAO = "Aplicar o COD_BARRA cadastrado na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_TIPO_ITEM(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
Dim TIPO_ITEM As String, TIPO_ITEM_TRIB$
Dim INCONSISTENCIA As String, SUGESTAO$, CFOP$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração PIS/COFINS
    TIPO_ITEM = Campos(dicTitulosApuracao("TIPO_ITEM") - i)
    CFOP = Campos(dicTitulosApuracao("CFOP") - i)
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Tributação PIS/COFINS
    TIPO_ITEM_TRIB = CamposTrib(dicTitulosTributacao("TIPO_ITEM") - i)
    
    Select Case True
        
        Case TIPO_ITEM <> TIPO_ITEM_TRIB
            INCONSISTENCIA = "TIPO_ITEM divergente: " & TIPO_ITEM & " (informado) vs " & TIPO_ITEM_TRIB & " (cadastrado) para a operação com CFOP " & CFOP
            SUGESTAO = "Aplicar o TIPO_ITEM cadastrado na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_IND_MOV(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
Dim IND_MOV As String, IND_MOV_TRIB$
Dim INCONSISTENCIA As String, SUGESTAO$, CFOP$
Dim i As Byte
    
    'Carregar campos do assistente de Apuração PIS/COFINS
    If LBound(Campos) = 0 Then i = 1
    IND_MOV = Util.ApenasNumeros(Campos(dicTitulosApuracao("IND_MOV") - i))
    CFOP = Campos(dicTitulosApuracao("CFOP") - i)
    
    'Carregar campos do assistente de Tributação PIS/COFINS
    If LBound(Campos) = 0 Then i = 1
    IND_MOV_TRIB = Util.ApenasNumeros(CamposTrib(dicTitulosTributacao("IND_MOV") - i))
    
    Select Case True
        
        Case IND_MOV <> IND_MOV_TRIB
            INCONSISTENCIA = "IND_MOV divergente: " & IND_MOV & " (informado) vs " & fnExcel.FormatarValores(IND_MOV_TRIB) & " (cadastrado) para a operação com CFOP " & CFOP
            SUGESTAO = "Aplicar o indicador de movimento cadastrado na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function
