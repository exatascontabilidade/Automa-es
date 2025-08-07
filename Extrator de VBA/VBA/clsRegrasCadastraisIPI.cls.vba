Attribute VB_Name = "clsRegrasCadastraisIPI"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function ValidarCampo_CST_IPI(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim CST_IPI$, CFOP$, INCONSISTENCIA$, SUGESTAO$
Dim ALIQ_IPI As Double
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do relatório de tributação IPI
    CST_IPI = Util.ApenasNumeros(Campos(dicTitulos("CST_IPI") - i))
    ALIQ_IPI = Util.ApenasNumeros(Campos(dicTitulos("ALIQ_IPI") - i))
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    
    If INCONSISTENCIA = "" Then
        
        Select Case True
            
            Case (CFOP Like "#406" Or CFOP Like "#551") And Not CST_IPI Like "49*"
                INCONSISTENCIA = "CST_IPI incompatível com operação de aquisição de ativo imobilizado"
                SUGESTAO = "Informar CST_IPI " & 49 & "para a operação"
                
            Case (CFOP Like "#407" Or CFOP Like "#556") And Not CST_IPI Like "49*"
                INCONSISTENCIA = "CST_IPI incompatível com operação de aquisição para uso e consumo"
                SUGESTAO = "Informar CST_IPI " & 49 & "para a operação"
                
            Case CFOP > 4000 And CST_IPI < 50
                INCONSISTENCIA = "CST_IPI de entrada informado em operação de saída"
                SUGESTAO = "Informar CST_IPI de entrada"
                
            Case CFOP < 4000 And CST_IPI > 49
                INCONSISTENCIA = "CST_IPI de saída informado em operação de entrada"
                SUGESTAO = "Informar CST_IPI de saída"
                
        End Select
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO)
    
End Function

Public Function ValidarCampo_ALIQ_IPI(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim CST_IPI$, CFOP$, INCONSISTENCIA$, SUGESTAO$
Dim fatCFOP As Boolean, devCompCFOP As Boolean
Dim ALIQ_IPI As Double
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do relatório inteligente de COFINS/COFINS
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    CST_IPI = VBA.Format(Util.ApenasNumeros(Campos(dicTitulos("CST_IPI") - i)), "00")
    ALIQ_IPI = fnExcel.FormatarPercentuais(Campos(dicTitulos("ALIQ_IPI") - i))
    
    If CFOP <> "" Then
        
        Select Case True
            
            Case (CFOP Like "#406" Or CFOP Like "#551") And ALIQ_IPI > 0
                INCONSISTENCIA = "Operação de aquisição de ativo imobilizado com ALIQ_IPI maior que 0"
                SUGESTAO = "Zerar campo ALIQ_IPI"
                
            Case (CFOP Like "#407" Or CFOP Like "#556") And ALIQ_IPI > 0
                INCONSISTENCIA = "Operação de aquisição de uso e consumo com ALIQ_IPI maior que 0"
                SUGESTAO = "Zerar campo ALIQ_IPI"
                
            Case CST_IPI Like "#1" And ALIQ_IPI > 0
                INCONSISTENCIA = "CST IPI (" & CST_IPI & " - Operação Tributável com Alíquota Zero) com campo ALIQ_IPI maior que 0"
                SUGESTAO = "Zerar campo ALIQ_IPI"
                
            Case CST_IPI Like "#2" And ALIQ_IPI > 0
                INCONSISTENCIA = "CST IPI (" & CST_IPI & " - Operação Isenta) com campo ALIQ_IPI maior que 0"
                SUGESTAO = "Zerar campo ALIQ_IPI"
                
            Case CST_IPI Like "#3" And ALIQ_IPI > 0
                INCONSISTENCIA = "CST IPI (" & CST_IPI & " - Operação Não-Tributada) com campo ALIQ_IPI maior que 0"
                SUGESTAO = "Zerar campo ALIQ_IPI"
                
            Case CST_IPI Like "#4" And ALIQ_IPI > 0
                INCONSISTENCIA = "CST IPI (" & CST_IPI & " - Operação Imune) com campo ALIQ_IPI maior que 0"
                SUGESTAO = "Zerar campo ALIQ_IPI"
                
            Case CST_IPI Like "#5" And ALIQ_IPI > 0
                INCONSISTENCIA = "CST IPI (" & CST_IPI & " - Operação com Suspensão) com campo ALIQ_IPI maior que 0"
                SUGESTAO = "Zerar campo ALIQ_IPI"
                
        End Select
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO)
        
End Function
