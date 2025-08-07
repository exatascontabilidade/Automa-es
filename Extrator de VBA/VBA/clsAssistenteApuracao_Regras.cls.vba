Attribute VB_Name = "clsAssistenteApuracao_Regras"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Sub ExecutarValidacoesGerais(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
    
    Call VerificarCampo_COD_SIT(Campos, dicTitulos)
    
End Sub

Private Sub VerificarCampo_COD_SIT(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim COD_SIT As String, SER$, INCONSISTENCIA$, SUGESTAO$
Dim vCOD_SIT As Boolean, tCOD_SIT As Boolean
Dim i As Integer
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    COD_SIT = Util.ApenasNumeros(Campos(dicTitulos("COD_SIT") - i))
    SER = Util.ApenasNumeros(Campos(dicTitulos("SER") - i))
    
    If SER = "890" And Not COD_SIT Like "08*" Then
        
        INCONSISTENCIA = "Nota Fiscal Avulsa (SER = 890) com campo COD_SIT diferente de 08 - Regime Especial ou Norma Específica"
        SUGESTAO = "Alterar valor do campo COD_SIT para: 08 - Regime Especial ou Norma Específica"
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Sub

