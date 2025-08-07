Attribute VB_Name = "cls0000"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function ImportarParaExcel(ByVal Registro As String, ByRef dicDados As Dictionary)

Dim Campos
    
    Campos = fnSPED.ExtrairCamposRegistro(Registro)
    With Campos0000
        
        .REG = Util.FormatarTexto(Campos(1))
        .COD_VER = Util.FormatarTexto(Campos(2))
        .COD_FIN = Util.FormatarTexto(Campos(3))
        .DT_INI = Util.FormatarData(Campos(4))
        .DT_FIN = Util.FormatarData(Campos(5))
        .NOME = Util.FormatarTexto(Campos(6))
        .CNPJ = Util.FormatarTexto(Campos(7))
        .CPF = Util.FormatarTexto(Campos(8))
        .UF = Util.FormatarTexto(Campos(9))
        .IE = Util.FormatarTexto(Campos(10))
        .COD_MUN = Util.FormatarTexto(Campos(11))
        .IM = Util.FormatarTexto(Campos(12))
        .SUFRAMA = Util.FormatarTexto(Campos(13))
        .IND_PERFIL = Util.FormatarTexto(Campos(14))
        .IND_ATIV = Util.FormatarTexto(Campos(15))
        .CHV_PAI = ""
        
        If .CNPJ <> "" Then
            .ARQUIVO = Util.FormatarTexto(VBA.Format(.DT_INI, "mm/yyyy") & "-" & VBA.Replace(.CNPJ, "'", "") & VBA.Replace(.CPF, "'", ""))
        Else
            .ARQUIVO = Util.FormatarTexto(VBA.Format(.DT_INI, "mm/yyyy") & "-" & VBA.Replace(.CPF, "'", ""))
        End If
        
        .CHV_REG = fnSPED.GerarChaveRegistro(.ARQUIVO)
        dicDados(.CHV_REG) = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI, "", .COD_VER, .COD_FIN, .DT_INI, .DT_FIN, .NOME, _
                                   .CNPJ, .CPF, .UF, .IE, .COD_MUN, .IM, .SUFRAMA, .IND_PERFIL, .IND_ATIV)
        
    End With
    
End Function

