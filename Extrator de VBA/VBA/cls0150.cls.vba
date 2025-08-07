Attribute VB_Name = "cls0150"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function ImportarParaAnalise(ByVal Registro As String, ByRef dicDados As Dictionary)

Dim Campos

    Campos = fnSPED.ExtrairCamposRegistro(Registro)
    With Campos0150

        .REG = Util.FormatarTexto(Campos(1))
        .COD_PART = Util.FormatarTexto(Campos(2))
        .NOME = Util.FormatarTexto(Campos(3))
        .COD_PAIS = Util.FormatarTexto(Campos(4))
        .CNPJ = Util.FormatarTexto(Campos(5))
        .CPF = Util.FormatarTexto(Campos(6))
        .IE = Util.FormatarTexto(Campos(7))
        .COD_MUN = Util.FormatarTexto(Campos(8))
        .SUFRAMA = Util.FormatarTexto(Campos(9))
        .END = Util.FormatarTexto(Campos(10))
        .NUM = Util.FormatarTexto(Campos(11))
        .COMPL = Util.FormatarTexto(Campos(12))
        .BAIRRO = Util.FormatarTexto(Campos(13))

        .CHV_REG = Util.RemoverAspaSimples(.COD_PART)
        dicDados(.CHV_REG) = Array(.REG, Campos0000.ARQUIVO, .COD_PART, .NOME, .COD_PAIS, _
                                   .CNPJ, .CPF, .IE, .COD_MUN, .SUFRAMA, .END, .NUM, .COMPL, .BAIRRO)

    End With

End Function
