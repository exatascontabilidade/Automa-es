Attribute VB_Name = "clsC185"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function AgruparRegistros(Optional ByVal OmitirMsg As Boolean)

Dim Campos As Variant, Campo, nCampo, Titulos
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim dicC185 As New Dictionary
Dim COD_ENFOQUE As String
Dim Chave As Variant
Dim Inicio As Date
    
    Inicio = Now()
    
    If Util.ChecarAusenciaDados(regC185_Contr, False) Then Exit Function
    Call Util.AtualizarBarraStatus("Agrupando dados do registro C185, por favor aguarde...")
    
    If regC185_Contr.AutoFilterMode Then regC185_Contr.AutoFilter.ShowAllData
    Set dicTitulos = Util.MapearTitulos(regC185_Contr, 3)
    Set Dados = Util.DefinirIntervalo(regC185_Contr, 4, 3)
    
    With CamposC185Contr
        
        For Each Linha In Dados.Rows
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                .CHV_REG = Campos(dicTitulos("CHV_REG"))
                .CHV_PAI = Campos(dicTitulos("CHV_PAI_FISCAL"))
                .CFOP = Campos(dicTitulos("CFOP"))
                .CST_COFINS = Campos(dicTitulos("CST_COFINS"))
                .ALIQ_COFINS = Campos(dicTitulos("ALIQ_COFINS"))
                
                COD_ENFOQUE = fnSPED.GerarChaveRegistro(.CHV_PAI, .CFOP, .CST_COFINS, .ALIQ_COFINS)
                If dicC185.Exists(COD_ENFOQUE) Then
                    
                    For Each Chave In dicTitulos.Keys()
                        
                        If Chave Like "VL_*" Then
                            Campos(dicTitulos(Chave)) = dicC185(COD_ENFOQUE)(dicTitulos(Chave)) + CDbl(Campos(dicTitulos(Chave)))
                        End If
                        
                    Next Chave
                    
                End If
                
                dicC185(COD_ENFOQUE) = Campos
                
            End If
            
        Next Linha
        
    End With
    
    Call Util.LimparDados(regC185_Contr, 4, False)
    Call Util.ExportarDadosDicionario(regC185_Contr, dicC185)
    
    Call dicC185.RemoveAll
    
    Application.StatusBar = False
    
    If OmitirMsg Then Exit Function
    Call Util.MsgInformativa("Valores atualizados com sucesso!", "Agrupamento de registro do C185", Inicio)
    
End Function

