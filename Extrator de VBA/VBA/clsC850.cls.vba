Attribute VB_Name = "clsC850"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function AtualizarImpostosC800()
    
Dim Dados As Range, Linha As Range
Dim Campos As Variant, Campo, nCampo, Titulos
Dim dicC800 As New Dictionary
Dim dicC850 As New Dictionary
Dim dicTitulos As New Dictionary
Dim Chave As String
Dim Impostos As Variant
Dim Inicio As Date
    
    Inicio = Now()
    Application.StatusBar = "Atualizando os valores dos impostos no registro C800, por favor aguarde..."
    Impostos = Array("VL_ICMS")
    If regC850.AutoFilterMode Then regC850.AutoFilter.ShowAllData
    Set dicTitulos = Util.MapearTitulos(regC850, 3)
    Set Dados = Util.DefinirIntervalo(regC850, 4, 3)
    
    'Carrega dados do C850 e soma os valores dos impostos
    For Each Linha In Dados.Rows
        Campos = Application.index(Linha.Value2, 0, 0)
        Chave = Campos(dicTitulos("CHV_PAI_FISCAL"))
        For Each nCampo In Impostos
            Campo = Campos(dicTitulos(nCampo))
            If Campo = "" Then Campo = 0
                        
            If Not dicC850.Exists(Chave) Then
                Set dicC850(Chave) = New Dictionary
                If IsNumeric(Campo) Then dicC850(Chave)(nCampo) = Campo
            Else
                If IsNumeric(Campo) Then dicC850(Chave)(nCampo) = CDbl(dicC850(Chave)(nCampo)) + CDbl(Campo)
            End If
            
        Next nCampo
        
    Next Linha
    
    'Atualiza os valores dos impostos no registro C800
    If regC800.AutoFilterMode Then regC800.AutoFilter.ShowAllData
    Set dicTitulos = Util.MapearTitulos(regC800, 3)
    Set Dados = Util.DefinirIntervalo(regC800, 4, 3)
    
    For Each Linha In Dados.Rows
        Campos = Application.index(Linha.Value2, 0, 0)
        
        Chave = Campos(dicTitulos("CHV_REG"))
        If dicC850.Exists(Chave) Then
            
            If Campos(dicTitulos("CHV_CFE")) <> "" Then Campos(dicTitulos("CHV_CFE")) = "'" & Campos(dicTitulos("CHV_CFE"))
            If Campos(dicTitulos("CNPJ_CPF")) <> "" Then Campos(dicTitulos("CNPJ_CPF")) = "'" & Campos(dicTitulos("CNPJ_CPF"))
            
            For Each nCampo In Impostos
                Campos(dicTitulos(nCampo)) = CDbl(dicC850(Chave)(nCampo))
            Next nCampo
        
        End If
        
        dicC800(Chave) = Campos
        
    Next Linha
    
    Call Util.LimparDados(regC800, 4, False)
    Call Util.ExportarDadosDicionario(regC800, dicC800)
    
    Call dicC850.RemoveAll
    Call dicC800.RemoveAll
    
    Application.StatusBar = False
    Call Util.MsgInformativa("Impostos atualizados com sucesso!", "Atualização de impostos do C800", Inicio)
              
End Function

Public Function AgruparRegistros(Optional ByVal OmitirMsg As Boolean)
    
Dim Campos As Variant, Campo, nCampo, Titulos
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim dicC850 As New Dictionary
Dim Chave As Variant
Dim Inicio As Date
    
    Inicio = Now()
    
    If Util.ChecarAusenciaDados(regC850, False) Then Exit Function
    Call Util.AtualizarBarraStatus("Agrupando dados do registro C850, por favor aguarde...")
    
    If regC850.AutoFilterMode Then regC850.AutoFilter.ShowAllData
    Set dicTitulos = Util.MapearTitulos(regC850, 3)
    Set Dados = Util.DefinirIntervalo(regC850, 4, 3)
    
    With CamposC850
        
        For Each Linha In Dados.Rows
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                .CHV_REG = Campos(dicTitulos("CHV_REG"))
                .CHV_PAI = Campos(dicTitulos("CHV_PAI_FISCAL"))
                .CFOP = Campos(dicTitulos("CFOP"))
                .CST_ICMS = Campos(dicTitulos("CST_ICMS"))
                .ALIQ_ICMS = Campos(dicTitulos("ALIQ_ICMS"))
                
                .COD_ENFOQUE = fnSPED.GerarChaveRegistro(.CHV_PAI, .CFOP, .CST_ICMS, .ALIQ_ICMS)
                If dicC850.Exists(.COD_ENFOQUE) Then
                    
                    For Each Chave In dicTitulos.Keys()
                        
                        If Chave Like "VL_*" Then
                            Campos(dicTitulos(Chave)) = dicC850(.COD_ENFOQUE)(dicTitulos(Chave)) + CDbl(Campos(dicTitulos(Chave)))
                        End If
                        
                    Next Chave
                    
                End If
                
                dicC850(.COD_ENFOQUE) = Campos
                
            End If
            
        Next Linha
        
    End With
    
    Call Util.LimparDados(regC850, 4, False)
    Call Util.ExportarDadosDicionario(regC850, dicC850)
    
    Call dicC850.RemoveAll
    
    Application.StatusBar = False
    
    If OmitirMsg Then Exit Function
    Call Util.MsgInformativa("Valores atualizados com sucesso!", "Agrupamento de registro do C850", Inicio)
    
End Function
