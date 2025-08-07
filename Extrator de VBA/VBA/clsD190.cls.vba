Attribute VB_Name = "clsD190"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function AgruparRegistros()
    
Dim Dados As Range, Linha As Range
Dim Campos As Variant, Campo, nCampo, Titulos
Dim dicD190 As New Dictionary
Dim dicTitulos As New Dictionary
Dim Chave As Variant
Dim Valores As Variant
Dim Inicio As Date
    
    Inicio = Now()
    
    If Util.ChecarAusenciaDados(regD190, False) Then Exit Function
    Call Util.AtualizarBarraStatus("Agrupando dados do registro D190, por favor aguarde...")
    
    Valores = Array("VL_OPR", "VL_BC_ICMS", "VL_BC_ICMS_ST", "VL_ICMS", "VL_ICMS_ST", "VL_RED_BC", "VL_IPI")
    
    If regD190.AutoFilterMode Then regD190.AutoFilter.ShowAllData
    Set dicTitulos = Util.MapearTitulos(regD190, 3)
    Set Dados = Util.DefinirIntervalo(regD190, 4, 3)
    
    With CamposD190
        
        'Carrega dados do D190 e soma os valores dos campos selecionados
        For Each Linha In Dados.Rows
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                .CHV_REG = Campos(dicTitulos("CHV_REG"))
                .CHV_PAI = Campos(dicTitulos("CHV_PAI_FISCAL"))
                .CFOP = Campos(dicTitulos("CFOP"))
                .CST_ICMS = Campos(dicTitulos("CST_ICMS"))
                .ALIQ_ICMS = Campos(dicTitulos("ALIQ_ICMS"))
                
                .COD_ENFOQUE = fnSPED.GerarChaveRegistro(.CHV_PAI, .CFOP, .CST_ICMS, .ALIQ_ICMS)
                If dicD190.Exists(.COD_ENFOQUE) Then
                    
                    'Soma valores valores do D190 para registros com a mesma chave
                    For Each Chave In dicTitulos.Keys()
                        
                        If Chave Like "VL_*" Then
                            Campos(dicTitulos(Chave)) = dicD190(.COD_ENFOQUE)(dicTitulos(Chave)) + CDbl(Campos(dicTitulos(Chave)))
                        End If
                        
                    Next Chave
                    
                End If
                
                'Call fnExcel.DefinirTipoCampos(Campos, dicTitulos)
                dicD190(.COD_ENFOQUE) = Campos
            
            End If
            
        Next Linha
        
    End With
    
    'Atualiza os dados do registro D190
    Call Util.LimparDados(regD190, 4, False)
    Call Util.ExportarDadosDicionario(regD190, dicD190)
    
    Call dicD190.RemoveAll
    
    Application.StatusBar = False
    Call Util.MsgInformativa("Valores atualizados com sucesso!", "Agrupamento de registro do D190", Inicio)
    
End Function

Public Function AtualizarImpostosD100(Optional ByVal OmitirMsg As Boolean)
    
Dim Dados As Range, Linha As Range
Dim Campos As Variant, Campo, nCampo, Titulos
Dim dicD100 As New Dictionary
Dim dicD190 As New Dictionary
Dim dicTitulos As New Dictionary
Dim dicTitulosD190 As New Dictionary
Dim Chave As Variant
Dim Impostos As Variant
Dim Inicio As Date
    
    Inicio = Now()
    
    If Util.ChecarAusenciaDados(regD190, False) Then Exit Function
    Call Util.AtualizarBarraStatus("Atualizando os valores dos impostos no registro D100, por favor aguarde...")
    
    Impostos = Array("VL_BC_ICMS", "VL_ICMS")
    If regD190.AutoFilterMode Then regD190.AutoFilter.ShowAllData
    Set dicTitulosD190 = Util.MapearTitulos(regD190, 3)
    Set Dados = Util.DefinirIntervalo(regD190, 4, 3)
    
    With CamposD190
        
        'Carrega dados do D190 e soma os valores dos campos selecionados
        For Each Linha In Dados.Rows
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                .CHV_PAI = Campos(dicTitulosD190("CHV_PAI_FISCAL"))
                
                If dicD190.Exists(.CHV_PAI) Then
                
                    'Soma valores valores do D190 para registros com a mesma chave
                    For Each Chave In dicTitulosD190.Keys()
                        
                        If Chave Like "VL_*" And Chave <> "VL_OPR" And Chave <> "VL_RED_BC" Then
                            Campos(dicTitulosD190(Chave)) = dicD190(.CHV_PAI)(dicTitulosD190(Chave)) + CDbl(Campos(dicTitulosD190(Chave)))
                        End If
                        
                    Next Chave
                                        
                End If
                
                'Call fnExcel.DefinirTipoCampos(Campos, dicTitulosD190)
                dicD190(.CHV_PAI) = Campos
            
            End If
            
        Next Linha
    
    End With
    
    With CamposD100
        
        'Atualiza os valores dos impostos no registro D100
        If regD100.AutoFilterMode Then regD100.AutoFilter.ShowAllData
        Set dicTitulos = Util.MapearTitulos(regD100, 3)
        Set Dados = Util.DefinirIntervalo(regD100, 4, 3)
                
        For Each Linha In Dados.Rows
        
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                .CHV_REG = Campos(dicTitulos("CHV_REG"))
                If dicD190.Exists(.CHV_REG) Then
                    Chave = dicD190(.CHV_REG)
                    For Each nCampo In Impostos
                        Campos(dicTitulos(nCampo)) = CDbl(dicD190(.CHV_REG)(dicTitulosD190(nCampo)))
                    Next nCampo
                
                End If
                                
                dicD100(.CHV_REG) = Campos
            
            End If
            
        Next Linha
    
    End With
    
    Call Util.LimparDados(regD100, 4, False)
    Call Util.ExportarDadosDicionario(regD100, dicD100)
    
    Call dicD190.RemoveAll
    Call dicD100.RemoveAll
    
    Application.StatusBar = False
    If Not OmitirMsg Then Call Util.MsgInformativa("Impostos atualizados com sucesso!", "Atualização de impostos do D100", Inicio)
              
End Function

Public Function CalcularReducaoBaseICMS()

Dim Dados As Range, Linha As Range
Dim Campos As Variant, Campo, nCampo, Titulos
Dim arrD190 As New ArrayList
Dim dicTitulos As New Dictionary
Dim Chave As String
Dim CST_ICMS As String
Dim Inicio As Date
    
    Inicio = Now()
    
    If Util.ChecarAusenciaDados(regD190, False) Then Exit Function
    Call Util.AtualizarBarraStatus("Calculando a redução de base do ICMS para o registro D190, por favor aguarde...")
    
    If regD190.AutoFilterMode Then regD190.AutoFilter.ShowAllData
    Set dicTitulos = Util.MapearTitulos(regD190, 3)
    Set Dados = Util.DefinirIntervalo(regD190, 4, 3)
    
    'Carrega dados do D190 e calcula o valor da redução de base para os registros com CST_ICMS terminando com 20 ou 70
    For Each Linha In Dados.Rows
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            CST_ICMS = Campos(dicTitulos("CST_ICMS"))
            If CST_ICMS Like "*20" Or CST_ICMS Like "*70" Then
                Campos(dicTitulos("VL_RED_BC")) = VBA.Round(Campos(dicTitulos("VL_OPR")) - Campos(dicTitulos("VL_BC_ICMS")), 2)
            Else
                Campos(dicTitulos("VL_RED_BC")) = 0
            End If
            If Campos(dicTitulos("VL_RED_BC")) < 0 Then Campos(dicTitulos("VL_RED_BC")) = 0
            
            Call fnExcel.DefinirTipoCampos(Campos, dicTitulos)
            arrD190.Add Campos
        
        End If
        
    Next Linha
    
    'Atualiza os dados do registro D190
    Call Util.LimparDados(regD190, 4, False)
    Call Util.ExportarDadosArrayList(regD190, arrD190)
    
    Call arrD190.Clear
    
    Application.StatusBar = False
    Call Util.MsgInformativa("Redução de base calculada com sucesso!", "Redução de base do ICMS no D190", Inicio)
    
End Function
