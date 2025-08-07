Attribute VB_Name = "clsC190"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function AtualizarImpostosC100(Optional ByVal OmitirMsg As Boolean)
    
Dim Dados As Range, Linha As Range
Dim Campos As Variant, Campo, nCampo, Titulos
Dim dicC100 As New Dictionary
Dim dicC190 As New Dictionary
Dim dicTitulos As New Dictionary
Dim dicTitulosC190 As New Dictionary
Dim Chave As Variant
Dim Impostos As Variant
Dim Inicio As Date
    
    If Not OmitirMsg Then Inicio = Now()
    
    If Util.ChecarAusenciaDados(regC190, OmitirMsg) Then Exit Function
    Call Util.AtualizarBarraStatus("Atualizando os valores dos impostos no registro C100, por favor aguarde...")
    
    Impostos = Array("VL_BC_ICMS", "VL_BC_ICMS_ST", "VL_ICMS", "VL_ICMS_ST", "VL_IPI")
    If regC190.AutoFilterMode Then regC190.AutoFilter.ShowAllData
    Set dicTitulosC190 = Util.MapearTitulos(regC190, 3)
    Set Dados = Util.DefinirIntervalo(regC190, 4, 3)
    
    With CamposC190
        
        'Carrega dados do C190 e soma os valores dos campos selecionados
        For Each Linha In Dados.Rows
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                .CHV_PAI = Campos(dicTitulosC190("CHV_PAI_FISCAL"))
                
                If dicC190.Exists(.CHV_PAI) Then
                
                    'Soma valores valores do C190 para registros com a mesma chave
                    For Each Chave In dicTitulosC190.Keys()
                        
                        If Chave Like "VL_*" And Chave <> "VL_OPR" And Chave <> "VL_RED_BC" Then
                            Campos(dicTitulosC190(Chave)) = dicC190(.CHV_PAI)(dicTitulosC190(Chave)) + CDbl(Campos(dicTitulosC190(Chave)))
                        End If
                        
                    Next Chave
                                        
                End If
                
                'Call fnExcel.DefinirTipoCampos(Campos, dicTitulosC190)
                dicC190(.CHV_PAI) = Campos
            
            End If
            
        Next Linha
    
    End With
    
    With CamposC100
        
        'Atualiza os valores dos impostos no registro C100
        If regC100.AutoFilterMode Then regC100.AutoFilter.ShowAllData
        Set dicTitulos = Util.MapearTitulos(regC100, 3)
        Set Dados = Util.DefinirIntervalo(regC100, 4, 3)
                
        For Each Linha In Dados.Rows
        
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                .CHV_REG = Campos(dicTitulos("CHV_REG"))
                If dicC190.Exists(.CHV_REG) Then
                    Chave = dicC190(.CHV_REG)
                    For Each nCampo In Impostos
                        Campos(dicTitulos(nCampo)) = CDbl(dicC190(.CHV_REG)(dicTitulosC190(nCampo)))
                    Next nCampo
                
                End If
                
                'Call fnExcel.DefinirTipoCampos(Campos, dicTitulos)
                
                dicC100(.CHV_REG) = Campos
            
            End If
            
        Next Linha
    
    End With
    
    Call Util.LimparDados(regC100, 4, False)
    Call Util.ExportarDadosDicionario(regC100, dicC100)
    
    Call dicC190.RemoveAll
    Call dicC100.RemoveAll
    
    Application.StatusBar = False
    If Not OmitirMsg Then Call Util.MsgInformativa("Impostos atualizados com sucesso!", "Atualização de impostos do C100", Inicio)
              
End Function

Public Function AgruparRegistros(Optional ByVal OmitirMsg As Boolean = False)
    
Dim Dados As Range, Linha As Range
Dim Campos As Variant, Campo, nCampo, Titulos
Dim dicC190 As New Dictionary
Dim dicTitulos As New Dictionary
Dim Chave As Variant
Dim Valores As Variant
Dim Inicio As Date
    
    If Not OmitirMsg Then Inicio = Now()
    
    If Util.ChecarAusenciaDados(regC190, OmitirMsg) Then Exit Function
    Call Util.AtualizarBarraStatus("Agrupando dados do registro C190, por favor aguarde...")
    
    Valores = Array("VL_OPR", "VL_BC_ICMS", "VL_BC_ICMS_ST", "VL_ICMS", "VL_ICMS_ST", "VL_RED_BC", "VL_IPI")
    
    If regC190.AutoFilterMode Then regC190.AutoFilter.ShowAllData
    Set dicTitulos = Util.MapearTitulos(regC190, 3)
    Set Dados = Util.DefinirIntervalo(regC190, 4, 3)
    
    With CamposC190
        
        'Carrega dados do C190 e soma os valores dos campos selecionados
        For Each Linha In Dados.Rows
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                .CHV_REG = Campos(dicTitulos("CHV_REG"))
                .CHV_PAI = Campos(dicTitulos("CHV_PAI_FISCAL"))
                .CFOP = Campos(dicTitulos("CFOP"))
                .CST_ICMS = Campos(dicTitulos("CST_ICMS"))
                .ALIQ_ICMS = Campos(dicTitulos("ALIQ_ICMS"))
                
                .COD_ENFOQUE = fnSPED.GerarChaveRegistro(.CHV_PAI, .CFOP, .CST_ICMS, .ALIQ_ICMS)
                If dicC190.Exists(.COD_ENFOQUE) Then
                    
                    'Soma valores valores do C190 para registros com a mesma chave
                    For Each Chave In dicTitulos.Keys()
                        
                        If Chave Like "VL_*" Then
                            Campos(dicTitulos(Chave)) = dicC190(.COD_ENFOQUE)(dicTitulos(Chave)) + CDbl(Campos(dicTitulos(Chave)))
                        End If
                        
                    Next Chave
                    
                End If
                
                'Call fnExcel.DefinirTipoCampos(Campos, dicTitulos)
                dicC190(.COD_ENFOQUE) = Campos
            
            End If
            
        Next Linha
        
    End With
    
    'Atualiza os dados do registro C190
    Call Util.LimparDados(regC190, 4, False)
    Call Util.ExportarDadosDicionario(regC190, dicC190)
    
    Call dicC190.RemoveAll
    
    Application.StatusBar = False
    
    If OmitirMsg Then Exit Function
    Call Util.MsgInformativa("Valores atualizados com sucesso!", "Agrupamento de registro do C190", Inicio)
    
End Function

Public Function CalcularReducaoBaseICMS()

Dim Dados As Range, Linha As Range
Dim Campos As Variant, Campo, nCampo, Titulos
Dim arrC190 As New ArrayList
Dim dicTitulos As New Dictionary
Dim Chave As String
Dim CST_ICMS As String
Dim Inicio As Date
    
    Inicio = Now()
    
    If Util.ChecarAusenciaDados(regC190, False) Then Exit Function
    Call Util.AtualizarBarraStatus("Calculando a redução de base do ICMS para o registro C190, por favor aguarde...")
    
    If regC190.AutoFilterMode Then regC190.AutoFilter.ShowAllData
    Set dicTitulos = Util.MapearTitulos(regC190, 3)
    Set Dados = Util.DefinirIntervalo(regC190, 4, 3)
    
    'Carrega dados do C190 e calcula o valor da redução de base para os registros com CST_ICMS terminando com 20 ou 70
    For Each Linha In Dados.Rows
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            CST_ICMS = Campos(dicTitulos("CST_ICMS"))
            If CST_ICMS Like "*20" Or CST_ICMS Like "*70" Then
                Campos(dicTitulos("VL_RED_BC")) = VBA.Round(Campos(dicTitulos("VL_OPR")) - Campos(dicTitulos("VL_BC_ICMS")) - Campos(dicTitulos("VL_ICMS_ST")) - Campos(dicTitulos("VL_IPI")), 2)
            Else
                Campos(dicTitulos("VL_RED_BC")) = 0
            End If
            If Campos(dicTitulos("VL_RED_BC")) < 0 Then Campos(dicTitulos("VL_RED_BC")) = 0
            
            'Call fnExcel.DefinirTipoCampos(Campos, dicTitulos)
            arrC190.Add Campos
        
        End If
        
    Next Linha
    
    'Atualiza os dados do registro C190
    Call Util.LimparDados(regC190, 4, False)
    Call Util.ExportarDadosArrayList(regC190, arrC190)
    
    Call arrC190.Clear
    
    Application.StatusBar = False
    Call Util.MsgInformativa("Redução de base calculada com sucesso!", "Redução de base do ICMS no C190", Inicio)
    
End Function
