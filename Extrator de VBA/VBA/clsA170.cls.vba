Attribute VB_Name = "clsA170"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function AtualizarImpostosA100(Optional ByVal OmitirMsg As Boolean)
    
Dim Dados As Range, Linha As Range
Dim Campos As Variant, Campo, nCampo, Titulos
Dim dicA100 As New Dictionary
Dim dicA170 As New Dictionary
Dim dicTitulos As New Dictionary
Dim dicTitulosA170 As New Dictionary
Dim Chave As Variant, Valores
Dim CHV_PAI As String
Dim Inicio As Date
    
    Inicio = Now()
    Application.StatusBar = "Atualizando os valores dos Valores no registro A100, por favor aguarde..."
    Valores = Array("VL_ITEM", "VL_DESC", "VL_BC_PIS", "VL_PIS", "VL_BC_COFINS", "VL_COFINS")
    If regA170.AutoFilterMode Then regA170.AutoFilter.ShowAllData
    Set dicTitulosA170 = Util.MapearTitulos(regA170, 3)
    Set Dados = Util.DefinirIntervalo(regA170, 4, 3)
    If Dados Is Nothing Then Exit Function
            
    'Carrega dados do A170 e soma os valores dos campos selecionados
    For Each Linha In Dados.Rows
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
        
            CHV_PAI = Campos(dicTitulosA170("CHV_PAI_FISCAL"))
            
            If dicA170.Exists(CHV_PAI) Then
            
                'Soma valores valores do A170 para registros com a mesma chave
                For Each Chave In dicTitulosA170.Keys()
                    
                    If Chave Like "VL_*" Then
                        If Campos(dicTitulosA170(Chave)) = "" Then Campos(dicTitulosA170(Chave)) = 0
                        Campos(dicTitulosA170(Chave)) = dicA170(CHV_PAI)(dicTitulosA170(Chave)) + CDbl(Campos(dicTitulosA170(Chave)))
                    End If
                    
                Next Chave
                                    
            End If
            
            dicA170(CHV_PAI) = Campos
        
        End If
        
    Next Linha
    
    'Atualiza os valores dos Valores no registro A100
    If regA100.AutoFilterMode Then regA100.AutoFilter.ShowAllData
    Set dicTitulos = Util.MapearTitulos(regA100, 3)
    Set Dados = Util.DefinirIntervalo(regA100, 4, 3)
    
    For Each Linha In Dados.Rows
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            Chave = Campos(dicTitulos("CHV_REG"))
            If dicA170.Exists(Chave) Then
                
                For Each nCampo In Valores
                    If nCampo = "VL_ITEM" Then
                        Campos(dicTitulos("VL_DOC")) = CDbl(dicA170(Chave)(dicTitulosA170("VL_ITEM")))
                    Else
                        Campos(dicTitulos(nCampo)) = CDbl(dicA170(Chave)(dicTitulosA170(nCampo)))
                    End If
                    
                Next nCampo
            
            End If
            
            dicA100(Chave) = Campos
        
        End If
        
    Next Linha
    
    Call Util.LimparDados(regA100, 4, False)
    Call Util.ExportarDadosDicionario(regA100, dicA100)
    
    Call dicA170.RemoveAll
    Call dicA100.RemoveAll
    
    Application.StatusBar = False
    If Not OmitirMsg Then Call Util.MsgInformativa("Valores atualizados com sucesso!", "Atualização de Valores do A100", Inicio)
              
End Function
