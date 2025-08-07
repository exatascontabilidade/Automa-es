Attribute VB_Name = "AssistenteInventario_Validacoes"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private dtoValidacoes As DTOsClasseValidacoesInventario

Private Type DTOsClasseValidacoesInventario
    
    dicTitulosSaldoInventario As Dictionary
    arrVerificacoes As ArrayList
    CamposInventario As Variant
    dicTitulos As Dictionary
    Campos As Variant
    
End Type

Public Sub ReprocessarRegrasICMS(ByVal arrCampos As Variant)

Dim Campos As Variant
    
    Call CarregarFuncoesVerificacao
    
    For Each Campos In arrCampos
        
        'Call ValidarRegrasInventario(Campos)
        
    Next Campos
    
End Sub

Public Sub ValidarRegrasICMS(ByVal Campos As Variant, ByRef dicTitulosSaldoInventario As Dictionary)
    
Dim Verificacao As Variant
    
    With dtoValidacoes
        
        .CamposInventario = Campos
        Set .dicTitulosSaldoInventario = dicTitulosSaldoInventario
        
    End With
    
    For Each Verificacao In dtoValidacoes.arrVerificacoes
        
        CallByName Me, CStr(Verificacao), VbMethod
        
        If dtoSaldoInventario.INCONSISTENCIA <> "" Then
            
            Call RegistrarInconsistencia
            Exit Sub
            
        End If
        
    Next Verificacao
    
End Sub

Public Sub CarregarFuncoesVerificacao()
    
    With dtoValidacoes
        
        Set .arrVerificacoes = New ArrayList
        
        .arrVerificacoes.Add "VerificarSaldoInventario"
        .arrVerificacoes.Add "VerificarEntradasSaidas"
        .arrVerificacoes.Add "VerificarMargem"
        
    End With
    
End Sub

Public Sub VerificarSaldoInventario()
    
    With dtoSaldoInventario
        
        Select Case True
            
            Case .QTD_FINAL < 0
                .INCONSISTENCIA = "Produto com saldo negativo"
                .SUGESTAO = "Verifique se o saldo de estoque inicial foi informado"
                
        End Select
        
    End With
    
End Sub

Public Sub VerificarEntradasSaidas()
    
    With dtoSaldoInventario
        
        Select Case True
            
            Case .QTD_ENT = 0 And .QTD_SAI > 0
                .INCONSISTENCIA = "Produto com saída sem nenhuma entrada"
                .SUGESTAO = "Verifique se as entradas estão com o código do contribuinte"
                
            Case .QTD_SAI = 0 And .QTD_ENT > 0
                .INCONSISTENCIA = "Produto com entradas sem nenhuma saída"
                .SUGESTAO = "Importe os XMLS de saída para apurar saldo corretamente"
                
        End Select
        
    End With
    
End Sub

Public Sub VerificarMargem()
    
    With dtoSaldoInventario
        
        Select Case True
            
            Case .ALIQ_MARGEM < 0
                .INCONSISTENCIA = "Produto com margem negativa"
                .SUGESTAO = "Investigar causas"
                
        End Select
        
    End With
    
End Sub

Private Sub RegistrarInconsistencia()
    
    With dtoValidacoes
        
        .CamposInventario(.dicTitulosSaldoInventario("INCONSISTENCIA")) = dtoSaldoInventario.INCONSISTENCIA
        .CamposInventario(.dicTitulosSaldoInventario("SUGESTAO")) = dtoSaldoInventario.SUGESTAO
        
    End With
    
End Sub

Public Function ResetarDTOs()

Dim dtoVazio As DTOsClasseValidacoesInventario
    
    LSet dtoValidacoes = dtoVazio
    
End Function
