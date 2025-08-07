Attribute VB_Name = "DTO_SPED_Contribuicoes"
Option Explicit

Public SPEDContribuicoes As RegistrosSPEDContribuicoes

Type RegistrosSPEDContribuicoes
    
    'Bloco 0
    dicDados0000 As New Dictionary
    dicTitulos0000 As New Dictionary
    
    dicDados0001 As New Dictionary
    dicTitulos0001 As New Dictionary
    
    dicDados0100 As New Dictionary
    dicTitulos0100 As New Dictionary
    
    dicDados0110 As New Dictionary
    dicTitulos0110 As New Dictionary
    
    dicDados0140 As New Dictionary
    dicTitulos0140 As New Dictionary
    
    dicDados0150 As New Dictionary
    dicTitulos0150 As New Dictionary
    
    dicDados0190 As New Dictionary
    dicTitulos0190 As New Dictionary
    
    dicDados0200 As New Dictionary
    dicTitulos0200 As New Dictionary
    
    
    'Bloco C
    dicDadosC001 As New Dictionary
    dicTitulosC001 As New Dictionary
    
    dicDadosC010 As New Dictionary
    dicTitulosC010 As New Dictionary
    
    dicDadosC100 As New Dictionary
    dicTitulosC100 As New Dictionary
    
    dicDadosC170 As New Dictionary
    dicTitulosC170 As New Dictionary
    
End Type

Public Sub CarregarDadosRegistro0000()
    
    With SPEDContribuicoes
        
        Call .dicDados0000.RemoveAll
        Call .dicTitulos0000.RemoveAll
        
        Set .dicDados0000 = Util.CriarDicionarioRegistro(reg0000_Contr, "ARQUIVO")
        Set .dicTitulos0000 = Util.MapearTitulos(reg0000_Contr, 3)
        
    End With
    
End Sub

Public Sub CarregarDadosRegistro0001()
    
    With SPEDContribuicoes
        
        Call .dicDados0001.RemoveAll
        Call .dicTitulos0001.RemoveAll
        
        Set .dicDados0001 = Util.CriarDicionarioRegistro(reg0001, "ARQUIVO")
        Set .dicTitulos0001 = Util.MapearTitulos(reg0001, 3)
        
    End With
    
End Sub

Public Sub CarregarDadosRegistro0100()
    
    With SPEDContribuicoes
        
        Call .dicDados0100.RemoveAll
        Call .dicTitulos0100.RemoveAll
        
        Set .dicDados0100 = Util.CriarDicionarioRegistro(reg0100, "ARQUIVO")
        Set .dicTitulos0100 = Util.MapearTitulos(reg0100, 3)
        
    End With
    
End Sub

Public Sub CarregarDadosRegistro0110()
    
    With SPEDContribuicoes
        
        Call .dicDados0110.RemoveAll
        Call .dicTitulos0110.RemoveAll
        
        Set .dicDados0110 = Util.CriarDicionarioRegistro(reg0110, "ARQUIVO")
        Set .dicTitulos0110 = Util.MapearTitulos(reg0110, 3)
        
    End With
    
End Sub

Public Sub CarregarDadosRegistro0140()
    
    With SPEDContribuicoes
        
        Call .dicDados0140.RemoveAll
        Call .dicTitulos0140.RemoveAll
        
        Set .dicDados0140 = Util.CriarDicionarioRegistro(reg0140, "ARQUIVO", "CNPJ")
        Set .dicTitulos0140 = Util.MapearTitulos(reg0140, 3)
        
    End With
    
End Sub

Public Sub CarregarDadosRegistro0150()
    
    With SPEDContribuicoes
        
        Call .dicDados0150.RemoveAll
        Call .dicTitulos0150.RemoveAll
        
        Set .dicDados0150 = Util.CriarDicionarioRegistro(reg0150, "ARQUIVO", "CNPJ", "CPF")
        Set .dicTitulos0150 = Util.MapearTitulos(reg0150, 3)
        
    End With
    
End Sub

Public Sub CarregarDadosRegistro0190()
    
    With SPEDContribuicoes
        
        Call .dicDados0190.RemoveAll
        Call .dicTitulos0190.RemoveAll
        
        Set .dicDados0190 = Util.CriarDicionarioRegistro(reg0190, "ARQUIVO", "UNID")
        Set .dicTitulos0190 = Util.MapearTitulos(reg0190, 3)
        
    End With
    
End Sub

Public Sub CarregarDadosRegistro0200()
    
    With SPEDContribuicoes
        
        Call .dicDados0200.RemoveAll
        Call .dicTitulos0200.RemoveAll
        
        Set .dicDados0200 = Util.CriarDicionarioRegistro(reg0200, "CHV_PAI_FISCAL", "COD_ITEM")
        Set .dicTitulos0200 = Util.MapearTitulos(reg0200, 3)
        
    End With
    
End Sub

Public Sub CarregarDadosRegistroC001()
    
    With SPEDContribuicoes
        
        Call .dicDadosC001.RemoveAll
        Call .dicTitulosC001.RemoveAll
        
        Set .dicDadosC001 = Util.CriarDicionarioRegistro(regC001, "ARQUIVO")
        Set .dicTitulosC001 = Util.MapearTitulos(regC001, 3)
        
    End With
    
End Sub

Public Sub CarregarDadosRegistroC010()
    
    With SPEDContribuicoes
        
        Call .dicDadosC010.RemoveAll
        Call .dicTitulosC010.RemoveAll
                
        Set .dicDadosC010 = Util.CriarDicionarioRegistro(regC010, "ARQUIVO", "CNPJ")
        Set .dicTitulosC010 = Util.MapearTitulos(regC010, 3)
        
    End With
    
End Sub

Public Sub CarregarDadosRegistroC100()
    
    With SPEDContribuicoes
        
        Call .dicDadosC100.RemoveAll
        Call .dicTitulosC100.RemoveAll
                
        Set .dicDadosC100 = Util.CriarDicionarioRegistro(regC100, "IND_OPER", "IND_EMIT", "CHV_NFE")
        Set .dicTitulosC100 = Util.MapearTitulos(regC100, 3)
        
    End With
    
End Sub

Public Sub CarregarDadosRegistroC170()
    
    With SPEDContribuicoes
        
        Call .dicDadosC170.RemoveAll
        Call .dicTitulosC170.RemoveAll
        
        Set .dicDadosC170 = Util.CriarDicionarioRegistro(regC170, "CHV_PAI_FISCAL", "NUM_ITEM")
        Set .dicTitulosC170 = Util.MapearTitulos(regC170, 3)
        
    End With
    
End Sub

Public Function CarregarRegistrosImportados()
    
    Call CarregarDadosRegistro0000
    Call CarregarDadosRegistro0001
    Call CarregarDadosRegistro0100
    Call CarregarDadosRegistro0110
    Call CarregarDadosRegistro0140
    Call CarregarDadosRegistro0150
    Call CarregarDadosRegistro0190
    Call CarregarDadosRegistro0200
    Call CarregarDadosRegistroC001
    Call CarregarDadosRegistroC010
    Call CarregarDadosRegistroC100
    Call CarregarDadosRegistroC170
    
End Function

Public Function ResetarDadosSPEDContribuicoes()
    
    Dim CamposVazios As RegistrosSPEDContribuicoes
    LSet SPEDContribuicoes = CamposVazios
    
End Function

