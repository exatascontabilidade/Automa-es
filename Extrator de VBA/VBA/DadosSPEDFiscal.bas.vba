Attribute VB_Name = "DadosSPEDFiscal"
Option Explicit

Public SPEDFiscal As RegistrosSPEDFiscal

Type RegistrosSPEDFiscal
    
    'Bloco 0
    dicDados0000 As New Dictionary
    dicTitulos0000 As New Dictionary
    
    dicDados0001 As New Dictionary
    dicTitulos0001 As New Dictionary
    
    dicDados0150 As New Dictionary
    dicTitulos0150 As New Dictionary
    
    dicDados0200 As New Dictionary
    dicTitulos0200 As New Dictionary
    
    'Bloco C
    dicDadosC100 As New Dictionary
    dicTitulosC100 As New Dictionary
    
End Type

Public Sub CarregarDadosRegistro0000()
    
    With SPEDFiscal
        
        Call .dicDados0000.RemoveAll
        Call .dicTitulos0000.RemoveAll
        
        Set .dicDados0000 = Util.CriarDicionarioRegistro(reg0000, "ARQUIVO")
        Set .dicTitulos0000 = Util.MapearTitulos(reg0000, 3)
        
    End With
    
End Sub

Public Sub CarregarDadosRegistro0001()
    
    With SPEDFiscal
        
        Call .dicDados0001.RemoveAll
        Call .dicTitulos0001.RemoveAll
        
        Set .dicDados0001 = Util.CriarDicionarioRegistro(reg0001, "ARQUIVO")
        Set .dicTitulos0001 = Util.MapearTitulos(reg0001, 3)
        
    End With
    
End Sub

Public Sub CarregarDadosRegistro0150()
    
    With SPEDFiscal
        
        Call .dicDados0150.RemoveAll
        Call .dicTitulos0150.RemoveAll
        
        Set .dicDados0150 = Util.CriarDicionarioRegistro(reg0150, "ARQUIVO", "CNPJ", "CPF")
        Set .dicTitulos0150 = Util.MapearTitulos(reg0150, 3)
        
    End With
    
End Sub

Public Sub CarregarDadosRegistro0200()
    
    With SPEDFiscal
        
        Call .dicDados0200.RemoveAll
        Call .dicTitulos0200.RemoveAll
        
        Set .dicDados0200 = Util.CriarDicionarioRegistro(reg0200, "ARQUIVO", "COD_ITEM")
        Set .dicTitulos0200 = Util.MapearTitulos(reg0200, 3)
        
    End With
    
End Sub

Public Sub CarregarDadosRegistroC100()
    
    With SPEDFiscal
        
        Call .dicDadosC100.RemoveAll
        Call .dicTitulosC100.RemoveAll
                
        Set .dicDadosC100 = Util.CriarDicionarioRegistro(regC100)
        Set .dicTitulosC100 = Util.MapearTitulos(regC100, 3)
        
    End With
    
End Sub

Public Function ResetarDadosSPEDFiscal()
    
    Dim CamposVazios As RegistrosSPEDFiscal
    LSet SPEDFiscal = CamposVazios
    
End Function
