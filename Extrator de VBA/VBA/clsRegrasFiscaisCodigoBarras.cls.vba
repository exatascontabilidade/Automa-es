Attribute VB_Name = "clsRegrasFiscaisCodigoBarras"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function ValidarCodigoBarras(ByVal COD_BARRAS As String) As Boolean

Dim digitoVerificador As Integer
Dim Tamanho As Integer
Dim soma As Integer
Dim i As Integer
    
    Tamanho = VBA.Len(COD_BARRAS)
    
    ' Verifica se o tamanho do código é EAN-8 ou EAN-13
    If (Tamanho <> 8 And Tamanho <> 13) Or Not IsNumeric(COD_BARRAS) Then
        ValidarCodigoBarras = False
        Exit Function
    End If
    
    ' Calcula a soma dos dígitos
    For i = 1 To Tamanho - 1
    
        If (i Mod 2 = 0 And Tamanho = 13) Or (i Mod 2 <> 0 And Tamanho = 8) Then
            soma = soma + CInt(Mid(COD_BARRAS, i, 1)) * 3
        Else
            soma = soma + CInt(Mid(COD_BARRAS, i, 1))
        End If
        
    Next i
    
    ' Calcula o dígito verificador
    digitoVerificador = 10 - (soma Mod 10)
    If digitoVerificador = 10 Then digitoVerificador = 0
    
    ' Verifica se o dígito verificador é igual ao último dígito do código
    If digitoVerificador = CInt(Right(COD_BARRAS, 1)) Then
        ValidarCodigoBarras = True
    Else
        ValidarCodigoBarras = False
    End If
    
End Function

