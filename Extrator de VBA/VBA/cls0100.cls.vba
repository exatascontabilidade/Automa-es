Attribute VB_Name = "cls0100"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function ImportarDadosContador()

Dim Arq As Variant, Registro, Campos, Chave, CamposDic
Dim dicDados As New Dictionary
Dim i As Byte
    
    Arq = Util.SelecionarArquivo("txt", "Selecione o SPED com dados do contador")
    If VarType(Arq) = vbBoolean Then Exit Function
    
    Set dicDados = Util.CriarDicionarioRegistro(reg0100)
    
    Registro = fnSPED.ExtrairDadosContador(Arq)
    Campos = VBA.Split(Registro, "|")
    
    For Each Chave In dicDados.Keys()
        
        CamposDic = dicDados(Chave)
        For i = 2 To UBound(Campos)
            If (i + 4) > UBound(CamposDic) Then Exit For
            CamposDic(i + 4) = Campos(i)
        Next i
        
        dicDados(Chave) = CamposDic
        
    Next Chave
    
    Call Util.LimparDados(reg0100, 4, False)
    Call Util.ExportarDadosDicionario(reg0100, dicDados)
    
End Function
