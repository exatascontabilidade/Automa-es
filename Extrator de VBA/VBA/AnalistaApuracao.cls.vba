Attribute VB_Name = "AnalistaApuracao"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public dicTitulosApuracao As Dictionary
Public dicTitulosResumo As Dictionary
Public dicTitulos As Dictionary
Public Campos As Variant

Public Function AtribuirValor(ByVal Titulo As String, ByVal Valor As Variant)
    
    Campos(dicTitulos(Titulo)) = Valor
    
End Function

Public Function RedimensionarArray(ByVal NumCampos As Long)
    
    ReDim Campos(1 To NumCampos) As Variant
    
End Function

