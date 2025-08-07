Attribute VB_Name = "clsFuncoesSPEDContribuicoes"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Sub ListarSPEDsContribuicoes(ByVal Arqs As Variant, ByRef arrContribuicoes As ArrayList)

Dim Arq As Variant
    
    For Each Arq In Arqs
        
        If fnSPED.ClassificarSPED(Arq) = "Fiscal" Then arrContribuicoes.Add Arq
        
    Next Arq
    
End Sub
