Attribute VB_Name = "clsFuncoesSPEDFiscal"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function ImportarSPEDFiscal(Optional ByVal SelReg As String, Optional Periodo As String, Optional Unificar As Boolean)



End Function

Public Sub ListarSPEDsFiscais(ByVal Arqs As Variant, ByRef arrFiscal As ArrayList)

Dim Arq As Variant
    
    For Each Arq In Arqs
        
        If fnSPED.ClassificarSPED(Arq) = "Fiscal" Then arrFiscal.Add Arq
        
    Next Arq
    
End Sub
