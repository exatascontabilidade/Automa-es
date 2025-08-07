Attribute VB_Name = "clsE113"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function CriarAtualizarE113(ByRef dicE113 As Dictionary, ByVal cAjuste, ByVal vAjuste As Double, _
                                   ByVal cPart As String, ByVal Modelo As String, ByVal SERIE As String, _
                                   ByVal nNF As String, ByVal Emissao As String, ByVal cProd As String, ByVal chNFe As String)
   
    If Not dicE113.Exists(cAjuste) Then Set dicE113(cAjuste) = CreateObject("System.Collections.ArrayList")
    dicE113(cAjuste).Add Join(Array("", "E113", cPart, Modelo, SERIE, "", nNF, Emissao, cProd, vAjuste, chNFe, ""), "|")
    
End Function
