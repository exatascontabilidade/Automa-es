Attribute VB_Name = "clsRegrasFiscais"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
        Option Explicit

Public Geral As New clsRegrasFiscaisGerais
Public PISCOFINS As New clsRegrasFiscaisPisCofins
Public ApuracaoICMS As New clsRegrasApuracaoICMS
Public ApuracaoIPI As New clsRegrasApuracaoIPI
Public ApuracaoPISCOFINS As New clsRegrasApuracaoPISCOFINS
Public SPEDFiscal As New clsRegrasSPEDFiscal
Public SPEDContribuicoes As New clsRegrasSPEDContribuicoes
Public DivergenciasFiscais As New clsRegrasDivergencias

