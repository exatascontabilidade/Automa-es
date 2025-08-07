Attribute VB_Name = "clsSefazBA"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function IncluirAjustesSefazBA(ByVal Arqs As Variant, ByRef arrProdutosExcluidos As ArrayList, ByRef dicFornecSN As Dictionary)

Dim nReg As String
Dim EFD As New ArrayList
Dim dicAjustes As New Dictionary
Dim dicE113 As New Dictionary
Dim Registros As Variant, Registro, Arq
Dim nNF As String, Emissao$, chNFe$, cPart$, Modelo$, SERIE$

    For Each Arq In Arqs

        Registros = "" ' fnSPED.ImportarDadosEFD(Arq)
        For Each Registro In Registros

            nReg = Mid(Registro, 2, 4)
            Select Case True

                Case nReg = "C100"
                    nNF = fnSPED.ExtrairCampo(Registro, 8)
                    Emissao = fnSPED.ExtrairCampo(Registro, 10)
                    chNFe = fnSPED.ExtrairCampo(Registro, 9)
                    Modelo = fnSPED.ExtrairCampo(Registro, 5)
                    SERIE = fnSPED.ExtrairCampo(Registro, 7)
                    cPart = fnSPED.ExtrairCampo(Registro, 4)
                    EFD.Add Registro

                Case nReg = "C170"
                    Call rC170.CalcularAjustesDecretoAtacadistaBA(Registro, arrProdutosExcluidos, dicAjustes, dicE113, cPart, Modelo, SERIE, nNF, Emissao, chNFe)
                    Registro = rC170.CalcularCreditoPresumidoArt269Inc10BA(Registro, dicFornecSN, dicAjustes, dicE113, cPart, Modelo, SERIE, nNF, Emissao, chNFe)
                    Registro = rC170.CalcularCreditoAquisicaoSimplesNacionalBA(Registro, dicFornecSN, dicAjustes, dicE113, cPart, Modelo, SERIE, nNF, Emissao, chNFe)
                    EFD.Add Registro

                Case nReg = "E110"
                    EFD.Add Registro
                    Call IncluirAjustesE111eE113(dicAjustes, dicE113, EFD)

                Case nReg <> ""
                    EFD.Add Registro

            End Select

        Next Registro

        Call fnSPED.ExportarSPED(Replace(Arq, ".txt", " - ALTERADO.txt"), fnSPED.TotalizarRegistrosSPED(EFD))

        EFD.Clear
        Call dicAjustes.RemoveAll
        Call dicE113.RemoveAll

    Next Arq

End Function
