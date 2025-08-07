Attribute VB_Name = "AssistenteEstoque"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private dicTitulosAnalitico As Dictionary
Private arrRelatorioAnalitico As ArrayList
Private Fiscal As clsRegistrosSPED
Public CamposInventario As Variant
Private dicTitulos As Dictionary
Private Campos As Variant

Public Sub GerarRelatorioMovimentacaoEstoque()
    
Dim Msg As String
    
    Msg = "Precisa haver dados no registro C170 do SPED Fiscal para usar esse recurso."
    If Util.ChecarAusenciaDados(regC170, False, Msg) Then Exit Sub
    
    Inicio = Now()
    
    Call CarregarRegistrosSPED
    Call ExtrairMovimentacaoC170
    Call ExtrairCamposC100
    
    Call Util.LimparDados(relMovimentoEstoque, 4, False)
    Call Util.ExportarDadosArrayList(relMovimentoEstoque, arrRelatorioAnalitico)
    
    relMovimentoEstoque.Activate
    
    Call Util.MsgInformativa("Relatório analítico de movimentação de estoque gerado com sucesso", _
        "Relatório Analítico de Movimentação de Estoque", Inicio)
        
End Sub

Private Sub ExtrairMovimentacaoC170()

Dim nCampo As Variant
    
    Set arrRelatorioAnalitico = New ArrayList
    
    For Each Campos In dtoRegSPED.rC170.Items()
        
        Call RedimensionarArray(dicTitulos.Count)
        Call DTO_EstoqueInventario.ResetarDTO_MovimentoEstoque
            
        Call ExtrairCamposC170
        Call ExtrairCamposC100
        Call ExtrairCampos0150
        Call ExtrairCampos0200
        Call ExtrairCampos0220
        Call GerarCamposCalculados
        
        arrRelatorioAnalitico.Add CamposInventario
        
    Next Campos
    
End Sub

Private Sub ExtrairCampos0150()

Const TitulosIgnorados As String = "REG"
Dim Campos0150 As Variant
Dim Posicao As Integer
Dim nCampo As Variant
    
    With dtoMovimentoEstoque
        
        .ARQUIVO = ExtrairValor("ARQUIVO")
        .COD_PART = ExtrairValor("COD_PART")
        .Chave = .ARQUIVO & .COD_PART
        If dtoRegSPED.r0150.Exists(.Chave) Then
            
            Campos0150 = dtoRegSPED.r0150(.Chave)
            For Each nCampo In dicTitulosAnalitico
                
                If TitulosIgnorados Like "*" & nCampo & "*" Then GoTo Prx:
                If dtoTitSPED.t0150.Exists(nCampo) Then
                    
                    Posicao = dtoTitSPED.t0150(nCampo)
                    AtribuirValor nCampo, Campos0150(Posicao)
                    
                End If
Prx:
            Next nCampo
            
        End If
        
    End With
    
End Sub

Private Sub ExtrairCampos0200()

Const TitulosIgnorados As String = "REG"
Dim Campos0200 As Variant
Dim Posicao As Integer
Dim nCampo As Variant
    
    With dtoMovimentoEstoque
        
        .ARQUIVO = ExtrairValor("ARQUIVO")
        .COD_ITEM = ExtrairValor("COD_ITEM")
        .Chave = .ARQUIVO & .COD_ITEM
        If dtoRegSPED.r0200.Exists(.Chave) Then
            
            Campos0200 = dtoRegSPED.r0200(.Chave)
            For Each nCampo In dicTitulosAnalitico
                
                If nCampo = "CHV_REG" Then .CHV_REG = Campos0200(dtoTitSPED.t0200(nCampo))
                If nCampo = "UNID_INV" Then .UNID_INV = Campos0200(dtoTitSPED.t0200(nCampo))
                
                If TitulosIgnorados Like "*" & nCampo & "*" Then GoTo Prx:
                If dtoTitSPED.t0200.Exists(nCampo) Then
                    
                    Posicao = dtoTitSPED.t0200(nCampo)
                    AtribuirValor nCampo, Campos0200(Posicao)
                    
                End If
Prx:
            Next nCampo
            
        End If
        
    End With
    
End Sub

Private Sub ExtrairCampos0220()

Const TitulosIgnorados As String = "REG"
Dim Campos0220 As Variant
Dim Posicao As Integer
Dim nCampo As Variant
    
    With dtoMovimentoEstoque
        
        .Chave = fnSPED.GerarChaveRegistro(.CHV_REG, .UNID_COM)
        If dtoRegSPED.r0220.Exists(.Chave) Then
            
            Campos0220 = dtoRegSPED.r0220(.Chave)
            For Each nCampo In dicTitulosAnalitico
                
                If nCampo = "FAT_CONV" Then .FAT_CONV = Campos0220(dtoTitSPED.t0220(nCampo))
                
                If TitulosIgnorados Like "*" & nCampo & "*" Then GoTo Prx:
                If dtoTitSPED.t0220.Exists(nCampo) Then
                    
                    Posicao = dtoTitSPED.t0220(nCampo)
                    AtribuirValor nCampo, Campos0220(Posicao)
                    
                End If
Prx:
            Next nCampo
            
        End If
        
    End With
    
End Sub

Private Sub ExtrairCamposC100()

Const TitulosIgnorados As String = "REG, VL_DESC, VL_ICMS, VL_ICMS_ST, VL_IPI, VL_PIS, VL_COFINS"
Dim CamposC100 As Variant
Dim Posicao As Integer
Dim nCampo As Variant
    
    With dtoMovimentoEstoque
        
        .CHV_PAI = ExtrairValor("CHV_PAI_FISCAL")
        If dtoRegSPED.rC100.Exists(.CHV_PAI) Then
                        
            CamposC100 = dtoRegSPED.rC100(.CHV_PAI)
            Call ExtrairVL_DESP_C100(CamposC100, .CHV_PAI)
            For Each nCampo In dicTitulosAnalitico
                
                If nCampo = "DT_OPERACAO" Then .DT_DOC = CamposC100(dtoTitSPED.tC100("DT_DOC"))
                If nCampo = "DT_OPERACAO" Then .DT_E_S = CamposC100(dtoTitSPED.tC100("DT_E_S"))
                
                If TitulosIgnorados Like "*" & nCampo & "*" Then GoTo Prx:
                If dtoTitSPED.tC100.Exists(nCampo) Then
                    
                    Posicao = dtoTitSPED.tC100(nCampo)
                    AtribuirValor nCampo, CamposC100(Posicao)
                    
                End If
Prx:
            Next nCampo
            
        End If
        
    End With
    
End Sub

Private Sub ExtrairCamposC170()

Dim Posicao As Integer
Dim nCampo As Variant
    
    With dtoMovimentoEstoque
        
        For Each nCampo In dicTitulosAnalitico
            
            If nCampo = "CFOP" Then .CFOP = Campos(dtoTitSPED.tC170(nCampo))
            If nCampo = "UNID" Then .UNID_COM = Campos(dtoTitSPED.tC170(nCampo))
            If nCampo = "QTD" Then .QTD_COM = Campos(dtoTitSPED.tC170(nCampo))
            If nCampo = "VL_ITEM" Then .VL_ITEM = Campos(dtoTitSPED.tC170(nCampo))
            If nCampo = "VL_ICMS" Then .VL_ICMS = Campos(dtoTitSPED.tC170(nCampo))
            If nCampo = "VL_ICMS_ST" Then .VL_ICMS_ST = Campos(dtoTitSPED.tC170(nCampo))
            If nCampo = "VL_IPI" Then .VL_IPI = Campos(dtoTitSPED.tC170(nCampo))
            If nCampo = "VL_PIS" Then .VL_PIS = Campos(dtoTitSPED.tC170(nCampo))
            If nCampo = "VL_COFINS" Then .VL_COFINS = Campos(dtoTitSPED.tC170(nCampo))
            
            If dtoTitSPED.tC170.Exists(nCampo) Then
                
                Posicao = dtoTitSPED.tC170(nCampo)
                AtribuirValor nCampo, Campos(Posicao)
                
            End If
            
        Next nCampo
        
    End With
    
End Sub

Private Sub ExtrairVL_DESP_C100(ByRef Campos As Variant, ByVal CHV_REG As String)
    
    With dtoMovimentoEstoque
        
        .VL_MERC = fnExcel.FormatarValores(Campos(dtoTitSPED.tC100("VL_MERC")))
        .VL_FRT = fnExcel.FormatarValores(Campos(dtoTitSPED.tC100("VL_FRT")))
        .VL_SEG = fnExcel.FormatarValores(Campos(dtoTitSPED.tC100("VL_SEG")))
        .VL_OUT = fnExcel.FormatarValores(Campos(dtoTitSPED.tC100("VL_OUT_DA")))
        .VL_DESP = .VL_FRT + .VL_SEG + .VL_OUT
        
        If .VL_MERC > 0 Then AtribuirValor "VL_DESP", (.VL_ITEM / .VL_MERC) * .VL_DESP
        
    End With
    
End Sub

Private Sub GerarCamposCalculados()
    
    Call DefinirDataOperacao
    Call CalcularQtdInventario
    Call CalcularCusto
    Call CalcularCustoUnitarioComercial
    Call CalcularCustoUnitarioInventario
    
End Sub

Private Sub DefinirDataOperacao()
    
    With dtoMovimentoEstoque
        
        Select Case True
            
            Case .CFOP < 4000
                .DT_OPERACAO = .DT_E_S
                
            Case .CFOP > 4000
                .DT_OPERACAO = .DT_DOC
                
        End Select
        
        AtribuirValor "DT_OPERACAO", .DT_OPERACAO
        
    End With
    
End Sub

Private Sub CalcularQtdInventario()
    
    With dtoMovimentoEstoque
        
        If .FAT_CONV <> 0 Then .QTD_INV = .QTD_COM * .FAT_CONV Else .QTD_INV = .QTD_COM
        AtribuirValor "QTD_INV", .QTD_INV
        
    End With

End Sub

Private Sub CalcularCusto()
    
    With dtoMovimentoEstoque
        
        Select Case .CFOP
            
            Case .CFOP < 4000
                .VL_CUSTO = .VL_ITEM - .VL_DESC - .VL_ICMS + .VL_IPI - .VL_PIS - .VL_COFINS
                
            Case .CFOP > 4000
                .VL_CUSTO = .VL_ITEM + .VL_DESP - .VL_DESC
                
        End Select
        
        AtribuirValor "VL_CUSTO", .VL_CUSTO
        
    End With
    
End Sub

Private Sub CalcularCustoUnitarioComercial()
    
    With dtoMovimentoEstoque
        
        .VL_CUSTO_UNIT_COM = .VL_CUSTO / .QTD_COM
        AtribuirValor "VL_CUSTO_UNIT_COM", .VL_CUSTO_UNIT_COM
        
    End With
    
End Sub

Private Sub CalcularCustoUnitarioInventario()
    
    With dtoMovimentoEstoque
        
        .VL_CUSTO_UNIT_INV = .VL_CUSTO / .QTD_INV
        AtribuirValor "VL_CUSTO_UNIT_INV", .VL_CUSTO_UNIT_INV
        
    End With
    
End Sub

Private Sub CarregarRegistrosSPED()
    
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    
    Set dicTitulosAnalitico = Util.MapearTitulos(relMovimentoEstoque, 3)
    Set dicTitulos = dicTitulosAnalitico
    
    Set Fiscal = New clsRegistrosSPED
    
    With Fiscal
        
        Call .CarregarDadosRegistro0150("ARQUIVO", "COD_PART")
        Call .CarregarDadosRegistro0190
        Call .CarregarDadosRegistro0200("ARQUIVO", "COD_ITEM")
        Call .CarregarDadosRegistro0220
        Call .CarregarDadosRegistroC100
        Call .CarregarDadosRegistroC170
        
    End With
    
End Sub

Private Function AtribuirValor(ByVal Titulo As String, ByVal Valor As Variant)
    
    CamposInventario(dicTitulos(Titulo)) = Valor
    
End Function

Private Function ExtrairValor(ByVal Titulo As String)
    
    ExtrairValor = CamposInventario(dicTitulos(Titulo))
    
End Function

Private Function RedimensionarArray(ByVal NumCampos As Long)
    
    ReDim CamposInventario(1 To NumCampos) As Variant
    
End Function
