Attribute VB_Name = "AssistenteInventario"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private dtoInventario As DTOsClasseInventario

Private Type DTOsClasseInventario
    
    Validacoes As AssistenteInventario_Validacoes
    dicTitulosMovimentoEstoque As Dictionary
    dicTitulosSaldoInventario As Dictionary
    dicMovimentoEstoque As Dictionary
    dicSaldoInventario As Dictionary
    Fiscal As clsRegistrosSPED
    arrMovimentoEstoque As ArrayList
    CamposInventario As Variant
    dicTitulos As Dictionary
    DT_INVENTARIO As String
    Campos As Variant
    
End Type

Public Sub GerarRelatorioSaldoInventario()

Dim Msg As String
    
    Call Util.AtualizarBarraStatus("Gerando saldos de inventário, por favor aguarde...")
    
    With dtoInventario
        
        .DT_INVENTARIO = relSaldoInventario.Range("DT_INVENTARIO").value
        If Not ValidarDataInventario(.DT_INVENTARIO) Then Exit Sub
        
        Msg = "Precisa haver dados na movimentação de estoque para usar esse recurso."
        If Util.ChecarAusenciaDados(relMovimentoEstoque, False, Msg) Then Exit Sub
        
        Inicio = Now()
        
        Call CarregarRegistros
        Call .Validacoes.CarregarFuncoesVerificacao
        
        Call CalcularSaldoInventario
        
        Call Util.LimparDados(relSaldoInventario, 4, False)
        
        Call Util.AtualizarBarraStatus("Exportando saldos de inventário...")
        Call Util.ExportarDadosDicionario(relSaldoInventario, .dicSaldoInventario)
        Call FuncoesFormatacao.DestacarInconsistencias(relSaldoInventario)
        
        Call .Validacoes.ResetarDTOs
        Call ResetarDTOs
        
        Call Util.MsgInformativa("Saldo de inventário calculado com sucesso!", "Análise de Saldo de Inventário", Inicio)
        
    End With
    
End Sub

Private Sub CalcularSaldoInventario()
    
Dim DT_OPERACAO As String, COD_ITEM$
    
    Set dtoInventario.dicSaldoInventario = New Dictionary
    
    With dtoSaldoInventario
        
        For Each Campos In dtoInventario.arrMovimentoEstoque
            
            Call Util.AntiTravamento(a, 50, "Gerando saldos de inventário, por favor aguarde...")
            DT_OPERACAO = Campos(dtoInventario.dicTitulosMovimentoEstoque("DT_OPERACAO"))
            
            If ValidarInclusaoOperacao(DT_OPERACAO) Then
            
                Call RedimensionarArray(dtoInventario.dicTitulos.Count)
                Call DTO_EstoqueInventario.ResetarDTO_MovimentoEstoque
                
                Call ExtrairCamposMovimentoEstoque
                Call ExtrairSaldoInventarioBlocoH
                
                If dtoInventario.dicSaldoInventario.Exists(.COD_ITEM) Then Call AtualizarRegistroInventario(.COD_ITEM)
                
                Call AtualizarSaldoInventario
                COD_ITEM = .COD_ITEM
                
                With dtoInventario
                    
                    Call CarregarDadosDTO
                        'Call .Validacoes.ValidarRegrasInventario(.CamposInventario, .dicTitulosSaldoInventario)
                    Call IncluirDadosDTOParaCamposInventario
                    
                End With
                
                dtoInventario.dicSaldoInventario(.COD_ITEM) = dtoInventario.CamposInventario
                
            End If
            
        Next Campos
        
    End With
    
End Sub

Private Function ValidarDataInventario(ByVal DT_INVENTARIO As String) As Boolean
    
    On Error GoTo Tratar:
    
    Select Case True
        
        Case DT_INVENTARIO = ""
            Call AletarDataInventarioAusente
            
        Case Not VBA.IsDate(CDate(DT_INVENTARIO))
            Call AletarDataInventarioInvalida
            
        Case VBA.IsDate(DT_INVENTARIO)
            ValidarDataInventario = True
            
    End Select
    
Exit Function
Tratar:
    
    Select Case True
        
        Case Err.Number = 13
            Call AletarDataInventarioInvalida
            
    End Select
    
End Function

Private Sub AletarDataInventarioInvalida()
    
    Call Util.MsgAlerta("Informe uma data VÁLIDA de inventário para gera o relatório.", "Data de inventário inválida")
    relSaldoInventario.Range("DT_INVENTARIO").Select
    
End Sub

Private Sub AletarDataInventarioAusente()
    
    Call Util.MsgAlerta("É preciso informar uma data de inventário para gera o relatório.", "Informe uma data de inventário")
    relSaldoInventario.Range("DT_INVENTARIO").Select
    
End Sub

Private Function ValidarInclusaoOperacao(ByVal DT_OPERACAO As String) As Boolean
    
    With dtoInventario
        
        Select Case True
            
            Case DT_OPERACAO = "", Not VBA.IsDate(CDate(DT_OPERACAO))
                Exit Function
                
            Case CDate(DT_OPERACAO) <= CDate(.DT_INVENTARIO)
                ValidarInclusaoOperacao = True
                
        End Select
    
    End With
    
End Function

Private Sub ExtrairCamposMovimentoEstoque()

Dim nCampo As Variant
    
    With dtoSaldoInventario
        
        For Each nCampo In dtoInventario.dicTitulosMovimentoEstoque.Keys()
            
            If nCampo = "COD_ITEM" Then .COD_ITEM = Campos(dtoInventario.dicTitulosMovimentoEstoque(nCampo))
            
            Select Case CStr(nCampo)
                
                Case "CFOP"
                    .CFOP = Campos(dtoInventario.dicTitulosMovimentoEstoque("CFOP"))
                    .QTD_INV = Campos(dtoInventario.dicTitulosMovimentoEstoque("QTD_INV"))
                    .VL_ITEM = Campos(dtoInventario.dicTitulosMovimentoEstoque("VL_ITEM"))
                    
                    Call AtribuirQuantidade
                
                Case "INCONSISTENCIA", "SUGESTAO"
                    AtribuirValor nCampo, Empty
                    
                Case Else
                    If dtoInventario.dicTitulosSaldoInventario.Exists(nCampo) Then _
                        AtribuirValor nCampo, Campos(dtoInventario.dicTitulosMovimentoEstoque(nCampo))
                    
            End Select
            
        Next nCampo
        
    End With
    
End Sub

Private Sub ExtrairSaldoInventarioBlocoH()

Const TitulosIgnorados As String = "REG, VL_DESC, VL_ICMS, VL_ICMS_ST, VL_IPI, VL_PIS, VL_COFINS"
Dim CamposH010 As Variant
Dim Posicao As Integer
Dim nCampo As Variant
    
    With dtoSaldoInventario
        
        .COD_ITEM = ExtrairValor("COD_ITEM")
        If dtoRegSPED.rH010.Exists(.COD_ITEM) Then
            
            CamposH010 = dtoRegSPED.rH010(.COD_ITEM)
            
            .QTD_INICIAL = CamposH010(dtoTitSPED.tH010("QTD"))
            .VL_INICIAL = CamposH010(dtoTitSPED.tH010("VL_ITEM"))
            
            AtribuirValor "QTD_INICIAL", .QTD_INICIAL
            AtribuirValor "VL_INICIAL", .VL_INICIAL
            
        End If
        
    End With

End Sub

Private Sub AtribuirQuantidade()
    
    With dtoSaldoInventario
        
        Select Case True
            
            Case .CFOP < 4000
                AtribuirValor "QTD_ENT", .QTD_INV
                AtribuirValor "VL_ENT", .VL_ITEM
                AtribuirValor "VL_UNIT_ENT", CalcularValorUnitario
                
                AtribuirValor "QTD_SAI", 0
                AtribuirValor "VL_SAI", 0
                AtribuirValor "VL_UNIT_SAI", 0
                
            Case .CFOP > 4000
                AtribuirValor "QTD_SAI", .QTD_INV
                AtribuirValor "VL_SAI", .VL_ITEM
                If .QTD_ENT = 164 Then Stop
                AtribuirValor "VL_UNIT_SAI", CalcularValorUnitario
                
                AtribuirValor "QTD_ENT", 0
                AtribuirValor "VL_ENT", 0
                AtribuirValor "VL_UNIT_ENT", 0
                
        End Select
        
    End With
    
End Sub

Private Sub CarregarRegistros()
    
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    
    With dtoInventario
        
        Set .arrMovimentoEstoque = Util.CriarArrayListRegistro(relMovimentoEstoque)
        Set .dicTitulosMovimentoEstoque = Util.MapearTitulos(relMovimentoEstoque, 3)
        Set .dicTitulosSaldoInventario = Util.MapearTitulos(relSaldoInventario, 3)
        Set .dicTitulos = .dicTitulosSaldoInventario
        
        Set .Fiscal = New clsRegistrosSPED
        Set .Validacoes = New AssistenteInventario_Validacoes
        
        With .Fiscal
            
            Call .CarregarDadosRegistroH010("COD_ITEM", "IND_PROP", "COD_PART")
            
        End With
        
    End With
    
End Sub

Private Function CalcularValorUnitario() As Double
    
    With dtoSaldoInventario
        
        If .QTD_INV > 0 Then CalcularValorUnitario = VBA.Round(.VL_ITEM / .QTD_INV, 2) Else CalcularValorUnitario = .VL_ITEM
        
    End With
    
End Function

Private Sub AtualizarRegistroInventario(ByVal Chave As String)

Dim CamposSaldo As Variant, nCampo
Dim i As Integer, j As Integer
Dim vCampo As Double
    
    With dtoSaldoInventario
        
        CamposSaldo = dtoInventario.dicSaldoInventario(Chave)
        i = Util.VerificarPosicaoInicialArray(CamposSaldo)
        j = Util.VerificarPosicaoInicialArray(dtoInventario.CamposInventario)
        For Each nCampo In dtoInventario.dicTitulosSaldoInventario.Keys()
            
            If nCampo Like "VL_*" Or nCampo Like "QTD_*" Then
                
                vCampo = CDbl(dtoInventario.CamposInventario(dtoInventario.dicTitulosSaldoInventario(CStr(nCampo)) - i)) _
                    + CDbl(CamposSaldo(dtoInventario.dicTitulosSaldoInventario(CStr(nCampo)) - i))
                AtribuirValor CStr(nCampo), vCampo
                
            End If
            
        Next nCampo
        
    End With
    
End Sub

Private Sub AtualizarSaldoInventario()
    
    Call CalcularSaldoFinal
    Call CalcularMargem
    
End Sub

Private Sub CalcularSaldoFinal()
    
    With dtoSaldoInventario
        
        .QTD_INICIAL = dtoInventario.CamposInventario(dtoInventario.dicTitulosSaldoInventario("QTD_INICIAL"))
        .QTD_ENT = dtoInventario.CamposInventario(dtoInventario.dicTitulosSaldoInventario("QTD_ENT"))
        .QTD_SAI = dtoInventario.CamposInventario(dtoInventario.dicTitulosSaldoInventario("QTD_SAI"))
        
        .QTD_FINAL = .QTD_INICIAL + .QTD_ENT - .QTD_SAI
        
        AtribuirValor "QTD_FINAL", .QTD_FINAL
        
    End With
    
End Sub

Private Sub CalcularMargem()
    
    With dtoSaldoInventario
        
        .VL_UNIT_ENT = dtoInventario.CamposInventario(dtoInventario.dicTitulosSaldoInventario("VL_UNIT_ENT"))
        .VL_UNIT_SAI = dtoInventario.CamposInventario(dtoInventario.dicTitulosSaldoInventario("VL_UNIT_SAI"))
        
        .VL_MARGEM = .VL_UNIT_SAI - .VL_UNIT_ENT
        If .VL_UNIT_ENT > 0 Then .ALIQ_MARGEM = .VL_MARGEM / .VL_UNIT_ENT
        
        AtribuirValor "VL_MARGEM", .VL_MARGEM
        AtribuirValor "ALIQ_MARGEM", .ALIQ_MARGEM
        
    End With
    
End Sub

Private Function AtribuirValor(ByVal Titulo As String, ByVal Valor As Variant)
    
    With dtoInventario
        
        .CamposInventario(.dicTitulos(Titulo)) = Valor
        
    End With
    
End Function

Private Function ExtrairValor(ByVal Titulo As String)
    
    With dtoInventario
        
        ExtrairValor = .CamposInventario(.dicTitulos(Titulo))
        
    End With
    
End Function

Private Function RedimensionarArray(ByVal NumCampos As Long)
    
    ReDim dtoInventario.CamposInventario(1 To NumCampos) As Variant
    
End Function

Private Sub CarregarDadosDTO()
    
    With dtoInventario
        
        dtoSaldoInventario.DT_INV = .CamposInventario(.dicTitulosSaldoInventario("DT_INV"))
        dtoSaldoInventario.COD_ITEM = .CamposInventario(.dicTitulosSaldoInventario("COD_ITEM"))
        dtoSaldoInventario.DESCR_ITEM = .CamposInventario(.dicTitulosSaldoInventario("DESCR_ITEM"))
        dtoSaldoInventario.QTD_INICIAL = .CamposInventario(.dicTitulosSaldoInventario("QTD_INICIAL"))
        dtoSaldoInventario.QTD_ENT = .CamposInventario(.dicTitulosSaldoInventario("QTD_ENT"))
        dtoSaldoInventario.QTD_SAI = .CamposInventario(.dicTitulosSaldoInventario("QTD_SAI"))
        dtoSaldoInventario.QTD_FINAL = .CamposInventario(.dicTitulosSaldoInventario("QTD_FINAL"))
        dtoSaldoInventario.VL_INICIAL = .CamposInventario(.dicTitulosSaldoInventario("VL_INICIAL"))
        dtoSaldoInventario.VL_ENT = .CamposInventario(.dicTitulosSaldoInventario("VL_ENT"))
        dtoSaldoInventario.VL_SAI = .CamposInventario(.dicTitulosSaldoInventario("VL_SAI"))
        dtoSaldoInventario.VL_UNIT_ENT = .CamposInventario(.dicTitulosSaldoInventario("VL_UNIT_ENT"))
        dtoSaldoInventario.VL_UNIT_SAI = .CamposInventario(.dicTitulosSaldoInventario("VL_UNIT_SAI"))
        dtoSaldoInventario.VL_MARGEM = .CamposInventario(.dicTitulosSaldoInventario("VL_MARGEM"))
        dtoSaldoInventario.ALIQ_MARGEM = .CamposInventario(.dicTitulosSaldoInventario("ALIQ_MARGEM"))
        dtoSaldoInventario.INCONSISTENCIA = .CamposInventario(.dicTitulosSaldoInventario("INCONSISTENCIA"))
        dtoSaldoInventario.SUGESTAO = .CamposInventario(.dicTitulosSaldoInventario("SUGESTAO"))
        
    End With
    
End Sub

Private Sub IncluirDadosDTOParaCamposInventario()
    
    With dtoInventario
        
        .CamposInventario(.dicTitulosSaldoInventario("DT_INV")) = dtoSaldoInventario.DT_INV
        .CamposInventario(.dicTitulosSaldoInventario("COD_ITEM")) = dtoSaldoInventario.COD_ITEM
        .CamposInventario(.dicTitulosSaldoInventario("DESCR_ITEM")) = dtoSaldoInventario.DESCR_ITEM
        .CamposInventario(.dicTitulosSaldoInventario("QTD_INICIAL")) = dtoSaldoInventario.QTD_INICIAL
        .CamposInventario(.dicTitulosSaldoInventario("QTD_ENT")) = dtoSaldoInventario.QTD_ENT
        .CamposInventario(.dicTitulosSaldoInventario("QTD_SAI")) = dtoSaldoInventario.QTD_SAI
        .CamposInventario(.dicTitulosSaldoInventario("QTD_FINAL")) = dtoSaldoInventario.QTD_FINAL
        .CamposInventario(.dicTitulosSaldoInventario("VL_INICIAL")) = dtoSaldoInventario.VL_INICIAL
        .CamposInventario(.dicTitulosSaldoInventario("VL_ENT")) = dtoSaldoInventario.VL_ENT
        .CamposInventario(.dicTitulosSaldoInventario("VL_SAI")) = dtoSaldoInventario.VL_SAI
        .CamposInventario(.dicTitulosSaldoInventario("VL_UNIT_ENT")) = dtoSaldoInventario.VL_UNIT_ENT
        .CamposInventario(.dicTitulosSaldoInventario("VL_UNIT_SAI")) = dtoSaldoInventario.VL_UNIT_SAI
        .CamposInventario(.dicTitulosSaldoInventario("VL_MARGEM")) = dtoSaldoInventario.VL_MARGEM
        .CamposInventario(.dicTitulosSaldoInventario("ALIQ_MARGEM")) = dtoSaldoInventario.ALIQ_MARGEM
        .CamposInventario(.dicTitulosSaldoInventario("INCONSISTENCIA")) = dtoSaldoInventario.INCONSISTENCIA
        .CamposInventario(.dicTitulosSaldoInventario("SUGESTAO")) = dtoSaldoInventario.SUGESTAO
        
    End With
    
End Sub

Public Sub ListarCFOPs()

Dim arrRelatorio As New ArrayList
Dim arrCFOPs As New ArrayList
Dim CFOP As Variant
    
    Set arrCFOPs = Util.ListarValoresUnicos(relMovimentoEstoque, 4, 3, "CFOP")
    
    For Each CFOP In arrCFOPs
        
        If CFOP Like "#9##" Then arrRelatorio.Add Array(CFOP, "NÃO") Else arrRelatorio.Add Array(CFOP, "SIM")
        
    Next CFOP
    
    Call Util.LimparDados(configInventario, 4, False)
    Call Util.ExportarDadosArrayList(configInventario, arrRelatorio)
    
End Sub

Public Function ResetarDTOs()

Dim dtoVazio As DTOsClasseInventario
    
    LSet dtoInventario = dtoVazio
    
End Function
