Attribute VB_Name = "clsRegrasFiscaisGerais"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public CodigoBarras As New clsRegrasFiscaisCodigoBarras
Private ValidacoesNCM As New clsRegrasFiscaisNCM
Public ValidacoesCFOP As New clsRegrasFiscaisCFOP

Public Function ValidarCEST(ByVal CEST As String) As Boolean

    If dicTabelaCEST.Count = 0 Then Call Util.CarregarTabelaCEST(dicTabelaCEST)
    
    If CEST <> "" Then
        If dicTabelaCEST.Exists(CEST) Then ValidarCEST = True
    End If
    
End Function

Public Function VerificarCEST(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim CEST As String
Dim i As Integer
    
    If LBound(Campos) = 0 Then i = 1
    
    CEST = fnSPED.FormatarCEST(Util.ApenasNumeros(Campos(dicTitulos("CEST") - i)))

    Select Case True
        
        Case CEST = ""
            Call Util.GravarSugestao(Campos, dicTitulos, _
                INCONSISTENCIA:="O campo COD_CEST não foi informado", _
                SUGESTAO:="informar um valor válido para o campo COD_CEST", _
                dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
                
        Case VBA.Len(CEST) < 7 And VBA.Len(CEST) > 0
            Call Util.GravarSugestao(Campos, dicTitulos, _
                INCONSISTENCIA:="O CEST precisa ter 7 dígitos.", _
                SUGESTAO:="Adicionar zeros a esquerda do CEST", _
                dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
                
        Case Else
            If Not ValidarCEST(CEST) Then
                Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="O CEST informado é inválido", _
                    SUGESTAO:="Informar um valor válido para o campo COD_CEST", _
                    dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
            End If
            
    End Select
    
End Function

Public Function VerificarImportacaoXML(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim NUM_ITEM_NF As String
Dim i As Integer
    
    If LBound(Campos) = 0 Then i = 1
    
    NUM_ITEM_NF = Campos(dicTitulos("NUM_ITEM_NF") - i)
    
    Select Case True
        
        Case NUM_ITEM_NF = ""
            Call Util.GravarSugestao(Campos, dicTitulos, _
                INCONSISTENCIA:="O XML dessa operação não foi importado", _
                SUGESTAO:="Inclua o XML dessa operação na pasta e gere o relatório novamente", _
                dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
                
    End Select
    
End Function

Public Function ValidarCampo_TIPO_ITEM(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim CFOP As Integer, i%
Dim TIPO_ITEM As String, INCONSISTENCIA$, SUGESTAO$
Dim vCompraUsoConsumoSemST As Boolean, vCompraUsoConsumoComST As Boolean, vCompraAtivoImobilizadoSemST As Boolean, _
    vCompraAtivoImobilizadoComST As Boolean, vCompraRevendaSemST As Boolean, vCompraRevendaComST As Boolean, _
    vCompraCombustiveisConsumo As Boolean
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega informações do relatório
    TIPO_ITEM = Util.ApenasNumeros(Campos(dicTitulos("TIPO_ITEM") - i))
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    
    'Verifica informações dos CFOPS
    vCompraRevendaSemST = ValidacoesCFOP.ValidarCFOPCompraRevendaSemST(CFOP)
    vCompraRevendaComST = ValidacoesCFOP.ValidarCFOPCompraRevendaComST(CFOP)
    vCompraUsoConsumoSemST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoSemST(CFOP)
    vCompraUsoConsumoComST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoComST(CFOP)
    vCompraAtivoImobilizadoSemST = ValidacoesCFOP.ValidarCFOPCompraAtivoImobilizadoSemST(CFOP)
    vCompraAtivoImobilizadoComST = ValidacoesCFOP.ValidarCFOPCompraAtivoImobilizadoComST(CFOP)
    vCompraCombustiveisConsumo = ValidacoesCFOP.ValidarCFOPCompraCombustiveisConsumo(CFOP)
    'Operação sujeita a ST com CST_ICMS incorreto
    If TIPO_ITEM = "" Then
        
        Call Util.GravarSugestao(Campos, dicTitulos, _
            INCONSISTENCIA:="O campo 'TIPO_ITEM' não foi informado", _
            SUGESTAO:="Informe um tipo de item para o produto", _
            dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
            
    'Verifica se o valor informado para o campos TIPO_ITEM está compatível com operações de aquisição para revenda
    Else
        
        Select Case True
            
            Case (vCompraRevendaSemST Or vCompraRevendaComST) And Not TIPO_ITEM Like "00*"
                INCONSISTENCIA = "O valor do campo TIPO_ITEM é incompatível com a operação de compra para revenda"
                SUGESTAO = "Alterar o valor do campo TIPO_ITEM para 00"
                
            Case (vCompraUsoConsumoSemST Or vCompraUsoConsumoComST Or vCompraCombustiveisConsumo) And Not TIPO_ITEM Like "07*"
                INCONSISTENCIA = "O valor do campo TIPO_ITEM é incompatível com a operação de compra para uso e consumo"
                SUGESTAO = "Alterar o valor do campo TIPO_ITEM para 07"
                
            Case (vCompraAtivoImobilizadoSemST Or vCompraAtivoImobilizadoComST) And Not TIPO_ITEM Like "08*"
                INCONSISTENCIA = "O valor do campo TIPO_ITEM é incompatível com a operação de compra para o ativo imobilizado"
                SUGESTAO = "Alterar o valor do campo TIPO_ITEM para 08"
                
        End Select
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_IND_MOV(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim IND_MOV As String, INCONSISTENCIA$, SUGESTAO$
Dim i As Integer
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    IND_MOV = Util.ApenasNumeros(Campos(dicTitulos("IND_MOV") - i))
    
    'Operação sujeita a ST com CST_ICMS incorreto
    If IND_MOV = "" Then
        INCONSISTENCIA = "O campo 'IND_MOV' não foi informado"
        SUGESTAO = "Informar se há movimentação física usando: 0 - SIM / 1 - NÃO"
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
            
End Function

Public Function ValidarCampo_COD_CTA(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
    
Dim COD_CTA As String
Dim i As Byte

    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    COD_CTA = Util.RemoverAspaSimples(Campos(dicTitulos("COD_CTA") - i))

    Select Case True
        
        Case COD_CTA = ""
            Call Util.GravarSugestao(Campos, dicTitulos, _
                INCONSISTENCIA:="O campo 'COD_CTA' não foi informado", _
                SUGESTAO:="Informar o código da conta analítica para a operação", _
                dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
                
    End Select
  
End Function

Public Function ValidarCampo_VL_DESC(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim VL_ITEM As Double, VL_DESP#, VL_DESC#
Dim INCONSISTENCIA As String, SUGESTAO$
Dim i As Integer
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    VL_ITEM = fnExcel.ConverterValores(Campos(dicTitulos("VL_ITEM") - i), True, 2)
    VL_DESP = fnExcel.ConverterValores(Campos(dicTitulos("VL_DESP") - i), True, 2)
    VL_DESC = fnExcel.ConverterValores(Campos(dicTitulos("VL_DESC") - i), True, 2)
    
    'Operação sujeita a ST com CST_ICMS incorreto
    If VL_DESC > (VL_ITEM + VL_DESP) Then
        INCONSISTENCIA = "O campo 'VL_DESC' está maior que o somatório dos campos VL_ITEM e VL_DESP"
        SUGESTAO = "Investigar o motivo da inconsistência"
    
    ElseIf VL_DESC > VL_ITEM Then
        INCONSISTENCIA = "O campo 'VL_DESC' está maior que o campo VL_ITEM "
        SUGESTAO = "Investigar o motivo da inconsistência"
    
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
            
End Function

