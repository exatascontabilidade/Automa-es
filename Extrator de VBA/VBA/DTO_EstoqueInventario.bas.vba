Attribute VB_Name = "DTO_EstoqueInventario"
Option Explicit

Public dtoMovimentoEstoque As CamposMovimentoEstoque
Public dtoSaldoInventario As CamposSaldoInventario

Public Type CamposMovimentoEstoque
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    Chave As String
    COD_MOD As String
    SER As String
    NUM_DOC As String
    CHV_NFE As String
    DT_DOC As String
    DT_E_S As String
    DT_OPERACAO As String
    NUM_ITEM As Integer
    COD_ITEM As String
    DESCR_ITEM As String
    COD_BARRA As String
    COD_NCM As String
    EX_IPI As String
    CEST As String
    TIPO_ITEM As String
    COD_PART As String
    NOME_RAZAO_SOCIAL As String
    IND_MOV As String
    CFOP As String
    QTD_COM As Double
    UNID_COM As String
    FAT_CONV As Double
    QTD_INV As Double
    UNID_INV As String
    VL_ITEM As Double
    VL_DESC As Double
    VL_PIS As Double
    VL_COFINS As Double
    VL_UNIT_COM As Double
    VL_UNIT_INV As Double
    CST_ICMS As String
    VL_BC_ICMS As Double
    VL_ICMS As Double
    VL_ICMS_ST As Double
    VL_IPI As Double
    VL_CUSTO As Double
    VL_CUSTO_UNIT_COM As Double
    VL_CUSTO_UNIT_INV As Double
    VL_MERC As Double
    VL_FRT As Double
    VL_SEG As Double
    VL_OUT As Double
    VL_DESP As Double
    INCONSISTENCIA As String
    SUGESTAO As String
    
End Type

Public Type CamposSaldoInventario
    
    Chave As String
    DT_OPERACAO As String
    DT_INV As String
    COD_ITEM As String
    DESCR_ITEM As String
    VL_ITEM As Double
    VL_INICIAL As Double
    VL_ENT As Double
    VL_SAI As Double
    VL_UNIT_ENT As Double
    VL_UNIT_SAI As Double
    VL_MARGEM As Double
    ALIQ_MARGEM As Double
    QTD_INICIAL As Double
    QTD_ENT As Double
    QTD_SAI As Double
    QTD_FINAL As Double
    QTD_INV As Double
    IND_PROP As String
    COD_PART As String
    CFOP As String
    INCONSISTENCIA As String
    SUGESTAO As String
    
End Type

Public Function ResetarDTO_MovimentoEstoque()

Dim EstoqueVazio As CamposMovimentoEstoque
    
    LSet dtoMovimentoEstoque = EstoqueVazio
    
End Function

Public Function ResetarDTO_SaldoInventario()

Dim SaldoVazio As CamposSaldoInventario
    
    LSet dtoSaldoInventario = SaldoVazio
    
End Function
