Attribute VB_Name = "DTO_EFDContribuicoes"
Option Explicit

Public fnC180Contrib As New clsC180Contrib
Public fnC190Contrib As New clsC190Contrib

'Bloco 0
Public Campos0000_Contr As CamposReg0000
Public Campos0110 As CamposReg0110


'Bloco A
Public CamposA001 As CamposRegA001
Public CamposA010 As CamposRegA010
Public CamposA100 As CamposRegA100
Public CamposA170 As CamposRegA170


'Bloco C
Public CamposC180Contr As CamposRegC180Contr
Public CamposC181Contr As CamposRegC181Contr
Public CamposC185Contr As CamposRegC185Contr
Public CamposC190Contr As CamposRegC190Contr


'Bloco D
Public CamposD010 As CamposRegD010
Public CamposD101_Contr As CamposRegD101Contr
Public CamposD105 As CamposRegD105
Public CamposD200 As CamposRegD200
Public CamposD201 As CamposRegD201
Public CamposD205 As CamposRegD205


'Bloco M
Public CamposM001 As CamposRegM001
Public CamposM100 As CamposRegM100
Public CamposM200 As CamposRegM200
Public CamposM210 As CamposRegM210
Public CamposM500 As CamposRegM500
Public CamposM600 As CamposRegM600
Public CamposM610 As CamposRegM610
Public CamposM990 As CamposRegM990

'#############################################################################################################################################
' REGISTRO DO BLOCO 0 DO SPED CONTRIBUIÇÕES
'#############################################################################################################################################

Public Type CamposReg0000
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_VER As String
    TIPO_ESCRIT As String
    IND_SIT_ESP As String
    NUM_REC_ANTERIOR As String
    DT_INI As String
    DT_FIN As String
    NOME As String
    CNPJ As String
    CPF As String
    UF As String
    IE As String
    COD_MUN As String
    SUFRAMA As String
    IND_NAT_PJ As String
    IND_ATIV As String
    
End Type

Public Type CamposReg0110
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_INC_TRIB As String
    IND_APRO_CRED As String
    COD_TIPO_CONT As String
    IND_REG_CUM As String
    
End Type

Public Type CamposRegA001
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_MOV As String
    
End Type

Public Type CamposRegA010
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    CNPJ As String
    
End Type


Public Type CamposRegA100
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_OPER As String
    IND_EMIT As String
    COD_PART As String
    COD_SIT As String
    SER As String
    SUB As String
    NUM_DOC As String
    CHV_NFSE As String
    DT_DOC As String
    DT_EXE_SERV As String
    VL_DOC As String
    IND_PGTO As String
    VL_DESC As String
    VL_BC_PIS As String
    VL_PIS As String
    VL_BC_COFINS As String
    VL_COFINS As String
    VL_PIS_RET As String
    VL_COFINS_RET As String
    VL_ISS As String
    
End Type

Public Type CamposRegA170
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    NUM_ITEM As String
    COD_ITEM As String
    DESCR_COMPL As String
    VL_ITEM As String
    VL_DESC As String
    NAT_BC_CRED As String
    IND_ORIG_CRED As String
    CST_PIS As String
    VL_BC_PIS As String
    ALIQ_PIS As String
    VL_PIS As String
    CST_COFINS As String
    VL_BC_COFINS As String
    ALIQ_COFINS As String
    VL_COFINS As String
    COD_CTA As String
    COD_CCUS As String
    
End Type

Public Type CamposRegC180Contr
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_MOD As String
    DT_DOC_INI As String
    DT_DOC_FIN As String
    COD_ITEM As String
    COD_NCM As String
    EX_IPI As String
    VL_TOT_ITEM As Double
    
End Type

Public Type CamposRegC181Contr
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    CST_PIS As String
    CFOP As String
    VL_ITEM As String
    VL_DESC As String
    VL_BC_PIS As String
    ALIQ_PIS As String
    QUANT_BC_PIS As String
    ALIQ_PIS_QUANT As String
    VL_PIS As String
    COD_CTA As String
    
End Type

Public Type CamposRegC185Contr
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    CST_COFINS As String
    CFOP As String
    VL_ITEM As String
    VL_DESC As String
    VL_BC_COFINS As String
    ALIQ_COFINS As String
    QUANT_BC_COFINS As String
    ALIQ_COFINS_QUANT As String
    VL_COFINS As String
    COD_CTA As String
    
End Type

Public Type CamposRegC190Contr
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_MOD As String
    DT_REF_INI As String
    DT_REF_FIN As String
    COD_ITEM As String
    COD_NCM As String
    EX_IPI As String
    VL_TOT_ITEM As Double
    
End Type

Public Type CamposRegD101Contr
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_NAT_FRT As String
    VL_ITEM As String
    CST_PIS As String
    NAT_BC_CRED As String
    VL_BC_PIS As String
    ALIQ_PIS As String
    VL_PIS As String
    COD_CTA As String
    
End Type

Public Type CamposRegD010
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    CNPJ As String
    
End Type

Public Type CamposRegD105
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_NAT_FRT As String
    VL_ITEM As String
    CST_COFINS As String
    NAT_BC_CRED As String
    VL_BC_COFINS As String
    ALIQ_COFINS As String
    VL_COFINS As String
    COD_CTA As String
    
End Type

Public Type CamposRegD200
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_MOD As String
    COD_SIT As String
    SER As String
    SUB As String
    NUM_DOC_INI As String
    NUM_DOC_FIN As String
    CFOP As String
    DT_REF As String
    VL_DOC As String
    VL_DESC As String
    
End Type

Public Type CamposRegD201
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    CST_PIS As String
    VL_ITEM As String
    VL_BC_PIS As String
    ALIQ_PIS As String
    VL_PIS As String
    COD_CTA As String
    
End Type

Public Type CamposRegD205
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    CST_COFINS As String
    VL_ITEM As String
    VL_BC_COFINS As String
    ALIQ_COFINS As String
    VL_COFINS As String
    COD_CTA As String
    
End Type

