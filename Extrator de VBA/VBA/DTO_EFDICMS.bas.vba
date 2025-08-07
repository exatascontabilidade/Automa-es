Attribute VB_Name = "DTO_EFDICMS"
Option Explicit

'Bloco 0 - EFD-ICMS/IPI
Public Campos0000 As CamposReg0000
Public Campos0001 As CamposReg0001
Public Campos0002 As CamposReg0002
Public Campos0005 As CamposReg0005
Public Campos0015 As CamposReg0015
Public Campos0100 As CamposReg0100
Public Campos0140 As CamposReg0140
Public Campos0150 As CamposReg0150
Public Campos0175 As CamposReg0175
Public Campos0190 As CamposReg0190
Public Campos0200 As CamposReg0200
Public Campos0205 As CamposReg0205
Public Campos0206 As CamposReg0206
Public Campos0210 As CamposReg0210
Public Campos0220 As CamposReg0220
Public Campos0221 As CamposReg0221
Public Campos0300 As CamposReg0300
Public Campos0305 As CamposReg0305
Public Campos0400 As CamposReg0400
Public Campos0450 As CamposReg0450
Public Campos0460 As CamposReg0460
Public Campos0500 As CamposReg0500
Public Campos0600 As CamposReg0600
Public Campos0990 As CamposReg0990


'Bloco B
Public CamposB001 As CamposRegB001
Public CamposB990 As CamposRegB990


'Bloco C
Public CamposC001 As CamposRegC001
Public CamposC010 As CamposRegC010
Public CamposC100 As CamposRegC100
Public CamposC101 As CamposRegC101
Public CamposC105 As CamposRegC105
Public CamposC110 As CamposRegC110
Public CamposC111 As CamposRegC111
Public CamposC112 As CamposRegC112
Public CamposC113 As CamposRegC113
Public CamposC114 As CamposRegC114
Public CamposC115 As CamposRegC115
Public CamposC116 As CamposRegC116
Public CamposC120 As CamposRegC120
Public CamposC130 As CamposRegC130
Public CamposC140 As CamposRegC140
Public CamposC141 As CamposRegC141
Public CamposC160 As CamposRegC160
Public CamposC165 As CamposRegC165
Public CamposC170 As CamposRegC170
Public CamposC171 As CamposRegC171
Public CamposC172 As CamposRegC172
Public CamposC173 As CamposRegC173
Public CamposC174 As CamposRegC174
Public CamposC175 As CamposRegC175
Public CamposC175Contrib As CamposregC175_Contr
Public CamposC176 As CamposRegC176
Public CamposC177 As CamposRegC177
Public CamposC178 As CamposRegC178
Public CamposC179 As CamposRegC179
Public CamposC180 As CamposRegC180
Public CamposC181 As CamposRegC181
Public Camposc185 As CamposRegC185
Public CamposC186 As CamposRegC186
Public CamposC190 As CamposRegC190
Public CamposC191 As CamposRegC191
Public CamposC195 As CamposRegC195
Public CamposC197 As CamposRegC197
Public CamposC400 As CamposRegC400
Public CamposC405 As CamposRegC405
Public CamposC410 As CamposRegC410
Public CamposC420 As CamposRegC420
Public CamposC425 As CamposRegC425
Public CamposC430 As CamposRegC430
Public CamposC460 As CamposRegC460
Public CamposC465 As CamposRegC465
Public CamposC470 As CamposRegC470
Public CamposC480 As CamposRegC480
Public CamposC490 As CamposRegC490
Public CamposC495 As CamposRegC495
Public CamposC500 As CamposRegC500
Public CamposC590 As CamposRegC590
Public CamposC800 As CamposRegC800
Public CamposC810 As CamposRegC810
Public CamposC815 As CamposRegC815
Public CamposC850 As CamposRegC850
Public CamposC855 As CamposRegC855
Public CamposC857 As CamposRegC857
Public CamposC860 As CamposRegC860
Public CamposC870 As CamposRegC870
Public CamposC880 As CamposRegC880
Public CamposC890 As CamposRegC890
Public CamposC895 As CamposRegC895
Public CamposC897 As CamposRegC897
Public CamposC990 As CamposRegC990


'Bloco D
Public CamposD001 As CamposRegD001
Public CamposD100 As CamposRegD100
Public CamposD101 As CamposRegD101
Public CamposD190 As CamposRegD190
Public CamposD195 As CamposRegD195
Public CamposD197 As CamposRegD197
Public CamposD500 As CamposRegD500
Public CamposD510 As CamposRegD510
Public CamposD530 As CamposRegD530
Public CamposD590 As CamposRegD590
Public CamposD990 As CamposRegD990


'Bloco E
Public CamposE001 As CamposRegE001
Public CamposE100 As CamposRegE100
Public CamposE110 As CamposRegE110
Public CamposE111 As CamposRegE111
Public CamposE112 As CamposRegE112
Public CamposE113 As CamposRegE113
Public CamposE115 As CamposRegE115
Public CamposE116 As CamposRegE116
Public CamposE200 As CamposRegE200
Public CamposE210 As CamposRegE210
Public CamposE220 As CamposRegE220
Public CamposE230 As CamposRegE230
Public CamposE240 As CamposRegE240
Public CamposE250 As CamposRegE250
Public CamposE300 As CamposRegE300
Public CamposE310 As CamposRegE310
Public CamposE311 As CamposRegE311
Public CamposE312 As CamposRegE312
Public CamposE313 As CamposRegE313
Public CamposE316 As CamposRegE316
Public CamposE500 As CamposRegE500
Public CamposE510 As CamposRegE510
Public CamposE520 As CamposRegE520
Public CamposE530 As CamposRegE530
Public CamposE531 As CamposRegE531
Public CamposE990 As CamposRegE990


'Bloco G
Public CamposG001 As CamposRegG001
Public CamposG110 As CamposRegG110
Public CamposG125 As CamposRegG125
Public CamposG126 As CamposRegG126
Public CamposG130 As CamposRegG130
Public CamposG140 As CamposRegG140
Public CamposG990 As CamposRegG990


'Bloco H
Public CamposH001 As CamposRegH001
Public CamposH005 As CamposRegH005
Public CamposH010 As CamposRegH010
Public CamposH020 As CamposRegH020
Public CamposH030 As CamposRegH030
Public CamposH990 As CamposRegH990


'Bloco K
Public CamposK001 As CamposRegK001
Public CamposK010 As CamposRegK010
Public CamposK100 As CamposRegK100
Public CamposK200 As CamposRegK200
Public CamposK230 As CamposRegK230
Public CamposK280 As CamposRegK280
Public CamposK990 As CamposRegK990


'Bloco 1
Public Campos1001 As CamposReg1001
Public Campos1010 As CamposReg1010
Public Campos1100 As CamposReg1100
Public Campos1105 As CamposReg1105
Public Campos1300 As CamposReg1300
Public Campos1310 As CamposReg1310
Public Campos1320 As CamposReg1320
Public Campos1350 As CamposReg1350
Public Campos1360 As CamposReg1360
Public Campos1370 As CamposReg1370
Public Campos1390 As CamposReg1390
Public Campos1391 As CamposReg1391
Public Campos1400 As CamposReg1400
'Public Campos1500 As CamposReg1500
'Public Campos1510 As CamposReg1510
'Public Campos1600 As CamposReg1600
Public Campos1601 As CamposReg1601
'Public Campos1700 As CamposReg1700
'Public Campos1710 As CamposReg1710
'Public Campos1800 As CamposReg1800
'Public Campos1900 As CamposReg1900
'Public Campos1910 As CamposReg1910
'Public Campos1920 As CamposReg1920
'Public Campos1921 As CamposReg1921
'Public Campos1922 As CamposReg1922
'Public Campos1923 As CamposReg1923
'Public Campos1925 As CamposReg1925
'Public Campos1926 As CamposReg1926
'Public Campos1960 As CamposReg1960
'Public Campos1970 As CamposReg1970
'Public Campos1975 As CamposReg1975
'Public Campos1980 As CamposReg1980
Public Campos1990 As CamposReg1990

'###############################################
' REGISTRO DO BLOCO 0 DO SPED FISCAL
'###############################################

Public Type CamposReg0000
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_VER As String
    COD_FIN As String
    DT_INI As String
    DT_FIN As String
    NOME As String
    CNPJ As String
    CPF As String
    UF As String
    IE As String
    COD_MUN As String
    IM As String
    SUFRAMA As String
    IND_PERFIL As String
    IND_ATIV As String
    
End Type

Public Type CamposReg0001
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_MOV As String
    
End Type

Public Type CamposReg0002
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    CLAS_ESTAB_IND As String
    
End Type

Public Type CamposReg0005
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    FANTASIA As String
    CEP As String
    END As String
    NUM As String
    COMPL As String
    BAIRRO As String
    FONE As String
    FAX As String
    EMAIL As String
    
End Type

Public Type CamposReg0015
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    UF_ST As String
    IE_ST As String
    
End Type

Public Type CamposReg0100
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    NOME As String
    CPF As String
    CRC As String
    CNPJ As String
    CEP As String
    END As String
    NUM As String
    COMPL As String
    BAIRRO As String
    FONE As String
    FAX As String
    EMAIL As String
    COD_MUN As String
    
End Type

Public Type CamposReg0140
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_EST As String
    NOME As String
    CNPJ As String
    UF As String
    IE As String
    COD_MUN As String
    IM As String
    SUFRAMA As String
    
End Type

Public Type CamposReg0150
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_PART As String
    NOME As String
    COD_PAIS As String
    CNPJ As String
    CPF As String
    IE As String
    COD_MUN As String
    SUFRAMA As String
    END As String
    NUM As String
    COMPL As String
    BAIRRO As String
    
End Type

Public Type CamposReg0175
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    DT_ALT As String
    NR_CAMPO As String
    CONT_ANT As String
    
End Type

Public Type CamposReg0190
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    UNID As String
    DESCR As String
    
End Type

Public Type CamposReg0200
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_ITEM As String
    DESCR_ITEM As String
    COD_BARRA As String
    COD_ANT_ITEM As String
    UNID_INV As String
    TIPO_ITEM As String
    COD_NCM As String
    EX_IPI As String
    COD_GEN As String
    COD_LST As String
    ALIQ_ICMS As String
    CEST As String
    
End Type

Public Type CamposReg0205
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    DESCR_ANT_ITEM As String
    DT_INI As String
    DT_FIM As String
    COD_ANT_ITEM As String
    
End Type

Public Type CamposReg0206
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_COMB As String
    
End Type

Public Type CamposReg0210

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_ITEM_COMP As String
    QTD_COMP As String
    PERDA As String
    
End Type

Public Type CamposReg0220
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    UNID_CONV As String
    FAT_CONV As String
    COD_BARRA As String
    
End Type

Public Type CamposReg0221
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_ITEM_ATOMICO As String
    QTD_CONTIDA As String
    
End Type

Public Type CamposReg0300
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_IND_BEM As String
    IDENT_MERC As String
    DESCR_ITEM As String
    COD_PRNC As String
    COD_CTA As String
    NR_PARC As String
    
End Type

Public Type CamposReg0305
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_CCUS As String
    FUNC As String
    VIDA_UTIL As String
    
End Type

Public Type CamposReg0400
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_NAT As String
    DESCR_NAT As String
    
End Type

Public Type CamposReg0450
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_INF As String
    TXT As String
    
End Type

Public Type CamposReg0460
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_OBS As String
    TXT As String
    
End Type

Public Type CamposReg0500
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    DT_ALT As String
    COD_NAT_CC As String
    IND_CTA As String
    NÍVEL As String
    COD_CTA As String
    NOME_CTA As String
        
End Type

Public Type CamposReg0600
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    DT_ALT As String
    COD_CCUS As String
    CCUS As String
        
End Type

Public Type CamposReg0990
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    QTD_LIN_0 As String
    
End Type


'###############################################
' REGISTRO DO BLOCO 'B' DO SPED FISCAL
'###############################################

Public Type CamposRegB001
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_DAD As String
    
End Type

Public Type CamposRegB990
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    QTD_LIN_B As String
    
End Type


'###############################################
' REGISTRO DO BLOCO 'C' DO SPED FISCAL
'###############################################

Public Type CamposRegC001
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_MOV As String
    
End Type

Public Type CamposRegC010
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    CNPJ As String
    IND_ESCRI As String
    
End Type

Public Type CamposRegC100
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_OPER As String
    IND_EMIT As String
    COD_PART As String
    COD_MOD As String
    COD_SIT As String
    SER As String
    NUM_DOC As String
    CHV_NFE As String
    DT_DOC As String
    DT_E_S As String
    VL_DOC As String
    VL_DOC_CALC As String
    IND_PGTO As String
    VL_DESC As String
    VL_ABAT_NT As String
    VL_MERC As String
    IND_FRT As String
    VL_FRT As String
    VL_SEG As String
    VL_OUT_DA As String
    VL_DESP As String
    VL_BC_ICMS As String
    VL_ICMS As String
    VL_FCP As String
    VL_BC_ICMS_ST As String
    VL_ICMS_ST As String
    VL_FCP_ST As String
    VL_IPI As String
    VL_PIS As String
    VL_COFINS As String
    VL_PIS_ST As String
    VL_COFINS_ST As String
    
End Type

Public Type CamposRegC101

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    VL_FCP_UF_DEST As String
    VL_ICMS_UF_DEST As String
    VL_ICMS_UF_REM As String
    
End Type

Public Type CamposRegC105

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    OPER As String
    UF As String
    
End Type

Public Type CamposRegC110
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_INF As String
    TXT_COMPL As String
    
End Type

Public Type CamposRegC111
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    NUM_PROC As String
    IND_PROC As String
    
End Type

Public Type CamposRegC112
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_DA As String
    UF As String
    NUM_DA As String
    COD_AUT As String
    VL_DA As String
    DT_VCTO As String
    DT_PGTO As String
    
End Type

Public Type CamposRegC113
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_OPER As String
    IND_EMIT As String
    COD_PART As String
    COD_MOD As String
    SER As String
    SUB As String
    NUM_DOC As String
    DT_DOC As String
    CHV_DOCE As String
    
End Type

Public Type CamposRegC114
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_MOD As String
    ECF_FAB As String
    ECF_CX As String
    NUM_DOC As String
    DT_DOC As String
    
End Type

Public Type CamposRegC115
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_CARGA As String
    CNPJ_COL As String
    IE_COL As String
    CPF_COL As String
    COD_MUN_COL As String
    CNPJ_ENTG As String
    IE_ENTG As String
    CPF_ENTG As String
    COD_MUN_ENTG As String
    
End Type

Public Type CamposRegC116
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_MOD As String
    NR_SAT As String
    CHV_CFE As String
    NUM_CFE As String
    DT_DOC As String
    
End Type

Public Type CamposRegC120
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_DOC_IMP As String
    NUM_DOC_IMP As String
    PIS_IMP As String
    COFINS_IMP As String
    NUM_ACDRAW As String
    
End Type

Public Type CamposRegC130
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    VL_SERV_NT As String
    VL_BC_ISSQN As String
    VL_ISSQN As String
    VL_BC_IRRF As String
    VL_ As String
    VL_BC_PREV As String
    VL_PREV As String

End Type

Public Type CamposRegC140
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_EMIT As String
    IND_TIT As String
    DESC_TIT As String
    NUM_TIT As String
    QTD_PARC As String
    VL_TIT As String
    
End Type

Public Type CamposRegC141
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    NUM_PARC As String
    DT_VCTO As String
    VL_PARC As String
    
End Type

Public Type CamposRegC160
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_PART As String
    VEIC_ID As String
    QTD_VOL As String
    PESO_BRT As String
    PESO_LIQ As String
    UF_ID As String
    
End Type

Public Type CamposRegC165
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_PART As String
    VEIC_ID As String
    COD_AUT As String
    NR_PASSE As String
    HORA As String
    TEMPER As String
    QTD_VOL As String
    PESO_BRT As String
    PESO_LIQ As String
    NOM_MOT As String
    CPF As String
    UF_ID As String
    
End Type

Public Type CamposRegC170
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    NUM_ITEM As String
    COD_ITEM As String
    DESCR_COMPL As String
    QTD As String
    UNID As String
    VL_ITEM As String
    VL_DESC As String
    IND_MOV As String
    CST_ICMS As String
    CFOP As String
    COD_NAT As String
    VL_BC_ICMS As String
    ALIQ_ICMS As String
    VL_ICMS As String
    VL_BC_ICMS_ST As String
    ALIQ_ST As String
    VL_ICMS_ST As String
    IND_APUR As String
    CST_IPI As String
    COD_ENQ As String
    VL_BC_IPI As String
    ALIQ_IPI As String
    VL_IPI As String
    CST_PIS As String
    VL_BC_PIS As String
    ALIQ_PIS As String
    QUANT_BC_PIS As String
    ALIQ_PIS_QUANT As String
    VL_PIS As String
    CST_COFINS As String
    VL_BC_COFINS As String
    ALIQ_COFINS As String
    QUANT_BC_COFINS As String
    ALIQ_COFINS_QUANT As String
    VL_COFINS As String
    COD_CTA As String
    VL_ABAT_NT As String
    DESCR_ITEM  As String
    TIPO_ITEM As String
    
End Type

Public Type CamposRegC171
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    NUM_TANQUE As String
    QTDE As String
    
End Type

Public Type CamposRegC172
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    VL_BC_ISSQN As String
    ALIQ_ISSQN As String
    VL_ISSQN As String
    
End Type

Public Type CamposRegC173
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    LOTE_MED As String
    QTD_ITEM As String
    DT_FAB As String
    DT_VAL As String
    IND_MED As String
    TP_PROD As String
    VL_TAB_MAX As String
    
End Type

Public Type CamposRegC174
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_ARM As String
    NUM_ARM As String
    DESCR_COMPL As String
    
End Type

Public Type CamposRegC175
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_VEIC_OPER As String
    CNPJ As String
    UF As String
    CHASSI_VEIC As String
    
End Type

Public Type CamposregC175_Contr
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    CFOP As String
    VL_OPER As String
    VL_DESC As String
    CST_PIS As String
    VL_BC_PIS As String
    ALIQ_PIS As String
    QUANT_BC_PIS As String
    ALIQ_PIS_QUANT As String
    VL_PIS As String
    CST_COFINS As String
    VL_BC_COFINS As String
    ALIQ_COFINS As String
    QUANT_BC_COFINS As String
    ALIQ_COFINS_QUANT As String
    VL_COFINS As String
    COD_CTA As String
    INFO_COMPL As String
    VL_ICMS As String
    
End Type

Public Type CamposRegC176
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_MOD_ULT_E As String
    NUM_DOC_ULT_E As String
    SER_ULT_E As String
    DT_ULT_E As String
    COD_PART_ULT_E As String
    QUANT_ULT_E As String
    VL_UNIT_ULT_E As String
    VL_UNIT_BC_ST As String
    CHAVE_NFE_ULT_E As String
    NUM_ITEM_ULT_E As String
    VL_UNIT_BC_ICMS_ULT_E As String
    ALIQ_ICMS_ULT_E As String
    VL_UNIT_LIMITE_BC_ICMS_ULT_E As String
    VL_UNIT_ICMS_ULT_E As String
    ALIQ_ST_ULT_E As String
    VL_UNIT_RES As String
    COD_RESP_RET As String
    COD_MOT_RES As String
    CHAVE_NFE_RET As String
    COD_PART_NFE_RET As String
    SER_NFE_RET As String
    NUM_NFE_RET As String
    ITEM_NFE_RET As String
    COD_DA As String
    NUM_DA As String
    VL_UNIT_RES_FCP_ST As String
    
End Type

Public Type CamposRegC177
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_INF_ITEM As String
    
End Type

Public Type CamposRegC178
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    CL_ENQ As String
    VL_UNID As String
    QUANT_PAD As String
    
End Type

Public Type CamposRegC179
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    BC_ST_ORIG_DEST As String
    ICMS_ST_REP As String
    ICMS_ST_COMPL As String
    BC_RET As String
    ICMS_RET As String
    
End Type

Public Type CamposRegC180
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_RESP_RET As String
    QUANT_CONV As String
    UNID As String
    VL_UNIT_CONV As String
    VL_UNIT_ICMS_OP_CONV As String
    VL_UNIT_BC_ICMS_ST_CONV As String
    VL_UNIT_ICMS_ST_CONV As String
    VL_UNIT_FCP_ST_CONV As String
    COD_DA As String
    NUM_DA As String
    
End Type

Public Type CamposRegC181
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_MOT_REST_COMPL As String
    QUANT_CONV As String
    UNID As String
    COD_MOD_SAIDA As String
    SERIE_SAIDA As String
    ECF_FAB_SAIDA As String
    NUM_DOC_SAIDA As String
    CHV_DFE_SAIDA As String
    DT_DOC_SAIDA As String
    NUM_ITEM_SAIDA As String
    VL_UNIT_CONV_SAIDA As String
    VL_UNIT_ICMS_OP_ESTOQUE_CONV_SAIDA As String
    VL_UNIT_ICMS_ST_ESTOQUE_CONV_SAIDA As String
    VL_UNIT_FCP_ICMS_ST_ESTOQUE_CONV_SAIDA As String
    VL_UNIT_ICMS_NA_OPERACAO_CONV_SAIDA As String
    VL_UNIT_ICMS_OP_CONV_SAIDA As String
    VL_UNIT_ICMS_ST_CONV_REST As String
    VL_UNIT_FCP_ST_CONV_REST As String
    VL_UNIT_ICMS_ST_CONV_COMPL As String
    VL_UNIT_FCP_ST_CONV_COMPL As String
    
End Type

Public Type CamposRegC185
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    NUM_ITEM As String
    COD_ITEM As String
    CST_ICMS As String
    CFOP As String
    COD_MOT_REST_COMPL As String
    QUANT_CONV As String
    UNID As String
    VL_UNIT_CONV As String
    VL_UNIT_ICMS_NA_OPERACAO_CONV As String
    VL_UNIT_ICMS_OP_CONV As String
    VL_UNIT_ICMS_OP_ESTOQUE_CONV As String
    VL_UNIT_ICMS_ST_ESTOQUE_CONV As String
    VL_UNIT_FCP_ICMS_ST_ESTOQUE_CONV As String
    VL_UNIT_ICMS_ST_CONV_REST As String
    VL_UNIT_FCP_ST_CONV_REST As String
    VL_UNIT_ICMS_ST_CONV_COMPL As String
    VL_UNIT_FCP_ST_CONV_COMPL As String
    
End Type

Public Type CamposRegC186
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    NUM_ITEM As String
    COD_ITEM As String
    CST_ICMS As String
    CFOP As String
    COD_MOT_REST_COMPL As String
    QUANT_CONV As String
    UNID As String
    COD_MOD_ENTRADA As String
    SERIE_ENTRADA As String
    NUM_DOC_ENTRADA As String
    CHV_DFE_ENTRADA As String
    DT_DOC_ENTRADA As String
    NUM_ITEM_ENTRADA As String
    VL_UNIT_CONV_ENTRADA As String
    VL_UNIT_ICMS_OP_CONV_ENTRADA As String
    VL_UNIT_BC_ICMS_ST_CONV_ENTRADA As String
    VL_UNIT_ICMS_ST_CONV_ENTRADA As String
    VL_UNIT_FCP_ST_CONV_ENTRADA As String
    
End Type

Public Type CamposRegC190

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_ENFOQUE As String
    CST_ICMS As String
    CFOP As String
    ALIQ_ICMS As String
    VL_OPR As String
    VL_BC_ICMS As String
    VL_ICMS As String
    VL_BC_ICMS_ST As String
    VL_ICMS_ST As String
    VL_RED_BC As String
    VL_IPI As String
    COD_OBS As String
    UF As String
    
End Type

Public Type CamposRegC191

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    VL_FCP_OP As String
    VL_FCP_ST As String
    VL_FCP_RET As String
    
End Type

Public Type CamposRegC195

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_OBS As String
    TXT_COMPL As String
    
End Type

Public Type CamposRegC197

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_AJ As String
    DESCR_COMPL_AJ As String
    COD_ITEM As String
    VL_BC_ICMS As String
    ALIQ_ICMS As String
    VL_ICMS As String
    VL_OUTROS As String
    
End Type


Public Type CamposRegC400
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_MOD As String
    ECF_MOD As String
    ECF_FAB As String
    ECF_CX As String
                
End Type


Public Type CamposRegC405
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    DT_DOC As String
    CRO As String
    CRZ As String
    NUM_COO_FIN As String
    GT_FIN As String
    VL_BRT As String
                
End Type


Public Type CamposRegC410
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    VL_PIS As String
    VL_COFINS As String
                
End Type


Public Type CamposRegC420
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_TOT_PAR As String
    VLR_ACUM_TOT As String
    NR_TOT As String
    DESCR_NR_TOT As String
                
End Type


Public Type CamposRegC425
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_ITEM As String
    QTD As String
    UNID As String
    VL_ITEM As String
    VL_PIS As String
    VL_COFINS As String
                
End Type


Public Type CamposRegC430
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_MOT_REST_COMPL As String
    QUANT_CONV As String
    UNID As String
    VL_UNIT_CONV As String
    VL_UNIT_ICMS_NA_OPERACAO_CONV As String
    VL_UNIT_ICMS_OP_CONV As String
    VL_UNIT_ICMS_OP_ESTOQUE_CONV As String
    VL_UNIT_ICMS_ST_ESTOQUE_CONV As String
    VL_UNIT_FCP_ICMS_ST_ESTOQUE_CONV As String
    VL_UNIT_ICMS_ST_CONV_REST As String
    VL_UNIT_FCP_ST_CONV_REST As String
    VL_UNIT_ICMS_ST_CONV_COMPL As String
    VL_UNIT_FCP_ST_CONV_COMPL As String
    CST_ICMS As String
    CFOP As String
                
End Type


Public Type CamposRegC460
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_MOD As String
    COD_SIT As String
    NUM_DOC As String
    DT_DOC As String
    VL_DOC As String
    VL_PIS As String
    VL_COFINS As String
    CPF_CNPJ As String
    NOM_ADQ As String
                
End Type


Public Type CamposRegC465
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    CHV_CFE As String
    NUM_CCF As String
                
End Type


Public Type CamposRegC470
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_ITEM As String
    QTD As String
    QTD_CANC As String
    UNID As String
    VL_ITEM As String
    CST_ICMS As String
    CFOP As String
    ALIQ_ICMS As String
    VL_PIS As String
    VL_COFINS As String
    
End Type


Public Type CamposRegC480
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_MOT_REST_COMPL As String
    QUANT_CONV As String
    UNID As String
    VL_UNIT_CONV As String
    VL_UNIT_ICMS_NA_OPERACAO_CONV As String
    VL_UNIT_ICMS_OP_CONV As String
    VL_UNIT_ICMS_OP_ESTOQUE_CONV As String
    VL_UNIT_ICMS_ST_ESTOQUE_CONV As String
    VL_UNIT_FCP_ICMS_ST_ESTOQUE_CONV As String
    VL_UNIT_ICMS_ST_CONV_REST As String
    VL_UNIT_FCP_ST_CONV_REST As String
    VL_UNIT_ICMS_ST_CONV_COMPL As String
    VL_UNIT_FCP_ST_CONV_COMPL As String
    CST_ICMS As String
    CFOP As String
    
End Type


Public Type CamposRegC490
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_ENFOQUE As String
    CST_ICMS As String
    CFOP As String
    ALIQ_ICMS As String
    VL_OPR As String
    VL_BC_ICMS As String
    VL_ICMS As String
    COD_OBS As String
    UF As String
    
End Type


Public Type CamposRegC495
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    ALIQ_ICMS As String
    COD_ITEM As String
    QTD As String
    QTD_CANC As String
    UNID As String
    VL_ITEM As String
    VL_DESC As String
    VL_CANC As String
    VL_ACMO As String
    VL_BC_ICMS As String
    VL_ICMS As String
    VL_ISEN As String
    VL_NT As String
    VL_ICMS_ST As String
    
End Type

Public Type CamposRegC500
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_OPER As String
    IND_EMIT As String
    COD_PART As String
    COD_MOD As String
    COD_SIT As String
    SER As String
    SUB As String
    COD_CONS As String
    NUM_DOC As String
    DT_DOC As String
    DT_E_S As String
    VL_DOC As String
    VL_DESC As String
    VL_FORN As String
    VL_SERV_NT As String
    VL_TERC As String
    VL_DA As String
    VL_BC_ICMS As String
    VL_ICMS As String
    VL_BC_ICMS_ST As String
    VL_ICMS_ST As String
    COD_INF As String
    VL_PIS As String
    VL_COFINS As String
    TP_LIGACAO As String
    COD_GRUPO_TENSAO As String
    CHV_DOCE As String
    FIN_DOCE As String
    CHV_DOCE_REF As String
    IND_DEST As String
    COD_MUN_DEST As String
    COD_CTA As String
    COD_MOD_DOC_REF As String
    HASH_DOC_REF As String
    SER_DOC_REF As String
    NUM_DOC_REF As String
    MES_DOC_REF As String
    ENER_INJET As String
    OUTRAS_DED As String
    
End Type

Public Type CamposRegC590
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    CST_ICMS As String
    CFOP As String
    ALIQ_ICMS As String
    VL_OPR As String
    VL_BC_ICMS As String
    VL_ICMS As String
    VL_BC_ICMS_ST As String
    VL_ICMS_ST As String
    VL_RED_BC As String
    COD_OBS As String
    COD_ENFOQUE As String
    UF As String
    
End Type

Public Type CamposRegC800

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_MOD As String
    COD_SIT As String
    NUM_CFE As String
    DT_DOC As String
    VL_CFE As String
    VL_PIS As String
    VL_COFINS As String
    CNPJ_CPF As String
    NR_SAT As String
    CHV_CFE As String
    VL_DESC As String
    VL_MERC As String
    VL_OUT_DA As String
    VL_ICMS As String
    VL_PIS_ST As String
    VL_COFINS_ST As String
    UF As String
    
End Type

Public Type CamposRegC810

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    NUM_ITEM As String
    COD_ITEM As String
    QTD As String
    UNID As String
    VL_ITEM As String
    CST_ICMS As String
    CFOP As String

End Type

Public Type CamposRegC815

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_MOT_REST_COMPL As String
    QUANT_CONV As String
    UNID As String
    VL_UNIT_CONV As String
    VL_UNIT_ICMS_NA_OPERACAO_CONV As String
    VL_UNIT_ICMS_OP_CONV As String
    VL_UNIT_ICMS_OP_ESTOQUE_CONV As String
    VL_UNIT_ICMS_ST_ESTOQUE_CONV As String
    VL_UNIT_FCP_ICMS_ST_ESTOQUE_CONV As String
    VL_UNIT_ICMS_ST_CONV_REST As String
    VL_UNIT_FCP_ST_CONV_REST As String
    VL_UNIT_ICMS_ST_CONV_COMPL As String
    VL_UNIT_FCP_ST_CONV_COMPL As String

End Type


Public Type CamposRegC850

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    CST_ICMS As String
    CFOP As String
    ALIQ_ICMS As String
    VL_OPR As String
    VL_BC_ICMS As String
    VL_ICMS As String
    COD_OBS As String
    COD_ENFOQUE As String
    UF As String
    
End Type

Public Type CamposRegC855

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_OBS As String
    TXT_COMPL As String

End Type

Public Type CamposRegC857

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_AJ As String
    DESCR_COMPL_AJ As String
    COD_ITEM As String
    VL_BC_ICMS As String
    ALIQ_ICMS As String
    VL_ICMS As String
    VL_OUTROS As String

End Type

Public Type CamposRegC860

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_MOD As String
    NR_SAT As String
    DT_DOC As String
    DOC_INI As String
    DOC_FIM As String

End Type

Public Type CamposRegC870

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_ITEM As String
    QTD As String
    UNID As String
    CST_ICMS As String
    CFOP As String

End Type

Public Type CamposRegC880

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_MOT_REST_COMPL As String
    QUANT_CONV As String
    UNID As String
    VL_UNIT_CONV As String
    VL_UNIT_ICMS_NA_OPERACAO_CONV As String
    VL_UNIT_ICMS_OP_CONV As String
    VL_UNIT_ICMS_OP_ESTOQUE_CONV As String
    VL_UNIT_ICMS_ST_ESTOQUE_CONV As String
    VL_UNIT_FCP_ICMS_ST_ESTOQUE_CONV As String
    VL_UNIT_ICMS_ST_CONV_REST As String
    VL_UNIT_FCP_ST_CONV_REST As String
    VL_UNIT_ICMS_ST_CONV_COMPL As String
    VL_UNIT_FCP_ST_CONV_COMPL As String

End Type

Public Type CamposRegC890

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    CST_ICMS As String
    CFOP As String
    ALIQ_ICMS As String
    VL_OPR As String
    VL_BC_ICMS As String
    VL_ICMS As String
    COD_OBS As String
    UF As String

End Type

Public Type CamposRegC895

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_OBS As String
    TXT_COMPL As String

End Type

Public Type CamposRegC897

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_AJ As String
    DESCR_COMPL_AJ As String
    COD_ITEM As String
    VL_BC_ICMS As String
    ALIQ_ICMS As String
    VL_ICMS As String
    VL_OUTROS As String

End Type

Public Type CamposRegC990
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    QTD_LIN_C As String
    
End Type

'###############################################
' REGISTRO DO BLOCO 'D' DO SPED FISCAL
'###############################################

Public Type CamposRegD001
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_MOV As String
    
End Type

Public Type CamposRegD100
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_OPER As String
    IND_EMIT As String
    COD_PART As String
    COD_MOD As String
    COD_SIT As String
    SER As String
    SUB As String
    NUM_DOC As String
    CHV_CTE As String
    DT_DOC As Variant
    DT_A_P As Variant
    TP_CTe As String
    CHV_CTE_REF As String
    VL_DOC As String
    VL_DESC As String
    IND_FRT As String
    VL_SERV As String
    VL_BC_ICMS As String
    VL_ICMS As String
    VL_NT As String
    COD_INF As String
    COD_CTA As String
    COD_MUN_ORIG As String
    COD_MUN_DEST As String
        
End Type

Public Type CamposRegD101

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    VL_FCP_UF_DEST As String
    VL_ICMS_UF_DEST As String
    VL_ICMS_UF_REM As String
    
End Type

Public Type CamposRegD190

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    CST_ICMS As String
    CFOP As String
    ALIQ_ICMS As String
    VL_OPR As String
    VL_BC_ICMS As String
    VL_ICMS As String
    VL_RED_BC As String
    COD_OBS As String
    COD_ENFOQUE As String
    UF As String
    
End Type

Public Type CamposRegD195

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_OBS As String
    TXT_COMPL As String
    
End Type

Public Type CamposRegD197

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_AJ As String
    DESCR_COMPL_AJ As String
    COD_ITEM As String
    VL_BC_ICMS As String
    ALIQ_ICMS As String
    VL_ICMS As String
    VL_OUTROS As String
    
End Type

Public Type CamposRegD500

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_OPER As String
    IND_EMIT As String
    COD_PART As String
    COD_MOD As String
    COD_SIT As String
    SER As String
    SUB As String
    NUM_DOC As String
    DT_DOC As String
    DT_A_P As String
    VL_DOC As String
    VL_DESC As String
    VL_SERV As String
    VL_SERV_NT As String
    VL_TERC As String
    VL_DA As String
    VL_BC_ICMS As String
    VL_ICMS As String
    COD_INF As String
    VL_PIS As String
    VL_COFINS As String
    COD_CTA As String
    TP_ASSINANTE As String
    
End Type

Public Type CamposRegD510

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    NUM_ITEM As String
    COD_ITEM As String
    COD_CLASS As String
    QTD As String
    UNID As String
    VL_ITEM As String
    VL_DESC As String
    CST_ICMS As String
    CFOP As String
    VL_BC_ICMS As String
    ALIQ_ICMS As String
    VL_ICMS As String
    VL_BC_ICMS_UF As String
    VL_ICMS_UF As String
    IND_REC As String
    COD_PART As String
    VL_PIS As String
    VL_COFINS As String
    COD_CTA As String

End Type

Public Type CamposRegD530

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_SERV As String
    DT_INI_SERV As String
    DT_FIN_SERV As String
    PER_FISCAL As String
    COD_AREA As String
    TERMINAL As String
    
End Type

Public Type CamposRegD590

    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    CST_ICMS As String
    CFOP As String
    ALIQ_ICMS As String
    VL_OPR As String
    VL_BC_ICMS As String
    VL_ICMS As String
    VL_BC_ICMS_UF As String
    VL_ICMS_UF As String
    VL_RED_BC As String
    COD_OBS As String
    COD_ENFOQUE As String
    UF As String
    
End Type

Public Type CamposRegD990
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    QTD_LIN_D As String
    
End Type



'###############################################
' REGISTRO DO BLOCO 'E' DO SPED FISCAL
'###############################################

Public Type CamposRegE001
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_MOV As String
    
End Type

Public Type CamposRegE100
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    DT_INI As String
    DT_FIN As String
    
End Type

Public Type CamposRegE110
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    VL_TOT_DEBITOS As String
    VL_AJ_DEBITOS As String
    VL_TOT_AJ_DEBITOS As String
    VL_ESTORNOS_CRED As String
    VL_TOT_CREDITOS As String
    VL_AJ_CREDITOS As String
    VL_TOT_AJ_CREDITOS As String
    VL_ESTORNOS_DEB As String
    VL_SLD_CREDOR_ANT As String
    VL_SLD_APURADO As String
    VL_TOT_DED As String
    VL_ICMS_RECOLHER As String
    VL_SLD_CREDOR_TRANSPORTAR As String
    DEB_ESP As String
    
End Type

Public Type CamposRegE111
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_AJ_APUR As String
    DESCR_COMPL_AJ As String
    VL_AJ_APUR As String
    
End Type

Public Type CamposRegE112
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    NUM_DA As String
    NUM_PROC As String
    IND_PROC As String
    PROC As String
    TXT_COMPL As String
    
End Type

Public Type CamposRegE113
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_PART As String
    COD_MOD As String
    SER As String
    SUB As String
    NUM_DOC As String
    DT_DOC As String
    COD_ITEM As String
    VL_AJ_ITEM As String
    CHV_DOCE As String
    
End Type

Public Type CamposRegE115
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_INF_ADIC As String
    VL_INF_ADIC As String
    DESCR_COMPL_AJ As String
        
End Type

Public Type CamposRegE116
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_OR As String
    VL_OR As String
    DT_VCTO As String
    COD_REC As String
    NUM_PROC As String
    IND_PROC As String
    PROC As String
    TXT_COMPL As String
    MES_REF As String
    
End Type

Public Type CamposRegE200
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    UF As String
    DT_INI As String
    DT_FIN As String
    
End Type

Public Type CamposRegE210
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_MOV_ST As String
    VL_SLD_CRED_ANT_ST As String
    VL_DEVOL_ST As String
    VL_RESSARC_ST As String
    VL_OUT_CRED_ST As String
    VL_AJ_CREDITOS_ST As String
    VL_RETENÇAO_ST As String
    VL_OUT_DEB_ST As String
    VL_AJ_DEBITOS_ST As String
    VL_SLD_DEV_ANT_ST As String
    VL_DEDUÇÕES_ST As String
    VL_ICMS_RECOL_ST As String
    VL_SLD_CRED_ST_TRANSPORTAR As String
    DEB_ESP_ST As String
    
End Type

Public Type CamposRegE220
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_AJ_APUR As String
    DESCR_COMPL_AJ As String
    VL_AJ_APUR As String
    
End Type

Public Type CamposRegE230
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    NUM_DA As String
    NUM_PROC As String
    IND_PROC As String
    PROC As String
    TXT_COMPL As String

End Type

Public Type CamposRegE240
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_PART As String
    COD_MOD As String
    SER As String
    SUB As String
    NUM_DOC As String
    DT_DOC As String
    COD_ITEM As String
    VL_AJ_ITEM As String
    CHV_DOCE As String

End Type

Public Type CamposRegE250
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_OR As String
    VL_OR As String
    DT_VCTO As String
    COD_REC As String
    NUM_PROC As String
    IND_PROC As String
    PROC As String
    TXT_COMPL As String
    MES_REF As String
    
End Type

Public Type CamposRegE300
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    UF As String
    DT_INI As String
    DT_FIN As String
    
End Type

Public Type CamposRegE310
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_MOV_FCP_DIFAL As String
    VL_SLD_CRED_ANT_DIFAL As String
    VL_TOT_DEBITOS_DIFAL As String
    VL_OUT_DEB_DIFAL As String
    VL_TOT_CREDITOS_DIFAL As String
    VL_OUT_CRED_DIFAL As String
    VL_SLD_DEV_ANT_DIFAL As String
    VL_DEDUCOES_DIFAL As String
    VL_RECOL_DIFAL As String
    VL_SLD_CRED_TRANSPORTAR_DIFAL As String
    DEB_ESP_DIFAL As String
    VL_SLD_CRED_ANT_FCP As String
    VL_TOT_DEB_FCP As String
    VL_OUT_DEB_FCP As String
    VL_TOT_CRED_FCP As String
    VL_OUT_CRED_FCP As String
    VL_SLD_DEV_ANT_FCP As String
    VL_DEDUCOES_FCP As String
    VL_RECOL_FCP As String
    VL_SLD_CRED_TRANSPORTAR_FCP As String
    DEB_ESP_FCP As String
    
End Type

Public Type CamposRegE311
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_AJ_APUR As String
    DESCR_COMPL_AJ As String
    VL_AJ_APUR As String
    
End Type

Public Type CamposRegE312
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    NUM_DA As String
    NUM_PROC As String
    IND_PROC As String
    PROC As String
    TXT_COMPL As String
    
End Type

Public Type CamposRegE313
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_PART As String
    COD_MOD As String
    SER As String
    SUB As String
    NUM_DOC As String
    CHV_DOCE As String
    DT_DOC As String
    COD_ITEM As String
    VL_AJ_ITEM As String
    
End Type

Public Type CamposRegE316
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_OR As String
    VL_OR As String
    DT_VCTO As String
    COD_REC As String
    NUM_PROC As String
    IND_PROC As String
    PROC As String
    TXT_COMPL As String
    MES_REF As String
    
End Type

Public Type CamposRegE500
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_APUR As String
    DT_INI As String
    DT_FIN As String
    
End Type

Public Type CamposRegE510
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    CFOP As String
    CST_IPI As String
    VL_CONT_IPI As String
    VL_BC_IPI As String
    VL_IPI As String
    
End Type

Public Type CamposRegE520
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    VL_SD_ANT_IPI As String
    VL_DEB_IPI As String
    VL_CRED_IPI As String
    VL_OD_IPI As String
    VL_OC_IPI As String
    VL_SC_IPI As String
    VL_SD_IPI As String
    
End Type

Public Type CamposRegE530
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_AJ As String
    VL_AJ As String
    COD_AJ As String
    IND_DOC As String
    NUM_DOC As String
    DESCR_AJ As String

End Type

Public Type CamposRegE531
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_PART As String
    COD_MOD As String
    SER As String
    SUB As String
    NUM_DOC As String
    DT_DOC As String
    COD_ITEM As String
    VL_AJ_ITEM As String
    CHV_NFE As String
    
End Type

Public Type CamposRegE990
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    QTD_LIN_E As String
    
End Type


'###############################################
' REGISTRO DO BLOCO 'G' DO SPED FISCAL
'###############################################

Public Type CamposRegG001
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_MOV As String
    
End Type

Public Type CamposRegG110
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    DT_INI As String
    DT_FIN As String
    SALDO_IN_ICMS As String
    SOM_PARC As String
    VL_TRIB_EXP As String
    VL_TOTAL As String
    IND_PER_SAI As String
    ICMS_APROP As String
    SOM_ICMS_OC As String
    
End Type

Public Type CamposRegG125
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_IND_BEM As String
    DT_MOV As String
    TIPO_MOV As String
    VL_IMOB_ICMS_OP As String
    VL_IMOB_ICMS_ST As String
    VL_IMOB_ICMS_FRT As String
    VL_IMOB_ICMS_DIF As String
    NUM_PARC As String
    VL_PARC_PASS As String
    
End Type


Public Type CamposRegG126
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    DT_INI As String
    DT_FIM As String
    NUM_PARC As String
    VL_PARC_PASS As String
    VL_TRIB_OC As String
    VL_TOTAL As String
    IND_PER_SAI As String
    VL_PARC_APROP As String
    
End Type

Public Type CamposRegG130
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_EMIT As String
    COD_PART As String
    COD_MOD As String
    SERIE As String
    NUM_DOC As String
    CHV_NFE_CTE As String
    DT_DOC As String
    NUM_DA As String
    
End Type

Public Type CamposRegG140
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    NUM_ITEM As String
    COD_ITEM As String
    QTDE As String
    UNID As String
    VL_ICMS_OP_APLICADO As String
    VL_ICMS_ST_APLICADO As String
    VL_ICMS_FRT_APLICADO As String
    VL_ICMS_DIF_APLICADO As String
    
End Type

Public Type CamposRegG990
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    QTD_LIN_G As String
    
End Type



'###############################################
' REGISTRO DO BLOCO 'H' DO SPED FISCAL
'###############################################

Public Type CamposRegH001
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_MOV As String
    
End Type

Public Type CamposRegH005
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    DT_INV As String
    VL_INV As String
    MOT_INV As String
    
End Type

Public Type CamposRegH010
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_ITEM As String
    UNID As String
    QTD As String
    VL_UNIT As String
    VL_ITEM As String
    IND_PROP As String
    COD_PART As String
    TXT_COMPL As String
    COD_CTA As String
    VL_ITEM_IR As String
    
End Type

Public Type CamposRegH020
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    CST_ICMS As String
    BC_ICMS As String
    VL_ICMS As String
    
End Type

Public Type CamposRegH030
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    VL_ICMS_OP As String
    VL_BC_ICMS_ST As String
    VL_ICMS_ST As String
    VL_FCP As String
    
End Type

Public Type CamposRegH990
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    QTD_LIN_H As String
    
End Type


'###############################################
' REGISTRO DO BLOCO 'K' DO SPED FISCAL
'###############################################

Public Type CamposRegK001
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_MOV As String
    
End Type

Public Type CamposRegK010
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_TP_LEIAUTE As String
    
End Type

Public Type CamposRegK100
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    DT_INI As String
    DT_FIN As String
    
End Type

Public Type CamposRegK200
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    DT_EST As String
    COD_ITEM As String
    QTD As String
    IND_EST As String
    COD_PART As String
    
End Type

Public Type CamposRegK280
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    DT_EST As String
    COD_ITEM As String
    QTD_COR_POS As String
    QTD_COR_NEG As String
    IND_EST As String
    COD_PART As String
    
End Type

Public Type CamposRegK230
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    DT_INI_OP As String
    DT_FIN_OP As String
    COD_DOC_OP As String
    COD_ITEM As String
    QTD_ENC As String
    
End Type

Public Type CamposRegK990
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    QTD_LIN_K As String
    
End Type

Public Type CamposRegM001
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_MOV As String
    
End Type

Public Type CamposRegM100
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_CRED As String
    IND_CRED_ORI As String
    VL_BC_PIS As String
    ALIQ_PIS As String
    QUANT_BC_PIS As String
    ALIQ_PIS_QUANT As String
    VL_CRED As String
    VL_AJUS_ACRES As String
    VL_AJUS_REDUC As String
    VL_CRED_DIF As String
    VL_CRED_DISP As String
    IND_DESC_CRED As String
    VL_CRED_DESC As String
    SLD_CRED As String
    
End Type

Public Type CamposRegM200
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    VL_TOT_CONT_NC_PER As String
    VL_TOT_CRED_DESC As String
    VL_TOT_CRED_DESC_ANT As String
    VL_TOT_CONT_NC_DEV As String
    VL_RET_NC As String
    VL_OUT_DED_NC As String
    VL_CONT_NC_REC As String
    VL_TOT_CONT_CUM_PER As String
    VL_RET_CUM As String
    VL_OUT_DED_CUM As String
    VL_CONT_CUM_REC As String
    VL_TOT_CONT_REC As String
    
End Type

Public Type CamposRegM210
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_CONT As String
    VL_REC_BRT As String
    VL_BC_CONT As String
    VL_AJUS_ACRES_BC_PIS As String
    VL_AJUS_REDUC_BC_PIS As String
    VL_BC_CONT_AJUS As String
    ALIQ_PIS As String
    QUANT_BC_PIS As String
    ALIQ_PIS_QUANT As String
    VL_CONT_APUR As String
    VL_AJUS_ACRES As String
    VL_AJUS_REDUC As String
    VL_CONT_DIFER As String
    VL_CONT_DIFER_ANT As String
    VL_CONT_PER As String
    
End Type

Public Type CamposRegM500
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_CRED As String
    IND_CRED_ORI As String
    VL_BC_COFINS As String
    ALIQ_COFINS As String
    QUANT_BC_COFINS As String
    ALIQ_COFINS_QUANT As String
    VL_CRED As String
    VL_AJUS_ACRES As String
    VL_AJUS_REDUC As String
    VL_CRED_DIFER As String
    VL_CRED_DISP As String
    IND_DESC_CRED As String
    VL_CRED_DESC As String
    SLD_CRED As String
    
End Type

Public Type CamposRegM600
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    VL_TOT_CONT_NC_PER As String
    VL_TOT_CRED_DESC As String
    VL_TOT_CRED_DESC_ANT As String
    VL_TOT_CONT_NC_DEV As String
    VL_RET_NC As String
    VL_OUT_DED_NC As String
    VL_CONT_NC_REC As String
    VL_TOT_CONT_CUM_PER As String
    VL_RET_CUM As String
    VL_OUT_DED_CUM As String
    VL_CONT_CUM_REC As String
    VL_TOT_CONT_REC As String
    
End Type

Public Type CamposRegM610
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_CONT As String
    VL_REC_BRT As String
    VL_BC_CONT As String
    VL_AJUS_ACRES_BC_COFINS As String
    VL_AJUS_REDUC_BC_COFINS As String
    VL_BC_CONT_AJUS As String
    ALIQ_COFINS As String
    QUANT_BC_COFINS As String
    ALIQ_COFINS_QUANT As String
    VL_CONT_APUR As String
    VL_AJUS_ACRES As String
    VL_AJUS_REDUC As String
    VL_CONT_DIFER As String
    VL_CONT_DIFER_ANT As String
    VL_CONT_PER As String
    
End Type

Public Type CamposRegM990
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    QTD_LIN_M As String
    
End Type

'###############################################
' REGISTRO DO BLOCO '1' DO SPED FISCAL
'###############################################

Public Type CamposReg1001
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_MOV As String

End Type

Public Type CamposReg1010
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_EXP As String
    IND_CCRF As String
    IND_COMB As String
    IND_USINA As String
    IND_VA As String
    IND_EE As String
    IND_CART As String
    IND_FORM As String
    IND_AER As String
    IND_GIAF1 As String
    IND_GIAF3 As String
    IND_GIAF4 As String
    IND_REST_RESSARC_COMPL_ICMS As String

End Type

Public Type CamposReg1100
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    IND_DOC As String
    NRO_DE As String
    DT_DE As String
    NAT_EXP As String
    NRO_RE As String
    DT_RE As String
    CHC_EMB As String
    DT_CHC As String
    DT_AVB As String
    TP_CHC As String
    PAIS As String
    
End Type

Public Type CamposReg1105
    
    CHV_PAI As String
    CHV_REG As String
    REG As String
    COD_MOD As String
    SERIE As String
    NUM_DOC As String
    CHV_NFE As String
    DT_DOC As String
    COD_ITEM As String

End Type

Public Type CamposReg1300
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_ITEM As String
    DT_FECH As String
    ESTQ_ABERT As String
    VOL_ENTR As String
    VOL_DISP As String
    VOL_SAIDAS As String
    ESTQ_ESCR As String
    VAL_AJ_PERDA As String
    VAL_AJ_GANHO As String
    FECH_FISICO As String
    
End Type

Public Type CamposReg1310
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    NUM_TANQUE As String
    ESTQ_ABERT As String
    VOL_ENTR As String
    VOL_DISP As String
    VOL_SAIDAS As String
    ESTQ_ESCR As String
    VAL_AJ_PERDA As String
    VAL_AJ_GANHO As String
    FECH_FISICO As String
    
End Type

Public Type CamposReg1320
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    NUM_BICO As String
    NR_INTERV As String
    MOT_INTERV As String
    NOM_INTERV As String
    CNPJ_INTERV As String
    CPF_INTERV As String
    VAL_FECHA As String
    VAL_ABERT As String
    VOL_AFERI As String
    VOL_VENDAS As String
        
End Type

Public Type CamposReg1350
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    SERIE As String
    FABRICANTE As String
    Modelo As String
    TIPO_MEDICAO As String
        
End Type

Public Type CamposReg1360
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    NUM_LACRE As String
    DT_APLICACAO As String
        
End Type

Public Type CamposReg1370
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    NUM_BICO As String
    COD_ITEM As String
    NUM_TANQUE As String
                
End Type

Public Type CamposReg1390
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_PROD As String
                
End Type

Public Type CamposReg1400
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_ITEM_IPM As String
    MUN As String
    Valor As String
    
End Type

Public Type CamposReg1391
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    DT_REGISTRO As String
    QTD_MOID As String
    ESTQ_INI As String
    QTD_PRODUZ As String
    ENT_ANID_HID As String
    OUTR_ENTR As String
    PERDA As String
    CONS As String
    SAI_ANI_HID As String
    SAÍDAS As String
    ESTQ_FIN As String
    ESTQ_INI_MEL As String
    PROD_DIA_MEL As String
    UTIL_MEL As String
    PROD_ALC_MEL As String
    OBS As String
    COD_ITEM As String
    TP_RESIDUO As String
    QTD_RESIDUO As String
                
End Type

Public Type CamposReg1601
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    COD_PART_IP As String
    COD_PART_IT As String
    TOT_VS As String
    TOT_ISS As String
    TOT_OUTROS As String
                
End Type

Public Type CamposReg1990
    
    REG As String
    ARQUIVO As String
    CHV_REG As String
    CHV_PAI As String
    CHV_PAI_FISCAL As String
    CHV_PAI_CONTRIBUICOES As String
    QTD_LIN_1 As String
    
End Type

Public Sub ResetarCampos0150()

Dim CamposVazios As CamposReg0150
    
    LSet Campos0150 = CamposVazios
    
End Sub

Public Sub ResetarCampos0190()

Dim CamposVazios As CamposReg0190
    
    LSet Campos0190 = CamposVazios
    
End Sub

Public Sub ResetarCampos0200()

Dim CamposVazios As CamposReg0200
    
    LSet Campos0200 = CamposVazios
    
End Sub

Public Sub ResetarCamposC100()

Dim CamposVazios As CamposRegC100
    
    LSet CamposC100 = CamposVazios
    
End Sub

Public Sub ResetarCamposC170()

Dim CamposVazios As CamposRegC170
    
    LSet CamposC170 = CamposVazios
    
End Sub
