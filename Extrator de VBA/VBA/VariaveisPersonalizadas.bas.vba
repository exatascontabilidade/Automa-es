Attribute VB_Name = "VariaveisPersonalizadas"
Option Explicit

Public Type CamposrelICMS
    
    REG As String
    CHV_REG As String
    COD_MOD As String
    CFOP As String
    CST_ICMS As String
    ALIQ_ICMS As String
    VL_OPR As String
    VL_BC_ICMS As String
    VL_ICMS As String
    ALIQ_FCP As String
    VL_FCP As String
    VL_BC_ICMS_ST As String
    VL_ICMS_ST As String
    VL_RED_BC As String
    VL_IPI As String
    VL_ISENTAS As String
    VL_OUTRAS As String
    OBSERVACOES As String
    UF As String
    
End Type

Public Type CamposLivroIPI
    
    CHV_REG As String
    CFOP As String
    CST_IPI As String
    ALIQ_IPI As String
    VL_OPERACAO As String
    VL_BC_IPI As String
    VL_IPI As String
    OBSERVACOES As String

End Type

Public Type CamposLivroPISCOFINS
        
    CHV_REG As String
    CFOP As String
    CST_PIS As String
    CST_COFINS As String
    ALIQ_PIS As String
    ALIQ_COFINS As String
    VL_CONTABIL As String
    VL_BC_PIS As String
    VL_BC_COFINS As String
    VL_PIS As String
    VL_COFINS As String
    OBSERVACOES As String
    
End Type

Public Type CamposrelInteligenteDivergencias
    
    CHV_REG As String
    DOC_CONTRIB As String
    COD_PART As String
    DOC_PART As String
    Modelo As String
    Operacao As String
    TP_EMISSAO As String
    Situacao As String
    SERIE As String
    NUM_DOC As String
    DT_DOC As String
    CHV_NFE As String
    DT_EMI As String
    TP_PAGAMENTO As String
    TP_FRETE As String
    VL_DOC As String
    VL_PROD As String
    VL_FRETE As String
    VL_SEG As String
    VL_OUTRO As String
    VL_DESC As String
    VL_ABATIMENTO As String
    VL_BC_ICMS As String
    VL_ICMS As String
    VL_BC_ICMS_ST As String
    VL_ICMS_ST As String
    VL_IPI As String
    VL_PIS As String
    VL_COFINS As String
    STATUS_ANALISE As String
    OBSERVACOES As String
    
End Type

Public Type CamposRessarcimento

    CHV_REG As String
    COD_MOD As String
    NUM_DOC As String
    SERIE As String
    DT_DOC As String
    CNPJ_EMIT As String
    QTD As String
    UNID As String
    VL_UNIT_PROD As String
    VL_UNIT_BC_ST As String
    CHV_NFE As String
    N_ITEM As String
    COD_ITEM As String
    DESCRICAO_ITEM As String
    VL_UNIT_BC_ICMS As String
    ALIQ_ICMS As String
    VL_UNIT_ICMS As String
    ALIQ_ST As String
    VL_UNIT_RESSARCIMENTO As String
    OBSERVACOES As String
    CFOP As String
    
End Type

Public Type DadosNotasFiscais
    
    cBarra As String
    CEST As String
    CFOP As String
    Chave As String
    chNFe As String
    chCTe As String
    chDoce As String
    CPF As String
    CNPJ As String
    CNPJEmit As String
    CNPJDest As String
    CNPJPart As String
    CNPJFornec As String
    CNPJPrest As String
    CNPJRem As String
    CNPJExp As String
    CNPJRec As String
    CNPJTomador As String
    cPart As String
    vSubstituto As String
    vSubstituido As String
    cMun As String
    Filho As String
    Transmissao As String
    Registro As String
    RazaoFornec As String
    RazaoPrest As String
    InscEstadual As String
    cProd As String
    cAjuste As String
    cResponsavel As String
    indEmit As String
    ARQUIVO As String
    DataSPED As Date
    Movimento As String
    tpGIA As String
    tpAjuste As String
    tpEvento As String
    vISS As String
    vContabil As String
    vContabilContrib As String
    vContabilNaoContrib As String
    cTributaria As String
    cANP As String
    CNAE As String
    DESCRICAO As String
    Referencia As String
    ItemServico As String
    vBCICMSDet As String
    vTotal As String
    vISSRetido As String
    vPISRetido As String
    vAjuste As String
    vCOFINSRetido As String
    vIRPJRetido As String
    vCSLLRetido As String
    vINSSRetido As String
    vLiquido As String
    vICMSDeson As String
    CSTICMS As String
    CodSituacao As String
    CodFornec As String
    redBCICMS As String
    vBCICMS As String
    vBCICMSContrib As String
    vBCICMSNaoContrib As String
    vST As String
    vICMSDet As String
    vOutroDet As String
    pISS As String
    itemST As String
    vBCISS As String
    AnexoSN As String
    vServico As String
    vDeducoes As String
    exTIPI As String
    Competencia As String
    dhEmi As String
    dtEmi As String
    dtEnt As Variant
    dtLancamento As String
    tpItem As String
    FatConv As String
    indFCP As String
    idDest As String
    item As String
    ItemPai As String
    ItemR1000 As String
    ItemR1100 As String
    ItemR1300 As String
    ItemR1500 As String
    DivergNF As String
    Hash As String
    pMVA As String
    Modelo As String
    SERIE As String
    NCM As String
    NITEM As String
    nNF As String
    OBSERVACOES As String
    pCargaTrib As String
    pFCP As String
    pICMS As String
    pICMSST As String
    pRedBC As String
    pRedBCST As String
    qCom As String
    qInv As String
    vConfICMS As String
    vFECOEPRessarcir As String
    vSTRecRessarcir As String
    vResultRecRessarcir As String
    bcICMS As String
    QTD As String
    vIsentas As String
    vRedBCICMS As String
    vOutras As String
    vOutrasDet As String
    RazaoEmit As String
    RazaoDest As String
    RazaoPart As String
    tpNF As String
    tpOperacao As String
    vOperacao As String
    vTotProd As String
    vUnit As String
    vIPI As String
    pIPI As String
    vICMSEfetivo As String
    vTotICMSEfetivo As String
    vMinUnit As String
    tpEmissao As String
    uCom As String
    uInv As String
    UFDest As String
    UFEmit As String
    UF As String
    vICMS As String
    vICMSPetroleo As String
    vICMSOutro As String
    vICMSST As String
    vICMSSTDet As String
    vICMSTotal As String
    vFCP As String
    vFCPST As String
    vNFSPED As String
    vNF As String
    vBCST As String
    vProd As String
    vFrete As String
    vSeg As String
    vDesc As String
    vOutro As String
    vMedBCST As String
    vComplementoICMS As String
    vUnCom As String
    xProd As String
    Status As String
    StatusSPED As String
    vTotBCST As String
    vTotICMS As String
    vUnitMedICMS As String
    vMedICMS As String
    qTotEntradas As Double
    qTotSaidas As Double
    Tomador As String
    
End Type

Public Type DadosConhecimentos

    nCTe As String
    dhEmi As String
    CNPJEmit As String
    RazaoEmit As String
    vCTe As String
    chCTe As String
    UF As String
    UFOrig As String
    Stituacao As String
    tpOperacao As String
    dtLancamento As String
    DivCTe As String
    OBSERVACOES As String
    
End Type

Public Type RegistrosEFD
    
    'Bloco 0
    dic0000 As New Dictionary
    dic0001 As New Dictionary
    dic0002 As New Dictionary
    dic0005 As New Dictionary
    dic0015 As New Dictionary
    dic0100 As New Dictionary
    dic0150 As New Dictionary
    dic0175 As New Dictionary
    dic0190 As New Dictionary
    dic0200 As New Dictionary
    dic0205 As New Dictionary
    dic0206 As New Dictionary
    dic0210 As New Dictionary
    dic0220 As New Dictionary
    dic0221 As New Dictionary
    dic0300 As New Dictionary
    dic0305 As New Dictionary
    dic0400 As New Dictionary
    dic0450 As New Dictionary
    dic0460 As New Dictionary
    dic0500 As New Dictionary
    dic0600 As New Dictionary
    dic0990 As New Dictionary
    
    
    'Bloco B
    dicB001 As New Dictionary
    dicB020 As New Dictionary
    dicB025 As New Dictionary
    dicB030 As New Dictionary
    dicB035 As New Dictionary
    dicB350 As New Dictionary
    dicB420 As New Dictionary
    dicB440 As New Dictionary
    dicB460 As New Dictionary
    dicB470 As New Dictionary
    dicB500 As New Dictionary
    dicB510 As New Dictionary
    dicB990 As New Dictionary
    
        
    'Bloco C
    dicC001 As New Dictionary
    dicC100 As New Dictionary
    dicC101 As New Dictionary
    dicC105 As New Dictionary
    dicC110 As New Dictionary
    dicC111 As New Dictionary
    dicC112 As New Dictionary
    dicC113 As New Dictionary
    dicC114 As New Dictionary
    dicC115 As New Dictionary
    dicC116 As New Dictionary
    dicC120 As New Dictionary
    dicC130 As New Dictionary
    dicC140 As New Dictionary
    dicC141 As New Dictionary
    dicC160 As New Dictionary
    dicC165 As New Dictionary
    dicC170 As New Dictionary
    dicC171 As New Dictionary
    dicC172 As New Dictionary
    dicC173 As New Dictionary
    dicC174 As New Dictionary
    dicC175 As New Dictionary
    dicC175Contrib As New Dictionary
    dicC176 As New Dictionary
    dicC177 As New Dictionary
    dicC178 As New Dictionary
    dicC179 As New Dictionary
    dicC180 As New Dictionary
    dicC181 As New Dictionary
    dicC185 As New Dictionary
    dicC186 As New Dictionary
    dicC190 As New Dictionary
    dicC191 As New Dictionary
    dicC195 As New Dictionary
    dicC197 As New Dictionary
    dicC300 As New Dictionary
    dicC310 As New Dictionary
    dicC320 As New Dictionary
    dicC321 As New Dictionary
    dicC330 As New Dictionary
    dicC350 As New Dictionary
    dicC370 As New Dictionary
    dicC380 As New Dictionary
    dicC390 As New Dictionary
    dicC400 As New Dictionary
    dicC405 As New Dictionary
    dicC410 As New Dictionary
    dicC420 As New Dictionary
    dicC425 As New Dictionary
    dicC430 As New Dictionary
    dicC460 As New Dictionary
    dicC465 As New Dictionary
    dicC470 As New Dictionary
    dicC480 As New Dictionary
    dicC490 As New Dictionary
    dicC495 As New Dictionary
    dicC500 As New Dictionary
    dicC510 As New Dictionary
    dicC590 As New Dictionary
    dicC591 As New Dictionary
    dicC595 As New Dictionary
    dicC597 As New Dictionary
    dicC600 As New Dictionary
    dicC601 As New Dictionary
    dicC610 As New Dictionary
    dicC690 As New Dictionary
    dicC700 As New Dictionary
    dicC790 As New Dictionary
    dicC791 As New Dictionary
    dicC800 As New Dictionary
    dicC810 As New Dictionary
    dicC815 As New Dictionary
    dicC850 As New Dictionary
    dicC855 As New Dictionary
    dicC857 As New Dictionary
    dicC860 As New Dictionary
    dicC870 As New Dictionary
    dicC880 As New Dictionary
    dicC890 As New Dictionary
    dicC895 As New Dictionary
    dicC897 As New Dictionary
    dicC990 As New Dictionary
    
    
    'Bloco D
    dicD001 As New Dictionary
    dicD100 As New Dictionary
    dicD101 As New Dictionary
    dicD110 As New Dictionary
    dicD120 As New Dictionary
    dicD130 As New Dictionary
    dicD140 As New Dictionary
    dicD150 As New Dictionary
    dicD160 As New Dictionary
    dicD161 As New Dictionary
    dicD162 As New Dictionary
    dicD170 As New Dictionary
    dicD180 As New Dictionary
    dicD190 As New Dictionary
    dicD195 As New Dictionary
    dicD197 As New Dictionary
    dicD300 As New Dictionary
    dicD301 As New Dictionary
    dicD310 As New Dictionary
    dicD350 As New Dictionary
    dicD355 As New Dictionary
    dicD360 As New Dictionary
    dicD365 As New Dictionary
    dicD370 As New Dictionary
    dicD390 As New Dictionary
    dicD400 As New Dictionary
    dicD410 As New Dictionary
    dicD411 As New Dictionary
    dicD420 As New Dictionary
    dicD500 As New Dictionary
    dicD510 As New Dictionary
    dicD530 As New Dictionary
    dicD590 As New Dictionary
    dicD600 As New Dictionary
    dicD610 As New Dictionary
    dicD690 As New Dictionary
    dicD695 As New Dictionary
    dicD696 As New Dictionary
    dicD697 As New Dictionary
    dicD700 As New Dictionary
    dicD730 As New Dictionary
    dicD731 As New Dictionary
    dicD735 As New Dictionary
    dicD737 As New Dictionary
    dicD750 As New Dictionary
    dicD760 As New Dictionary
    dicD761 As New Dictionary
    dicD990 As New Dictionary
    
    
    'Bloco E
    dicE001 As New Dictionary
    dicE100 As New Dictionary
    dicE110 As New Dictionary
    dicE111 As New Dictionary
    dicE112 As New Dictionary
    dicE113 As New Dictionary
    dicE115 As New Dictionary
    dicE116 As New Dictionary
    dicE200 As New Dictionary
    dicE210 As New Dictionary
    dicE220 As New Dictionary
    dicE230 As New Dictionary
    dicE240 As New Dictionary
    dicE250 As New Dictionary
    dicE300 As New Dictionary
    dicE310 As New Dictionary
    dicE311 As New Dictionary
    dicE312 As New Dictionary
    dicE313 As New Dictionary
    dicE316 As New Dictionary
    dicE500 As New Dictionary
    dicE510 As New Dictionary
    dicE520 As New Dictionary
    dicE530 As New Dictionary
    dicE531 As New Dictionary
    dicE990 As New Dictionary
    
    
    'Bloco G
    dicG001 As New Dictionary
    dicG110 As New Dictionary
    dicG125 As New Dictionary
    dicG126 As New Dictionary
    dicG130 As New Dictionary
    dicG140 As New Dictionary
    dicG990 As New Dictionary
    
    
    'Bloco H
    dicH001 As New Dictionary
    dicH005 As New Dictionary
    dicH010 As New Dictionary
    dicH020 As New Dictionary
    dicH030 As New Dictionary
    dicH990 As New Dictionary
    
    
    'Bloco K
    dicK001 As New Dictionary
    dicK010 As New Dictionary
    dicK100 As New Dictionary
    dicK200 As New Dictionary
    dicK210 As New Dictionary
    dicK215 As New Dictionary
    dicK220 As New Dictionary
    dicK230 As New Dictionary
    dicK235 As New Dictionary
    dicK250 As New Dictionary
    dicK255 As New Dictionary
    dicK260 As New Dictionary
    dicK265 As New Dictionary
    dicK270 As New Dictionary
    dicK275 As New Dictionary
    dicK280 As New Dictionary
    dicK290 As New Dictionary
    dicK291 As New Dictionary
    dicK292 As New Dictionary
    dicK300 As New Dictionary
    dicK301 As New Dictionary
    dicK302 As New Dictionary
    dicK990 As New Dictionary
    
    
    'Bloco M
    dicM001 As New Dictionary
    dicM100 As New Dictionary
    dicM200 As New Dictionary
    dicM210 As New Dictionary
    dicM500 As New Dictionary
    dicM600 As New Dictionary
    dicM610 As New Dictionary
    dicM990 As New Dictionary
    
    
    'Bloco 1
    dic1001 As New Dictionary
    dic1010 As New Dictionary
    dic1100 As New Dictionary
    dic1105 As New Dictionary
    dic1110 As New Dictionary
    dic1200 As New Dictionary
    dic1210 As New Dictionary
    dic1250 As New Dictionary
    dic1255 As New Dictionary
    dic1300 As New Dictionary
    dic1310 As New Dictionary
    dic1320 As New Dictionary
    dic1350 As New Dictionary
    dic1360 As New Dictionary
    dic1370 As New Dictionary
    dic1390 As New Dictionary
    dic1391 As New Dictionary
    dic1400 As New Dictionary
    dic1500 As New Dictionary
    dic1510 As New Dictionary
    dic1600 As New Dictionary
    dic1601 As New Dictionary
    dic1700 As New Dictionary
    dic1710 As New Dictionary
    dic1800 As New Dictionary
    dic1900 As New Dictionary
    dic1910 As New Dictionary
    dic1920 As New Dictionary
    dic1921 As New Dictionary
    dic1922 As New Dictionary
    dic1923 As New Dictionary
    dic1925 As New Dictionary
    dic1926 As New Dictionary
    dic1960 As New Dictionary
    dic1970 As New Dictionary
    dic1975 As New Dictionary
    dic1980 As New Dictionary
    dic1990 As New Dictionary
    
    
    'Bloco 9
    dic9000 As New Dictionary
    dic9900 As New Dictionary
    dic9990 As New Dictionary
    dic9999 As New Dictionary
    
End Type

