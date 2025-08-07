Attribute VB_Name = "clsDominioSistemas"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function Gerar0000(ByVal Registro As String) As String

Dim Campos As Variant
Dim CNPJ As String
    
        CNPJ = fnSPED.ExtrairCampo(Registro, 7)
        
        Campos = Array("0000", CNPJ)
        
    Gerar0000 = fnSPED.GerarRegistro(Campos)

End Function

Public Function Gerar0100(ByVal Registro As String) As String

Dim Campos As Variant
            
    
    Campos = Array("0100")
    Gerar0100 = fnSPED.GerarRegistro(Campos)

End Function

Public Function Gerar1000(ByVal Registro As String) As String

Dim Campos As Variant
Dim Especie As Integer, codExDIEF&, cAcumulador&, CFOP&, Segmento&, nDoc&, nDocFin&, CFOPEx&, cTRansCredito&, _
cObs&, cAntTrib&, cRessarcimento&, MunOrigem&, cModelo&, codSit&, subSerie&, Operacao&, CFOPDoc&, tpCTe&, modImportacao&, _
cInfComp&, clConsumo&, tpLigacao&, grTensao&, tpAssinante&, KWHConsumo&, tpDocImportacao&, NatFretePISCOFINS&, CSTPISCOFINS&, _
BaseCredPISCOFINS&, CFPS&, NatPISCOFINS&, LancSCP&, tpServ&, MunDest&, tpServMaoObra&, indServMaoObra&, tpTitulo&

Dim InscFornec As String, SERIE$, OBS$, tpFrete$, tpEmit$, cRecISSRet$, cRecIRRFRet$, fgCRF$, fgIRRF$, InsEstFornec$, _
InscMunFornec$, codOperacao$, nParecer$, nDI$, bFiscal$, chNFe$, cRecFETHAB$, respFETHAB$, CTeRef$, infCompl$, _
nDraw$, chNFSe$, nProcesso$, OrigProcesso$, CSTIPI$, nDocArrecad$, Identificacao$

Dim dtEnt As String, dtEmi As String, dtVistoNF As String, Competencia As String, dtParecer As String, dtEscrituracao As String

Dim vContabil As Double, vExDIEF#, vFrete#, vSeg#, vDesp#, vPIS#, vCOFINS#, vDARE#, pDARE#, vBCST#, EntIsentas#, _
OutEntIsentas#, vFreteBase#, vProd#, vDedReceita#, ConsumoGasEnergia#, vCobradoTerceiros#, vServPISCOFINS#, _
bcPISCOFINS#, pPIS#, pCOFINS#, vPedagio#, vIPI#, vICMSST#
    
    Especie = 0
    InscFornec = ""
    codExDIEF = 0
    cAcumulador = 0
    CFOP = 0
    Segmento = 0
    nDoc = fnSPED.ExtrairCampo(Registro, 8)
    SERIE = ""
    nDocFin = fnSPED.ExtrairCampo(Registro, 8)
    dtEnt = fnSPED.ExtrairCampo(Registro, 11)
    If dtEnt <> "" Then dtEnt = CDate(Util.FormatarData(dtEnt))
    dtEmi = fnSPED.ExtrairCampo(Registro, 10)
    If dtEmi <> "" Then dtEmi = CDate(Util.FormatarData(dtEmi))
    vContabil = Util.FormatarValores(fnSPED.ExtrairCampo(Registro, 12))
    vExDIEF = 0
    OBS = ""
    If dtEmi <> "" Then tpFrete = IdentificarTipoFrete(VBA.Left(fnSPED.ExtrairCampo(Registro, 17), 1), dtEmi)
    tpEmit = IdentificarTipoEmitente(VBA.Left(fnSPED.ExtrairCampo(Registro, 3), 1))
    CFOPEx = 0
    cTRansCredito = 0
    cRecISSRet = ""
    cRecIRRFRet = ""
    cObs = 0
    dtVistoNF = 0
    fgCRF = ""
    fgIRRF = ""
    vFrete = 0
    vSeg = 0
    vDesp = 0
    vPIS = 0
    cAntTrib = 0
    vCOFINS = 0
    vDARE = 0
    pDARE = 0
    vBCST = 0
    EntIsentas = 0
    OutEntIsentas = 0
    vFreteBase = 0
    cRessarcimento = 0
    vProd = 0
    MunOrigem = 0
    cModelo = IdentificarTipoModelo(VBA.Left(fnSPED.ExtrairCampo(Registro, 6), 2))
    codSit = 0
    subSerie = 0
    InsEstFornec = ""
    InscMunFornec = ""
    codOperacao = ""
    vDedReceita = 0
    Competencia = 0
    Operacao = 0
    nParecer = ""
    dtParecer = 0
    nDI = ""
    bFiscal = ""
    chNFe = fnSPED.ExtrairCampo(Registro, 9)
    cRecFETHAB = ""
    respFETHAB = ""
    CFOPDoc = 0
    tpCTe = 0
    CTeRef = ""
    modImportacao = 0
    cInfComp = 0
    clConsumo = 0
    tpLigacao = 0
    grTensao = 0
    tpAssinante = 0
    KWHConsumo = 0
    ConsumoGasEnergia = 0
    vCobradoTerceiros = 0
    tpDocImportacao = 0
    nDraw = ""
    NatFretePISCOFINS = 0
    CSTPISCOFINS = 0
    BaseCredPISCOFINS = 0
    vServPISCOFINS = 0
    bcPISCOFINS = 0
    pPIS = 0
    pCOFINS = 0
    chNFSe = ""
    nProcesso = ""
    OrigProcesso = ""
    dtEscrituracao = 0
    CFPS = 0
    NatPISCOFINS = 0
    CSTIPI = ""
    LancSCP = 0
    tpServ = 0
    MunDest = 0
    vPedagio = 0
    vIPI = 0
    vICMSST = 0
    tpServMaoObra = 0
    indServMaoObra = 0
    nDocArrecad = ""
    tpTitulo = 0
    Identificacao = ""
    
    Campos = Array("1000", Especie, InscFornec, codExDIEF, cAcumulador, CFOP, Segmento, nDoc, SERIE, _
                   nDocFin, dtEnt, dtEmi, vContabil, vExDIEF, OBS, tpFrete, tpEmit, CFOPEx, cTRansCredito, _
                   cRecISSRet, cRecIRRFRet, cObs, dtVistoNF, fgCRF, fgIRRF, vFrete, vSeg, vDesp, vPIS, _
                   cAntTrib, vCOFINS, vDARE, pDARE, vBCST, EntIsentas, OutEntIsentas, vFreteBase, _
                   cRessarcimento, vProd, MunOrigem, cModelo, codSit, subSerie, InsEstFornec, _
                   InscMunFornec, codOperacao, vDedReceita, Competencia, Operacao, nParecer, dtParecer, _
                   nDI, bFiscal, chNFe, cRecFETHAB, respFETHAB, CFOPDoc, tpCTe, CTeRef, modImportacao, _
                   cInfComp, clConsumo, tpLigacao, grTensao, tpAssinante, KWHConsumo, ConsumoGasEnergia, _
                   vCobradoTerceiros, tpDocImportacao, nDraw, NatFretePISCOFINS, CSTPISCOFINS, BaseCredPISCOFINS, _
                   vServPISCOFINS, bcPISCOFINS, pPIS, pCOFINS, chNFe, nProcesso, OrigProcesso, dtEscrituracao, _
                   CFPS, NatPISCOFINS, CSTIPI, LancSCP, tpServ, MunDest, vPedagio, vIPI, _
                   vICMSST, tpServMaoObra, indServMaoObra, nDocArrecad, tpTitulo, Identificacao)
                   
    Gerar1000 = fnSPED.GerarRegistro(Campos)
    
End Function

Public Function Gerar2000(ByVal Registro As String) As String

Dim Campos As Variant
    
    Campos = Array("2000")
    Gerar2000 = fnSPED.GerarRegistro(Campos)

End Function

Public Function IdentificarTipoFrete(ByVal tpFrete As String, ByVal Data As Date)

    If Data >= CDate("2018-01-01") Then
    
        Select Case tpFrete
            
            Case "0"
                IdentificarTipoFrete = "C"
                
            Case "1"
                IdentificarTipoFrete = "F"
                
            Case "2"
                IdentificarTipoFrete = "T"
                
            Case "3"
                IdentificarTipoFrete = "R"
                
            Case "4"
                IdentificarTipoFrete = "D"
                
            Case "9"
                IdentificarTipoFrete = "S"
                
        End Select
    
    ElseIf Data >= CDate("2012-01-01") Then
        
        Select Case tpFrete
            
            Case "0"
                IdentificarTipoFrete = "R"
                
            Case "1"
                IdentificarTipoFrete = "D"
                
            Case "2"
                IdentificarTipoFrete = "T"
                
            Case "9"
                IdentificarTipoFrete = "S"
                
        End Select
    
    Else
    
        Select Case tpFrete
            
            Case "0"
                IdentificarTipoFrete = "T"
                
            Case "1"
                IdentificarTipoFrete = "R"
                
            Case "2"
                IdentificarTipoFrete = "D"
                
            Case "9"
                IdentificarTipoFrete = "S"
                
        End Select
        
    End If

End Function

Public Function IdentificarTipoEmitente(ByVal tpEmit As String)

    Select Case tpEmit
        
        Case "0"
            IdentificarTipoEmitente = "P"
            
        Case "1"
            IdentificarTipoEmitente = "T"
            
    End Select

End Function

Public Function IdentificarTipoModelo(ByVal tpModel As String)
    
    Select Case tpModel
        
        Case "00"
            IdentificarTipoModelo = "0"
            
        Case "01"
            IdentificarTipoModelo = "1"
            
        Case "02", "03"
            IdentificarTipoModelo = "2"
            
        Case "04"
            IdentificarTipoModelo = "7"
            
        Case "05"
            IdentificarTipoModelo = "8"
            
        Case "06"
            IdentificarTipoModelo = "6"
            
        Case "07"
            IdentificarTipoModelo = "10"
            
        Case "08"
            IdentificarTipoModelo = "9"
            
    End Select

End Function
