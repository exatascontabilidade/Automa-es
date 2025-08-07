Attribute VB_Name = "clsAssistenteApuracao"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Campo As Variant
Private COD_MOD As String
Public UFContrib As String
Public GerenciadorSPED As New clsRegistrosSPED
Public IPI As New clsAssistenteApuracaoIPI
Public ICMS As New clsAssistenteApuracaoICMS
Public PISCOFINS As New clsAssistenteApuracaoPISCOFINS

'Títulos do Relatório
Public dicTitulos As Dictionary
Public dicTitulosLivro As Dictionary
Public dicTitulosApuracao As Dictionary

Public Sub ExtrairDados0150(ByVal CHV_PAI As String, ByVal COD_PART As String, Optional Contrib As Boolean, Optional IPI As Boolean)

Dim Chave As String, UF_PART$, NOME$, TIPO_PART$
    
    With dtoRegSPED
        
        If .r0150 Is Nothing Then Call CarregarDadosRegistro0150(Contrib)
        If Not Util.ValidarDicionario(.r0150) Then Exit Sub
                
        Chave = Util.UnirCampos(CHV_PAI, COD_PART)
        If .r0150.Exists(Chave) Then
            
            UF_PART = ExtrairUF_PART(Chave)
            NOME = .r0150(Chave)(dtoTitSPED.t0150("NOME"))
            TIPO_PART = ExtrairTIPO_PART(Chave)
            
            AtribuirValor "TIPO_PART", TIPO_PART
            AtribuirValor "NOME_RAZAO", NOME
            AtribuirValor "UF_PART", UF_PART
            If Not Contrib And Not IPI Then AtribuirValor "CONTRIBUINTE", ExtrairCONTRIBUINTE(Chave)
            
        Else
            
            AtribuirValor "TIPO_PART", "PF"
            AtribuirValor "UF_PART", UFContrib
            If Not Contrib And Not IPI Then AtribuirValor "CONTRIBUINTE", "NÃO"
            
        End If
        
    End With
    
End Sub

Public Sub ExtrairDados0200(ByVal CHV_PAI As String, ByVal COD_ITEM As String, Optional Contrib As Boolean, Optional IPI As Boolean)

Dim Chave As String, UF_PART$, NOME$, TIPO_PART$
    
    With dtoRegSPED
        
        If .r0200 Is Nothing Then Call CarregarDadosRegistro0200(Contrib)
        If Not Util.ValidarDicionario(.r0200) Then Exit Sub
                
        Chave = Util.UnirCampos(CHV_PAI, COD_ITEM)
        If .r0200.Exists(Chave) Then
            
            AtribuirValor "DESCR_ITEM", .r0200(Chave)(dtoTitSPED.t0200("DESCR_ITEM"))
            AtribuirValor "COD_BARRA", fnExcel.FormatarTexto(.r0200(Chave)(dtoTitSPED.t0200("COD_BARRA")))
            AtribuirValor "COD_NCM", fnExcel.FormatarTexto(.r0200(Chave)(dtoTitSPED.t0200("COD_NCM")))
            AtribuirValor "EX_IPI", fnExcel.FormatarTexto(.r0200(Chave)(dtoTitSPED.t0200("EX_IPI")))
            AtribuirValor "TIPO_ITEM", .r0200(Chave)(dtoTitSPED.t0200("TIPO_ITEM"))
            
            If dicTitulos.Exists("CEST") Then AtribuirValor "CEST", fnExcel.FormatarTexto(.r0200(Chave)(dtoTitSPED.t0200("CEST")))
            
        Else
            
            AtribuirValor "DESCR_ITEM", "ITEM NÃO IDENTIFICADO"
            
        End If
        
    End With
    
End Sub

Public Sub ExtrairDados0200_oLD(ByVal ARQUIVO As String, ByVal COD_ITEM As String, _
    Optional ByVal CNPJEstabelecimento As String, Optional ByVal SPEDContrib As Boolean = False)

Dim i As Byte
Dim CHV_REG As String, CHV_PAI$, Chave$
    
    With dtoRegSPED
        
        If .r0200 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0200
        If Not Util.ValidarDicionario(.r0200) Then Exit Sub
        
        Chave = Util.UnirCampos(ARQUIVO, CNPJEstabelecimento)
        Select Case True
            
            Case CNPJEstabelecimento <> ""
                If .r0140 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0140("ARQUIVO", "CNPJ")
                If .r0140.Exists(Chave) Then CHV_PAI = .r0140(Chave)(dtoTitSPED.t0140("CHV_REG"))
                
            Case Else
                If .r0001 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0001("ARQUIVO")
                If .r0001.Exists(ARQUIVO) Then CHV_PAI = .r0001(ARQUIVO)(dtoTitSPED.t0001("CHV_REG"))
                
        End Select
        
        CHV_REG = Util.UnirCampos(CHV_PAI, COD_ITEM)
        If .r0200.Exists(CHV_REG) Then
            
            AtribuirValor "DESCR_ITEM", .r0200(CHV_REG)(dtoTitSPED.t0200("DESCR_ITEM"))
            AtribuirValor "COD_BARRA", fnExcel.FormatarTexto(.r0200(CHV_REG)(dtoTitSPED.t0200("COD_BARRA")))
            AtribuirValor "COD_NCM", fnExcel.FormatarTexto(.r0200(CHV_REG)(dtoTitSPED.t0200("COD_NCM")))
            AtribuirValor "EX_IPI", fnExcel.FormatarTexto(.r0200(CHV_REG)(dtoTitSPED.t0200("EX_IPI")))
            AtribuirValor "TIPO_ITEM", .r0200(CHV_REG)(dtoTitSPED.t0200("TIPO_ITEM"))
            
            If dicTitulos.Exists("CEST") Then AtribuirValor "CEST", fnExcel.FormatarTexto(.r0200(CHV_REG)(dtoTitSPED.t0200("CEST")))
            
        Else
            
            AtribuirValor "DESCR_ITEM", "ITEM NÃO IDENTIFICADO"
            
        End If
    
    End With
    
End Sub

Public Sub ExtrairDados0400(ByVal ARQUIVO As String, ByVal COD_NAT As String)
    
Dim CHV_REG As String
    
    With dtoRegSPED
        
        If .r0400 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0400("ARQUIVO", "COD_NAT")
        If Not Util.ValidarDicionario(.r0400) Then Exit Sub
        
        CHV_REG = Util.UnirCampos(ARQUIVO, COD_NAT)
        If .r0400.Exists(CHV_REG) Then
            
            AtribuirValor "COD_NAT", COD_NAT & " - " & .r0400(CHV_REG)(dtoTitSPED.t0400("DESCR_NAT"))
            
        Else
            
            If COD_NAT <> "" Then AtribuirValor "COD_NAT", COD_NAT
            
        End If
    
    End With
    
End Sub

Public Sub ExtrairDadosA100(ByVal CHV_REG As String)
    
Dim i As Byte
    
    With dtoRegSPED
        
        If .rA100 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroA100
        If Not Util.ValidarDicionario(.rA100) Then Exit Sub
        
        If .rA100.Exists(CHV_REG) Then
            
            AtribuirValor "CHV_NFE", fnExcel.FormatarTexto(.rA100(CHV_REG)(dtoTitSPED.tA100("CHV_NFSE")))
            AtribuirValor "NUM_DOC", fnExcel.FormatarTexto(.rA100(CHV_REG)(dtoTitSPED.tA100("NUM_DOC")))
            AtribuirValor "SER", fnExcel.FormatarTexto(.rA100(CHV_REG)(dtoTitSPED.tA100("SER")))
            AtribuirValor "IND_OPER", fnExcel.FormatarTexto(.rA100(CHV_REG)(dtoTitSPED.tA100("IND_OPER")))
            AtribuirValor "DT_DOC", fnExcel.FormatarData(.rA100(CHV_REG)(dtoTitSPED.tA100("DT_DOC")))
            AtribuirValor "DT_ENT_SAI", fnExcel.FormatarData(.rA100(CHV_REG)(dtoTitSPED.tA100("DT_EXE_SERV")))
            AtribuirValor "COD_PART", fnExcel.FormatarTexto(.rA100(CHV_REG)(dtoTitSPED.tA100("COD_PART")))
            
        End If
        
    End With
    
End Sub

Public Sub ExtrairDadosC100(ByVal CHV_REG As String)
    
Dim i As Byte
    
    With dtoRegSPED
        
        If .rC100 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC100
        If Not Util.ValidarDicionario(.rC100) Then Exit Sub
            
        If .rC100.Exists(CHV_REG) Then
            
            COD_MOD = .rC100(CHV_REG)(dtoTitSPED.tC100("COD_MOD"))
            
            AtribuirValor "COD_MOD", COD_MOD
            AtribuirValor "COD_SIT", fnExcel.FormatarTexto(.rC100(CHV_REG)(dtoTitSPED.tC100("COD_SIT")))
            AtribuirValor "CHV_NFE", fnExcel.FormatarTexto(.rC100(CHV_REG)(dtoTitSPED.tC100("CHV_NFE")))
            AtribuirValor "NUM_DOC", fnExcel.FormatarTexto(.rC100(CHV_REG)(dtoTitSPED.tC100("NUM_DOC")))
            AtribuirValor "SER", fnExcel.FormatarTexto(.rC100(CHV_REG)(dtoTitSPED.tC100("SER")))
            AtribuirValor "IND_OPER", fnExcel.FormatarTexto(.rC100(CHV_REG)(dtoTitSPED.tC100("IND_OPER")))
            AtribuirValor "DT_DOC", fnExcel.FormatarData(.rC100(CHV_REG)(dtoTitSPED.tC100("DT_DOC")))
            AtribuirValor "DT_ENT_SAI", fnExcel.FormatarData(.rC100(CHV_REG)(dtoTitSPED.tC100("DT_E_S")))
            AtribuirValor "COD_PART", fnExcel.FormatarTexto(.rC100(CHV_REG)(dtoTitSPED.tC100("COD_PART")))
            
        End If
    
    End With
    
End Sub

Public Sub ExtrairDadosC177(ByVal CHV_REG As String)
    
    With dtoRegSPED
        
        If .rC177 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC177("CHV_PAI_FISCAL")
        If Not Util.ValidarDicionario(.rC177) Then Exit Sub
        
        If .rC177.Exists(CHV_REG) Then AtribuirValor "COD_INF_ITEM", .rC177(CHV_REG)(dtoTitSPED.tC177("COD_INF_ITEM"))
        
    End With
    
End Sub

Public Sub ExtrairDadosC180(ByVal CHV_REG As String)

Dim i As Byte
    
    With dtoRegSPED
        
        If .rC180 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC180
        If Not Util.ValidarDicionario(.rC180) Then Exit Sub
        
        If .rC180.Exists(CHV_REG) Then
            
            AtribuirValor "COD_MOD", fnExcel.FormatarTexto(.rC180(CHV_REG)(dtoTitSPED.tC180("COD_MOD")))
            AtribuirValor "DT_DOC", fnExcel.FormatarData(.rC180(CHV_REG)(dtoTitSPED.tC180("DT_DOC_INI")))
            AtribuirValor "DT_ENT_SAI", fnExcel.FormatarData(.rC180(CHV_REG)(dtoTitSPED.tC180("DT_DOC_FIN")))
            AtribuirValor "COD_ITEM", fnExcel.FormatarTexto(.rC180(CHV_REG)(dtoTitSPED.tC180("COD_ITEM")))
            AtribuirValor "COD_NCM", fnExcel.FormatarTexto(.rC180(CHV_REG)(dtoTitSPED.tC180("COD_NCM")))
            AtribuirValor "EX_IPI", fnExcel.FormatarTexto(.rC180(CHV_REG)(dtoTitSPED.tC180("EX_IPI")))
            
        End If
        
    End With
    
End Sub

Public Function ExtrairCOD_PART_A100(ByVal CHV_REG As String) As String
    
    With dtoRegSPED
        
        If .rA100 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroA100
        ExtrairCOD_PART_A100 = .rA100(CHV_REG)(dtoTitSPED.tA100("COD_PART"))
        
    End With
    
End Function

Public Function ExtrairCOD_PART_C100(ByVal CHV_REG As String) As String
    
    With dtoRegSPED
        
        If .rC100 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC100
        
        If .rC100.Exists(CHV_REG) Then ExtrairCOD_PART_C100 = .rC100(CHV_REG)(dtoTitSPED.tC100("COD_PART"))
        
    End With
    
End Function

Public Function ExtrairCHV_0001(ByVal ARQUIVO As String) As String
    
    With dtoRegSPED
        
        If .r0001 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0001("ARQUIVO")
        If .r0001.Exists(ARQUIVO) Then ExtrairCHV_0001 = .r0001(ARQUIVO)(dtoTitSPED.t0001("CHV_REG"))
        
    End With
    
End Function

Public Function ExtrairCHV_0140(ByVal Chave As String) As String
    
    With dtoRegSPED
        
        If .r0140 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0140("ARQUIVO", "CNPJ")
        If .r0140.Exists(Chave) Then ExtrairCHV_0140 = .r0140(Chave)(dtoTitSPED.t0140("CHV_REG"))
        
    End With
    
End Function

Public Function ExtrairCHV_C001(ByVal ARQUIVO As String) As String
        
    With dtoRegSPED
        
        If .rC001 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC001("ARQUIVO")
        If .rC001.Exists(ARQUIVO) Then ExtrairCHV_C001 = .rC001(ARQUIVO)(dtoTitSPED.tC001("CHV_REG"))
        
    End With
    
End Function

Public Function ExtrairVL_DESP_C100(ByVal CHV_REG As String, ByVal VL_ITEM As Double) As Double

Dim VL_MERC As Double, VL_FRT#, VL_SEG#, VL_OUT#, VL_DESP#
    
    With dtoRegSPED
        
        If .rC100 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC100
        
        If .rC100.Exists(CHV_REG) Then
            
            VL_MERC = fnExcel.FormatarValores(.rC100(CHV_REG)(dtoTitSPED.tC100("VL_MERC")))
            VL_FRT = fnExcel.FormatarValores(.rC100(CHV_REG)(dtoTitSPED.tC100("VL_FRT")))
            VL_SEG = fnExcel.FormatarValores(.rC100(CHV_REG)(dtoTitSPED.tC100("VL_SEG")))
            VL_OUT = fnExcel.FormatarValores(.rC100(CHV_REG)(dtoTitSPED.tC100("VL_OUT_DA")))
            VL_DESP = fnExcel.FormatarValores(VL_FRT + VL_SEG + VL_OUT, True, 2)
            
            If VL_MERC > 0 Then ExtrairVL_DESP_C100 = fnExcel.FormatarValores((VL_ITEM / VL_MERC) * VL_DESP, True, 2)
            
        End If
        
    End With
    
End Function

'TODO: Criar rotina para extrair o campo TIPO_PART do registro C180, com base no modelo do documento fiscal
Public Function ExtrairTIPO_PART(ByVal CHV_REG As String) As String

Dim CNPJ As String, CPF$
    
    With dtoRegSPED
        
        If .r0150 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0150
        
        If .r0150.Exists(CHV_REG) Then
            
            CPF = Util.ApenasNumeros(.r0150(CHV_REG)(dtoTitSPED.t0150("CPF")))
            CNPJ = Util.ApenasNumeros(.r0150(CHV_REG)(dtoTitSPED.t0150("CNPJ")))
            
            If CPF <> "" Then ExtrairTIPO_PART = "PF"
            If CNPJ <> "" Then ExtrairTIPO_PART = "PJ"
            Exit Function
            
        End If
        
        ExtrairTIPO_PART = "PF"
    
    End With
    
End Function

Public Function ExtrairUF_PART(ByVal CHV_REG As String) As String

Dim UF As String
    
    With dtoRegSPED
        
        If .r0150 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0150
        If Not Util.ValidarDicionario(.r0150) Then ExtrairUF_PART = UFContrib: Exit Function
        
        If .r0150.Exists(CHV_REG) Then
            UF = .r0150(CHV_REG)(dtoTitSPED.t0150("COD_MUN"))
        End If
        
        If UF = "" Then ExtrairUF_PART = UFContrib Else ExtrairUF_PART = Util.ConverterIBGE_UF(UF)
        
    End With
    
End Function

Public Function ExtrairUF_Estabelecimento(ByVal ARQUIVO As String, ByVal CNPJ As String) As String

Dim UF As String, Chave$
    
    With dtoRegSPED
    
        If .r0140 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0140("ARQUIVO", "CNPJ")
        If Not Util.ValidarDicionario(.r0140) Then ExtrairUF_Estabelecimento = UFContrib
        
        Chave = Util.UnirCampos(ARQUIVO, CNPJ)
        If .r0140.Exists(Chave) Then
            UF = .r0140(Chave)(dtoTitSPED.t0140("UF"))
        End If
        
        If UF = "" Then ExtrairUF_Estabelecimento = UFContrib Else ExtrairUF_Estabelecimento = UF
        
    End With
    
End Function

Public Function ExtrairUFContribuinte(ByVal ARQUIVO As String, Optional ByVal Contrib As Boolean = False) As String

Dim UF As String
    
    With dtoRegSPED
        
        If .r0000 Is Nothing Then Call CarregarDadosRegistro0000(Contrib)
        If .r0000.Exists(ARQUIVO) Then UFContrib = .r0000(ARQUIVO)(dtoTitSPED.t0000("UF"))
        
        ExtrairUFContribuinte = UFContrib
        
    End With
    
End Function

Public Sub CarregarDadosRegistro0000(Optional Contrib As Boolean = False)
    
    Select Case Contrib
        
        Case True
            GerenciadorSPED.Contrib = Contrib
            Call GerenciadorSPED.CarregarDadosRegistro0000("ARQUIVO")
            
        Case Else
            GerenciadorSPED.Contrib = Contrib
            Call GerenciadorSPED.CarregarDadosRegistro0000("ARQUIVO")
            
    End Select
    
End Sub

Private Sub CarregarDadosRegistro0150(Optional Contrib As Boolean = False)
    
    Select Case Contrib
        
        Case True
            Call GerenciadorSPED.CarregarDadosRegistro0150("CHV_PAI_CONTRIBUICOES", "COD_PART")
            
        Case Else
            Call GerenciadorSPED.CarregarDadosRegistro0150("CHV_PAI_FISCAL", "COD_PART")
            
    End Select
        
End Sub

Private Sub CarregarDadosRegistro0200(Optional Contrib As Boolean = False)
    
    Select Case Contrib
        
        Case True
            Call GerenciadorSPED.CarregarDadosRegistro0200("CHV_PAI_CONTRIBUICOES", "COD_ITEM")
            
        Case Else
            Call GerenciadorSPED.CarregarDadosRegistro0200("CHV_PAI_FISCAL", "COD_ITEM")
            
    End Select
        
End Sub

'Private Sub CarregarDadosRegistro0500()
'
'    Set dicDados0500 = Util.CriarDicionarioRegistro(reg0500, "ARQUIVO", "COD_CTA")
'    Set dicTitulos0500 = Util.MapearTitulos(reg0500, 3)
'
'End Sub

Public Function AtribuirValor(ByVal Titulo As String, ByVal Valor As Variant)
    
    Campo(dicTitulos(Titulo)) = Valor
    
End Function

Public Function ExtrairValor(ByVal Titulo As String)
    
    ExtrairValor = Campo(dicTitulos(Titulo))
    
End Function

Public Function ExtrairREGIME_TRIBUTARIO(ByVal ARQUIVO As String, Optional AtribuirValorCampo As Boolean)
        
Dim COD_INC_TRIB As String
    
    With dtoRegSPED
        
        If .r0110 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0110("ARQUIVO")
        If Not Util.ValidarDicionario(.r0110) Then
            AtribuirValor "REGIME_TRIBUTARIO", "DEFINA UM REGIME PARA A OPERAÇÃO"
            Exit Function
        End If
        
        If .r0110.Exists(ARQUIVO) Then COD_INC_TRIB = .r0110(ARQUIVO)(dtoTitSPED.t0110("COD_INC_TRIB"))
        ExtrairREGIME_TRIBUTARIO = COD_INC_TRIB
        
        If AtribuirValorCampo Then
            
            Select Case VBA.Left(COD_INC_TRIB, 1)
                
                Case "1", "2"
                    AtribuirValor "REGIME_TRIBUTARIO", RegrasCadastrais.PIS_COFINS.ValidarEnumeracao_REGIME_TRIBUTARIO(COD_INC_TRIB)
                    
                Case Else
                    AtribuirValor "REGIME_TRIBUTARIO", "DEFINA UM REGIME PARA A OPERAÇÃO"
                    ExtrairREGIME_TRIBUTARIO = ""
                    
            End Select
            
        End If
        
    End With
    
End Function

Public Function ExtrairCNPJ_ESTABELECIMENTO(ByVal REG As String, ByVal CHV_REG As String, ByVal ARQUIVO As String) As String

Dim CNPJ As String
    
    With dtoRegSPED
        
        Select Case REG
            
            Case "A170"
                CNPJ = ExtrairCNPJ_ESTABELECIMENTO_A100(CHV_REG)
                
            Case "C170"
                CNPJ = ExtrairCNPJ_ESTABELECIMENTO_C100(CHV_REG, ARQUIVO)
                
            Case "C181", "C1815"
                CNPJ = ExtrairCNPJ_ESTABELECIMENTO_C180(CHV_REG)
                
            Case "F100"
                CNPJ = ExtrairCNPJ_ESTABELECIMENTO_F010(CHV_REG)
                
        End Select
                
        If CNPJ = "" And .r0000.Exists(ARQUIVO) Then CNPJ = .r0000(ARQUIVO)(dtoTitSPED.t0000("CNPJ"))
        ExtrairCNPJ_ESTABELECIMENTO = CNPJ
        
    End With
    
End Function

Public Function ExtrairCNPJ_ESTABELECIMENTO_A100(ByVal CHV_A100 As String) As String

Dim CHV_A010 As String
    
    With dtoRegSPED
        
        If .rA010 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroA010
        If .rA010 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroA100
        
        If .rA010.Exists(CHV_A100) Then
            
            CHV_A010 = .rA010(CHV_A100)(dtoTitSPED.tA100("CHV_PAI_FISCAL"))
            If .rA010.Exists(CHV_A010) Then ExtrairCNPJ_ESTABELECIMENTO_A100 = .rA010(CHV_A010)(dtoTitSPED.tA010("CNPJ"))
            
        End If
        
    End With
    
End Function

Public Function ExtrairCNPJ_ESTABELECIMENTO_C100(ByVal CHV_C100 As String, Optional ARQUIVO As String) As String

Dim CHV_C010 As String
    
    With dtoRegSPED
        
        If .rC010 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC010
        If Not Util.ValidarDicionario(.rC010) Then
            ExtrairCNPJ_ESTABELECIMENTO_C100 = ExtrairCNPJ_CONTRIBUINTE(ARQUIVO)
            Exit Function
        End If
        
        If .rC100 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC100
        If Not Util.ValidarDicionario(.rC100) Then Exit Function
        
        If .rC100.Exists(CHV_C100) Then
            
            CHV_C010 = .rC100(CHV_C100)(dtoTitSPED.tC100("CHV_PAI_CONTRIBUICOES"))
            If .rC010.Exists(CHV_C010) Then ExtrairCNPJ_ESTABELECIMENTO_C100 = .rC010(CHV_C010)(dtoTitSPED.tC010("CNPJ"))
            
        End If
        
    End With
    
End Function

Public Function ExtrairCNPJ_ESTABELECIMENTO_C180(ByVal CHV_C180 As String) As String

Dim CHV_C010 As String
    
    With dtoRegSPED
        
        If .rC010 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC010
        If .rC180 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroC180
        
        If .rC180.Exists(CHV_C180) Then
            
            CHV_C010 = .rC180(CHV_C180)(dtoTitSPED.tC180("CHV_PAI_CONTRIBUICOES"))
            If .rC010.Exists(CHV_C010) Then ExtrairCNPJ_ESTABELECIMENTO_C180 = .rC010(CHV_C010)(dtoTitSPED.tC010("CNPJ"))
            
        End If
    
    End With
    
End Function

Public Function ExtrairCNPJ_CONTRIBUINTE(ByVal ARQUIVO As String) As String
    
    ExtrairCNPJ_CONTRIBUINTE = VBA.Split(ARQUIVO, "-")(1)
    
End Function

Public Function ExtrairCNPJ_ESTABELECIMENTO_F010(ByVal CHV_F010 As String) As String
    
    With dtoRegSPED
    
        If .rF010 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroF010
        If Not Util.ValidarDicionario(.rF010) Then Exit Function
        
        If .rF010.Exists(CHV_F010) Then ExtrairCNPJ_ESTABELECIMENTO_F010 = .rF010(CHV_F010)(dtoTitSPED.tF010("CNPJ"))
    
    End With
    
End Function

Public Function ExtrairCONTRIBUINTE(ByVal CHV_REG As String) As String

Dim IE As String
    
    With dtoRegSPED
        
        If .r0150 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0150
        If Not Util.ValidarDicionario(.r0150) Then ExtrairCONTRIBUINTE = "NÃO": Exit Function
        
        If .r0150.Exists(CHV_REG) Then IE = Util.ApenasNumeros(.r0150(CHV_REG)(dtoTitSPED.t0150("IE")))
        If IE <> "" Then ExtrairCONTRIBUINTE = "SIM" Else ExtrairCONTRIBUINTE = "NÃO"
    
    End With
    
End Function

Public Function RedimensionarArray(ByVal NumCampos As Long)
    
    ReDim Campo(1 To NumCampos) As Variant
    
End Function
