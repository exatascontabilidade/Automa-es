Attribute VB_Name = "clsD100"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private EnumContrib As clsEnumeracoesSPEDContribuicoes
Private GerenciadorSPED As clsRegistrosSPED

Public Function IncluirMunicipios(ByVal Registro As String, ByRef dicMunicipiosCTe As Dictionary, _
                                  ByRef arrOcorrencias As ArrayList, ByRef EFD As ArrayList) As String

Dim chCTe As String
Dim Campos As Variant

    Campos = fnSPED.ExtrairCamposRegistro(Registro)
    
    chCTe = Campos(10)
    If dicMunicipiosCTe.Exists(chCTe) Then
        Campos(24) = dicMunicipiosCTe(chCTe)(0)
        Campos(25) = dicMunicipiosCTe(chCTe)(1)
    Else
        arrOcorrencias.Add "A chave de acesso: " & chCTe & ", não foi encontrada na pasta dos xmls."
    End If
    
    IncluirMunicipios = Join(Campos, "|")
    
End Function

Public Function ListarResumosD190()

Dim arrChaves As New ArrayList
Dim Dados As Variant, Intervalo
Dim i As Long, UltLin As Long
Dim Cels As Range, Cel As Range
Dim ARQUIVO As String

    Intervalo = regD100.Range("A3:" & Util.ConverterNumeroColuna(regD100.Range("A3").END(xlToRight).Column) & "3")
    UltLin = regD100.Range("A" & Rows.Count).END(xlUp).Row
    If UltLin > 3 Then

        Set Cels = regD100.Range("A4:" & Util.ConverterNumeroColuna(regD100.Range("A3").END(xlToRight).Column) & UltLin)
        For Each Cel In Cels.Rows

            If Cel.EntireRow.Hidden = False Then
                
                With CamposD100
                
                    .CHV_CTE = regD100.Cells(Cel.Row, EncontrarColuna("CHV_CTE", Intervalo))
                    If Not arrChaves.contains(.CHV_CTE) Then arrChaves.Add .CHV_CTE
                
                End With
            
            End If
            
        Next Cel
        
        If regD190.AutoFilterMode Then regD190.AutoFilter.ShowAllData
        Intervalo = regD190.Range("A3:" & Util.ConverterNumeroColuna(regD190.Range("A3").END(xlToRight).Column) & "3")
        If arrChaves.Count > 0 Then
            
            regD190.Range("$A$3:$" & Util.ConverterNumeroColuna(regD190.Range("A3").END(xlToRight).Column) & "$1048576").AutoFilter Field:=CInt(EncontrarColuna("CHV_PAI_FISCAL", Intervalo)), Criteria1:=arrChaves.toArray, Operator:=xlFilterValues
            regD190.Activate
        
        End If
        
    Else
    
        Call Util.MsgAlerta("É necessário selecionar pelo menos uma nota para usar essa função!", "Nenhuma nota selecionada")
    
    End If

End Function

Public Function GerarRegistroD200()

Dim dicTitulosD100 As New Dictionary
Dim dicTitulosD101Contr As New Dictionary
Dim dicTitulosD105 As New Dictionary
Dim dicTitulosD190 As New Dictionary
Dim dicTitulosD200 As New Dictionary
Dim dicTitulosD201 As New Dictionary
Dim dicTitulosD205 As New Dictionary
Dim dicDadosD100 As New Dictionary
Dim dicDadosD101Contr As New Dictionary
Dim dicDadosD105 As New Dictionary
Dim dicDadosD190 As New Dictionary
Dim dicDadosD200 As New Dictionary
Dim dicDadosD201 As New Dictionary
Dim dicDadosD205 As New Dictionary
Dim Campos As Variant, CamposDic, CamposD190
Dim CHV_REG As String, ARQUIVO$, NUM_DOC_INI$, NUM_DOC_FIN$, Chave$
Dim VL_DOC As Double, VL_ICMS#
Dim i As Long
    
    'Inicia contagem de tempo
    Inicio = Now()
    
    'Carrega dados do PIS e COFINS do D100 e filhos
    Set dicDadosD100 = Util.CriarDicionarioRegistro(regD100)
    Set dicDadosD101Contr = Util.CriarDicionarioRegistro(regD101_Contr, "CHV_PAI_FISCAL")
    Set dicDadosD105 = Util.CriarDicionarioRegistro(regD105, "CHV_PAI_FISCAL")
    Set dicDadosD190 = Util.CriarDicionarioRegistro(regD190, "CHV_PAI_FISCAL")
        
    'Carrega títulos do PIS e COFINS do D100 e filhos
    Set dicTitulosD100 = Util.MapearTitulos(regD100, 3)
    Set dicTitulosD101Contr = Util.MapearTitulos(regD101_Contr, 3)
    Set dicTitulosD105 = Util.MapearTitulos(regD105, 3)
    Set dicTitulosD190 = Util.MapearTitulos(regD190, 3)
    Set dicTitulosD200 = Util.MapearTitulos(regD200, 3)
    Set dicTitulosD201 = Util.MapearTitulos(regD201, 3)
    Set dicTitulosD205 = Util.MapearTitulos(regD205, 3)
    
    'Percorre registros do D100
    For Each Campos In dicDadosD100.Items()
        
        With CamposD200
            
            CHV_REG = Campos(dicTitulosD100("CHV_REG"))
            ARQUIVO = Campos(dicTitulosD100("ARQUIVO"))
            .REG = "D200"
            .COD_MOD = Campos(dicTitulosD100("COD_MOD"))
            .COD_SIT = Campos(dicTitulosD100("COD_SIT"))
            .SER = Campos(dicTitulosD100("SER"))
            .SUB = Campos(dicTitulosD100("SUB"))
            .NUM_DOC_INI = Campos(dicTitulosD100("NUM_DOC"))
            .NUM_DOC_FIN = Campos(dicTitulosD100("NUM_DOC"))
            .DT_REF = Campos(dicTitulosD100("DT_DOC"))
            .VL_DOC = Campos(dicTitulosD100("VL_DOC"))
            .VL_DESC = Campos(dicTitulosD100("VL_DESC"))
            .CHV_PAI = Campos(dicTitulosD100("CHV_PAI_FISCAL"))
            
            'Carrega dados do D190
            If dicDadosD190.Exists(CHV_REG) Then
            
                CamposDic = dicDadosD190(CHV_REG)
                .CFOP = CamposDic(dicTitulosD190("CFOP"))
                VL_ICMS = CamposDic(dicTitulosD190("VL_ICMS"))
                
            End If
            
            'Carrega dados do D101
            If dicDadosD101Contr.Exists(CHV_REG) Then
            
                CamposDic = dicDadosD101Contr(CHV_REG)
                
                CamposD201.CST_PIS = CamposDic(dicTitulosD101Contr("CST_PIS"))
                CamposD201.ALIQ_PIS = CamposDic(dicTitulosD101Contr("ALIQ_PIS"))
                CamposD201.COD_CTA = CamposDic(dicTitulosD101Contr("COD_CTA"))
                
            End If
            
            'Carrega dados do D105
            If dicDadosD105.Exists(CHV_REG) Then
            
                CamposDic = dicDadosD105(CHV_REG)
                
                CamposD205.CST_COFINS = CamposDic(dicTitulosD105("CST_COFINS"))
                CamposD205.ALIQ_COFINS = CamposDic(dicTitulosD105("ALIQ_COFINS"))
                CamposD205.COD_CTA = CamposDic(dicTitulosD105("COD_CTA"))
                
            End If
            
            'Define a chave do registro D200
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .COD_MOD, .SER, .COD_SIT, .CFOP, .DT_REF)
            'fnSPED.GerarChaveRegistro(.CHV_PAI, .COD_MOD, .COD_SIT, .SER, .SUB, .NUM_DOC_INI, .NUM_DOC_FIN, .CFOP, .DT_REF)

            Call IncuirRegistroD201(CDbl(.VL_DOC), VL_ICMS, dicDadosD201, ARQUIVO, dicTitulosD201)
            Call IncuirRegistroD205(CDbl(.VL_DOC), VL_ICMS, dicDadosD205, ARQUIVO, dicTitulosD205)
            
            
            Chave = fnSPED.GerarChaveRegistro(.CHV_PAI, .COD_MOD, .SER, .COD_SIT, .CFOP, .DT_REF)
            If dicDadosD200.Exists(Chave) Then
                
                'carrega campos do dicionário
                CamposDic = dicDadosD200(Chave)
                If LBound(CamposDic) = 0 Then i = 1 Else i = 0
                
                'carrega informações a serem atualizadas no registro
                NUM_DOC_INI = Util.ApenasNumeros(CamposDic(dicTitulosD200("NUM_DOC_INI") - i))
                NUM_DOC_FIN = Util.ApenasNumeros(CamposDic(dicTitulosD200("NUM_DOC_FIN") - i))
                VL_DOC = CamposDic(dicTitulosD200("VL_DOC") - i)
                
                .VL_DOC = VL_DOC + CDbl(.VL_DOC)
                If NUM_DOC_INI < .NUM_DOC_INI Then .NUM_DOC_INI = NUM_DOC_INI
                If NUM_DOC_FIN > .NUM_DOC_FIN Then .NUM_DOC_FIN = NUM_DOC_FIN
                
            End If
                        
            Campos = Array(.REG, ARQUIVO, .CHV_REG, .CHV_PAI, "", .COD_MOD, .COD_SIT, Util.FormatarTexto(.SER), Util.FormatarTexto(.SUB), _
                 Util.FormatarTexto(.NUM_DOC_INI), Util.FormatarTexto(.NUM_DOC_FIN), .CFOP, .DT_REF, CDbl(.VL_DOC), CDbl(.VL_DESC))
            
            dicDadosD200(Chave) = Campos
            
        End With
        
    Next Campos
    
    'Deleta informações dos registros
    Call Util.LimparDados(regD200, 4, False)
    Call Util.LimparDados(regD201, 4, False)
    Call Util.LimparDados(regD205, 4, False)
    
    'Atualiza dados dos registros D200 e filhos
    Call Util.ExportarDadosDicionario(regD200, dicDadosD200)
    Call Util.ExportarDadosDicionario(regD201, dicDadosD201)
    Call Util.ExportarDadosDicionario(regD205, dicDadosD205)
    
    Call Util.MsgInformativa("Registros D200 e filhos gerados com sucesso!", "Geração do D200 e filhos", Inicio)
    
End Function

Public Sub IncuirRegistroD201(ByRef vRec As Double, ByRef vICMS As Double, ByRef dicDadosD201 As Dictionary, _
    ByVal ARQUIVO As String, ByRef dicTitulosD201 As Dictionary)
    
Dim Campos As Variant, dicCampos
Dim i As Byte

    With CamposD201
        
        .REG = "D201"
        .VL_ITEM = vRec
        .VL_BC_PIS = vRec - vICMS
        .VL_PIS = VBA.Round(.VL_BC_PIS * .ALIQ_PIS, 2)
        .CHV_PAI = CamposD200.CHV_REG
        
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .CST_PIS, .ALIQ_PIS, .COD_CTA)
        If dicDadosD201.Exists(.CHV_REG) Then
        
            dicCampos = dicDadosD201(.CHV_REG)
            If LBound(dicCampos) = 0 Then i = 1 Else i = 0
            
            .VL_ITEM = dicCampos(dicTitulosD201("VL_ITEM") - i) + CDbl(.VL_ITEM)
            .VL_BC_PIS = dicCampos(dicTitulosD201("VL_BC_PIS") - i) + CDbl(.VL_BC_PIS)
            .VL_PIS = dicCampos(dicTitulosD201("VL_PIS") - i) + CDbl(.VL_PIS)
            
        End If
        
        Campos = Array(.REG, ARQUIVO, .CHV_REG, .CHV_PAI, "", Util.FormatarTexto(.CST_PIS), _
            CDbl(.VL_ITEM), CDbl(.VL_BC_PIS), CDbl(.ALIQ_PIS), CDbl(.VL_PIS), .COD_CTA)
            
        dicDadosD201(.CHV_REG) = Campos
        
    End With
    
End Sub

Public Sub IncuirRegistroD205(ByRef vRec As Double, ByRef vICMS As Double, ByRef dicDadosD205 As Dictionary, _
    ByVal ARQUIVO As String, ByRef dicTitulosD205 As Dictionary)
    
Dim Campos As Variant, dicCampos
Dim i As Byte
    
    With CamposD205
        
        .REG = "D205"
        .VL_ITEM = vRec
        .VL_BC_COFINS = vRec - vICMS
        .VL_COFINS = VBA.Round(.VL_BC_COFINS * .ALIQ_COFINS, 2)
        .CHV_PAI = CamposD200.CHV_REG
        
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .CST_COFINS, .ALIQ_COFINS, .COD_CTA)
        If dicDadosD205.Exists(.CHV_REG) Then
            
            dicCampos = dicDadosD205(.CHV_REG)
            If LBound(dicCampos) = 0 Then i = 1 Else i = 0
            
            .VL_ITEM = dicCampos(dicTitulosD205("VL_ITEM") - i) + CDbl(.VL_ITEM)
            .VL_BC_COFINS = dicCampos(dicTitulosD205("VL_BC_COFINS") - i) + CDbl(.VL_BC_COFINS)
            .VL_COFINS = dicCampos(dicTitulosD205("VL_COFINS") - i) + CDbl(.VL_COFINS)
            
        End If
        
        Campos = Array(.REG, ARQUIVO, .CHV_REG, .CHV_PAI, "", Util.FormatarTexto(.CST_COFINS), _
            CDbl(.VL_ITEM), CDbl(.VL_BC_COFINS), CDbl(.ALIQ_COFINS), CDbl(.VL_COFINS), .COD_CTA)
            
        dicDadosD205(.CHV_REG) = Campos
        
    End With
    
End Sub

Public Sub GerarRegistrosD101_D105_PISCOFINS()
    
    If Util.ChecarAusenciaDados(regD100, False) Then Exit Sub
    
    Call Util.DesabilitarControles
        Inicio = Now()
        
        Call CarregarRegistros
            
            Call GerarRegistrosD101_D105
            
            Call Util.AtualizarBarraStatus("Exportando registros, por favor aguarde...")
            Call ExportarRegistros
            
        Call DescarregarObjetos
        
        Call Util.AtualizarBarraStatus("Registros gerados com sucesso!")
        Call Util.MsgInformativa("Registros D101/D105 gerado com sucesso!", "Geração dos Registros D101/D105", Inicio)
        
        Call Util.AtualizarBarraStatus(False)
    
    Call Util.HabilitarControles
    
End Sub

Private Sub GerarRegistrosD101_D105()

Dim Campos As Variant
    
    With dtoRegSPED
        
        a = 0
        Comeco = Timer()
        For Each Campos In .rD100.Items()
            
            Call Util.AntiTravamento(a, 50, "Gerando registros D101/D105, por favor aguarde...", .rD100.Count, Comeco)
            Call CriarRegistroD101(Campos)
            Call CriarRegistroD105(Campos)
            
        Next Campos
    
    End With
    
End Sub

Private Sub CarregarRegistros()
    
    Call Util.AtualizarBarraStatus("Carregando dados dos registros, por favor aguarde...")
    
    Set EnumContrib = New clsEnumeracoesSPEDContribuicoes
    Set GerenciadorSPED = New clsRegistrosSPED
    
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    
    With dtoRegSPED
        
        If .r0110 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistro0110("ARQUIVO")
        If .rD100 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroD100
        
        If .rD101_Contr Is Nothing Then Set .rD101_Contr = New Dictionary ' Call GerenciadorSPED.CarregarDadosRegistroD101_Contr("CHV_PAI_CONTRIBUICOES", "CST_PIS", "ALIQ_PIS", "COD_CTA")
        If .rD105 Is Nothing Then Set .rD105 = New Dictionary 'Call GerenciadorSPED.CarregarDadosRegistroD105("CHV_PAI_CONTRIBUICOES", "CST_COFINS", "ALIQ_COFINS", "COD_CTA")
        
    End With
    
End Sub

Private Sub DescarregarObjetos()
    
    Call Util.AtualizarBarraStatus("Descarregando objetos, por favor aguarde...")
    
    Set EnumContrib = Nothing
    Set GerenciadorSPED = Nothing
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    
End Sub

Public Sub CriarRegistroD101(ByVal Campos As Variant)

Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    With CamposD101_Contr
        
        .REG = "D101"
        .ARQUIVO = Campos(dtoTitSPED.tD100("ARQUIVO") - i)
        .IND_NAT_FRT = ""
        .VL_ITEM = Campos(dtoTitSPED.tD100("VL_DOC") - i)
        .CST_PIS = ExtrairCST_PIS_COFINS_AquisicaoFrete(.ARQUIVO)
        .NAT_BC_CRED = ""
        .VL_BC_PIS = VBA.Round(.VL_ITEM - Campos(dtoTitSPED.tD100("VL_ICMS") - i), 2)
        .ALIQ_PIS = ExtrairALIQ_PIS_AquisicaoFrete(.ARQUIVO)
        .VL_PIS = VBA.Round(.VL_BC_PIS * .ALIQ_PIS, 2)
        .COD_CTA = fnExcel.FormatarTexto(Campos(dtoTitSPED.tD100("COD_CTA") - i))
        .CHV_PAI_CONTRIBUICOES = Campos(dtoTitSPED.tD100("CHV_REG") - i)
        
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI_CONTRIBUICOES, .CST_PIS, .ALIQ_PIS, .COD_CTA)
        
        Campos = Array(.REG, .ARQUIVO, .CHV_REG, "", .CHV_PAI_CONTRIBUICOES, .IND_NAT_FRT, CDbl(.VL_ITEM), _
            .CST_PIS, .NAT_BC_CRED, CDbl(.VL_BC_PIS), CDbl(.ALIQ_PIS), CDbl(.VL_PIS), .COD_CTA)
            
        dtoRegSPED.rD101_Contr(.CHV_REG) = Campos
        
    End With
    
End Sub

Public Sub CriarRegistroD105(ByVal Campos As Variant)

Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    With CamposD105
        
        .REG = "D105"
        .ARQUIVO = Campos(dtoTitSPED.tD100("ARQUIVO") - i)
        .IND_NAT_FRT = ""
        .VL_ITEM = Campos(dtoTitSPED.tD100("VL_DOC") - i)
        .CST_COFINS = ExtrairCST_PIS_COFINS_AquisicaoFrete(.ARQUIVO)
        .NAT_BC_CRED = ""
        .VL_BC_COFINS = VBA.Round(.VL_ITEM - Campos(dtoTitSPED.tD100("VL_ICMS") - i), 2)
        .ALIQ_COFINS = ExtrairALIQ_COFINS_AquisicaoFrete(.ARQUIVO)
        .VL_COFINS = VBA.Round(.VL_BC_COFINS * .ALIQ_COFINS, 2)
        .COD_CTA = fnExcel.FormatarTexto(Campos(dtoTitSPED.tD100("COD_CTA") - i))
        .CHV_PAI_CONTRIBUICOES = Campos(dtoTitSPED.tD100("CHV_REG") - i)
        
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI_CONTRIBUICOES, .CST_COFINS, .ALIQ_COFINS, .COD_CTA)
        
        Campos = Array(.REG, .ARQUIVO, .CHV_REG, "", .CHV_PAI_CONTRIBUICOES, .IND_NAT_FRT, CDbl(.VL_ITEM), _
            .CST_COFINS, .NAT_BC_CRED, CDbl(.VL_BC_COFINS), CDbl(.ALIQ_COFINS), CDbl(.VL_COFINS), .COD_CTA)
            
        dtoRegSPED.rD105(.CHV_REG) = Campos
        
    End With
    
End Sub

Private Function ExtrairCST_PIS_COFINS_AquisicaoFrete(ByVal ARQUIVO As String) As String

Dim i As Byte
Dim Campos As Variant
Dim COD_INC_TRIB As String
        
    Campos = dtoRegSPED.r0110(ARQUIVO)
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    COD_INC_TRIB = VBA.Left(Util.ApenasNumeros(Campos(dtoTitSPED.t0110("COD_INC_TRIB") - i)), 1)
    
    Select Case COD_INC_TRIB
        
        Case "1"
            ExtrairCST_PIS_COFINS_AquisicaoFrete = EnumContrib.ValidarEnumeracao_CST_PIS_COFINS("50")
            
        Case "2"
            ExtrairCST_PIS_COFINS_AquisicaoFrete = EnumContrib.ValidarEnumeracao_CST_PIS_COFINS("70")
            
    End Select
    
End Function

Private Function ExtrairALIQ_PIS_AquisicaoFrete(ByVal ARQUIVO As String) As Double

Dim i As Byte
Dim Campos As Variant
Dim COD_INC_TRIB As String
        
    Campos = dtoRegSPED.r0110(ARQUIVO)
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    COD_INC_TRIB = VBA.Left(Util.ApenasNumeros(Campos(dtoTitSPED.t0110("COD_INC_TRIB") - i)), 1)
    
    Select Case COD_INC_TRIB
        
        Case "1"
            ExtrairALIQ_PIS_AquisicaoFrete = 0.0165
            
        Case "2"
            ExtrairALIQ_PIS_AquisicaoFrete = 0
            
    End Select
    
End Function

Private Function ExtrairALIQ_COFINS_AquisicaoFrete(ByVal ARQUIVO As String) As Double

Dim i As Byte
Dim Campos As Variant
Dim COD_INC_TRIB As String
        
    Campos = dtoRegSPED.r0110(ARQUIVO)
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    COD_INC_TRIB = VBA.Left(Util.ApenasNumeros(Campos(dtoTitSPED.t0110("COD_INC_TRIB") - i)), 1)
    
    Select Case COD_INC_TRIB
        
        Case "1"
            ExtrairALIQ_COFINS_AquisicaoFrete = 0.076
            
        Case "2"
            ExtrairALIQ_COFINS_AquisicaoFrete = 0
            
    End Select
    
End Function

Private Function ExportarRegistros()
    
Dim i As Long
Dim Plan As Worksheet
Dim Registro As Dictionary
Dim colPlanilhas As Collection
Dim colRegistros As Collection

    Set colPlanilhas = CarregarPlanilhasDestino
    Set colRegistros = ListarDadosRegistros
    
    For i = 1 To colPlanilhas.Count
        
        Set Plan = colPlanilhas.item(i)
        Set Registro = colRegistros.item(i)
        
        Call Util.AtualizarBarraStatus("Exportando dados do registro " & Plan.name)
        
        Call Util.LimparDados(Plan, 4, False)
        Call Util.ExportarDadosDicionario(Plan, Registro, "A4")
        
    Next i
        
End Function

Private Function CarregarPlanilhasDestino() As Collection

Dim colPlanilhas As New Collection
    
    colPlanilhas.Add regD101_Contr
    colPlanilhas.Add regD105
    
    Set CarregarPlanilhasDestino = colPlanilhas
    
End Function

Private Function ListarDadosRegistros() As Collection

Dim colRegistros As New Collection
    
    With dtoRegSPED
        
        colRegistros.Add .rD101_Contr
        colRegistros.Add .rD105
        
    End With
    
    Set ListarDadosRegistros = colRegistros
    
End Function
