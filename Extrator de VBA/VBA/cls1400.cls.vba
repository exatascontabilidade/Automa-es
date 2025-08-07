Attribute VB_Name = "cls1400"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function GerarSaidasCTeMunicipio()

Dim COD_MUN_ORIG As String, ARQUIVO$, COD_UF$, COD_SIT$, CHV_1001$, CHV_1400$
Dim Campos As Variant, CamposDic, Chave
Dim dicTitulos0000 As New Dictionary
Dim dicTitulosD100 As New Dictionary
Dim dicTitulos1001 As New Dictionary
Dim dicTitulos1400 As New Dictionary
Dim dicDados As New Dictionary
Dim dicDados0000 As New Dictionary
Dim dicDadosD100 As New Dictionary
Dim dicDados1001 As New Dictionary
Dim dicDados1400 As New Dictionary
Dim VL_DOC As Double
Dim i As Byte
    
    'Carrega os dados dos registros necessários
    Set dicDados0000 = Util.CriarDicionarioRegistro(reg0000, "ARQUIVO")
    Set dicTitulos0000 = Util.MapearTitulos(reg0000, 3)
    
    Set dicDadosD100 = Util.CriarDicionarioRegistro(regD100)
    Set dicTitulosD100 = Util.MapearTitulos(regD100, 3)
    
    Set dicDados1001 = Util.CriarDicionarioRegistro(reg1001, "ARQUIVO")
    Set dicTitulos1001 = Util.MapearTitulos(reg1001, 3)
    
    Set dicDados1400 = Util.CriarDicionarioRegistro(reg1400)
    Set dicTitulos1400 = Util.MapearTitulos(reg1400, 3)
    
    'Gera os dados do registro 1400 a partir das informações do registro D100
    For Each Campos In dicDadosD100.Items()
        
        ARQUIVO = Campos(dicTitulosD100("ARQUIVO"))
        COD_MUN_ORIG = Campos(dicTitulosD100("COD_MUN_ORIG"))
        COD_SIT = Campos(dicTitulosD100("COD_SIT"))
        VL_DOC = Campos(dicTitulosD100("VL_DOC"))
        
        If dicDados1001.Exists(ARQUIVO) Then CHV_1001 = dicDados1001(ARQUIVO)(dicTitulos1001("CHV_REG"))
        If dicDados1001.Exists(ARQUIVO) Then COD_UF = VBA.Left(dicDados0000(ARQUIVO)(dicTitulos0000("COD_MUN")), 2)
        
        CHV_1400 = fnSPED.GerarChaveRegistro(CHV_1001, "", COD_MUN_ORIG)
        If dicDados.Exists(CHV_1400) Then
                        
            CamposDic = dicDados(CHV_1400)
            If LBound(CamposDic) = 0 Then i = 1 Else i = 0
            
            VL_DOC = VL_DOC + CDbl(CamposDic(dicTitulos1400("VALOR") - i))
            
        End If
        
        'Adiciona dados do registro caso caso o valor do documento seja maior que zero
        If VL_DOC > 0 And COD_MUN_ORIG Like COD_UF & "*" Then
        
            Campos = Array("'1400", ARQUIVO, CHV_1400, CHV_1001, "", COD_MUN_ORIG, VL_DOC)
            dicDados(CHV_1400) = Campos
        
        End If
        
    Next Campos
    
    'Verifica se já existe um registro 1400 com a mesma chave
    For Each Chave In dicDados.Keys()
        dicDados1400(Chave) = dicDados(Chave)
    Next Chave
    
    'Limpa os dados do registro 1400
    Call Util.LimparDados(reg1400, 4, False)
    
    'Exporta os dados coletados para o registro 1400
    Call Util.ExportarDadosDicionario(reg1400, dicDados1400)
    
    'Verifica se houveram dados a exportar e emite mensagem ao usuário
    If dicDados.Count > 0 Then
        Call Util.MsgAviso("Registros 1400 gerados com sucesso!", "Geração do registro 1400")
    Else
        Call Util.MsgAlerta("Nenhum dado encontrado para gerar o registro 1400!", "Geração do registro 1400")
    End If
    
End Function

Public Function AgruparRegistros()

Dim dicAuxiliar As New Dictionary
Dim dicTitulos As New Dictionary
Dim Dados As Range, Linha As Range
Dim Campos As Variant
Dim CHV_REG As String
    
    Set dicTitulos = Util.MapearTitulos(reg1400, 3)
    
    Set Dados = Util.DefinirIntervalo(reg1400, 4, 3)
    If Dados Is Nothing Then Exit Function
    
    a = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 100, "Agrupando dados do registro 1400, por favor aguarde...", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            CHV_REG = Campos(dicTitulos("CHV_REG"))
            Call Util.SomarValoresDicionario(dicAuxiliar, Campos, CHV_REG)
            
        End If
        
    Next Linha
    
    Call Util.LimparDados(reg1400, 4, False)
    Call Util.ExportarDadosDicionario(reg1400, dicAuxiliar)
    
End Function
