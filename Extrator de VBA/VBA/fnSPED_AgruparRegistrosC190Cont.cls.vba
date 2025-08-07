Attribute VB_Name = "fnSPED_AgruparRegistrosC190Cont"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Implements IAgruparRegistros

Private dicTitulos As New Dictionary
Private dicDados As New Dictionary
Private arrDados As New ArrayList

Public Sub IAgruparRegistros_AgruparRegistros()
    
    If Util.ChecarAusenciaDados(regC190_Contr, False) Then Exit Sub
    Call Util.AtualizarBarraStatus("Agrupando registros, por favor aguarde...")
    
    Call IAgruparRegistros_CarregarDados
    Call IAgruparRegistros_ProcessarAgrupamento
    
    Call Util.LimparDados(regC190_Contr, 4, False)
    Call Util.ExportarDadosDicionario(regC190_Contr, dicDados)
    
    Call Util.MsgInformativa("Os registros C190 foram agrupados com sucesso!", "Agrupamento dos Registros C190", Inicio)
    
End Sub

Private Sub IAgruparRegistros_CarregarDados()
    
    Call Util.AtualizarBarraStatus("Coletando dados do registro C190, por favor aguarde...")
    Set dicTitulos = Util.MapearTitulos(regC190_Contr, 3)
    Set arrDados = Util.CriarArrayListRegistro(regC190_Contr)
    
End Sub

Private Sub IAgruparRegistros_ProcessarAgrupamento()

Dim Campos As Variant
Dim Chave As String
    
    For Each Campos In arrDados
        
        With CamposC190Contr
            
            Chave = IAgruparRegistros_ObterChaveOperacao(Campos)
            If dicDados.Exists(Chave) Then Campos = IAgruparRegistros_AtualizarRegistroOperacao(Campos, Chave)
            
            dicDados(Chave) = Campos
            
        End With
        
    Next Campos
    
End Sub

Private Function IAgruparRegistros_ObterChaveOperacao(ByRef Campos As Variant) As String

Dim Chave As String
    
    With CamposC190Contr
        
        .COD_MOD = Campos(dicTitulos("COD_MOD"))
        .COD_ITEM = Campos(dicTitulos("COD_ITEM"))
        .COD_NCM = Campos(dicTitulos("COD_NCM"))
        .EX_IPI = Campos(dicTitulos("EX_IPI"))
        
        IAgruparRegistros_ObterChaveOperacao = Util.UnirCampos(.COD_MOD, .COD_ITEM, .COD_NCM, .EX_IPI)
        
    End With
    
End Function

Private Function IAgruparRegistros_AtualizarRegistroOperacao(ByRef Campos As Variant, ByVal Chave As String) As Variant
    
Dim CamposDic As Variant
    
    CamposDic = dicDados(Chave)
    
    Campos = AtualizarValorOperacao(Campos, CamposDic)
    Campos = AtualizarDataInicialOperacao(Campos, CamposDic)
    Campos = AtualizarDataFinalOperacao(Campos, CamposDic)
    
    IAgruparRegistros_AtualizarRegistroOperacao = Campos
    
End Function

Private Function AtualizarValorOperacao(ByRef Campos As Variant, ByVal CamposDic As Variant) As Variant
    
Dim NomeCampo As String
    
    NomeCampo = "VL_TOT_ITEM"
    
    With CamposC190Contr
        
        .VL_TOT_ITEM = fnExcel.ConverterValores(CamposDic(dicTitulos(NomeCampo)), True, 2)
        .VL_TOT_ITEM = .VL_TOT_ITEM + fnExcel.ConverterValores(Campos(dicTitulos(NomeCampo)), True, 2)
        
        Campos(dicTitulos(NomeCampo)) = .VL_TOT_ITEM
        
        AtualizarValorOperacao = Campos
        
    End With
    
End Function

Private Function AtualizarDataInicialOperacao(ByRef Campos As Variant, ByVal CamposDic As Variant) As Variant
    
Dim NomeCampo As String
Dim DataInicial As String

    NomeCampo = "DT_REF_INI"
    DataInicial = Campos(dicTitulos(NomeCampo))
    
    With CamposC190Contr
                
        .DT_REF_INI = fnExcel.FormatarData(CamposDic(dicTitulos(NomeCampo)))
        
        If CDate(DataInicial) < CDate(.DT_REF_INI) Then .DT_REF_INI = DataInicial
        
        Campos(dicTitulos(NomeCampo)) = .DT_REF_INI
        
        AtualizarDataInicialOperacao = Campos
        
    End With
    
End Function

Private Function AtualizarDataFinalOperacao(ByRef Campos As Variant, ByVal CamposDic As Variant) As Variant
    
Dim NomeCampo As String
Dim DataFinal As String

    NomeCampo = "DT_REF_FIN"
    DataFinal = Campos(dicTitulos(NomeCampo))
    
    With CamposC190Contr
        
        .DT_REF_FIN = fnExcel.FormatarData(CamposDic(dicTitulos(NomeCampo)))
        If CDate(DataFinal) > CDate(.DT_REF_FIN) Then .DT_REF_FIN = DataFinal
        
        Campos(dicTitulos(NomeCampo)) = .DT_REF_FIN
        
        AtualizarDataFinalOperacao = Campos
        
    End With
    
End Function

