Attribute VB_Name = "clsC190Contrib_AtualizarNCM"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Implements Iprocessador

Private dicTitulosC190Contrib As New Dictionary
Private dicDadosC190Contrib As New Dictionary
Private arrDadosC190Contrib As New ArrayList
Private dicTitulos0200 As New Dictionary
Private dicDados0200 As New Dictionary
Private arrDados As New ArrayList

Public Sub IProcessador_Executar(ByVal TipoSPED As String, ByVal Registro As String)

Dim Seguir As Boolean
    
    Call Util.AtualizarBarraStatus("Atualizando NCMs, por favor aguarde...")
        
    Seguir = CarregarDadosRegistros
    If Not Seguir Then Exit Sub
    
    Call AtualizarNCM_C190_0200
    
    Call Util.LimparDados(regC190_Contr, 4, False)
    Call Util.ExportarDadosArrayList(regC190_Contr, arrDados)
    
    Call Util.MsgInformativa("Os NCMs do registro C190 foram atualizados com sucesso!", "Atualização de NCMs", Inicio)
    
    Call LimparObjetos
    
End Sub

Private Function CarregarDadosRegistros() As Boolean
    
    Call Util.AtualizarBarraStatus("Coletando dados do registro C190, por favor aguarde...")
    Set dicTitulosC190Contrib = Util.MapearTitulos(regC190_Contr, 3)
    Set arrDadosC190Contrib = Util.CriarArrayListRegistro(regC190_Contr)
    
    If arrDadosC190Contrib.Count = 0 Then
        
        CarregarDadosRegistros = False
        Call Util.MsgAlerta_AusenciaDados
        
        Exit Function
        
    End If
    
    Call Util.AtualizarBarraStatus("Coletando dados do registro 0200, por favor aguarde...")
    
    Set dicTitulos0200 = Util.MapearTitulos(reg0200, 3)
    Set dicDados0200 = Util.CriarDicionarioRegistro(reg0200, "ARQUIVO", "COD_ITEM")
    
    CarregarDadosRegistros = True
    
End Function

Private Function AtualizarNCM_C190_0200()

Dim Campos As Variant
    
    Call Util.AtualizarBarraStatus("Atualizando NCMs, por favor aguarde...")
    
    a = 0
    For Each Campos In arrDadosC190Contrib
        
        Call Util.AntiTravamento(a, 100)
        Call VerificarNCM_C190_0200(Campos)
        
    Next Campos

End Function

Private Function VerificarNCM_C190_0200(ByRef Campos As Variant)

Dim ARQUIVO As String, COD_ITEM$, COD_NCM$, Chave$, COD_NCM_0200$
    
    ARQUIVO = Campos(dicTitulosC190Contrib("ARQUIVO"))
    COD_ITEM = Campos(dicTitulosC190Contrib("COD_ITEM"))
    COD_NCM = Campos(dicTitulosC190Contrib("COD_NCM"))
    
    Chave = ARQUIVO & COD_ITEM
    If dicDados0200.Exists(Chave) Then
        
        COD_NCM_0200 = CarregarNCM_0200(Chave)
        Campos(dicTitulosC190Contrib("COD_NCM")) = COD_NCM_0200
        
    End If
    
    arrDados.Add Campos
    
End Function

Private Function CarregarNCM_0200(ByVal Chave As String)
    
Dim Campos As Variant
Dim COD_NCM As String
    
    Campos = dicDados0200(Chave)
    CarregarNCM_0200 = Campos(dicTitulos0200("COD_NCM"))
    
End Function

Private Sub LimparObjetos()
    
    Set dicTitulosC190Contrib = Nothing
    Set arrDadosC190Contrib = Nothing
    
    Set dicTitulos0200 = Nothing
    Set dicDados0200 = Nothing
    
    Set arrDados = Nothing
    
    Call Util.AtualizarBarraStatus(False)
    
End Sub
