Attribute VB_Name = "clsK200"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private dicTitulosK200 As New Dictionary
Private arrDadosK200 As New ArrayList
Private dicTitulos0200 As New Dictionary
Private dicDados0200 As New Dictionary
Private arrDados As New ArrayList

Public Function GerarSaldoEstoque()

Dim dicDados0200 As New Dictionary
Dim dicDados0220 As New Dictionary
Dim dicDadosC100 As New Dictionary
Dim dicDadosC170 As New Dictionary
Dim dicDadosK200 As New Dictionary
Dim dicTitulos0200 As New Dictionary
Dim dicTitulos0220 As New Dictionary
Dim dicTitulosC100 As New Dictionary
Dim dicTitulosC170 As New Dictionary
Dim dicTitulosK200 As New Dictionary
Dim Dados, Intervalo, Campos, Chave
Dim i As Long, UltLin&
Dim ARQUIVO As String, UND_COM As String, TIPO_ITEM As String
    
    Set dicDados0200 = Util.CriarDicionarioRegistro(reg0200)
    Set dicDados0220 = Util.CriarDicionarioRegistro(reg0220)
    Set dicDadosC100 = Util.CriarDicionarioRegistro(regC100)
    Set dicDadosC170 = Util.CriarDicionarioRegistro(regC170)
    Set dicDadosK200 = Util.CriarDicionarioRegistro(regK200)
    
    Call Util.IndexarCampos(dicDados0200("ARQUIVO|COD_ITEM"), dicTitulos0200)
    Call dicDados0200.Remove("ARQUIVO|COD_ITEM")
    
    Call Util.IndexarCampos(dicDados0220("ARQUIVO|CHV_PAI|UNID_CONVFAT_CONV"), dicTitulos0200)
    Call dicDados0220.Remove("ARQUIVO|CHV_PAI|UNID_CONVFAT_CONV")
    
    Call Util.IndexarCampos(dicDadosC100("CHV_NFE"), dicTitulosC100)
    Call dicDadosC100.Remove("CHV_NFE")
    
    Call Util.IndexarCampos(dicDadosC170("CHV_PAI|NUM_ITEM"), dicTitulosC170)
    Call dicDadosC170.Remove("CHV_PAI|NUM_ITEM")
    
    Call Util.IndexarCampos(dicDadosK200("ARQUIVO|CHV_PAI|DT_ESTCOD_ITEMIND_ESTCOD_PART"), dicTitulosK200)
    Call dicDadosK200.Remove("ARQUIVO|CHV_PAI|DT_ESTCOD_ITEMIND_ESTCOD_PART")
    
    a = Util.AntiTravamento(a)
    With CamposK200
        
        For Each Chave In dicDadosC170.Keys()
            
            Campos = VBA.Split(Chave, "|")
            
            .REG = "K200"
            .DT_EST = Util.FimMes(dicDadosC100(Campos(0))(dicTitulosC100("DT_DOC")))
            .COD_ITEM = dicDadosC170(Chave)(dicTitulosC170("COD_ITEM"))
            .QTD = dicDadosC170(Chave)(dicTitulosC170("QTD"))
            .IND_EST = "0"
            .COD_PART = ""
            
            UND_COM = dicDadosC170(Chave)(dicTitulosC170("UNID"))
            If dicDados0220.Exists(ARQUIVO & .COD_ITEM & UND_COM) Then .QTD = CDbl(.QTD) * CDbl(dicDados0220(.COD_ITEM & UND_COM)(dicTitulos0220("FAT_CONV")))
            
            Campos = Array("", .REG, .DT_EST, .COD_ITEM, .QTD, .IND_EST, .COD_PART, "")
            
            TIPO_ITEM = VBA.Left(dicDados0200(ARQUIVO & .COD_ITEM)(dicTitulos0200("TIPO_ITEM")), 2)
                        
            Select Case TIPO_ITEM
                    
                Case "00", "01", "02", "03", "04", "05", "06", "10"
                
                    .CHV_PAI = Dados(i, EncontrarColuna("CHV_PAI_FISCAL", Intervalo))
                    .CHV_REG = fnSPED.MontarChaveRegistro(ARQUIVO, .CHV_PAI, .DT_EST, .COD_ITEM, .IND_EST, .COD_PART)
                    
                    dicDadosK200(.CHV_REG) = fnSPED.GerarRegistro(Campos)
                    
            End Select
            
        Next Chave
        
    End With
    
End Function

Public Function ExcluirItem(ByVal Registro As String, ByRef dicDados0200 As Dictionary) As String

Dim cConta As String
Dim Campos As Variant
    
    Campos = fnSPED.ExtrairCamposRegistro(Registro)
    
    If Campos(4) > 0 Then
        
        Select Case dicDados0200(cConta)
            
            Case "00", "01", "02", "03", "04", "05", "06", "10"
                ExcluirItem = fnSPED.GerarRegistro(Campos)
                
        End Select
    
    End If
    
End Function

Public Sub ListarProdutosAusentes()

Dim Seguir As Boolean
Dim Plan As Worksheet
    
    Call Util.AtualizarBarraStatus("Listando produtos ausentes, por favor aguarde...")
    
    Inicio = Now()
    
    Seguir = CarregarDadosRegistros
    If Not Seguir Then Exit Sub
    
    Call VerificarProdutosAusentes
    
    Set Plan = FormatarPlanilhaDestino()
    
    Call Util.LimparDados(Plan, 2, False)
    Call Util.ExportarDadosArrayList(Plan, arrDados)
    
    Call Util.MsgInformativa("Os produtos ausentes foram listados com sucesso!", "Listagem de produtos ausentes", Inicio)
    
    Call LimparObjetos
    
End Sub

Private Function CarregarDadosRegistros() As Boolean
    
    Call Util.AtualizarBarraStatus("Coletando dados do registro K200, por favor aguarde...")
    Set dicTitulosK200 = Util.MapearTitulos(regK200, 3)
    Set arrDadosK200 = Util.CriarArrayListRegistro(regK200)
    
    If arrDadosK200.Count = 0 Then
        
        CarregarDadosRegistros = False
        Call AlertarUsuario
        
        Exit Function
        
    End If
    
    Call Util.AtualizarBarraStatus("Coletando dados do registro 0200, por favor aguarde...")
    
    Set dicTitulos0200 = Util.MapearTitulos(reg0200, 3)
    Set dicDados0200 = Util.CriarDicionarioRegistro(reg0200, "ARQUIVO", "COD_ITEM")
    
    CarregarDadosRegistros = True
    
End Function

Private Sub AlertarUsuario()
    
Dim Msg As String
    
    Msg = "Para utilizar essa função é necessários existir dados na planilha."
    Call Util.MsgAlerta(Msg, "Dados Inexistentes")
    
End Sub

Private Function VerificarProdutosAusentes()

Dim Campos As Variant
    
    Call Util.AtualizarBarraStatus("Verificando produtos ausentes, por favor aguarde...")
    
    a = 0
    For Each Campos In arrDadosK200
        
        Call Util.AntiTravamento(a, 100)
        Call ChecarAusenciaProduto(Campos)
        
    Next Campos

End Function

Private Function ChecarAusenciaProduto(ByRef Campos As Variant)

Dim ARQUIVO As String, COD_ITEM$, Chave$
    
    ARQUIVO = Campos(dicTitulosK200("ARQUIVO"))
    COD_ITEM = Campos(dicTitulosK200("COD_ITEM"))
    
    Chave = ARQUIVO & COD_ITEM
    If Not dicDados0200.Exists(Chave) Then arrDados.Add Array(COD_ITEM, "")
    
End Function

Private Function FormatarPlanilhaDestino() As Worksheet

Dim TitulosCabecalho As Variant
Dim PastaTrabalho As Workbook
Dim Plan As Worksheet
    
    Set PastaTrabalho = Workbooks.Add
    Set Plan = PastaTrabalho.ActiveSheet
    Plan.name = "Cadastros 0200 ausentes"
    
    TitulosCabecalho = Array("COD_ITEM")

    With Plan.Range("A1").Resize(1, UBound(TitulosCabecalho) + 1)

        .value = TitulosCabecalho
        .Font.Bold = True

    End With

    Set FormatarPlanilhaDestino = Plan
    
End Function

Private Sub LimparObjetos()
    
    Set dicTitulosK200 = Nothing
    Set arrDadosK200 = Nothing
    
    Set dicTitulos0200 = Nothing
    Set dicDados0200 = Nothing
    
    Set arrDados = Nothing
    
    Call Util.AtualizarBarraStatus(False)
    
End Sub


