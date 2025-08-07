Attribute VB_Name = "clsAssistenteTributacao"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function ReprocessarSugestoes()

Dim Dados As Range, Linha As Range
Dim arrRelatorio As New ArrayList
Dim dicTitulos As New Dictionary
Dim Campos As Variant
    
    Set dicTitulos = Util.MapearTitulos(assTributacaoICMS, 3)
    If assTributacaoICMS.AutoFilterMode Then assTributacaoICMS.AutoFilter.ShowAllData
    
    Set Dados = Util.DefinirIntervalo(assTributacaoICMS, 4, 3)
    If Dados Is Nothing Then Exit Function
    
    a = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 100, "Reprocessando sugestões, por favor aguarde...", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            Campos(dicTitulos("INCONSISTENCIA")) = Empty
            Campos(dicTitulos("SUGESTAO")) = Empty
            Call AnalisarTributacoes(Campos, dicTitulos)
            
            arrRelatorio.Add Campos
            
        End If
        
    Next Linha
    
    Call Util.LimparDados(assTributacaoICMS, 4, False)
    Call Util.ExportarDadosArrayList(assTributacaoICMS, arrRelatorio)
    Call FuncoesFormatacao.DestacarInconsistencias(assTributacaoICMS)
        
End Function

Public Function AceitarSugestoes()

Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim Campos As Variant
Dim dicDados As New Dictionary
Dim UltimaSugestao As String
    
    Inicio = Now()
    Application.StatusBar = "Implementando sugestões selecionadas, por favor aguarde..."
    
    Set dicTitulos = Util.MapearTitulos(assTributacaoICMS, 3)
    Set Dados = assTributacaoICMS.Range("A4").CurrentRegion
    If Dados Is Nothing Then
        Call Util.MsgAlerta("Não foi encontrada nenhuma inconsistência!", "Inconsistências Fiscais")
        Exit Function
    End If
    
    For Each Linha In Dados.Rows
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            If Linha.EntireRow.Hidden = False And Campos(dicTitulos("SUGESTAO")) <> "" And Linha.Row > 3 Then
VERIFICAR:
                 
'                If UltimaSugestao = Campos(dicTitulos("SUGESTAO")) Then Trava = Trava + 1
'                UltimaSugestao = Campos(dicTitulos("SUGESTAO"))
'                Debug.Print UltimaSugestao
                
                Select Case Campos(dicTitulos("SUGESTAO"))
                        
                    Case Else
                        
                End Select
                
            End If
            
            If Linha.Row > 3 Then dicDados(Campos(dicTitulos("CHV_REG"))) = Campos
            
        End If
        
    Next Linha
    
    If dicDados.Count = 0 Then
        Call Util.MsgAlerta("Não existem sugestões para processar!", "Sugestões Fiscais")
        Exit Function
    End If
    
    
    If assTributacaoICMS.AutoFilterMode Then assTributacaoICMS.AutoFilter.ShowAllData
    Call Util.LimparDados(assTributacaoICMS, 4, False)
    Call Util.ExportarDadosDicionario(assTributacaoICMS, dicDados)
    
    Call FuncoesFormatacao.AplicarFormatacao(ActiveSheet)
    Call FuncoesFormatacao.DestacarInconsistencias(assTributacaoICMS)
    
    Call Util.MsgInformativa("Sugestões aplicadas com sucesso!", "Inclusão de Sugestões", Inicio)
    Application.StatusBar = False
    
End Function

Public Function IgnorarInconsistencias()

Dim Dados As Range, Linha As Range
Dim CHV_REG As String, INCONSISTENCIA$
Dim dicTitulos As New Dictionary
Dim dicDados As New Dictionary
Dim Resposta As VbMsgBoxResult
Dim Campos As Variant

    Resposta = MsgBox("Tem certeza que deseja ignorar as inconsistências selecionadas?" & vbCrLf & _
                      "Essa operação NÃO pode ser desfeita.", vbExclamation + vbYesNo, "Ignorar Inconsistências")
    
    If Resposta = vbNo Then Exit Function

    Inicio = Now()
    Application.StatusBar = "Ignorando as sugestões selecionadas, por favor aguarde..."
    
    Set dicTitulos = Util.MapearTitulos(assTributacaoICMS, 3)
    Set Dados = assTributacaoICMS.Range("A4").CurrentRegion
    If Dados Is Nothing Then
        Call Util.MsgAlerta("Não foi encontrada nenhuma inconsistência!", "Inconsistências Fiscais")
        Exit Function
    End If
    
    For Each Linha In Dados.Rows
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
                        
            If Linha.EntireRow.Hidden = False And Campos(dicTitulos("INCONSISTENCIA")) <> "" And Linha.Row > 3 Then
                
                CHV_REG = Campos(dicTitulos("CHV_REG"))
                INCONSISTENCIA = Campos(dicTitulos("INCONSISTENCIA"))
                
                'Verifica se o registro já possui inconsistências ignoradas, caso não exista cria
                If Not dicInconsistenciasIgnoradas.Exists(CHV_REG) Then Set dicInconsistenciasIgnoradas(CHV_REG) = New ArrayList
                
                'Verifica se a inconsistência já foi ignorada e caso contrário adiciona ela na lista
                If Not dicInconsistenciasIgnoradas(CHV_REG).contains(INCONSISTENCIA) Then _
                    dicInconsistenciasIgnoradas(CHV_REG).Add INCONSISTENCIA
                
                Campos(dicTitulos("INCONSISTENCIA")) = Empty
                Campos(dicTitulos("SUGESTAO")) = Empty
                Call AnalisarTributacoes(Campos, dicTitulos)
                
            End If
            
            If Linha.Row > 3 Then dicDados(Campos(dicTitulos("CHV_REG"))) = Campos
            
        End If
        
    Next Linha

    If dicDados.Count = 0 Then
        Call Util.MsgAlerta("Não existem Inconsistêncais a ignorar!", "Ignorar Inconsistências")
        Exit Function
    End If
    
    If assTributacaoICMS.AutoFilterMode Then assTributacaoICMS.AutoFilter.ShowAllData
    Call Util.LimparDados(assTributacaoICMS, 4, False)
    Call Util.ExportarDadosDicionario(assTributacaoICMS, dicDados)
    
    Call FuncoesFormatacao.AplicarFormatacao(ActiveSheet)
    Call FuncoesFormatacao.DestacarInconsistencias(assTributacaoICMS)
    
    Call Util.MsgInformativa("Inconsistências ignoradas com sucesso!", "Ignorar Inconsistências", Inicio)
    Application.StatusBar = False
    
End Function

Public Function GerarAnaliseTributacao()
    'Stop
End Function

Public Function AnalisarTributacoes(ByRef Campos As Variant, ByRef dicTitulos As Dictionary) As Variant
    
Dim Registro As String
Dim i As Integer

    If UBound(Campos) = -1 Then
        AnalisarTributacoes = Campos
        Exit Function
    End If
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    Registro = Campos(dicTitulos("REG") - i)
    
    Select Case Registro
        
        Case "C170"
'            Call AnalisarTributacoesC170(Campos, dicTitulos)
            
'        Case "C175"
'            Call AnalisarTributacoesC175(Campos, dicTitulos)
'
'        Case "F100"
'            Call AnalisarTributacoesF100(Campos, dicTitulos)
'
'        Case "F120"
'            Call AnalisarTributacoesF120(Campos, dicTitulos)
            
    End Select
    
    AnalisarTributacoes = Campos
    
End Function

Public Sub AtualizarRegistros()

Dim Campos As Variant, Campos0200, CamposC100, CamposC170, CamposC177, dicCampos, regCampo
Dim CHV_C170 As String, CHV_C177$, CHV_C100$, CHV_0200$
Dim dicTitulos0200 As New Dictionary
Dim dicTitulosC100 As New Dictionary
Dim dicTitulosC170 As New Dictionary
Dim dicTitulosC177 As New Dictionary
Dim dicDados0200 As New Dictionary
Dim dicDadosC100 As New Dictionary
Dim dicDadosC170 As New Dictionary
Dim dicDadosC177 As New Dictionary
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
    
    Inicio = Now()
    Application.StatusBar = "Preparando dados para atualização do SPED, por favor aguarde..."
    
    Campos0200 = Array("REG", "COD_BARRA", "COD_NCM", "EX_IPI", "CEST", "TIPO_ITEM")
    CamposC100 = Array("CHV_NFE", "NUM_DOC", "SER")
    CamposC170 = Array("IND_MOV", "CFOP", "VL_ITEM", "CST_ICMS", "VL_BC_ICMS", "ALIQ_ICMS", "VL_ICMS", "VL_BC_ICMS_ST", "ALIQ_ST", "VL_ICMS_ST")
    CamposC177 = Array("COD_INF_ITEM")
    
    Set dicTitulos0200 = Util.MapearTitulos(reg0200, 3)
    Set dicDados0200 = Util.CriarDicionarioRegistro(reg0200, "ARQUIVO", "COD_ITEM")
    
    Set dicTitulosC100 = Util.MapearTitulos(regC100, 3)
    Set dicDadosC100 = Util.CriarDicionarioRegistro(regC100)
    
    Set dicTitulosC170 = Util.MapearTitulos(regC170, 3)
    Set dicDadosC170 = Util.CriarDicionarioRegistro(regC170)
    
    Set dicTitulos = Util.MapearTitulos(assApuracaoICMS, 3)
    If assApuracaoICMS.AutoFilterMode Then assApuracaoICMS.AutoFilter.ShowAllData
    
    Set Dados = Util.DefinirIntervalo(assApuracaoICMS, 4, 3)
    If Dados Is Nothing Then Exit Sub
    
    a = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
    
        Call Util.AntiTravamento(a, 100, "Preparando dados para atualização do SPED, por favor aguarde...", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            CHV_0200 = VBA.Join(Array(Campos(dicTitulos("ARQUIVO")), Campos(dicTitulos("COD_ITEM"))))
            CHV_C100 = Campos(dicTitulos("CHV_PAI_FISCAL"))
            CHV_C170 = Campos(dicTitulos("CHV_REG"))
            
            'Atualizar dados do 0200
            If dicDados0200.Exists(CHV_0200) Then
                
                dicCampos = dicDados0200(CHV_0200)
                For Each regCampo In Campos0200
                    
                    If regCampo = "CEST" Or regCampo = "COD_BARRA" Or regCampo = "COD_NCM" Or regCampo = "EX_TIPI" Then
                        Campos(dicTitulos(regCampo)) = Util.FormatarTexto(Campos(dicTitulos(regCampo)))
                    End If
                    
                    If regCampo = "REG" Then Campos(dicTitulos(regCampo)) = "'0200"
                    dicCampos(dicTitulos0200(regCampo)) = Campos(dicTitulos(regCampo))
                    
                Next regCampo
                
                dicDados0200(CHV_0200) = dicCampos
                
            End If
            
            'Atualizar dados do C100
            If dicDadosC100.Exists(CHV_C100) Then
                
                dicCampos = dicDadosC100(CHV_C100)
                For Each regCampo In CamposC100
                    
                    dicCampos(dicTitulosC100(regCampo)) = Campos(dicTitulos(regCampo))
                    
                Next regCampo
                
                dicDadosC100(CHV_C100) = dicCampos
                
            End If
            
            'Atualizar dados do C170
            If dicDadosC170.Exists(CHV_C170) Then
                
                dicCampos = dicDadosC170(CHV_C170)
                For Each regCampo In CamposC170
                    
                    If regCampo = "CST_ICMS" Or regCampo = "COD_BARRA" Then
                        Campos(dicTitulos(regCampo)) = fnExcel.FormatarTexto(Campos(dicTitulos(regCampo)))
                    End If
                    
                    If regCampo Like "VL_*" Then Campos(dicTitulos(regCampo)) = VBA.Round(Campos(dicTitulos(regCampo)), 2)
                    dicCampos(dicTitulosC170(regCampo)) = Campos(dicTitulos(regCampo))
                
                Next regCampo
                
                dicDadosC170(CHV_C170) = dicCampos
                
            End If
                        
            'Atualizar dados do C177
            CHV_C177 = fnSPED.GerarChaveRegistro(CHV_C170, "C177")
            If dicDadosC177.Exists(CHV_C177) Then
                
                dicCampos = dicDadosC177(CHV_C177)
                For Each regCampo In CamposC177

                    dicCampos(dicTitulosC177(regCampo)) = Campos(dicTitulos(regCampo))
                    
                Next regCampo
                
                dicDadosC177(CHV_C177) = dicCampos
                
            ElseIf Campos(dicTitulos("COD_INF_ITEM")) <> "" Then
                
                dicCampos = Array("C177", Campos(dicTitulos("ARQUIVO")), CHV_C177, CHV_C170, Campos(dicTitulos("COD_INF_ITEM")))
                dicDadosC177(CHV_C177) = dicCampos
                
            End If
            
        End If

    Next Linha
    
    Application.StatusBar = "Atualizando dados do registro 0200, por favor aguarde..."
    Call Util.LimparDados(reg0200, 4, False)
    Call Util.ExportarDadosDicionario(reg0200, dicDados0200)
    
    Application.StatusBar = "Atualizando dados do registro C100, por favor aguarde..."
    Call Util.LimparDados(regC100, 4, False)
    Call Util.ExportarDadosDicionario(regC100, dicDadosC100)
    
    Application.StatusBar = "Atualizando dados do registro C170, por favor aguarde..."
    Call Util.LimparDados(regC170, 4, False)
    Call Util.ExportarDadosDicionario(regC170, dicDadosC170)
    
    Application.StatusBar = "Atualizando dados do registro C177, por favor aguarde..."
    Call Util.LimparDados(regC177, 4, False)
    Call Util.ExportarDadosDicionario(regC177, dicDadosC177)
    
    Application.StatusBar = "Atualizando dados do registro C190, por favor aguarde..."
    Call rC170.GerarC190(True)
    
    Application.StatusBar = "Atualizando valores dos impostos no registro C100, por favor aguarde..."
    Call rC170.AtualizarImpostosC100(True)
    Call r0200.AtualizarCodigoGenero(True)
    
    Call FuncoesFormatacao.AplicarFormatacao(ActiveSheet)
    Call FuncoesFormatacao.DestacarInconsistencias(assTributacaoICMS)
        
    Application.StatusBar = "Atualização concluída com sucesso!"
    Call Util.MsgInformativa("Registros atualizados com sucesso!", "Atualização de dados", Inicio)
    Application.StatusBar = False
    
End Sub




