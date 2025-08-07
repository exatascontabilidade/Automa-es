Attribute VB_Name = "clsAssistenteEstoque"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Sub GerarRelatorio()

Dim CHV_REG As String, COD_ITEM$, UNID_COM$, UNID_INV$, CHV_REG_0200$, CHV_REG_C100$, CHV_REG_C170$, IND_OPER$, CHV_NFE$, REG$, ARQUIVO$, DT_OPER$, DESCR_ITEM$, COD_BARRA$, TIPO_ITEM$, IND_MOV$, UNID$, UND_INV$, CFOP$
Dim FAT_CONV As Double, QTD_COM#, VL_ITEM#, QTD_INV#, VL_UNIT_INV#, VL_UNIT_COM#, QTD#, VL_UNIT#
Dim dicTitulosC170 As New Dictionary
Dim dicTitulos0190 As New Dictionary
Dim dicTitulos0200 As New Dictionary
Dim dicTitulos0220 As New Dictionary
Dim dicTitulosC100 As New Dictionary
Dim dicDados0190 As New Dictionary
Dim dicDados0200 As New Dictionary
Dim dicDados0220 As New Dictionary
Dim dicDadosC100 As New Dictionary
Dim Dados As Range, Linha As Range
Dim arrRelatorio As New ArrayList
Dim dicTitulos As New Dictionary
Dim arrDados As New ArrayList
Dim Campos As Variant
    
    Inicio = Now()
    Application.StatusBar = "Gerando relatório inteligente de movimentação, por favor aguarde..."
    
    Set dicTitulos = Util.MapearTitulos(relInteligenteEstoque, 3)
    
    Set dicTitulos0190 = Util.MapearTitulos(reg0190, 3)
    Set dicDados0190 = Util.CriarDicionarioRegistro(reg0190, "ARQUIVO", "UNID")
    
    Set dicTitulos0200 = Util.MapearTitulos(reg0200, 3)
    Set dicDados0200 = Util.CriarDicionarioRegistro(reg0200, "ARQUIVO", "COD_ITEM")
    
    Set dicTitulos0220 = Util.MapearTitulos(reg0220, 3)
    Set dicDados0220 = Util.CriarDicionarioRegistro(reg0220)
    
    Set dicTitulosC100 = Util.MapearTitulos(regC100, 3)
    Set dicDadosC100 = Util.CriarDicionarioRegistro(regC100)
    
    Set dicTitulosC170 = Util.MapearTitulos(regC170, 3)
    Set Dados = Util.DefinirIntervalo(regC170, 4, 3)
    If Dados Is Nothing Then
        Call Util.MsgAlerta("Não existem dados nos registros C170", "Dados indisponíveis")
        Exit Sub
    End If
    
    a = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
    
        Call Util.AntiTravamento(a, 100, "Gerando relatório inteligente de movimentação, por favor aguarde...", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            REG = Campos(dicTitulosC170("REG"))
            ARQUIVO = Campos(dicTitulosC170("ARQUIVO"))
            CHV_REG_C100 = Campos(dicTitulosC170("CHV_PAI_FISCAL"))
            CHV_REG_C170 = Campos(dicTitulosC170("CHV_REG"))
            
            'Coleta dados do registro C100
            If dicDadosC100.Exists(CHV_REG_C100) Then
                
                CHV_NFE = dicDadosC100(CHV_REG_C100)(dicTitulosC100("CHV_NFE"))
                IND_OPER = dicDadosC100(CHV_REG_C100)(dicTitulosC100("IND_OPER"))
                Select Case Util.ApenasNumeros(dicDadosC100(CHV_REG_C100)(dicTitulosC100("COD_SIT")))
                    
                    Case "00", "01", "06", "07", "08"
                        
                        If VBA.Left(IND_OPER, 1) = "0" Then
                            DT_OPER = dicDadosC100(CHV_REG_C100)(dicTitulosC100("DT_E_S"))
                            
                        ElseIf VBA.Left(IND_OPER, 1) = "1" Then
                            DT_OPER = dicDadosC100(CHV_REG_C100)(dicTitulosC100("DT_DOC"))
                            
                        End If
                        
                    Case Else
                        
                        GoTo Prx:
                        
                End Select
                
            Else
                
                GoTo Prx:
                
            End If
            
            COD_ITEM = Campos(dicTitulosC170("COD_ITEM"))
                
            'Coleta dados do registro 0200
            CHV_REG = Util.UnirCampos(CStr(Campos(dicTitulosC170("ARQUIVO"))), CStr(Campos(dicTitulosC170("COD_ITEM"))))
            If dicDados0200.Exists(CHV_REG) Then
                
                CHV_REG_0200 = dicDados0200(CHV_REG)(dicTitulos0200("CHV_REG"))
                DESCR_ITEM = dicDados0200(CHV_REG)(dicTitulos0200("DESCR_ITEM"))
                COD_BARRA = dicDados0200(CHV_REG)(dicTitulos0200("COD_BARRA"))
                TIPO_ITEM = dicDados0200(CHV_REG)(dicTitulos0200("TIPO_ITEM"))
                UNID_INV = dicDados0200(CHV_REG)(dicTitulos0200("UNID_INV"))
                UNID_INV = ExtrairUnidade0190(ARQUIVO, UNID_INV, dicDados0190, dicTitulos0190)
                
            Else
                
                CHV_REG_0200 = ""
                DESCR_ITEM = "ITEM NÃO ITENTIFICADO"
                COD_BARRA = ""
                TIPO_ITEM = ""
                UNID_INV = ""
                
            End If
                
            IND_MOV = Campos(dicTitulosC170("IND_MOV"))
            CFOP = Campos(dicTitulosC170("CFOP"))
            VL_ITEM = Campos(dicTitulosC170("VL_ITEM"))
            QTD_COM = Campos(dicTitulosC170("QTD"))
            UNID_COM = Campos(dicTitulosC170("UNID"))
            UNID_COM = ExtrairUnidade0190(ARQUIVO, UNID_COM, dicDados0190, dicTitulos0190)
            
            If QTD_COM > 0 Then VL_UNIT_COM = VL_ITEM / QTD_COM
                
                'Coleta dados do registro 0220
                CHV_REG = fnSPED.GerarChaveRegistro(CStr(CHV_REG_0200), CStr(UNID_COM))
                If dicDados0220.Exists(CHV_REG) And CFOP < 4000 Then
                    
                    FAT_CONV = fnExcel.ConverterValores(dicDados0220(CHV_REG)(dicTitulos0220("FAT_CONV")))
                    
                Else
                    
                    FAT_CONV = 0
                    
                End If
                
            If FAT_CONV > 0 Then QTD_INV = QTD_COM * FAT_CONV Else QTD_INV = QTD_COM
            If QTD_INV > 0 Then VL_UNIT_INV = VL_ITEM / QTD_INV
            
        End If
        
        Campos = Array(REG, ARQUIVO, CHV_REG_C100, CHV_REG_C170, CHV_REG_0200, CHV_NFE, IND_OPER, DT_OPER, COD_ITEM, DESCR_ITEM, COD_BARRA, TIPO_ITEM, IND_MOV, CFOP, VL_ITEM, QTD_COM, UNID_COM, VL_UNIT_COM, FAT_CONV, QTD_INV, UNID_INV, VL_UNIT_INV, Empty, Empty)
        Campos = AnalisarInconsistencias(Campos, dicTitulos)
        Call fnExcel.DefinirTipoCampos(Campos, dicTitulos)
        arrRelatorio.Add Campos

Prx:
        FAT_CONV = 0
        
    Next Linha
    
    Call Util.LimparDados(relInteligenteEstoque, 4, False)
    Call Util.ExportarDadosArrayList(relInteligenteEstoque, arrRelatorio)
    
    Call FuncoesFormatacao.AplicarFormatacao(relInteligenteEstoque)
    Call FuncoesFormatacao.DestacarInconsistencias(relInteligenteEstoque)
    
    Call Util.MsgInformativa("Relatório de estoque gerado com sucesso", "Relatório Inteligente de Estoque", Inicio)
    
    Application.StatusBar = False
    
End Sub

Public Function ExtrairUnidade0190(ByVal ARQUIVO As String, ByVal UNID As String, ByRef dicDados0190 As Dictionary, ByRef dicTitulos0190 As Dictionary) As String

Dim CHV_REG As String, DESCR$
    
    CHV_REG = Util.UnirCampos(ARQUIVO, UNID)
    If dicDados0190.Exists(CHV_REG) Then
        
        UNID = dicDados0190(CHV_REG)(dicTitulos0190("UNID"))
        DESCR = dicDados0190(CHV_REG)(dicTitulos0190("DESCR"))
        
        If Util.ApenasNumeros(UNID) <> "" Then ExtrairUnidade0190 = UNID & " - " & DESCR Else ExtrairUnidade0190 = UNID
        
    End If
    
End Function

Public Function AnalisarInconsistencias(ByVal Campos As Variant, ByRef dicTitulos As Dictionary) As Variant

Dim CFOP As String, CST_ICMS$, TIPO_ITEM$, IND_MOV$, UND_COM$, UND_INV$, COD_ITEM$, DESCR_ITEM$, INCONSISTENCIA$, SUGESTAO$
Dim FAT_CONV As Double
Dim i As Integer
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    COD_ITEM = Util.RemoverAspaSimples(Campos(dicTitulos("COD_ITEM") - i))
    DESCR_ITEM = Util.RemoverAspaSimples(Campos(dicTitulos("DESCR_ITEM") - i))
    TIPO_ITEM = Util.ApenasNumeros(Campos(dicTitulos("TIPO_ITEM") - i))
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    IND_MOV = Util.ApenasNumeros(Campos(dicTitulos("IND_MOV") - i))
    UND_COM = fnExcel.FormatarTexto(Campos(dicTitulos("UNID") - i))
    UND_INV = fnExcel.FormatarTexto(Campos(dicTitulos("UND_INV") - i))
    FAT_CONV = fnExcel.ConverterValores(Util.ApenasNumeros(Campos(dicTitulos("FAT_CONV") - i)))
    
    Select Case True
        
        Case COD_ITEM = ""
            INCONSISTENCIA = "Não foi informado um item de correlação (campo COD_ITEM vazio)"
            SUGESTAO = "Informe uma correlação válida o item no relatório de correlações"
        
        Case DESCR_ITEM Like "*ITEM NÃO ITENTIFICADO"
            INCONSISTENCIA = "Item não cadastrado no registro 0200"
            SUGESTAO = "Cadastre o item no registro 0200"
            
        Case COD_ITEM = "SEM CORRELAÇÃO"
            INCONSISTENCIA = "Não foi feita a correlação de itens entre os produtos do XML e do 0200"
            SUGESTAO = "Item sem correlação. Acesse o relatório de Correlações"
                        
        Case UND_INV = ""
            INCONSISTENCIA = "Item sem unidade de medida cadastrada"
            SUGESTAO = "Cadastre uma unidade de medida para o item no registro 0200"
        
        Case UND_COM <> UND_INV And FAT_CONV = 0
            INCONSISTENCIA = "O campo FAT_CON precisa ser informado quando os campos UND_COM e UND_INV forem diferentes"
            SUGESTAO = "informar um valor maior que 0 para o campo FAT_CONV"
            
        Case UND_COM = UND_INV And FAT_CONV > 0
            INCONSISTENCIA = "Os campos UND_COM e UND_INV possuem o mesmo valor, neste caso o campo FAT_CONV deve ser 0"
            SUGESTAO = "Zerar o campo FAT_CONV"
            
        Case IND_MOV = ""
            INCONSISTENCIA = "O Campo de movimentação física do item 'IND_MOV' não foi informado"
            SUGESTAO = "Informe '0' para SIM e '1' para NÃO"
                        
    End Select
    
    'Registra a inconsistência caso exista
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA, SUGESTAO, dicInconsistenciasIgnoradas)
    
    AnalisarInconsistencias = Campos
    
End Function

Public Function AceitarSugestoes()

Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim Campos As Variant
Dim dicDados As New Dictionary
Dim UltimaSugestao As String

    Inicio = Now()
    Application.StatusBar = "Implementando sugestões selecionadas, por favor aguarde..."
    
    Set dicTitulos = Util.MapearTitulos(relInteligenteEstoque, 3)
    Set Dados = relInteligenteEstoque.Range("A4").CurrentRegion
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
    
    If relInteligenteEstoque.AutoFilterMode Then relInteligenteEstoque.AutoFilter.ShowAllData
    Call Util.LimparDados(relInteligenteEstoque, 4, False)
    Call Util.ExportarDadosDicionario(relInteligenteEstoque, dicDados)
    
    Call FuncoesFormatacao.AplicarFormatacao(relInteligenteEstoque)
    Call FuncoesFormatacao.DestacarInconsistencias(relInteligenteEstoque)
    
    Call Util.MsgInformativa("Sugestões aplicadas com sucesso!", "Inclusão de Sugestões", Inicio)
    Application.StatusBar = False
    
End Function

Public Function ReprocessarSugestoes()

Dim Dados As Range, Linha As Range
Dim arrRelatorio As New ArrayList
Dim dicTitulos As New Dictionary
Dim Campos As Variant
Dim FAT_CONV As Double, QTD_COM#, QTD_INV#, VL_ITEM#
    
    Set dicTitulos = Util.MapearTitulos(relInteligenteEstoque, 3)
    If relInteligenteEstoque.AutoFilterMode Then relInteligenteEstoque.AutoFilter.ShowAllData
    
    Set Dados = Util.DefinirIntervalo(relInteligenteEstoque, 4, 3)
    If Dados Is Nothing Then Exit Function
    
    a = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 100, "Reprocessando sugestões, por favor aguarde...", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            VL_ITEM = fnExcel.FormatarValores(Campos(dicTitulos("VL_ITEM")))
            FAT_CONV = fnExcel.FormatarValores(Campos(dicTitulos("FAT_CONV")))
            QTD_COM = fnExcel.FormatarValores(Campos(dicTitulos("QTD")))
            
            If FAT_CONV > 0 Then
                
                QTD_INV = QTD_COM * FAT_CONV
                Campos(dicTitulos("QTD_INV")) = QTD_INV
                If QTD_INV = 0 Then Campos(dicTitulos("VL_UNIT_INV")) = VL_ITEM _
                    Else Campos(dicTitulos("VL_UNIT_INV")) = fnExcel.ConverterValores(VL_ITEM / QTD_INV, True, 2)
                
            Else
                
                Campos(dicTitulos("QTD_INV")) = QTD_COM
                If QTD_COM = 0 Then Campos(dicTitulos("VL_UNIT_INV")) = VL_ITEM _
                    Else Campos(dicTitulos("VL_UNIT_INV")) = fnExcel.ConverterValores(VL_ITEM / QTD_COM, True, 2)
                
            End If
            
            Campos(dicTitulos("INCONSISTENCIA")) = Empty
            Campos(dicTitulos("SUGESTAO")) = Empty
            Campos = AnalisarInconsistencias(Campos, dicTitulos)
            arrRelatorio.Add Campos
            
        End If
        
    Next Linha
    
    Application.StatusBar = ""
    Call Util.LimparDados(relInteligenteEstoque, 4, False)
    Call Util.ExportarDadosArrayList(relInteligenteEstoque, arrRelatorio)
    
    Call FuncoesFormatacao.AplicarFormatacao(relInteligenteEstoque)
    Call FuncoesFormatacao.DestacarInconsistencias(relInteligenteEstoque)
    
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
    
    Set dicTitulos = Util.MapearTitulos(relInteligenteEstoque, 3)
    Set Dados = relInteligenteEstoque.Range("A4").CurrentRegion
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
                Call AnalisarInconsistencias(Campos, dicTitulos)
                
            End If
            
            If Linha.Row > 3 Then dicDados(Campos(dicTitulos("CHV_REG"))) = Campos
            
        End If
        
    Next Linha

    If dicDados.Count = 0 Then
        Call Util.MsgAlerta("Não existem Inconsistêncais a ignorar!", "Ignorar Inconsistências")
        Exit Function
    End If
    
    If relInteligenteEstoque.AutoFilterMode Then relInteligenteEstoque.AutoFilter.ShowAllData
    Call Util.LimparDados(relInteligenteEstoque, 4, False)
    Call Util.ExportarDadosDicionario(relInteligenteEstoque, dicDados)
    
    Call FuncoesFormatacao.AplicarFormatacao(ActiveSheet)
    Call FuncoesFormatacao.DestacarInconsistencias(relInteligenteEstoque)
    
    Call Util.MsgInformativa("Inconsistências ignoradas com sucesso!", "Ignorar Inconsistências", Inicio)
    Application.StatusBar = False
    
End Function

Public Sub AtualizarRegistros()

Dim dicTitulos0200 As New Dictionary
Dim dicTitulos0220 As New Dictionary
Dim dicTitulosC170 As New Dictionary
Dim dicDados0200 As New Dictionary
Dim dicDados0220 As New Dictionary
Dim dicDadosC170 As New Dictionary
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim Campos As Variant, Campos0200, Campos0220, CamposC170, dicCampos, regCampo, nCampo
Dim CHV_C170 As String, CHV_0200$, CHV_0220$, UNID$
Dim i As Byte
    
    Inicio = Now()
    Application.StatusBar = "Preparando dados para atualização do SPED, por favor aguarde..."
    
    Campos0200 = Array("COD_BARRA", "TIPO_ITEM")
    Campos0220 = Array("FAT_CONV", "UNID_CONV")
    CamposC170 = Array("IND_MOV", "CFOP", "VL_ITEM", "QTD", "UNID")
    
    Set dicTitulos0200 = Util.MapearTitulos(reg0200, 3)
    Set dicDados0200 = Util.CriarDicionarioRegistro(reg0200)
    
    Set dicTitulos0220 = Util.MapearTitulos(reg0220, 3)
    Set dicDados0220 = Util.CriarDicionarioRegistro(reg0220)
    
    Set dicTitulosC170 = Util.MapearTitulos(regC170, 3)
    Set dicDadosC170 = Util.CriarDicionarioRegistro(regC170, "CHV_REG")
    
    Set dicTitulos = Util.MapearTitulos(relInteligenteEstoque, 3)
    If relInteligenteEstoque.AutoFilterMode Then relInteligenteEstoque.AutoFilter.ShowAllData
    
    Set Dados = Util.DefinirIntervalo(relInteligenteEstoque, 4, 3)
    If Dados Is Nothing Then Exit Sub
    
    a = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 100, "Preparando dados para atualização do SPED, por favor aguarde...", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            'Atualizar dados do 0200
            CHV_0200 = Campos(dicTitulos("CHV_REG_0200"))
            If dicDados0200.Exists(CHV_0200) Then
                
                dicCampos = dicDados0200(CHV_0200)
                For Each regCampo In Campos0200
                    
                    dicCampos(dicTitulos0200(regCampo)) = Campos(dicTitulos(regCampo))
                    
                Next regCampo
                
                dicDados0200(CHV_0200) = dicCampos
                
            End If
            
            'Atualizar dados do 0220
            UNID = Campos(dicTitulos("UNID"))
            If UNID Like "* - *" Then UNID = VBA.Split(UNID, " - ")(0)
            CHV_0220 = fnSPED.GerarChaveRegistro(CHV_0200, UNID)
            If dicDados0220.Exists(CHV_0220) Then
                
                dicCampos = dicDados0220(CHV_0220)
                If LBound(dicCampos) = 0 Then i = 1 Else i = 0
                
                For Each regCampo In Campos0220
                    
                    If regCampo = "UNID_CONV" Then
                        
                        nCampo = "UNID"
                        UNID = Campos(dicTitulos(nCampo))
                        If UNID Like "* - *" Then Campos(dicTitulos(nCampo)) = VBA.Split(UNID, " - ")(0)
                        
                    Else
                        
                        nCampo = regCampo
                        
                    End If
                    
                    dicCampos(dicTitulos0220(regCampo) - i) = Campos(dicTitulos(nCampo))
                    
                Next regCampo
                
                dicDados0220(CHV_0220) = dicCampos
                
            ElseIf Campos(dicTitulos("FAT_CONV")) > 0 Then
                
                If regCampo = "UNID_CONV" Then
                    
                    nCampo = "UNID"
                    UNID = Campos(dicTitulos(nCampo))
                    If UNID Like "* - *" Then Campos(dicTitulos(nCampo)) = VBA.Split(UNID, " - ")(0)
                    
                Else
                    
                    nCampo = regCampo
                    
                End If
                
                UNID = Campos(dicTitulos("UNID"))
                If UNID Like "* - *" Then Campos(dicTitulos("UNID")) = VBA.Split(UNID, " - ")(0)
                
                dicCampos = Array("'0220", Campos(dicTitulos("ARQUIVO")), CHV_0220, Campos(dicTitulos("CHV_REG_0200")), _
                    "", "'" & Campos(dicTitulos("UNID")), Campos(dicTitulos("FAT_CONV")), "")
                
                dicDados0220(CHV_0220) = dicCampos
                
            End If
            
            'Atualizar dados do C170
            CHV_C170 = Campos(dicTitulos("CHV_REG"))
            If dicDadosC170.Exists(CHV_C170) Then
                
                dicCampos = dicDadosC170(CHV_C170)
                For Each regCampo In CamposC170
                    
                    If regCampo Like "*UNID*" Then
                        
                        UNID = Campos(dicTitulos(regCampo))
                        If UNID Like "* - *" Then Campos(dicTitulos(regCampo)) = VBA.Split(UNID, " - ")(0)
                        dicCampos(dicTitulosC170(regCampo)) = Campos(dicTitulos(regCampo))
                        
                    End If
                    
                    If regCampo Like "VL_*" Or regCampo Like "QTD" Then Campos(dicTitulos(regCampo)) = VBA.Round(Campos(dicTitulos(regCampo)), 2)
                    dicCampos(dicTitulosC170(regCampo)) = Campos(dicTitulos(regCampo))
                
                Next regCampo
                
                dicDadosC170(CHV_C170) = dicCampos
                
            End If
            
        End If

    Next Linha
    
    Application.StatusBar = "Atualizando dados do registro 0200, por favor aguarde..."
    Call Util.LimparDados(reg0200, 4, False)
    Call Util.ExportarDadosDicionario(reg0200, dicDados0200)
    
    Application.StatusBar = "Atualizando dados do registro 0220, por favor aguarde..."
    Call Util.LimparDados(reg0220, 4, False)
    Call Util.ExportarDadosDicionario(reg0220, dicDados0220)
    
    Application.StatusBar = "Atualizando dados do registro C170, por favor aguarde..."
    Call Util.LimparDados(regC170, 4, False)
    Call Util.ExportarDadosDicionario(regC170, dicDadosC170)
    
    DoEvents
    Application.StatusBar = "Atualizando dados do registro C190, por favor aguarde..."
    Call rC170.GerarC190(True)
    
    DoEvents
    Application.StatusBar = "Atualizando dados do registro C100, por favor aguarde..."
    Call rC170.AtualizarImpostosC100(True)
    
    Call FuncoesFormatacao.AplicarFormatacao(relInteligenteEstoque)
    Call FuncoesFormatacao.DestacarInconsistencias(relInteligenteEstoque)
    
    Application.StatusBar = "Atualização concluída com sucesso!"
    Call Util.MsgInformativa("Registros atualizados com sucesso!", "Atualização de dados", Inicio)
    Application.StatusBar = False
    
End Sub
