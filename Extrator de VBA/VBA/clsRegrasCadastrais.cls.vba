Attribute VB_Name = "clsRegrasCadastrais"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public ICMS As New clsRegrasCadastraisICMS
Public IPI As New clsRegrasCadastraisIPI
Public PIS_COFINS As New clsRegrasCadastraisPISCOFINS
Private ValidacoesCFOP As New clsRegrasFiscaisCFOP

Public Function VerificarCFOP(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim CFOP As String, TIPO_ITEM$
Dim i As Integer
    
    If LBound(Campos) = 0 Then i = 1
    
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    TIPO_ITEM = Util.ApenasNumeros(Campos(dicTitulos("TIPO_ITEM") - i))

    Select Case True
        
        Case CStr(CFOP) = "" And TIPO_ITEM <> "09"
            Call Util.GravarSugestao(Campos, dicTitulos, _
                INCONSISTENCIA:="O campo CFOP não foi informado e o campo TIPO_ITEM é diferente de: 09 - Serviços", _
                SUGESTAO:="informar um valor válido para o campo CFOP ou mudar o campos TIPO_ITEM para 09 - Serviços")
                
        Case Else
            If Not ValidarCFOP(CFOP) And CFOP <> "" Then
                Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="O CFOP informado não existe na tabela CFOP", _
                    SUGESTAO:="Informar um valor válido para o campo CFOP")
            End If
            
    End Select

End Function

Private Function ValidarCFOP(ByVal CFOP As String) As Boolean

    If dicTabelaCFOP.Count = 0 Then Call ValidacoesCFOP.CarregarTabelaCFOP
    
    If CFOP <> "" Then
        If dicTabelaCFOP.Exists(CFOP) Then ValidarCFOP = True
    End If
    
End Function

Public Function ValidarCST_ICMS(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
    
Dim CST_ICMS$, CFOP$, INCONSISTENCIA$, SUGESTAO$
Dim fatCFOP As Boolean, devCompCFOP As Boolean
Dim arrCST_ICMS As New ArrayList
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    CST_ICMS = Util.ApenasNumeros(Campos(dicTitulos("CST_ICMS") - i))
        
    'carrega lista de CSTs válidos
    Call ListarCST_ICMS(arrCST_ICMS)
        
    Select Case True
        
        Case CST_ICMS = ""
            INCONSISTENCIA = "CST_ICMS não foi informado"
            SUGESTAO = "Informar um valor válido para o campo CST_ICMS"
            
        Case CST_ICMS <> "" And Not arrCST_ICMS.contains(CST_ICMS)
            INCONSISTENCIA = "O CST_ICMS informado está inválido"
            SUGESTAO = "Informar um CST_ICMS válido"
            
    End Select

    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO)
    
End Function

Private Function ListarCST_ICMS(ByRef arrCST_ICMS As ArrayList)

Dim CST As Variant
Dim CST_ICMS As Variant
Dim ListaCST_ICMS As Variant
    
    CST_ICMS = "000,002,010,015,020,030,040,041,050,051,053,060,061,070,090,"
    CST_ICMS = CST_ICMS & "100,102,110,115,120,130,140,141,150,151,153,160,161,170,190,"
    CST_ICMS = CST_ICMS & "200,202,210,215,220,230,240,241,250,251,253,260,261,270,290,"
    CST_ICMS = CST_ICMS & "300,302,310,315,320,330,340,341,350,351,353,360,361,370,390,"
    CST_ICMS = CST_ICMS & "400,402,410,415,420,430,440,441,450,451,453,460,461,470,490,"
    CST_ICMS = CST_ICMS & "500,502,510,515,520,530,540,541,550,551,553,560,561,570,590,"
    CST_ICMS = CST_ICMS & "600,602,610,615,620,630,640,641,650,651,653,660,661,670,690,"
    CST_ICMS = CST_ICMS & "700,702,710,715,720,730,740,741,750,751,753,760,761,770,790,"
    CST_ICMS = CST_ICMS & "800,802,810,815,820,830,840,841,850,851,853,860,861,870,890"
    
    ListaCST_ICMS = VBA.Split(CST_ICMS, ",")
    For Each CST In ListaCST_ICMS
        
        If Not arrCST_ICMS.contains(CST) Then arrCST_ICMS.Add CST
        
    Next CST
    
End Function

Public Function ValidarCST_IPI(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
    
Dim CST_IPI$, CFOP$, INCONSISTENCIA$, SUGESTAO$
Dim fatCFOP As Boolean, devCompCFOP As Boolean
Dim arrCST_IPI As New ArrayList
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    CST_IPI = Util.ApenasNumeros(Campos(dicTitulos("CST_IPI") - i))
        
    'carrega lista de CSTs válidos
    Call ListarCST_IPI(arrCST_IPI)
        
    Select Case True
        
        Case CST_IPI = ""
            INCONSISTENCIA = "CST_IPI não foi informado"
            SUGESTAO = "Informar um valor válido para o campo CST_IPI"
            
        Case CST_IPI <> "" And Not arrCST_IPI.contains(CST_IPI)
            INCONSISTENCIA = "O CST_IPI informado está inválido"
            SUGESTAO = "Informar um CST_IPI válido"
            
    End Select

    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO)
    
End Function

Private Function ListarCST_IPI(ByRef arrCST_IPI As ArrayList)

Dim CST As Variant
Dim CST_IPI As Variant
Dim ListaCST_IPI As Variant
    
    CST_IPI = "00,01,02,03,04,05,49,50,51,52,53,54,55,99"
    
    ListaCST_IPI = VBA.Split(CST_IPI, ",")
    For Each CST In ListaCST_IPI
        
        If Not arrCST_IPI.contains(CST) Then arrCST_IPI.Add CST
        
    Next CST
    
End Function

Public Function ValidarCFOP_IPI(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim CFOP$, CST_IPI$, INCONSISTENCIA$, SUGESTAO$
Dim fatCFOP As Boolean, devCompCFOP As Boolean
Dim arrCFOP_IPI As New ArrayList
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    CST_IPI = Util.ApenasNumeros(Campos(dicTitulos("CST_IPI") - i))
    
    'carrega lista de CFOP válidos para o IPI
    Call ListarCFOP_IPI(arrCFOP_IPI)
    
    Select Case True
        
        Case CST_IPI = "" And arrCFOP_IPI.contains(CFOP)
            INCONSISTENCIA = "O CFOP informado não possui um CST_IPI associado"
            SUGESTAO = "Informar um CST_IPI para a operação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO)
    
End Function

Private Function ListarCFOP_IPI(ByRef arrCFOP_IPI As ArrayList)

Dim CST As Variant
Dim CFOP_IPI As Variant
Dim ListaCFOP_IPI As Variant
    
    CFOP_IPI = "1101,1111,1116,1120,1122,1124,1125,1135,1151,1208,1209,1212,1252,1401,1651,1901,1902,1903,1924,1925,"
    CFOP_IPI = CFOP_IPI & "2101,2111,2116,2120,2122,2124,2125,2135,2151,2208,2209,2212,2252,2401,2651,2901,2902,2903,2924,2925,"
    CFOP_IPI = CFOP_IPI & "3101,3127,3129,3202,3211,3212,3651,"
    CFOP_IPI = CFOP_IPI & "5101,5103,5105,5116,5124,5125,"
    CFOP_IPI = CFOP_IPI & "6101,6103,6105,6116,6124,6125,"
    CFOP_IPI = CFOP_IPI & "7101,7105,7127"
    
    ListaCFOP_IPI = VBA.Split(CFOP_IPI, ",")
    For Each CST In ListaCFOP_IPI
        
        If Not arrCFOP_IPI.contains(CST) Then arrCFOP_IPI.Add CST
        
    Next CST
    
End Function
