Attribute VB_Name = "FuncoesFiltragem"
Option Explicit

Public dicFiltrosSalvos As New Dictionary

Public Sub FiltrarEntradas()

Dim dicTitulos As New Dictionary
Dim Intervalo As Range
Dim UltLin As Long
    
    UltLin = Util.UltimaLinha(relICMS, "A")
    If UltLin > 3 Then
    
        If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilter.ShowAllData
        Set dicTitulos = Util.MapearTitulos(ActiveSheet, 3)
        Set Intervalo = Util.DefinirIntervalo(ActiveSheet, 3, 3)
        'Intervalo = ActiveSheet.Range("A3:" & Util.ConverterNumeroColuna(ActiveSheet.Range("A3").END(xlToRight).Column) & "3")
            
        Intervalo.AutoFilter Field:=CInt(dicTitulos("CFOP")), Criteria1:="<4000"
        
        With ActiveSheet.AutoFilter.Sort
            .SortFields.Clear
            .SortFields.Add2 Key:=Range("C3:C" & Rows.Count)
            .Apply
        End With
            
        Application.GoTo [A3]
    
    End If
    
End Sub

Public Sub FiltrarSaidas()

Dim dicTitulos As New Dictionary
Dim Intervalo As Range
Dim UltLin As Long
    
    UltLin = Util.UltimaLinha(relICMS, "A")
    If UltLin > 3 Then
        
        If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilter.ShowAllData
        Set dicTitulos = Util.MapearTitulos(ActiveSheet, 3)
        Set Intervalo = Util.DefinirIntervalo(ActiveSheet, 3, 3)
        'Intervalo = ActiveSheet.Range("A3:" & Util.ConverterNumeroColuna(ActiveSheet.Range("A3").END(xlToRight).Column) & "3")
            
        Intervalo.AutoFilter Field:=CInt(dicTitulos("CFOP")), Criteria1:=">4000"
        
        With ActiveSheet.AutoFilter.Sort
            .SortFields.Clear
            .SortFields.Add2 Key:=Range("C3:C" & Rows.Count)
            .Apply
        End With
        
        Application.GoTo [A3]
    
    End If
    
End Sub

Public Function FiltrarInconsistencias(ByRef Plan As Worksheet)
    
Dim dicTitulos As New Dictionary
Dim UltLin As Long, UltCol&
    
    UltLin = Util.UltimaLinha(Plan, "A")
    UltCol = Plan.Cells(3, Columns.Count).END(xlToLeft).Column
    
    Set dicTitulos = Util.MapearTitulos(Plan, 3)
    If Plan.AutoFilterMode Then Plan.AutoFilterMode = False
    
    With Plan.Cells(3, 1).Resize(UltLin - 3 + 1, UltCol)
        .AutoFilter Field:=dicTitulos("INCONSISTENCIA"), Criteria1:="<>"
    End With
    
End Function

Public Function AcessarNotaSelecionada()

Dim arrDocumentos As New ArrayList
Dim dicTitulos As New Dictionary
Dim Dados As Variant, Titulos
Dim Intervalo As Range
Dim i As Long, UltLin&
Dim CHV_NFE As String
        
        Set dicTitulos = Util.IndexarDados(Util.DefinirTitulos(relInteligenteDivergencias, 3))
        CHV_NFE = relInteligenteDivergencias.Cells(ActiveCell.Row, dicTitulos("CHV_NFE")).value
        
        If VBA.Len(CHV_NFE) <> 44 Then Call Util.MsgAlerta("Chave de acesso inválida!", "Nenhuma nota selecionada"): Exit Function
        If regC100.AutoFilterMode Then regC100.AutoFilter.ShowAllData
        Set Intervalo = Util.DefinirIntervalo(regC100, 3, 3)
        Set dicTitulos = Util.IndexarDados(Util.DefinirTitulos(regC100, 3))
        
        regC100.Activate
        Intervalo.AutoFilter Field:=CInt(dicTitulos("CHV_NFE")), Criteria1:=CHV_NFE
        Call Application.GoTo(regC100.Range("A3"))
        
End Function

Public Function AcessarEnfoqueDeclarante()

Dim Plan As Worksheet
Dim arrDocumentos As New ArrayList
Dim dicTitulos As New Dictionary
Dim Dados As Variant, Titulos
Dim Intervalo As Range
Dim i As Long, UltLin&
Dim REG As String, CFOP$, CST$, ALIQ$
        
        Set dicTitulos = Util.IndexarDados(Util.DefinirTitulos(relICMS, 3))
        REG = relICMS.Cells(ActiveCell.Row, dicTitulos("REG")).value
        CFOP = relICMS.Cells(ActiveCell.Row, dicTitulos("CFOP")).value
        CST = relICMS.Cells(ActiveCell.Row, dicTitulos("CST_ICMS")).value
        ALIQ = relICMS.Cells(ActiveCell.Row, dicTitulos("ALIQ_ICMS")).text
        
        If CFOP = "" Or CFOP = "CFOP" Then Call Util.MsgAlerta("Seleção inválida!", "Nenhum registro selecionado"): Exit Function
        Set Plan = Worksheets(REG)
        Set Intervalo = Util.DefinirIntervalo(Plan, 3, 3)
        Set dicTitulos = Util.IndexarDados(Util.DefinirTitulos(Plan, 3))
        
        Plan.Activate
        With Intervalo
        
            .AutoFilter Field:=CInt(dicTitulos("CFOP")), Criteria1:=CFOP
            .AutoFilter Field:=dicTitulos("CST_ICMS"), Criteria1:=CST
            .AutoFilter Field:=dicTitulos("ALIQ_ICMS"), Criteria1:=ALIQ
        
        End With
        
        Call Application.GoTo(Plan.Range("A3"))
        
End Function

Public Sub ListarNotasSemLancar()
    
    Select Case ActiveSheet.CodeName
        
        Case "EntNFe", "SaiNFe", "EntCTe", "SaiCTe", "SaiNFCe", "SaiCFe"
            ActiveSheet.Range("A3:L" & Rows.Count).AutoFilter Field:=10, Criteria1:=""
    
    End Select
    
End Sub

Public Sub ListarDivergencias()

Dim dicTitulos As New Dictionary
Dim Campo As String, Valor$
Dim Intervalo As Range
Dim Plan As Worksheet
Dim UltLin As Long
    
    UltLin = Util.UltimaLinha(ActiveSheet, "A")
    If UltLin > 3 Then
        
        Set Plan = ActiveSheet
        With Plan
        
            Set dicTitulos = Util.MapearTitulos(Worksheets(.name), 3)
            Set Intervalo = Util.DefinirIntervalo(Plan, 3, 3)
            
            If .name = "Divergências Fiscais" Then
                Campo = "STATUS_ANALISE"
                Valor = "DIVERGÊNCIA"
            
            ElseIf .name = "Divergências Fiscais" Then
                Campo = "OBSERVACOES"
                Valor = "<>"
                
            End If
            
            On Error Resume Next
            Intervalo.AutoFilter Field:=dicTitulos(Campo), Criteria1:=RGB(255, 0, 0), Operator:=xlFilterCellColor
        
        End With
        
    End If
    
End Sub

Public Function CriarFiltro(ByRef Plan As Worksheet)

Dim ultimaColuna As Long
Dim UltimaLinha As Long
Dim Intervalo As Range
    
    If Plan.AutoFilterMode Then Plan.AutoFilterMode = False
    
    ' Encontra a última coluna com dados na linha 3
    ultimaColuna = Plan.Cells(3, Plan.Columns.Count).END(xlToLeft).Column
    
    ' Define o intervalo para aplicar o autofiltro
    Set Intervalo = Plan.Range(Plan.Cells(3, 1), Plan.Cells(Rows.Count, ultimaColuna))
    
    ' Aplica o autofiltro ao intervalo
    Intervalo.AutoFilter
    
End Function

