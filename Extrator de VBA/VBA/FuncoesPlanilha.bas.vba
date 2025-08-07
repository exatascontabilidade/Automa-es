Attribute VB_Name = "FuncoesPlanilha"
Option Explicit

'Sub ProtegerPlanilha(ByRef Planilha As Worksheet)
'
'    Call Planilha.Protect(Password:="C1664B12A18ACF241F5403555175CBC7", _
'                          DrawingObjects:=True, _
'                          Contents:=True, _
'                          Scenarios:=True, _
'                          UserInterfaceOnly:=True, _
'                          AllowFormattingCells:=False, _
'                          AllowFormattingColumns:=True, _
'                          AllowFormattingRows:=True, _
'                          AllowInsertingColumns:=True, _
'                          AllowInsertingRows:=False, _
'                          AllowInsertingHyperlinks:=True, _
'                          AllowDeletingColumns:=False, _
'                          AllowDeletingRows:=True, _
'                          AllowSorting:=True, _
'                          AllowFiltering:=True, _
'                          AllowUsingPivotTables:=True)
'
'End Sub
'
'Sub DesprotegerPlanilha(ByRef Planilha As Worksheet)
'    Planilha.Unprotect "C1664B12A18ACF241F5403555175CBC7"
'End Sub

Public Function RemoverDuplicatas(ByRef Plan As Worksheet, ParamArray Titulos() As Variant)

Dim dicTitulos As New Dictionary
Dim arrTitulos As New ArrayList
Dim Intervalo As Range
Dim Titulo As Variant
    
    Set dicTitulos = Util.MapearTitulos(Plan, 3)
    Set Intervalo = Util.DefinirIntervalo(Plan, 3, 3)
    
    For Each Titulo In Titulos
        If dicTitulos.Exists(Titulo) Then arrTitulos.Add dicTitulos(Titulo)
    Next Titulo
    
    Intervalo.RemoveDuplicates Columns:=arrTitulos.toArray, Header:=xlYes
    
    MsgBox "Duplicidades removidas com sucesso!", vbInformation, "Remoção de registros duplicados"
    
End Function
