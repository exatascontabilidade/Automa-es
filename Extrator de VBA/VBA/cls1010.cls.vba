Attribute VB_Name = "cls1010"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function GerarRegistro(ByRef dicDados As Dictionary, ByVal ARQUIVO As String)

Dim Campos As Variant
Dim r1100 As String, r1200$, r1250$, r1300$, r1390$, r1400$, r1500$, r1600$, r1700$, r1800$, r1960$, r1970$, r1980$
    
Dim DataRef As String

    With regEFD
        
        If .dic1100.Count > 0 Then r1100 = "S" Else r1100 = "N"
        If .dic1200.Count > 0 Then r1200 = "S" Else r1200 = "N"
        If .dic1250.Count > 0 Then r1250 = "S" Else r1250 = "N"
        If .dic1300.Count > 0 Then r1300 = "S" Else r1300 = "N"
        If .dic1390.Count > 0 Then r1390 = "S" Else r1390 = "N"
        If .dic1400.Count > 0 Then r1400 = "S" Else r1400 = "N"
        If .dic1500.Count > 0 Then r1500 = "S" Else r1500 = "N"
        If .dic1600.Count > 0 Then r1600 = "S" Else r1600 = "N"
        If .dic1601.Count > 0 Then r1600 = "S" Else r1600 = "N"
        If .dic1700.Count > 0 Then r1700 = "S" Else r1700 = "N"
        If .dic1800.Count > 0 Then r1800 = "S" Else r1800 = "N"
        If .dic1960.Count > 0 Then r1960 = "S" Else r1960 = "N"
        If .dic1970.Count > 0 Then r1970 = "S" Else r1970 = "N"
        If .dic1980.Count > 0 Then r1980 = "S" Else r1980 = "N"
        
        DataRef = VBA.Right(VBA.Left(ARQUIVO, 7), 4) & "-" & VBA.Left(ARQUIVO, 2) & "-" & "01"
        If CDate(DataRef) >= CDate("2020-01-01") Then
            Campos = Array("", "1010", r1100, r1200, r1300, r1390, r1400, r1500, r1600, r1700, r1800, r1960, r1970, r1980, r1250, "")
        
        ElseIf CDate(DataRef) >= CDate("2019-01-01") Then
            Campos = Array("", "1010", r1100, r1200, r1300, r1390, r1400, r1500, r1600, r1700, r1800, r1960, r1970, r1980, "")
                        
        ElseIf CDate(DataRef) >= CDate("2012-07-01") Then
            Campos = Array("", "1010", r1100, r1200, r1300, r1390, r1400, r1500, r1600, r1700, r1800, "")
            
        End If
        
        If CDate(DataRef) >= CDate("2012-07-01") Then dicDados(ARQUIVO) = fnSPED.GerarRegistro(Campos)
        
    End With

End Function

