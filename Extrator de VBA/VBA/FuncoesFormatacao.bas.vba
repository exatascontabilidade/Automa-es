Attribute VB_Name = "FuncoesFormatacao"
Option Explicit

Sub DeletarFormatacao()
    On Error Resume Next
    ActiveSheet.Cells.FormatConditions.Delete
End Sub

Sub AplicarFormatacao(ByRef Plan As Worksheet)
'TODO: Criar regra para identificar erro e alerta e colorí-los com cores diferentes. Exemplo: ERRO: Vermelho, Alerta: Amarelo.
Dim rng As Range
    
    On Error Resume Next
    Call DeletarFormatacao
    
    'Defina a planilha e a faixa de células onde deseja aplicar a formatação condicional
    Set rng = Util.DefinirIntervalo(Plan, 4, 3)
    If Not rng Is Nothing Then
        
        'Aplica a formatação condicional
        With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=E($A4<>"""";LIN()>2;MOD(LIN();2)=0)")
            .Interior.Color = RGB(215, 215, 215)
            .SetLastPriority
        End With
        
    End If
    
End Sub

Sub FormatarInconsistencias(ByRef Plan As Worksheet)

Dim Formulas(0 To 10) As Variant
Dim formula As Variant
Dim rng As Range
Dim Cor As Long
    
    Formulas(0) = "=E($A4<>"""";ESQUERDA($D4;1)*1<4;DIREITA($D4;3)=""403"";$I4<>0)" 'Identifica CFOPs com fim 403 que possuem crédito de ICMS
    Formulas(1) = "=E($A4<>"""";ESQUERDA($D4;1)*1<4;DIREITA($D4;3)=""405"";$I4<>0)" 'Identifica CFOPs com fim 405 que possuem crédito de ICMS
    Formulas(2) = "=E($A4<>"""";ESQUERDA($D4;1)*1<4;DIREITA($D4;3)=""551"";$I4<>0)" 'Identifica CFOPs com fim 551 que possuem crédito de ICMS
    Formulas(3) = "=E($A4<>"""";ESQUERDA($D4;1)*1<4;DIREITA($D4;3)=""556"";$I4<>0)" 'Identifica CFOPs com fim 556 que possuem crédito de ICMS
    Formulas(4) = "=E($A4<>"""";ESQUERDA($D4;1)*1>4;DIREITA($D4;3)=""929"";$I4<>0)" 'Identifica CFOPs com fim 929 que possuem crédito de ICMS
    Formulas(5) = "=E($A4<>"""";ESQUERDA($D4;1)*1>4;DIREITA($D4;3)=""404"";$M4=0)" 'Identifica CFOPs da ST que possuem não possuem destaque da ST
    
    'Defina a planilha e a faixa de células onde deseja aplicar a formatação condicional
    Set rng = Util.DefinirIntervalo(Plan, 4, 3)
    If Not rng Is Nothing Then
    
        For Each formula In Formulas
        
            If formula <> "" Then
            
                With rng.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
                
                    .Font.Bold = True
                    .Font.Color = RGB(255, 255, 255)
                    .Interior.Color = RGB(255, 0, 0) ' Cor vermelha
                    .SetFirstPriority
                    
                End With
            
            End If
            
        Next formula
    
    End If
    
    Call DefinirRegrasFiscais(Plan)
    
End Sub

Private Sub DefinirRegrasFiscais(ByRef Plan As Worksheet)

Dim dicTitulos As New Dictionary
Dim UltLin As Long
Dim Lin As Variant
Dim rng As Range
    
    Set rng = Util.DefinirIntervalo(Plan, 4, 3)
    If Not rng Is Nothing Then
    
        Set dicTitulos = Util.MapearTitulos(Plan, 3)
        UltLin = Util.UltimaLinha(Plan, "A")
        
        'Lançar observações das divergências no livro fiscal
        For Each Lin In rng.Rows
            
            If Lin.Row = UltLin Then Exit Sub
            
            With rng
                                                                                                                                            
                'Regras fiscais para operações de entrada
                If VBA.Left(.Cells(Lin.Row, dicTitulos("CFOP")), 1) < 4 Then
                
                    Select Case True
                        
                        'Se o CFOP for de uma operação sujeita ao ST e o campo VL_ICMS for diferente de zero
                        Case (VBA.Right(.Cells(Lin.Row, dicTitulos("CFOP")), 3) = "403" Or VBA.Right(.Cells(Lin.Row, dicTitulos("CFOP")), 3) = "405") _
                              And (.Cells(Lin.Row, dicTitulos("VL_ICMS")) <> 0)
                            rng.Cells(Lin.Row, dicTitulos("OBSERVACOES")) = "Produto sujeito a ST com valor de ICMS"
                    
                        'Se o CFOP for de uma operação de uso e consumo e o campo VL_ICMS for diferente de zero
                        Case (VBA.Right(.Cells(Lin.Row, dicTitulos("CFOP")), 3) = "551") And (.Cells(Lin.Row, dicTitulos("VL_ICMS")) <> 0)
                            rng.Cells(Lin.Row, dicTitulos("OBSERVACOES")) = "Aproveitamento de ICMS em operação de ativo imobilizado"
                                        
                        'Se o CFOP for de uma operação de uso e consumo e o campo VL_ICMS for diferente de zero
                        Case (VBA.Right(.Cells(Lin.Row, dicTitulos("CFOP")), 3) = "556") And (.Cells(Lin.Row, dicTitulos("VL_ICMS")) <> 0)
                            rng.Cells(Lin.Row, dicTitulos("OBSERVACOES")) = "Aproveitamento de ICMS em operação de uso e consumo"
                            
                    End Select
                
                ElseIf VBA.Left(.Cells(Lin.Row, dicTitulos("CFOP")), 1) > 4 Then
                
                    Select Case True
                        
                        'Se o CFOP for de uma operação de emissão de NFe em decorrência de emissão de cupom fiscal e o campo VL_ICMS for diferente de zero
                        Case (VBA.Right(.Cells(Lin.Row, dicTitulos("CFOP")), 3) = "929") And (.Cells(Lin.Row, dicTitulos("VL_ICMS")) <> 0)
                            rng.Cells(Lin.Row, dicTitulos("OBSERVACOES")) = "Débito de ICMS em operação em decorrência de cupom fiscal"
                            
                        'Se o CFOP for de uma operação com ST e o campo VL_ICMS_ST for diferente igual a zero
                        Case (VBA.Right(.Cells(Lin.Row, dicTitulos("CFOP")), 3) = "401" Or VBA.Right(.Cells(Lin.Row, dicTitulos("CFOP")), 3) = "402" _
                            Or VBA.Right(.Cells(Lin.Row, dicTitulos("CFOP")), 3) = "403" Or VBA.Right(.Cells(Lin.Row, dicTitulos("CFOP")), 3) = "404") _
                            And (.Cells(Lin.Row, dicTitulos("VL_ICMS_ST")) = 0)
                            rng.Cells(Lin.Row, dicTitulos("OBSERVACOES")) = "Operação sujeita a substituição sem destaque do ST"
                            
                    End Select
                
                End If
                
            End With
            
        Next Lin
    
    End If
    
End Sub

Sub FormatarDivergencias(ByRef Plan As Worksheet)

Dim Formulas(0 To 10) As Variant
Dim formula As Variant
Dim Cor As Long, Lin&
Dim CorFonte As Long
Dim rng As Range
    
    Call DeletarFormatacao
    
    Formulas(0) = "=E($A4<>"""";$AA4=""DIVERGÊNCIA"")" 'Identifica notas fiscais com divergências
    Formulas(1) = "=E($A4<>"""";$AA4=""OK"")" 'Identifica notas fiscais já analisadas
    
    formula = "=E($A4<>"""";$B2>0)" 'Destaca quantidade de notas não importadas
    Cor = RGB(255, 0, 0)
    CorFonte = RGB(255, 255, 255)
    
    Set rng = Plan.Range("A2:B2")
    If Not rng Is Nothing Then
    
        With rng.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
            
            .Font.Bold = True
            .Font.Color = CorFonte
            .Interior.Color = Cor
            .SetFirstPriority
            
        End With
        
    End If
    
    Set rng = Util.DefinirIntervalo(Plan, 4, 3)
    If Not rng Is Nothing Then
    
        For Each formula In Formulas
            
            'Definindo cor interior
            If VBA.InStr(1, formula, "DIVERGÊNCIA") > 0 Then Cor = RGB(255, 0, 0)
            If VBA.InStr(1, formula, "OK") > 0 Then Cor = RGB(4, 128, 13)
            
            'Definindo cor da fonte
            If VBA.InStr(1, formula, "DIVERGÊNCIA") > 0 Then CorFonte = RGB(255, 255, 255)
            If VBA.InStr(1, formula, "OK") > 0 Then CorFonte = RGB(255, 255, 255)
            
            
            If formula <> "" Then
            
                With rng.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
                
                    .Font.Bold = True
                    .Font.Color = CorFonte
                    .Interior.Color = Cor
                    .SetFirstPriority
                    
                End With
            
            End If
            
        Next formula
    
    End If
    
End Sub

Function IdentificarEnderecoCampo(ByRef Plan As Worksheet, ByVal nLin As Long, Campo As String) As String
    
Dim rng As Range
Dim Cel As Range

    'Definir o intervalo na linha especificada
    Set rng = Plan.Rows(nLin)

    'Procura o valor na linha especificada
    Set Cel = rng.Find(What:=Campo, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext)

    'Verificar se o valor foi encontrado
    If Not Cel Is Nothing Then IdentificarEnderecoCampo = VBA.Replace(Cel.Address, "$", "")
    
End Function

Sub DestacarInconsistencias(ByRef Plan As Worksheet)

Dim Formulas(0 To 10) As Variant
Dim formula As Variant
Dim rng As Range
Dim Cor As Long
Dim CampoInconsistencia As String
    
    'Identifica o endereço da célula INCONSISTÊNCIA
    CampoInconsistencia = IdentificarEnderecoCampo(Plan, 3, "INCONSISTENCIA")
    CampoInconsistencia = VBA.Left(CampoInconsistencia, VBA.Len(CampoInconsistencia) - 1) & 4
    
    Formulas(0) = "=E($A4<>"""";$" & CampoInconsistencia & "<>"""")" 'Se a coluna inconsistências for diferente de vazio
    
    'Defina a planilha e a faixa de células onde deseja aplicar a formatação condicional
    Set rng = Util.DefinirIntervalo(Plan, 4, 3)
    If Not rng Is Nothing Then
    
        For Each formula In Formulas
        
            If formula <> "" Then
            
                With rng.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
                
                    .Font.Bold = True
                    .Font.Color = RGB(0, 0, 0) 'Fonte Preta
                    .Interior.Color = RGB(255, 255, 102) ' Amarela
                    .SetFirstPriority
                    
                End With
            
            End If
            
        Next formula
    
    End If

End Sub

Sub DestacarMelhorCorrelacao(ByRef Plan As Worksheet)

Dim Formulas(0 To 10) As Variant
Dim formula As Variant
Dim rng As Range
Dim Cor As Long
    
    Formulas(0) = "=E($A4<>"""";$AH4=""SIM"")" 'Se a coluna inconsistências for diferente de vazio
    
    'Defina a planilha e a faixa de células onde deseja aplicar a formatação condicional
    Set rng = Util.DefinirIntervalo(Plan, 4, 3)
    If Not rng Is Nothing Then
    
        For Each formula In Formulas
        
            If formula <> "" Then
            
                With rng.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
                
                    .Font.Bold = True
                    .Font.Color = RGB(0, 0, 0) 'Fonte Preta
                    .Interior.Color = 65535 ' Amarela
                    .SetFirstPriority
                    
                End With
            
            End If
            
        Next formula
    
    End If

End Sub
