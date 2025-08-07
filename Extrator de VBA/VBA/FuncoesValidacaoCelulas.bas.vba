Attribute VB_Name = "FuncoesValidacaoCelulas"
Option Explicit

Sub ValidarCelulaRegimeICMS()

Dim RegimeICMS As Range
    
    Set RegimeICMS = CadContrib.Range("RegimeICMS")
    
    If RegimeICMS Is Nothing Then
        MsgBox "A célula nomeada 'RegimeICMS' não foi encontrada.", vbExclamation
        Exit Sub
    End If

    With RegimeICMS.Validation
    
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="CONTA CORRENTE,SIMPLES NACIONAL"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Regime de ICMS"
        .ErrorTitle = "Opção Inválida"
        .InputMessage = "Escolha uma opção da lista ou deixe em branco."
        .ErrorMessage = "Por favor, selecione 'CONTA CORRENTE', 'SIMPLES NACIONAL' ou deixe o campo vazio."
        .ShowInput = True
        .ShowError = True
        
    End With

    MsgBox "Validação aplicada à célula 'RegimeICMS'.", vbInformation

End Sub

Sub ValidarCelulaRegimeIPI()

Dim RegimeIPI As Range
    
    Set RegimeIPI = CadContrib.Range("RegimeIPI")
    
    If RegimeIPI Is Nothing Then
        MsgBox "A célula nomeada 'RegimeIPI' não foi encontrada.", vbExclamation
        Exit Sub
    End If

    With RegimeIPI.Validation
    
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="CONTRIBUINTE,NÃO CONTRIBUINTE"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Regime de IPI"
        .ErrorTitle = "Opção Inválida"
        .InputMessage = "Escolha uma opção da lista ou deixe em branco."
        .ErrorMessage = "Por favor, selecione 'CONTRIBUINTE', 'NÃO CONTRIBUINTE' ou deixe o campo vazio."
        .ShowInput = True
        .ShowError = True
        
    End With

    MsgBox "Validação aplicada à célula 'RegimeIPI'.", vbInformation

End Sub

Sub ValidarCelulaRegimePISCOFINS()

Dim RegimePISCOFINS As Range
    
    Set RegimePISCOFINS = CadContrib.Range("RegimePISCOFINS")
    
    If RegimePISCOFINS Is Nothing Then
        MsgBox "A célula nomeada 'RegimePISCOFINS' não foi encontrada.", vbExclamation
        Exit Sub
    End If

    With RegimePISCOFINS.Validation
    
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="CUMULATIVO,NÃO CUMULATIVO"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Regime de PISCOFINS"
        .ErrorTitle = "Opção Inválida"
        .InputMessage = "Escolha uma opção da lista ou deixe em branco."
        .ErrorMessage = "Por favor, selecione 'CUMULATIVO', 'NÃO CUMULATIVO' ou deixe o campo vazio."
        .ShowInput = True
        .ShowError = True
        
    End With

    MsgBox "Validação aplicada à célula 'RegimePISCOFINS'.", vbInformation

End Sub

Sub ValidarDataCorteInventario()

Dim DataInventario As Range
    
    Set DataInventario = relSaldoInventario.Range("DT_INVENTARIO")
    
    With DataInventario.Validation
        
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="01/01/1900", Formula2:="12/31/2100"
        .IgnoreBlank = True
        .InputTitle = "Data de Corte do Inventário"
        .InputMessage = "Informe a data final (DD/MM/AAAA) para o relatório de saldo de inventário. O saldo será apurado até este dia."
        .ShowInput = True
        .ErrorTitle = "Data de Corte Inválida"
        .ErrorMessage = "A data informada para o corte do inventário está inválida. Utilize o formato DD/MM/AAAA ou deixe o campo vazio."
        .ShowError = True

    End With

    MsgBox "Validação de data de corte aplicada à célula " & DataInventario.Address(False, False) & ".", vbInformation

End Sub


