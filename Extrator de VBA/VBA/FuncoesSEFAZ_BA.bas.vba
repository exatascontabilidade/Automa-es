Attribute VB_Name = "FuncoesSEFAZ_BA"
Option Explicit

Public Function IncluirAjustesSefazBA()

Dim TXTS, XMLS
Dim Msg As String, Caminho$
Dim arrXMLs As New ArrayList
Dim dicFornecSN As New Dictionary
Dim arrProdutosExcluidos As New ArrayList
    
    TXTS = Util.SelecionarArquivos("txt", "Selecione os arquivos da EFD")
    Caminho = Util.SelecionarPasta("Selecione a pasta onde estão os XML")
    Call ImportarListaProdutosExcluidos(arrProdutosExcluidos)
    
    If (Caminho <> "") And (VarType(TXTS) <> 11) Then
        
        Call Util.ListarArquivos(arrXMLs, Caminho)
        Call fnXML.ColetarFornecedoresSN(dicFornecSN, arrXMLs)
        Call SefazBA.IncluirAjustesSefazBA(TXTS, arrProdutosExcluidos, dicFornecSN)
        Call Util.MsgAviso("Ajustes incluídos com sucesso!", "Lançamento de Ajustes Sefaz/BA")
        
    End If
    
End Function

Public Function CalcularCreditoPresumidoInsterestadualDecretoAtacadistaBA(ByVal vOperacao As Double, ByVal vICMS As Double) As Double
    
    'A base legal para o crédito presumido nas operações de saída insterestadual está no art. 2º do Decreto 7.799/2000 da SEFAZ/BA
    'Observação: O estorno de débito só se aplicará as operações com alíquota igual ou superior a 12% [Base Legal: Parágrafo Único do art. 2º do Decreto 7.799/2000 da SEFAZ/BA]
    'Importante!: O estorno não se aplica as operações com papel higiênico [Base Legal: Parágrafo Único do art. 2º-A do Decreto 7.799/2000 da SEFAZ/BA]
    'Data da consulta: 06/02/2023
    
    Select Case True
    
        Case (vOperacao > 0) And (Round(vICMS / vOperacao, 2) >= 0.12)
            CalcularCreditoPresumidoInsterestadualDecretoAtacadistaBA = Round(vICMS * 0.16667, 2)
    
    End Select
    
End Function

Public Function CalcularEstornoCreditoDecretoAtacadistaBA(ByVal vOperacao As Double, ByVal vICMS As Double) As Double
    
    'A base legal para o estorno de créditos que exceda 10% do valor da operação está no art. 6º do Decreto 7.799/2000 da SEFAZ/BA
    'Observação: O estorno de crédito não se aplicará as operações de entradas de mercadorias decorrentes de importação do exterior [Base Legal: §2º do art. 6º do Decreto 7.799/2000 da SEFAZ/BA]
    'Importante!: O estorno não se aplica as operações com papel higiênico [Base Legal: Parágrafo Único do art. 2º-A do Decreto 7.799/2000 da SEFAZ/BA]
    'Data da consulta: 06/02/2023
    
    Select Case True
    
        Case (vOperacao > 0) And (Round(vICMS / vOperacao, 2) > 0.1)
            CalcularEstornoCreditoDecretoAtacadistaBA = Round(vICMS - ((vOperacao * 0.1)), 2)
    
    End Select
    
End Function

Public Function CalcularCreditoPresumidoArt269Inc10BA(ByVal vOperacao As Double) As Double
    
    'A base legal para o crédito presumido para os contribuintes sujeitos ao regime de conta-corrente fiscal nas aquisições internas
    'de mercadorias junto a microempresa ou empresa de pequeno porte'industrial optante pelo Simples Nacional, desde que por elas produzidas,
    'em opção ao crédito fiscal informado no documento fiscal nos termos do art. 57:
        
        'a) serão concedidos os créditos nos percentuais a seguir indicados, aplicáveis sobre o valor da operação:
            
            '1 - 10% (dez por cento) nas aquisições junto às indústrias do setor têxtil, de artigos de vestuário e acessórios,
            'de couro e derivados, moveleiro, metalúrgico, de celulose e de produtos de papel;

            '2 - 12% (doze por cento) nas aquisições junto aos demais segmentos de indústrias;
        
        'b) na hipótese de previsão na legislação de redução da base de cálculo na operação subsequente,
        'o crédito presumido previsto neste inciso fica reduzido na mesma proporção;
        
        'c) excluem-se do disposto neste inciso as mercadorias enquadradas no regime de substituição tributária;
    
    'Data da consulta: 07/02/2023
    
    CalcularCreditoPresumidoArt269Inc10BA = Round(vOperacao * 0.1, 2)
        
End Function

Public Function CalcularCreditoAquisicaoSimplesNacionalBA(ByVal vOperacao As Double, pCredSN As Double) As Double
    CalcularCreditoAquisicaoSimplesNacionalBA = Round(vOperacao * pCredSN, 2)
End Function
