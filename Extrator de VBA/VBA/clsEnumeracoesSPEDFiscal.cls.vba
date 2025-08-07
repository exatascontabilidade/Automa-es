Attribute VB_Name = "clsEnumeracoesSPEDFiscal"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function ListarEnumeracoes()
    
Dim Enumeracoes As Variant, Enumeracao
    
    Enumeracoes = Array("TIPO_ITEM", "IND_OPER", "IND_EMIT", "COD_SIT", "IND_PGTO", "IND_FRT", "IND_MOV", "TP_CT_E", "COD_CONS", _
            "TP_LIGACAO", "COD_GRUPO_TENSAO", "IND_DEST", "FIN_DOCE", "TP_ASSINANTE", "IND_APUR", "IND_TIT", "MOT_INV", _
            "IND_PROP", "CST_PIS", "CST_COFINS", "TIPO_ESCRIT", "IND_SIT_ESP", "IND_NAT_PJ", "IND_ATIV", "COD_INC_TRIB", _
            "IND_APRO_CRED", "COD_TIPO_CONT", "IND_REG_CUM", "COD_FIN", "COD_DOC_IMP", "NAT_BC_CRED", "IND_ORIG_CRED", _
            "IND_CTA", "COD_NAT_CC", "MOT_INV", "COD_NAT", "CST_IPI")
    
    For Each Enumeracao In Enumeracoes
        If Not arrEnumeracoesSPEDFiscal.contains(Enumeracao) Then arrEnumeracoesSPEDFiscal.Add Enumeracao
    Next Enumeracao
    
End Function

Public Function RemoverEnumeracoes(ByVal Campo As Variant, ByVal nCampo As String)
    
    Select Case nCampo
        
        Case "TIPO_ITEM", "IND_OPER", "IND_EMIT", "COD_SIT", "IND_PGTO", "IND_FRT", "IND_MOV", "TP_CT_E", "COD_CONS", _
            "TP_LIGACAO", "IND_DEST", "FIN_DOCE", "TP_ASSINANTE", "IND_APUR", "IND_TIT", "MOT_INV", _
            "IND_PROP", "CST_PIS", "CST_COFINS", "TIPO_ESCRIT", "IND_SIT_ESP", "IND_NAT_PJ", "IND_ATIV", "COD_INC_TRIB", _
            "IND_APRO_CRED", "COD_TIPO_CONT", "IND_REG_CUM", "COD_FIN", "COD_DOC_IMP", "NAT_BC_CRED", "IND_ORIG_CRED", _
            "COD_NAT_CC", "MOT_INV", "CST_IPI"
            
            RemoverEnumeracoes = Util.ApenasNumeros(Util.RemoverAspaSimples(Campo))
            
        Case "IND_CTA"
            RemoverEnumeracoes = VBA.Left(Campo, 1)
        
        Case "COD_NAT", "COD_GRUPO_TENSAO"
            RemoverEnumeracoes = VBA.Left(Campo, VBA.InStr(1, Campo, "-") - 2)
        
        Case Else
            RemoverEnumeracoes = Campo
            
    End Select
    
End Function

Public Function ValidarEnumeracao_COD_VER(ByVal Periodo As String) As String

Dim Ano As String
Dim Mes As String
    
    Ano = VBA.Right(Periodo, 4)
    Mes = VBA.Left(Periodo, 2)
    dtIni = Ano & "-" & Mes & "-" & "01"
    
    Select Case True
    
        Case dtIni >= "2025-01-01"
            ValidarEnumeracao_COD_VER = "019"
    
        Case dtIni >= "2024-01-01"
            ValidarEnumeracao_COD_VER = "018"
        
        Case dtIni >= "2023-01-01"
            ValidarEnumeracao_COD_VER = "017"
            
        Case dtIni >= "2022-01-01"
            ValidarEnumeracao_COD_VER = "016"
            
        Case dtIni >= "2021-01-01"
            ValidarEnumeracao_COD_VER = "015"
            
        Case dtIni >= "2020-01-01"
            ValidarEnumeracao_COD_VER = "014"
            
        Case dtIni >= "2019-01-01"
            ValidarEnumeracao_COD_VER = "013"
            
        Case dtIni >= "2018-01-01"
            ValidarEnumeracao_COD_VER = "012"
            
        Case dtIni >= "2017-01-01"
            ValidarEnumeracao_COD_VER = "011"
            
        Case dtIni >= "2016-01-01"
            ValidarEnumeracao_COD_VER = "010"
            
        Case dtIni >= "2015-01-01"
            ValidarEnumeracao_COD_VER = "009"
            
        Case dtIni >= "2014-01-01"
            ValidarEnumeracao_COD_VER = "008"
            
        Case dtIni >= "2013-01-01"
            ValidarEnumeracao_COD_VER = "007"
    
        Case dtIni >= "2012-07-01"
            ValidarEnumeracao_COD_VER = "006"
    
        Case dtIni >= "2012-01-01"
            ValidarEnumeracao_COD_VER = "005"
    
        Case dtIni >= "2011-01-01"
            ValidarEnumeracao_COD_VER = "004"
    
        Case dtIni >= "2010-01-01"
            ValidarEnumeracao_COD_VER = "003"
    
        Case dtIni >= "2009-01-01"
            ValidarEnumeracao_COD_VER = "002"
            
        Case Else
            ValidarEnumeracao_COD_VER = "001"
    
    End Select

End Function

Public Function ValidarEnumeracao_COD_FIN(ByVal COD_FIN As String)
    
    Select Case Util.ApenasNumeros(COD_FIN)
        
        Case "0"
            ValidarEnumeracao_COD_FIN = "0 - Original"
            
        Case "1"
            ValidarEnumeracao_COD_FIN = "1 - Retificadora"
            
        Case Else
            ValidarEnumeracao_COD_FIN = COD_FIN & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_IND_ATIV(ByVal IND_ATIV As String)
    
    Select Case VBA.Val(IND_ATIV)
        
        Case "0"
            ValidarEnumeracao_IND_ATIV = "0 - Industrial ou equiparado a industrial"
            
        Case "1"
            ValidarEnumeracao_IND_ATIV = "1 - Outros"
            
        Case Else
            ValidarEnumeracao_IND_ATIV = IND_ATIV & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_IND_MOV(ByVal IND_MOV As String)
    
    Select Case VBA.Val(IND_MOV)
        
        Case "0"
            ValidarEnumeracao_IND_MOV = "0 - Bloco com dados informados"
            
        Case "1"
            ValidarEnumeracao_IND_MOV = "1 - Bloco sem dados informados"
            
        Case Else
            ValidarEnumeracao_IND_MOV = IND_MOV & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_TIPO_ITEM(ByVal TIPO_ITEM As String)
            
    Select Case VBA.Format(Util.ApenasNumeros(TIPO_ITEM), "00")
    
        Case "00"
            ValidarEnumeracao_TIPO_ITEM = "00 - Mercadoria para Revenda"
            
        Case "01"
            ValidarEnumeracao_TIPO_ITEM = "01 - Matéria-Prima"
            
        Case "02"
            ValidarEnumeracao_TIPO_ITEM = "02 - Embalagem"
            
        Case "03"
            ValidarEnumeracao_TIPO_ITEM = "03 - Produto em Processo"
            
        Case "04"
            ValidarEnumeracao_TIPO_ITEM = "04 - Produto Acabado"
            
        Case "05"
            ValidarEnumeracao_TIPO_ITEM = "05 - Subproduto"
            
        Case "06"
            ValidarEnumeracao_TIPO_ITEM = "06 - Produto Intermediário"
            
        Case "07"
            ValidarEnumeracao_TIPO_ITEM = "07 - Material de Uso e Consumo"
            
        Case "08"
            ValidarEnumeracao_TIPO_ITEM = "08 - Ativo Imobilizado"
            
        Case "09"
            ValidarEnumeracao_TIPO_ITEM = "09 - Serviços"
        
        Case "10"
            ValidarEnumeracao_TIPO_ITEM = "10 - Outros Insumos"
        
        Case "99"
            ValidarEnumeracao_TIPO_ITEM = "99 - Outras"
        
        Case ""
            ValidarEnumeracao_TIPO_ITEM = ""
            
        Case Else
            ValidarEnumeracao_TIPO_ITEM = TIPO_ITEM & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_IND_OPER(ByVal IND_OPER As Variant) As String
    
    Select Case VBA.Val(IND_OPER)
    
        Case "0"
            ValidarEnumeracao_IND_OPER = "0 - Entrada"
        
        Case "1"
            ValidarEnumeracao_IND_OPER = "1 - Saida"
        
        Case Is <> ""
            ValidarEnumeracao_IND_OPER = IND_OPER & " - Código Inválido"
                        
    End Select
    
End Function

Public Function ValidarEnumeracao_D100_IND_OPER(ByVal IND_OPER As Variant) As String
    
    Select Case VBA.Val(IND_OPER)
    
        Case "0"
            ValidarEnumeracao_D100_IND_OPER = "0 - Aquisição"
        
        Case "1"
            ValidarEnumeracao_D100_IND_OPER = "1 - Prestação"
        
        Case Is <> ""
            ValidarEnumeracao_D100_IND_OPER = IND_OPER & " - Código Inválido"
                        
    End Select
    
End Function

Public Function ValidarEnumeracao_IND_EMIT(ByVal IND_EMIT As Variant) As String
    
    Select Case VBA.Val(IND_EMIT)
    
        Case "0"
            ValidarEnumeracao_IND_EMIT = "0 - Própria"
        
        Case "1"
            ValidarEnumeracao_IND_EMIT = "1 - Terceiros"
            
        Case Is <> ""
            ValidarEnumeracao_IND_EMIT = IND_EMIT & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_IND_TIT(ByVal IND_TIT As Variant) As String
    
    Select Case VBA.Format(Util.ApenasNumeros(IND_TIT), "00")
    
        Case "00"
            ValidarEnumeracao_IND_TIT = "00 - Duplicata"
            
        Case "01"
            ValidarEnumeracao_IND_TIT = "01 - Cheque"
            
        Case "02"
            ValidarEnumeracao_IND_TIT = "02 - Promissória"
            
        Case "03"
            ValidarEnumeracao_IND_TIT = "03 - Recibo"
            
        Case "99"
            ValidarEnumeracao_IND_TIT = "99 - Outros"
        
        Case Else
            ValidarEnumeracao_IND_TIT = IND_TIT & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_COD_NAT_CC(ByVal COD_NAT_CC As Variant) As String
    
    Select Case VBA.Format(Util.ApenasNumeros(COD_NAT_CC), "00")
                
        Case "01"
            ValidarEnumeracao_COD_NAT_CC = "01 - Contas de ativo"
            
        Case "02"
            ValidarEnumeracao_COD_NAT_CC = "02 - Contas de passivo"
            
        Case "03"
            ValidarEnumeracao_COD_NAT_CC = "03 - Patrimônio líquido"
            
        Case "04"
            ValidarEnumeracao_COD_NAT_CC = "04 - Contas de resultado"
            
        Case "05"
            ValidarEnumeracao_COD_NAT_CC = "05 - Contas de compensação"
            
        Case "06"
            ValidarEnumeracao_COD_NAT_CC = "09 - Outras"
            
        Case Is <> ""
            ValidarEnumeracao_COD_NAT_CC = COD_NAT_CC & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_COD_SIT(ByVal COD_SIT As Variant) As String
    
    Select Case VBA.Format(Util.ApenasNumeros(COD_SIT), "00")
    
        Case "00"
            ValidarEnumeracao_COD_SIT = "00 - Documento Regular"
        
        Case "01"
            ValidarEnumeracao_COD_SIT = "01 - Documento Extemporâneo Regular"
        
        Case "02"
            ValidarEnumeracao_COD_SIT = "02 - Documento Cancelado"
            
        Case "03"
            ValidarEnumeracao_COD_SIT = "03 - Cancelado Extemporâneo"
        
        Case "04"
            ValidarEnumeracao_COD_SIT = "04 - Documento Denegado"
        
        Case "05"
            ValidarEnumeracao_COD_SIT = "05 - Numeração Inutilizada"
        
        Case "06"
            ValidarEnumeracao_COD_SIT = "06 - Documento Complementar"
            
        Case "07"
            ValidarEnumeracao_COD_SIT = "07 - Documento Extemporâneo Complementar"
            
        Case "08"
            ValidarEnumeracao_COD_SIT = "08 - Regime Especial ou Norma Específica"
            
        Case Is <> ""
            ValidarEnumeracao_COD_SIT = COD_SIT & " - Código Inválido"
            
    End Select
        
End Function

Public Function ValidarEnumeracao_IND_PGTO(ByVal IND_PGTO As String) As String
    
    Select Case VBA.Val(IND_PGTO)
    
        Case "0"
            ValidarEnumeracao_IND_PGTO = "0 - À Vista"
        
        Case "1"
            ValidarEnumeracao_IND_PGTO = "1 - A Prazo"
            
        Case "2"
            ValidarEnumeracao_IND_PGTO = "2 - Outros"
        
        Case "9"
            ValidarEnumeracao_IND_PGTO = "9 - Outros"
            
        Case Else
            ValidarEnumeracao_IND_PGTO = IND_PGTO & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_IND_FRT_INICIAL(ByVal IND_FRT As Variant) As String
    
    Select Case VBA.Val(IND_FRT)
        
        Case "0"
            ValidarEnumeracao_IND_FRT_INICIAL = "0 - Por conta de terceiros"
            
        Case "1"
            ValidarEnumeracao_IND_FRT_INICIAL = "1 - Por conta do emitente"
            
        Case "2"
            ValidarEnumeracao_IND_FRT_INICIAL = "2 - Por conta do destinatário"
            
        Case "3"
            ValidarEnumeracao_IND_FRT_INICIAL = "9 - Sem cobrança de frete"
            
        Case Is <> ""
            ValidarEnumeracao_IND_FRT_INICIAL = IND_FRT & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_IND_FRT_2012(ByVal IND_FRT As Variant) As String
    
    Select Case VBA.Val(IND_FRT)
        
        Case "0"
            ValidarEnumeracao_IND_FRT_2012 = "0 - Por conta do emitente"
            
        Case "1"
            ValidarEnumeracao_IND_FRT_2012 = "1 - Por conta do destinatário/remetente"
            
        Case "2"
            ValidarEnumeracao_IND_FRT_2012 = "2 - Por conta de terceiros"
            
        Case "3"
            ValidarEnumeracao_IND_FRT_2012 = "9 - Sem cobrança de frete"
            
        Case Is <> ""
            ValidarEnumeracao_IND_FRT_2012 = IND_FRT & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_IND_FRT(ByVal IND_FRT As Variant) As String
    
    Select Case VBA.Val(IND_FRT)
    
        Case "0"
            ValidarEnumeracao_IND_FRT = "0 - Contratação do Frete por conta do Remetente (CIF)"
        
        Case "1"
            ValidarEnumeracao_IND_FRT = "1 - Contratação do Frete por conta do Destinatário (FOB)"
            
        Case "2"
            ValidarEnumeracao_IND_FRT = "2 - Contratação do Frete por conta de Terceiros"
            
        Case "3"
            ValidarEnumeracao_IND_FRT = "3 - Transporte Próprio por conta do Remetente"
            
        Case "4"
            ValidarEnumeracao_IND_FRT = "4 - Transporte Próprio por conta do Destinatário"
                    
        Case "9"
            ValidarEnumeracao_IND_FRT = "9 - Sem Ocorrência de Transporte"
                    
        Case Is <> ""
            ValidarEnumeracao_IND_FRT = IND_FRT & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_TP_CT_E(ByVal TP_CT_E As String)
    
    Select Case VBA.Val(TP_CT_E)
        
        Case "0"
            ValidarEnumeracao_TP_CT_E = "0 - CT-e ou BP-e Normal"
        
        Case "1"
            ValidarEnumeracao_TP_CT_E = "1 - CT-e de Complemento de Valores"
        
        Case "2"
            ValidarEnumeracao_TP_CT_E = "2 - CT-e emitido em hipótese de anulação de débito"
        
        Case "3"
            ValidarEnumeracao_TP_CT_E = "3 - CTE substituto do CT-e anulado ou BP-e substituição"
    
    End Select
    
End Function

Public Function ValidarEnumeracao_COD_CONS(ByVal COD_CONS As String)
            
    Select Case VBA.Format(Util.ApenasNumeros(COD_CONS), "00")
    
        Case "01"
            ValidarEnumeracao_COD_CONS = "01 - Comercial"
    
        Case "02"
            ValidarEnumeracao_COD_CONS = "02 - Consumo Próprio"
    
        Case "03"
            ValidarEnumeracao_COD_CONS = "03 - Iluminação Pública"
    
        Case "04"
            ValidarEnumeracao_COD_CONS = "04 - Industrial"
    
        Case "05"
            ValidarEnumeracao_COD_CONS = "05 - Poder Público"
    
        Case "06"
            ValidarEnumeracao_COD_CONS = "06 - Residencial"
    
        Case "07"
            ValidarEnumeracao_COD_CONS = "07 - Rural"
    
        Case "08"
            ValidarEnumeracao_COD_CONS = "08 - Serviço Público"
        
        Case Else
            ValidarEnumeracao_COD_CONS = COD_CONS & " - Código Inválido"
                                            
    End Select
        
End Function

Public Function ValidarEnumeracao_C170_IND_MOV(ByVal IND_MOV As String)
        
    Select Case VBA.Val(IND_MOV)
        
        Case "0"
            ValidarEnumeracao_C170_IND_MOV = "0 - SIM"
            
        Case "1"
            ValidarEnumeracao_C170_IND_MOV = "1 - NÃO"
            
        Case Else
            ValidarEnumeracao_C170_IND_MOV = IND_MOV & " - Código Inválido"
            
    End Select
        
End Function

Public Function ValidarEnumeracao_IND_APUR(ByVal IND_APUR As String)
            
        Select Case VBA.Val(IND_APUR)
        
            Case "0"
                ValidarEnumeracao_IND_APUR = "0 - MENSAL"
                
            Case "1"
                ValidarEnumeracao_IND_APUR = "1 - DECENDIAL"
                
            Case Else
                ValidarEnumeracao_IND_APUR = IND_APUR & " - Código Inválido"
                
        End Select
    
End Function

Public Function ValidarEnumeracao_CST_PIS_COFINS(ByVal CST As String)
    
    CST = VBA.Format(Util.ApenasNumeros(CST), "00")
    Select Case CST
                
        Case "01"
            ValidarEnumeracao_CST_PIS_COFINS = "01 - Operação Tributável com Alíquota Básica"
            
        Case "02"
            ValidarEnumeracao_CST_PIS_COFINS = "02 - Operação Tributável com Alíquota Diferenciada"
            
        Case "03"
            ValidarEnumeracao_CST_PIS_COFINS = "03 - Operação Tributável com Alíquota por Unidade de Medida de Produto"
            
        Case "04"
            ValidarEnumeracao_CST_PIS_COFINS = "04 - Operação Tributável Monofásica – Revenda a Alíquota Zero"
            
        Case "05"
            ValidarEnumeracao_CST_PIS_COFINS = "05 - Operação Tributável por Substituição Tributária"
            
        Case "06"
            ValidarEnumeracao_CST_PIS_COFINS = "06 - Operação Tributável a Alíquota Zero"
            
        Case "07"
            ValidarEnumeracao_CST_PIS_COFINS = "07 - Operação Isenta da Contribuição"
            
        Case "08"
            ValidarEnumeracao_CST_PIS_COFINS = "08 - Operação sem Incidência da Contribuição"
            
        Case "09"
            ValidarEnumeracao_CST_PIS_COFINS = "09 - Operação com Suspensão da Contribuição"
        
        Case "49"
            ValidarEnumeracao_CST_PIS_COFINS = "49 - Outras Operações de Saída"
        
        Case "50"
            ValidarEnumeracao_CST_PIS_COFINS = "50 - Operação com Direito a Crédito – Vinculada Exclusivamente a Receita Tributada no Mercado Interno"
        
        Case "51"
            ValidarEnumeracao_CST_PIS_COFINS = "51 - Operação com Direito a Crédito – Vinculada Exclusivamente a Receita Não-Tributada no Mercado Interno"
                              
        Case "52"
            ValidarEnumeracao_CST_PIS_COFINS = "52 - Operação com Direito a Crédito – Vinculada Exclusivamente a Receita de Exportação"
                              
        Case "53"
            ValidarEnumeracao_CST_PIS_COFINS = "53 - Operação com Direito a Crédito – Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno"
        
        Case "54"
            ValidarEnumeracao_CST_PIS_COFINS = "54 - Operação com Direito a Crédito – Vinculada a Receitas Tributadas no Mercado Interno e de Exportação"
                              
        Case "55"
            ValidarEnumeracao_CST_PIS_COFINS = "55 - Operação com Direito a Crédito – Vinculada a Receitas Não Tributadas no Mercado Interno e de Exportação"
                              
        Case "56"
            ValidarEnumeracao_CST_PIS_COFINS = "56 - Operação com Direito a Crédito – Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno e de Exportação"
                              
        Case "60"
            ValidarEnumeracao_CST_PIS_COFINS = "60 - Crédito Presumido – Operação de Aquisição Vinculada Exclusivamente a Receita Tributada no Mercado Interno"
                              
        Case "61"
            ValidarEnumeracao_CST_PIS_COFINS = "61 - Crédito Presumido – Operação de Aquisição Vinculada Exclusivamente a Receita Não-Tributada no Mercado Interno"
                              
        Case "62"
            ValidarEnumeracao_CST_PIS_COFINS = "62 - Crédito Presumido – Operação de Aquisição Vinculada Exclusivamente a Receita de Exportação"
                              
        Case "63"
            ValidarEnumeracao_CST_PIS_COFINS = "63 - Crédito Presumido – Operação de Aquisição Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno"
                              
        Case "64"
            ValidarEnumeracao_CST_PIS_COFINS = "64 - Crédito Presumido – Operação de Aquisição Vinculada a Receitas Tributadas no Mercado Interno e de Exportação"
                              
        Case "65"
            ValidarEnumeracao_CST_PIS_COFINS = "65 - Crédito Presumido – Operação de Aquisição Vinculada a Receitas Não-Tributadas no Mercado Interno e de Exportação"
                              
        Case "66"
            ValidarEnumeracao_CST_PIS_COFINS = "66 - Crédito Presumido – Operação de Aquisição Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno e de Exportação"
                              
        Case "67"
            ValidarEnumeracao_CST_PIS_COFINS = "67 - Crédito Presumido – Outras Operações"
                              
        Case "70"
            ValidarEnumeracao_CST_PIS_COFINS = "70 - Operação de Aquisição sem Direito a Crédito"
                              
        Case "71"
            ValidarEnumeracao_CST_PIS_COFINS = "71 - Operação de Aquisição com Isenção"
                              
        Case "72"
            ValidarEnumeracao_CST_PIS_COFINS = "72 - Operação de Aquisição com Suspensão"
                              
        Case "73"
            ValidarEnumeracao_CST_PIS_COFINS = "73 - Operação de Aquisição a Alíquota Zero"
                              
        Case "74"
            ValidarEnumeracao_CST_PIS_COFINS = "74 - Operação de Aquisição sem Incidência da Contribuição"
                              
        Case "75"
            ValidarEnumeracao_CST_PIS_COFINS = "75 - Operação de Aquisição por Substituição Tributária"
                              
        Case "98"
            ValidarEnumeracao_CST_PIS_COFINS = "98 - Outras Operações de Entrada"
                              
        Case "99"
            ValidarEnumeracao_CST_PIS_COFINS = "99 - Outras Operações"
                              
        Case Else
            ValidarEnumeracao_CST_PIS_COFINS = CST & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_CST_IPI(ByVal CST As String)
    
    CST = VBA.Format(Util.ApenasNumeros(CST), "00")
    Select Case CST
                
        Case ""
            ValidarEnumeracao_CST_IPI = ""
            
        Case "00"
            ValidarEnumeracao_CST_IPI = "00 - Entrada com Recuperação de Crédito"
            
        Case "01"
            ValidarEnumeracao_CST_IPI = "01 - Entrada Tributável com Alíquota Zero"
            
        Case "02"
            ValidarEnumeracao_CST_IPI = "02 - Entrada Isenta"
            
        Case "03"
            ValidarEnumeracao_CST_IPI = "03 - Entrada Não-Tributada"
            
        Case "04"
            ValidarEnumeracao_CST_IPI = "04 - Entrada Imune"
            
        Case "05"
            ValidarEnumeracao_CST_IPI = "05 - Entrada com Suspensão"
            
        Case "49"
            ValidarEnumeracao_CST_IPI = "49 - Outras Entradas"
            
        Case "50"
            ValidarEnumeracao_CST_IPI = "50 - Saída Tributada"
            
        Case "51"
            ValidarEnumeracao_CST_IPI = "51 - Saída Tributável com Alíquota Zero"
            
        Case "52"
            ValidarEnumeracao_CST_IPI = "52 - Saída Isenta"
            
        Case "53"
            ValidarEnumeracao_CST_IPI = "53 - Saída Não-Tributada"
            
        Case "54"
            ValidarEnumeracao_CST_IPI = "54 - Saída Imune"
            
        Case "55"
            ValidarEnumeracao_CST_IPI = "55 - Saída com Suspensão"
            
        Case "99"
            ValidarEnumeracao_CST_IPI = "99 - Outras Saídas"
            
        Case Else
            ValidarEnumeracao_CST_IPI = CST & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_TP_LIGACAO(ByVal TP_LIGACAO As String)
    
    Select Case VBA.Val(TP_LIGACAO)
        
        Case "1"
            ValidarEnumeracao_TP_LIGACAO = "1 - Monofásico"
            
        Case "2"
            ValidarEnumeracao_TP_LIGACAO = "2 - Bifásico"
            
        Case "3"
            ValidarEnumeracao_TP_LIGACAO = "3 - Trifásico"
            
        Case Else
            ValidarEnumeracao_TP_LIGACAO = TP_LIGACAO & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_IND_PROP(ByVal IND_PROP As String)
    
    Select Case VBA.Val(IND_PROP)
        
        Case "0"
            ValidarEnumeracao_IND_PROP = "0 - Item de propriedade do informante e em seu poder"
            
        Case "1"
            ValidarEnumeracao_IND_PROP = "1 - Item de propriedade do informante em posse de terceiros"
            
        Case "2"
            ValidarEnumeracao_IND_PROP = "2 - Item de propriedade de terceiros em posse do informante"
            
        Case ""
        Case Else
            ValidarEnumeracao_IND_PROP = IND_PROP & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_COD_GRUPO_TENSAO(ByVal COD_GRUPO_TENSAO As String)
    
    Select Case VBA.Format(Util.ApenasNumeros(COD_GRUPO_TENSAO), "00")
        
        Case "01"
            ValidarEnumeracao_COD_GRUPO_TENSAO = "01 - A1 - Alta Tensão (230kV ou mais)"
        
        Case "02"
            ValidarEnumeracao_COD_GRUPO_TENSAO = "02 - A2 - Alta Tensão (88 a 138kV)"
            
        Case "03"
            ValidarEnumeracao_COD_GRUPO_TENSAO = "03 - A3 - Alta Tensão (69kV)"
            
        Case "04"
            ValidarEnumeracao_COD_GRUPO_TENSAO = "04 - A3a - Alta Tensão (30kV a 44kV)"
            
        Case "05"
            ValidarEnumeracao_COD_GRUPO_TENSAO = "05 - A4 - Alta Tensão (2,3kV a 25kV)"
            
        Case "06"
            ValidarEnumeracao_COD_GRUPO_TENSAO = "06 - AS - Alta Tensão Subterrâneo 06"
            
        Case "07"
            ValidarEnumeracao_COD_GRUPO_TENSAO = "07 - B1 - Residencial "
            
        Case "08"
            ValidarEnumeracao_COD_GRUPO_TENSAO = "08 - B1 - Residencial Baixa Renda"
            
        Case "09"
            ValidarEnumeracao_COD_GRUPO_TENSAO = "09 - B2 - Rural"
            
        Case "10"
            ValidarEnumeracao_COD_GRUPO_TENSAO = "10 - B2 - Cooperativa de Eletrificação Rural"
            
        Case "11"
            ValidarEnumeracao_COD_GRUPO_TENSAO = "11 - B2 - Serviço Público de Irrigação"
            
        Case "12"
            ValidarEnumeracao_COD_GRUPO_TENSAO = "12 - B3 - Demais Classes"
            
        Case "13"
            ValidarEnumeracao_COD_GRUPO_TENSAO = "13 - B4a - Iluminação Pública - rede de distribuição"
            
        Case "14"
            ValidarEnumeracao_COD_GRUPO_TENSAO = "14 - B4b - Iluminação Pública - bulbo de lâmpada"
            
        Case Else
            ValidarEnumeracao_COD_GRUPO_TENSAO = COD_GRUPO_TENSAO & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_MOT_INV(ByVal MOT_INV As String)
    
    Select Case VBA.Format(Util.ApenasNumeros(MOT_INV), "00")
        
        Case "01"
            ValidarEnumeracao_MOT_INV = "01 - No final no período"
        
        Case "02"
            ValidarEnumeracao_MOT_INV = "02 - Na mudança de forma de tributação da mercadoria (ICMS)"
            
        Case "03"
            ValidarEnumeracao_MOT_INV = "03 - Na solicitação da baixa cadastral, paralisação temporária e outras situações"
            
        Case "04"
            ValidarEnumeracao_MOT_INV = "04 - Na alteração de regime de pagamento – condição do contribuinte"
            
        Case "05"
            ValidarEnumeracao_MOT_INV = "05 - Por determinação dos fiscos"
            
        Case "06"
            ValidarEnumeracao_MOT_INV = "06 - Para controle das mercadorias sujeitas ao regime de substituição tributária – restituição/ ressarcimento/ complementação"
            
        Case ""
        Case Else
            ValidarEnumeracao_MOT_INV = MOT_INV & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_FIN_DOCE(ByVal FIN_DOCE As String)
            
    Select Case VBA.Val(FIN_DOCE)
    
        Case "1"
            ValidarEnumeracao_FIN_DOCE = "1 - Normal"
    
        Case "2"
            ValidarEnumeracao_FIN_DOCE = "2 - Substituição"
    
        Case "3"
            ValidarEnumeracao_FIN_DOCE = "3 - Normal com ajuste"
    
        Case Else
            ValidarEnumeracao_FIN_DOCE = FIN_DOCE & " - Código Inválido"
    
    End Select

End Function

Public Function ValidarEnumeracao_IND_DEST(ByVal IND_DEST As String)
    
    Select Case VBA.Val(IND_DEST)
        
        Case "1"
            ValidarEnumeracao_IND_DEST = "1 - Contribuinte do ICMS"
            
        Case "2"
            ValidarEnumeracao_IND_DEST = "2 - Contribuinte Isento de Inscrição no Cadastro de Contribuintes do ICMS"
            
        Case "9"
            ValidarEnumeracao_IND_DEST = "9 - Não Contribuinte"
            
        Case Else
            ValidarEnumeracao_IND_DEST = IND_DEST & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_TP_ASSINANTE(ByVal TP_ASSINANTE As String)
            
    Select Case VBA.Val(TP_ASSINANTE)
        
        Case "1"
            ValidarEnumeracao_TP_ASSINANTE = "1 - Comercial/Industrial"
            
        Case "2"
            ValidarEnumeracao_TP_ASSINANTE = "2 - Poder Público"
            
        Case "3"
            ValidarEnumeracao_TP_ASSINANTE = "3 - Residencial/Pessoa física"
            
        Case "4"
            ValidarEnumeracao_TP_ASSINANTE = "4 - Público"
            
        Case "5"
            ValidarEnumeracao_TP_ASSINANTE = "5 - Semi-Público"
            
        Case "6"
            ValidarEnumeracao_TP_ASSINANTE = "6 - Outros"
            
        Case Else
            ValidarEnumeracao_TP_ASSINANTE = TP_ASSINANTE & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_COD_NAT(ByVal ARQUIVO As Variant, ByVal COD_NAT As Variant) As String
    
Dim dicDados0400 As New Dictionary
Dim dicTitulos0400 As New Dictionary
Dim CHV_0400 As String

    Set dicTitulos0400 = Util.MapearTitulos(reg0400, 3)
    Set dicDados0400 = Util.CriarDicionarioRegistro(reg0400, "ARQUIVO", "COD_NAT")

    CHV_0400 = VBA.Replace(VBA.Join(Array(ARQUIVO, COD_NAT)), " ", "")
    If dicDados0400.Exists(CHV_0400) Then
    
        ValidarEnumeracao_COD_NAT = COD_NAT & " - " & dicDados0400(CHV_0400)(dicTitulos0400("DESCR_NAT"))
    
    Else
    
        ValidarEnumeracao_COD_NAT = COD_NAT & " - NATUREZA NÃO CADASTRADA"
    
    End If
    
End Function

Public Function ValidarEnumeracoes(ByVal nReg As String, ByVal Titulo As String, ByVal Valor As String)

Dim Val As String
    
    Val = Util.ApenasNumeros(Valor)
    
    Select Case True
        
        Case nReg = "0000"
            If Titulo Like "COD_FIN" Then Valor = ValidarEnumeracao_COD_FIN(Val)
            If Titulo Like "IND_ATIV" Then Valor = ValidarEnumeracao_IND_ATIV(Val)
            
        
        Case nReg Like "*001"
            If Titulo Like "IND_MOV" Then Valor = ValidarEnumeracao_IND_MOV(Val)
        
        
        Case nReg = "0200"
            If Titulo Like "TIPO_ITEM" Then Valor = ValidarEnumeracao_TIPO_ITEM(Val)
        
        
        Case nReg = "0500"
            If Titulo Like "COD_NAT_CC" Then Valor = ValidarEnumeracao_COD_NAT_CC(Val)
            
            
        Case nReg = "C100"
            If Titulo Like "IND_OPER" Then Valor = ValidarEnumeracao_IND_OPER(Val)
            If Titulo Like "IND_EMIT" Then Valor = ValidarEnumeracao_IND_EMIT(Val)
            If Titulo Like "COD_SIT" Then Valor = ValidarEnumeracao_COD_SIT(Val)
            If Titulo Like "IND_PGTO" Then Valor = ValidarEnumeracao_IND_PGTO(Val)
            If Titulo Like "IND_FRT" Then Valor = ValidarEnumeracao_IND_FRT(Val)
            
            
        Case nReg = "C140"
            If Titulo Like "IND_TIT" Then Valor = ValidarEnumeracao_IND_TIT(Val)
            
            
        Case nReg = "C170"
            'If Titulo  Like "CST_ICMS" Then Valor = ExtrairCST_CSOSN_ICMS(Val)
            If Titulo Like "IND_MOV" Then Valor = ValidarEnumeracao_C170_IND_MOV(Val)
            If Titulo Like "IND_APUR" Then Valor = ValidarEnumeracao_IND_APUR(Val)
            If Titulo Like "CST_IPI" Then Valor = ValidarEnumeracao_CST_IPI(Val)
            If Titulo Like "CST_PIS" Then Valor = ValidarEnumeracao_CST_PIS_COFINS(Val)
            If Titulo Like "CST_COFINS" Then Valor = ValidarEnumeracao_CST_PIS_COFINS(Val)
        
        
        Case nReg = "C175" Or nReg = "M400" Or nReg = "M800"
            If Titulo Like "CST_PIS" Then Valor = ValidarEnumeracao_CST_PIS_COFINS(Val)
            If Titulo Like "CST_COFINS" Then Valor = ValidarEnumeracao_CST_PIS_COFINS(Val)
            
        
        Case nReg = "C500"
            If Titulo Like "IND_OPER" Then Valor = ValidarEnumeracao_IND_OPER(Val)
            If Titulo Like "IND_EMIT" Then Valor = ValidarEnumeracao_IND_EMIT(Val)
            If Titulo Like "COD_SIT" Then Valor = ValidarEnumeracao_COD_SIT(Val)
            If Titulo Like "COD_CONS" Then Valor = ValidarEnumeracao_COD_CONS(Val)
            If Titulo Like "TP_LIGACAO" Then Valor = ValidarEnumeracao_TP_LIGACAO(Val)
            If Titulo Like "COD_GRUPO_TENSAO" Then Valor = ValidarEnumeracao_COD_GRUPO_TENSAO(Val)
            If Titulo Like "FIN_DOCE" Then Valor = ValidarEnumeracao_FIN_DOCE(Val)
            If Titulo Like "IND_DEST" Then Valor = ValidarEnumeracao_IND_DEST(Val)
            
            
        Case nReg = "C800"
            If Titulo Like "COD_SIT" Then Valor = ValidarEnumeracao_COD_SIT(Val)
            
            
        Case nReg = "D100"
            If Titulo Like "IND_OPER" Then Valor = ValidarEnumeracao_D100_IND_OPER(Val)
            If Titulo Like "IND_EMIT" Then Valor = ValidarEnumeracao_IND_EMIT(Val)
            If Titulo Like "COD_SIT" Then Valor = ValidarEnumeracao_COD_SIT(Val)
            If Titulo Like "TP_CT_E" Then Valor = ValidarEnumeracao_TP_CT_E(Val)
            If Titulo Like "IND_FRT" Then Valor = ValidarEnumeracao_IND_FRT(Val)
                
                
        Case nReg = "D500"
            If Titulo Like "IND_OPER" Then Valor = ValidarEnumeracao_D100_IND_OPER(Val)
            If Titulo Like "IND_EMIT" Then Valor = ValidarEnumeracao_IND_EMIT(Val)
            If Titulo Like "COD_SIT" Then Valor = ValidarEnumeracao_COD_SIT(Val)
            If Titulo Like "TP_ASSINANTE" Then Valor = ValidarEnumeracao_TP_ASSINANTE(Val)
                
        Case nReg = "E500"
            If Titulo Like "IND_APUR" Then Valor = ValidarEnumeracao_IND_APUR(Val)
            
        Case nReg = "E510"
            If Titulo Like "CST_IPI" Then Valor = ValidarEnumeracao_CST_IPI(Val)
            
            
        Case nReg = "H005"
            If Titulo Like "MOT_INV" Then Valor = ValidarEnumeracao_MOT_INV(Val)
                
                
        Case Else
            If Titulo Like "CST_PIS" Then Valor = ValidarEnumeracao_CST_PIS_COFINS(Val)
            If Titulo Like "CST_COFINS" Then Valor = ValidarEnumeracao_CST_PIS_COFINS(Val)


    End Select
    
    ValidarEnumeracoes = Valor
    
End Function
