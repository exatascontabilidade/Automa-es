Attribute VB_Name = "clsEnumeracoesSPEDContribuicoes"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function ValidarEnumeracao_COD_VER(ByVal Periodo As String) As String

Dim Ano As String
Dim Mes As String
    
    Ano = VBA.Right(Periodo, 4)
    Mes = VBA.Left(Periodo, 2)
    dtIni = Ano & "-" & Mes & "-" & "01"
    
    Select Case True
    
        Case dtIni >= "2020-01-01"
            ValidarEnumeracao_COD_VER = "006"
        
        Case dtIni >= "2019-01-01"
            ValidarEnumeracao_COD_VER = "005"
            
        Case dtIni >= "2018-01-01"
            ValidarEnumeracao_COD_VER = "004"
            
        Case dtIni >= "2012-01-01"
            ValidarEnumeracao_COD_VER = "003"
            
        Case dtIni >= "2011-04-01"
            ValidarEnumeracao_COD_VER = "002"
            
        Case Else
            ValidarEnumeracao_COD_VER = "001"
            
    End Select

End Function

Public Function ValidarEnumeracao_IND_ATIV(ByVal IND_ATIV As String)
    
    Select Case VBA.Val(IND_ATIV)
        
        Case "0"
            ValidarEnumeracao_IND_ATIV = "0 - Industrial ou equiparado a industrial"
            
        Case "1"
            ValidarEnumeracao_IND_ATIV = "1 - Prestador de serviços"
            
        Case "2"
            ValidarEnumeracao_IND_ATIV = "2 - Atividade de comércio"
            
        Case "3"
            ValidarEnumeracao_IND_ATIV = "3 - Pessoas jurídicas referidas nos §§ 6º, 8º e 9º do art. 3º da Lei nº 9.718, de 1998"
            
        Case "4"
            ValidarEnumeracao_IND_ATIV = "4 - Atividade imobiliária"
            
        Case "9"
            ValidarEnumeracao_IND_ATIV = "9 - Outros"
            
        Case Else
            ValidarEnumeracao_IND_ATIV = IND_ATIV & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_NAT_BC_CRED(ByVal NAT_BC_CRED As String) As String
            
    Select Case VBA.Format(Util.ApenasNumeros(NAT_BC_CRED), "00")
    
        Case "01"
            ValidarEnumeracao_NAT_BC_CRED = "01 - Aquisição de bens para revenda"
            
        Case "02"
            ValidarEnumeracao_NAT_BC_CRED = "02 - Aquisição de bens utilizados como insumo"
            
        Case "03"
            ValidarEnumeracao_NAT_BC_CRED = "03 - Aquisição de serviços utilizados como insumo"
            
        Case "04"
            ValidarEnumeracao_NAT_BC_CRED = "04 - Energia elétrica e térmica, inclusive sob a forma de vapor"
            
        Case "05"
            ValidarEnumeracao_NAT_BC_CRED = "05 - Aluguéis de prédios"
            
        Case "06"
            ValidarEnumeracao_NAT_BC_CRED = "06 - Aluguéis de máquinas e equipamentos"
            
        Case "07"
            ValidarEnumeracao_NAT_BC_CRED = "07 - Armazenagem de mercadoria e frete na operação de venda"
            
        Case "08"
            ValidarEnumeracao_NAT_BC_CRED = "08 - Contraprestações de arrendamento mercantil"
            
        Case "09"
            ValidarEnumeracao_NAT_BC_CRED = "09 - Máquinas, equipamentos e outros bens incorporados ao ativo imobilizado (crédito sobre encargos de depreciação)"
            
        Case "10"
            ValidarEnumeracao_NAT_BC_CRED = "10 - Máquinas, equipamentos e outros bens incorporados ao ativo imobilizado (crédito com base no valor de aquisição)"
            
        Case "11"
            ValidarEnumeracao_NAT_BC_CRED = "11 - Amortização e Depreciação de edificações e benfeitorias em imóveis"
            
        Case "12"
            ValidarEnumeracao_NAT_BC_CRED = "12 - Devolução de Vendas Sujeitas à Incidência Não-Cumulativa"
            
        Case "13"
            ValidarEnumeracao_NAT_BC_CRED = "13 - Outras Operações com Direito a Crédito (inclusive os créditos presumidos sobre receitas)"
            
        Case "14"
            ValidarEnumeracao_NAT_BC_CRED = "14 - Atividade de Transporte de Cargas – Subcontratação"
            
        Case "15"
            ValidarEnumeracao_NAT_BC_CRED = "15 - Atividade Imobiliária – Custo Incorrido de Unidade Imobiliária"
            
        Case "16"
            ValidarEnumeracao_NAT_BC_CRED = "16 - Atividade Imobiliária – Custo Orçado de unidade não concluída"
            
        Case "17"
            ValidarEnumeracao_NAT_BC_CRED = "17 - Atividade de Prestação de Serviços de Limpeza, Conservação e Manutenção – vale-transporte, vale-refeição ou vale-alimentação, fardamento ou uniforme"
            
        Case "18"
            ValidarEnumeracao_NAT_BC_CRED = "18 - Estoque de abertura de bens"
            
        Case ""
        Case Else
            ValidarEnumeracao_NAT_BC_CRED = NAT_BC_CRED & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_TIPO_ESCRIT(ByVal TIPO_ESCRIT As String)
    
    Select Case VBA.Val(TIPO_ESCRIT)
        
        Case "0"
            ValidarEnumeracao_TIPO_ESCRIT = "0 - Original"
            
        Case "1"
            ValidarEnumeracao_TIPO_ESCRIT = "1 - Retificadora"
            
        Case Else
            ValidarEnumeracao_TIPO_ESCRIT = TIPO_ESCRIT & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_IND_SIT_ESP(ByVal IND_SIT_ESP As String)
    
    Select Case VBA.Val(IND_SIT_ESP)
        
        Case "0"
            ValidarEnumeracao_IND_SIT_ESP = "0 - Abertura"
            
        Case "1"
            ValidarEnumeracao_IND_SIT_ESP = "1 - Cisão"
            
        Case "2"
            ValidarEnumeracao_IND_SIT_ESP = "2 - Fusão"
            
        Case "3"
            ValidarEnumeracao_IND_SIT_ESP = "3 - Incorporação"
            
        Case "4"
            ValidarEnumeracao_IND_SIT_ESP = "4 - Encerramento"
            
        Case Else
            ValidarEnumeracao_IND_SIT_ESP = IND_SIT_ESP & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_IND_NAT_PJ(ByVal IND_NAT_PJ As Variant) As String
    
    Select Case VBA.Format(Util.ApenasNumeros(IND_NAT_PJ), "00")
    
        Case "00"
            ValidarEnumeracao_IND_NAT_PJ = "00 - Pessoa jurídica em geral (não participante de SCP como sócia ostensiva)"
            
        Case "01"
            ValidarEnumeracao_IND_NAT_PJ = "01 - Sociedade cooperativa (não participante de SCP como sócia ostensiva)"
            
        Case "02"
            ValidarEnumeracao_IND_NAT_PJ = "02 - Entidade sujeita ao PIS/Pasep exclusivamente com base na Folha de Salários"
            
        Case "03"
            ValidarEnumeracao_IND_NAT_PJ = "03 - Pessoa jurídica em geral participante de SCP como sócia ostensiva"
            
        Case "04"
            ValidarEnumeracao_IND_NAT_PJ = "04 - Sociedade cooperativa participante de SCP como sócia ostensiva"
        
        Case "05"
            ValidarEnumeracao_IND_NAT_PJ = "05 - Sociedade em Conta de Participação - SCP"

        Case Else
            ValidarEnumeracao_IND_NAT_PJ = IND_NAT_PJ & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_COD_INC_TRIB(ByVal COD_INC_TRIB As String)
    
    Select Case VBA.Val(COD_INC_TRIB)
            
        Case "1"
            ValidarEnumeracao_COD_INC_TRIB = "1 - Escrituração de operações com incidência exclusivamente no regime não-cumulativo"
            
        Case "2"
            ValidarEnumeracao_COD_INC_TRIB = "2 - Escrituração de operações com incidência exclusivamente no regime cumulativo"
            
        Case "3"
            ValidarEnumeracao_COD_INC_TRIB = "3 - Escrituração de operações com incidência nos regimes não-cumulativo e cumulativo"

        Case Else
            ValidarEnumeracao_COD_INC_TRIB = COD_INC_TRIB & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_IND_APRO_CRED(ByVal IND_APRO_CRED As String)
    
    Select Case VBA.Val(IND_APRO_CRED)
            
        Case "1"
            ValidarEnumeracao_IND_APRO_CRED = "1 - Método de Apropriação Direta"
            
        Case "2"
            ValidarEnumeracao_IND_APRO_CRED = "2 - Método de Rateio Proporcional (Receita Bruta)"
    
        Case Else
            ValidarEnumeracao_IND_APRO_CRED = IND_APRO_CRED & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_COD_TIPO_CONT(ByVal COD_TIPO_CONT As String)
    
    Select Case VBA.Val(COD_TIPO_CONT)
            
        Case "1"
            ValidarEnumeracao_COD_TIPO_CONT = "1 - Apuração da Contribuição Exclusivamente a Alíquota Básica"
            
        Case "2"
            ValidarEnumeracao_COD_TIPO_CONT = "2 - Apuração da Contribuição a Alíquotas Específicas (Diferenciadas e/ou por Unidade de Medida de Produto)"

        Case Else
            ValidarEnumeracao_COD_TIPO_CONT = COD_TIPO_CONT & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_IND_REG_CUM(ByVal IND_REG_CUM As String)
    
    Select Case VBA.Val(IND_REG_CUM)
            
        Case "1"
            ValidarEnumeracao_IND_REG_CUM = "1 - Regime de Caixa – Escrituração consolidada (Registro F500)"
            
        Case "2"
            ValidarEnumeracao_IND_REG_CUM = "2 - Regime de Competência - Escrituração consolidada (Registro F550)"

        Case "9"
            ValidarEnumeracao_IND_REG_CUM = "9 - Regime de Competência - Escrituração detalhada, com base nos registros dos Blocos A, C, D e F"

        Case Else
            ValidarEnumeracao_IND_REG_CUM = IND_REG_CUM & " - Código Inválido"
            
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

Public Function ValidarEnumeracao_COD_DOC_IMP(ByVal COD_DOC_IMP As String)
    
    Select Case VBA.Val(COD_DOC_IMP)
        
        Case "0"
            ValidarEnumeracao_COD_DOC_IMP = "0 - Declaração de Importação"
            
        Case "1"
            ValidarEnumeracao_COD_DOC_IMP = "1 - Declaração Simplificada de Importação"
            
        Case Else
            ValidarEnumeracao_COD_DOC_IMP = COD_DOC_IMP & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_A100_IND_OPER(ByVal IND_OPER As Variant) As String
    
    Select Case VBA.Val(IND_OPER)
    
        Case "0"
            ValidarEnumeracao_A100_IND_OPER = "0 - Serviço Contratado pelo Estabelecimento"
        
        Case "1"
            ValidarEnumeracao_A100_IND_OPER = "1 - Serviço Prestado pelo Estabelecimento"
        
        Case Is <> ""
            ValidarEnumeracao_A100_IND_OPER = IND_OPER & " - Código Inválido"
                        
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

Public Function ValidarEnumeracao_IND_ORIG_CRED(ByVal IND_ORIG_CRED As Variant) As String
    
    Select Case VBA.Val(IND_ORIG_CRED)
        
        Case "0"
            ValidarEnumeracao_IND_ORIG_CRED = "0 - Operação no Mercado Interno"
            
        Case "1"
            ValidarEnumeracao_IND_ORIG_CRED = "1 - Operação de Importação"
            
        Case Is <> ""
            ValidarEnumeracao_IND_ORIG_CRED = IND_ORIG_CRED & " - Código Inválido"
            
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

Public Function ValidarEnumeracao_COD_NAT_CC(ByVal COD_NAT_CC As String)
            
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
                        
        Case "09"
            ValidarEnumeracao_COD_NAT_CC = "09 - Outras"
                
        Case ""
            ValidarEnumeracao_COD_NAT_CC = ""
            
        Case Else
            ValidarEnumeracao_COD_NAT_CC = COD_NAT_CC & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_IND_CTA(ByVal IND_CTA As Variant) As String
    
    Select Case VBA.UCase(VBA.Left(Util.RemoverAspaSimples(IND_CTA), 1))
        
        Case "S"
            ValidarEnumeracao_IND_CTA = "S - Sintética (grupo de contas)"
            
        Case "A"
            ValidarEnumeracao_IND_CTA = "A - Analítica (conta)"
            
        Case Is <> ""
            ValidarEnumeracao_IND_CTA = IND_CTA & " - Código Inválido"
            
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

Public Function ValidarEnumeracao_IND_PGTO(ByVal IND_PGTO As String) As String
    
    Select Case CStr(VBA.Val(IND_PGTO))
        
        Case "0"
            ValidarEnumeracao_IND_PGTO = "0 - À Vista"
        
        Case "1"
            ValidarEnumeracao_IND_PGTO = "1 - A Prazo"
            
        Case "2"
            ValidarEnumeracao_IND_PGTO = "2 - Outros"
        
        Case Is <> ""
            ValidarEnumeracao_IND_PGTO = IND_PGTO & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_IND_PGTO_A100(ByVal IND_PGTO As String) As String
    
    Select Case CStr(VBA.Val(IND_PGTO))
        
        Case "0"
            ValidarEnumeracao_IND_PGTO_A100 = "0 - À Vista"
        
        Case "1"
            ValidarEnumeracao_IND_PGTO_A100 = "1 - A Prazo"
            
        Case "9"
            ValidarEnumeracao_IND_PGTO_A100 = "9 - Sem pagamento"
        
        Case Is <> ""
            ValidarEnumeracao_IND_PGTO_A100 = IND_PGTO & " - Código Inválido"
            
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

Public Function ValidarEnumeracoes(ByVal nReg As String, ByVal Titulo As String, ByVal Valor As String)

Dim Val As String
    
    Select Case True
        
        Case nReg = "0000"
            If Titulo Like "TIPO_ESCRIT" Then Valor = ValidarEnumeracao_TIPO_ESCRIT(Valor)
            If Titulo Like "IND_SIT_ESP" Then Valor = ValidarEnumeracao_IND_SIT_ESP(Valor)
            If Titulo Like "IND_NAT_PJ" Then Valor = ValidarEnumeracao_IND_NAT_PJ(Valor)
            If Titulo Like "IND_ATIV" Then Valor = ValidarEnumeracao_IND_ATIV(Valor)
            
        Case nReg Like "*001"
            If Titulo Like "IND_MOV" Then Valor = ValidarEnumeracao_IND_MOV(Valor)
            
        Case nReg = "0200"
            If Titulo Like "TIPO_ITEM" Then Valor = ValidarEnumeracao_TIPO_ITEM(Valor)
            
        Case nReg = "0500"
            If Titulo Like "COD_NAT_CC" Then Valor = ValidarEnumeracao_COD_NAT_CC(Valor)
            If Titulo Like "IND_CTA" Then Valor = ValidarEnumeracao_IND_CTA(Valor)
            
        Case nReg = "C100"
            If Titulo Like "IND_OPER" Then Valor = ValidarEnumeracao_IND_OPER(Valor)
            If Titulo Like "IND_EMIT" Then Valor = ValidarEnumeracao_IND_EMIT(Valor)
            If Titulo Like "COD_SIT" Then Valor = ValidarEnumeracao_COD_SIT(Valor)
            If Titulo Like "IND_PGTO" Then Valor = ValidarEnumeracao_IND_PGTO(Valor)
            If Titulo Like "IND_FRT" Then Valor = ValidarEnumeracao_IND_FRT(Valor)
            
        Case nReg = "A100"
            If Titulo Like "IND_OPER" Then Valor = ValidarEnumeracao_IND_OPER(Valor)
            If Titulo Like "IND_EMIT" Then Valor = ValidarEnumeracao_IND_EMIT(Valor)
            If Titulo Like "COD_SIT" Then Valor = ValidarEnumeracao_COD_SIT(Valor)
            If Titulo Like "IND_PGTO" Then Valor = ValidarEnumeracao_IND_PGTO_A100(Valor)
            
        Case nReg = "C170"
            If Titulo Like "IND_MOV" Then Valor = ValidarEnumeracao_C170_IND_MOV(Valor)
            If Titulo Like "IND_APUR" Then Valor = ValidarEnumeracao_IND_APUR(Valor)
            If Titulo Like "CST_IPI" Then Valor = ValidarEnumeracao_CST_IPI(Valor)
            If Titulo Like "CST_PIS" Then Valor = ValidarEnumeracao_CST_PIS_COFINS(Valor)
            If Titulo Like "CST_COFINS" Then Valor = ValidarEnumeracao_CST_PIS_COFINS(Valor)
            
        Case nReg = "C175" Or nReg = "M400" Or nReg = "M800"
            If Titulo Like "CST_PIS" Then Valor = ValidarEnumeracao_CST_PIS_COFINS(Valor)
            If Titulo Like "CST_COFINS" Then Valor = ValidarEnumeracao_CST_PIS_COFINS(Valor)
            
        Case nReg = "C500"
            If Titulo Like "COD_SIT" Then Valor = ValidarEnumeracao_COD_SIT(Valor)
            
        Case nReg = "C501"
            If Titulo Like "NAT_BC_CRED" Then Valor = ValidarEnumeracao_NAT_BC_CRED(Valor)
            
        Case nReg = "C505"
            If Titulo Like "NAT_BC_CRED" Then Valor = ValidarEnumeracao_NAT_BC_CRED(Valor)
            
        Case nReg = "D101", nReg = "D501"
            If Titulo Like "NAT_BC_CRED" Then Valor = ValidarEnumeracao_NAT_BC_CRED(Valor)
            
        Case nReg = "D200"
            If Titulo Like "COD_SIT" Then Valor = ValidarEnumeracao_COD_SIT(Valor)
            
        Case nReg = "D105", nReg = "D505"
            If Titulo Like "NAT_BC_CRED" Then Valor = ValidarEnumeracao_NAT_BC_CRED(Valor)
            
        Case nReg = "M105"
            If Titulo Like "NAT_BC_CRED" Then Valor = ValidarEnumeracao_NAT_BC_CRED(Valor)
            If Titulo Like "CST_PIS" Then Valor = ValidarEnumeracao_CST_PIS_COFINS(Valor)
            
        Case nReg = "M505"
            If Titulo Like "NAT_BC_CRED" Then Valor = ValidarEnumeracao_NAT_BC_CRED(Valor)
            If Titulo Like "CST_COFINS" Then Valor = ValidarEnumeracao_CST_PIS_COFINS(Valor)
            
        Case nReg = "D100"
            If Titulo Like "IND_OPER" Then Valor = ValidarEnumeracao_D100_IND_OPER(Valor)
            If Titulo Like "IND_EMIT" Then Valor = ValidarEnumeracao_IND_EMIT(Valor)
            If Titulo Like "COD_SIT" Then Valor = ValidarEnumeracao_COD_SIT(Valor)
            If Titulo Like "TP_CT_E" Then Valor = ValidarEnumeracao_TP_CT_E(Valor)
            If Titulo Like "IND_FRT" Then Valor = ValidarEnumeracao_IND_FRT(Valor)
            
        Case nReg = "D500"
            If Titulo Like "IND_OPER" Then Valor = ValidarEnumeracao_D100_IND_OPER(Valor)
            If Titulo Like "IND_EMIT" Then Valor = ValidarEnumeracao_IND_EMIT(Valor)
            If Titulo Like "COD_SIT" Then Valor = ValidarEnumeracao_COD_SIT(Valor)
            If Titulo Like "TP_ASSINANTE" Then Valor = ValidarEnumeracao_TP_ASSINANTE(Valor)
            
        Case Else
            If Titulo Like "CST_PIS" Then Valor = ValidarEnumeracao_CST_PIS_COFINS(Valor)
            If Titulo Like "CST_COFINS" Then Valor = ValidarEnumeracao_CST_PIS_COFINS(Valor)
            
    End Select
    
    ValidarEnumeracoes = Valor
    
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
