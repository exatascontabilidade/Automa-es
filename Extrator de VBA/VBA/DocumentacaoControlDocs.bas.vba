Attribute VB_Name = "DocumentacaoControlDocs"
Option Explicit

#If VBA7 Then
    ' Código para sistemas de 32 e 64 bits no Office 2010 e posterior
    Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#Else
    ' Código para sistemas de 32 bits no Office 2007 e anteriores
    Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#End If

Function ControlPressionado() As Boolean
    ' Esta função verifica se a tecla Control está pressionada
    ControlPressionado = (GetKeyState(vbKeyControl) And &H8000) <> 0
End Function

Public Function AcessarDocumentacao(ByRef control As IRibbonControl)

Dim URL As String
Dim urlBase As String

    urlBase = "https://controldocs-doc.escoladaautomacaofiscal.com.br/documentacao/"
    
    Select Case control.id
        
        'Botão Assinatura ControlDocs
        Case "btnAssinaturaControlDocs"
            URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs"
        
        '#Grupo Assinar ControlDocs
            
            '#Planos Básicos
                'Botão Assinar Plano Básico Mensal
                Case "btnBasicoMensal"
                    URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-assinar-controldocs/botao-assinar-controldocs/planos-basicos/plano-basico-mensal-controldocs"
                
                'Botão Assinar Plano Básico Semestral
                Case "btnBasicoSemestral"
                    URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-assinar-controldocs/botao-assinar-controldocs/planos-basicos/plano-basico-semestral-controldocs"
                
                'Botão Assinar Plano Básico Anual
                Case "btnBasicoAnual"
                    URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-assinar-controldocs/botao-assinar-controldocs/planos-basicos/plano-basico-anual-controldocs"
            
            '#Planos Plus
                'Botão Assinar Plano Plus Mensal
                Case "btnPlusMensal"
                    URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-assinar-controldocs/botao-assinar-controldocs/planos-plus/plano-plus-mensal-controldocs"
                
                'Botão Assinar Plano Plus Semestral
                Case "btnPlusSemestral"
                    URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-assinar-controldocs/botao-assinar-controldocs/planos-plus/plano-plus-semestral-controldocs"
                
                'Botão Assinar Plano Plus Anual
                Case "btnPlusAnual"
                    URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-assinar-controldocs/botao-assinar-controldocs/planos-plus/plano-plus-anual-controldocs"
            
            '#Planos Premium
                'Botão Assinar Plano Premium Mensal
                Case "btnPremiumMensal"
                    URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-assinar-controldocs/botao-assinar-controldocs/planos-Premium/plano-Premium-mensal-controldocs"
                
                'Botão Assinar Plano Premium Semestral
                Case "btnPremiumSemestral"
                    URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-assinar-controldocs/botao-assinar-controldocs/planos-Premium/plano-Premium-semestral-controldocs"
                
                'Botão Assinar Plano Premium Anual
                Case "btnPremiumAnual"
                    URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-assinar-controldocs/botao-assinar-controldocs/planos-Premium/plano-Premium-anual-controldocs"
            
            '#Assinatura Experimental
                'Botão Obter Assinatura Experimental
                Case "btnExperimentarControlDocs"
                    URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-assinar-controldocs/botao-obter-assinatura-experimental"
        
        '# Grupo Recursos de Assinatura
        
            'Botão Autenticar Usuário
            Case "btnAutenticarUsuario"
                URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-recursos-de-assinatura/botao-autenticar-usuario"
        
            'Botão Consultar Assinatura
            Case "btnConsultarAssinatura"
                URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-recursos-de-assinatura/botao-consultar-dados-da-assinatura"
        
            'Botão Limpar Dados Assinatura
            Case "btnLimparDadosAssinatura"
                URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-recursos-de-assinatura/botao-limpar-dados"
            
            
        'Botão Cadastro do Contribuinte
        Case "btnCadContrib"
            URL = "guia-controldocs/grupo-navegacao-rapida/cadastro-do-contribuinte"
            
            'Botão Extrair Cadastro da Web
            Case "btnExtCadWeb"
            URL = "guia-controldocs/grupo-navegacao-rapida/cadastro-do-contribuinte/extrair-cadastro-da-web"
                        
            'Botão Extrair Cadastro do SPED
            Case "btnExtCadSPED"
            URL = "guia-controldocs/grupo-navegacao-rapida/cadastro-do-contribuinte/extrair-cadastro-do-sped-fiscal"
                        
                        
        'Botão Recursos ControlDocs
        Case "btnRecursosControlDocs"
            URL = "guia-controldocs/grupo-navegacao-rapida/recursos-controldocs"
            
            'Botão Acessar Plataforma Educacional
            Case "btnControlDocs"
            URL = "guia-controldocs/grupo-navegacao-rapida/recursos-controldocs/grupo-recursos-de-ajuda/acessar-plataforma-educacional"
            
            'Botão Documentação ControlDocs
            Case "btnDocControlDocs"
            URL = "guia-controldocs/grupo-navegacao-rapida/recursos-controldocs/grupo-recursos-de-ajuda/documentacao-controldocs"
            
            'Botão Download da Versão Atual
            Case "btnDonwloadControlDocs"
            URL = "guia-controldocs/grupo-navegacao-rapida/recursos-controldocs/grupo-recursos-de-ajuda/download-da-versao-atual"
            
            'Botão Suporte Via WhatsApp
            Case "btnSuporte"
            URL = "guia-controldocs/grupo-navegacao-rapida/recursos-controldocs/grupo-recursos-de-ajuda/suporte-via-whatsapp"
            
            'Botão Sugerir Melhorias
            Case "btnSugestoes"
            URL = "guia-controldocs/grupo-navegacao-rapida/recursos-controldocs/grupo-recursos-de-ajuda/sugerir-melhorias"
            
            
        'Botão Configurações e Personalizações
        Case "btnConfiguracoesControlDocs"
            URL = "guia-controldocs/grupo-navegacao-rapida/configuracoes-e-personalizacoes"
            
            'CheckBox Remover Linhas de Grade
            Case "chLinhasGrade"
            URL = "guia-controldocs/grupo-navegacao-rapida/configuracoes-e-personalizacoes/grupo-personalizacao/remover-linhas-de-grade"
            
            'Botão Resetar ControlDocs
            Case "btnImportarExcel"
            URL = "guia-controldocs/grupo-navegacao-rapida/configuracoes-e-personalizacoes/grupo-configuracoes/botao-importar-dados-versao-anterior"
            
            'Botão Resetar ControlDocs
            Case "btnResetarControlDocs"
            URL = "guia-controldocs/grupo-navegacao-rapida/configuracoes-e-personalizacoes/grupo-configuracoes/resetar-controldocs"
            
            
        Case Else
            Call Util.MsgAviso("Esse recurso ainda não foi documentado." & vbCrLf & _
                "Caso precise de informações contate nosso suporte.", "Documentação ControlDocs")
            Exit Function
            
    End Select
    
    Call FuncoesLinks.AbrirUrl(urlBase & URL)
    
End Function
