Attribute VB_Name = "VariaveisConstantesGlobais"
Option Explicit

'---------------------------------------------
'Constantes
'---------------------------------------------

Public Const urlConfigComputador As String = "https://club.escoladaautomacaofiscal.com.br/138376-controldocs-academy/3218745-configurando-o-seu-computador"
Public Const urlConfigExcel As String = "https://club.escoladaautomacaofiscal.com.br/138376-controldocs-academy/3218746-configurando-o-excel"
Public Const URL As String = "https://script.google.com/macros/s/AKfycbxytaETr19GKodQZNzzqrc6KmdtKhBZ7kuM8SvxUORf_BfQJA-PIGMdX1EObv6whk3YYA/exec"
Public Const urlDocumentacao As String = "https://controldocs-doc.escoladaautomacaofiscal.com.br/"
Public Const urlTestesControlDocs As String = "http://127.0.0.1:8787/controldocs"
Public Const urlControlDocs As String = "https://api.escoladaautomacaofiscal.com.br/controldocs"
Public Const urlTutoriais As String = "https://controldocs.escoladaautomacaofiscal.com.br/tutoriais"
Public Const DownloadControlDocs As String = "https://downloadcontroldocs.escoladaautomacaofiscal.com.br"
Public Const urlSuporte As String = "https://wa.me/5571996699194?text=Sou+usu%C3%A1rio+do+ControlDocs+e+preciso+de+suporte."
Public Const urlSugestoes As String = "https://forms.gle/MUzX1DcAxeGZ21Fu6"
Public Const TokenControlDocs As String = "b9fdc103fbee50954822c1215a5f10b8777531dddb5317973f41ed393866bafa"
Public Const TokenApiTelegram As String = "f76b9989b2e64c14f6dae47a6fd5b18b66de36928af1cda7b0933ae9eb0d3656"
Public Const Tracking As String = "&utm_source=ControlDocs&utm_medium=Ferramenta&SCK=Ferramenta"
Public Const urlAssinaturaIndividualMensal As String = "https://buy.stripe.com/fZe8xX9AA8ddfWoaEF"
Public Const urlAssinaturaIndividualAnual As String = "https://buy.stripe.com/dR66pPeUU1OPh0s5km"
Public Const urlAssinaturaEmpresarialMensal As String = "https://buy.stripe.com/7sI5lLaEE511dOg147"
Public Const urlAssinaturaEmpresarialAnual As String = "https://buy.stripe.com/28o5lLdQQ9hhdOgbIN"
Public Const Planos As String = "basico,plus,premium,ultra, enterprise"
Public DivNotas As New AssistenteDivergenciasNotas

'Variáveis Personalizadas
Public relatICMS As CamposrelICMS
Public relIPI As CamposLivroIPI
Public assApuracaoPISCOFINS As CamposLivroPISCOFINS
Public resICMS As CamposRessarcimento


'Controles
Public ExportarC170Contribuicoes As Boolean
Public ExportarC175Contruicoes As Boolean
Public CadItensFornecProprios As Boolean
Public IgnorarEmissoesProprias As Boolean
Public ApropriarCreditosICMS As Boolean
Public ExportarC170Proprios As Boolean
Public ExportarC140Filhos As Boolean
Public ExportarPISCOFINS As Boolean
Public ImportarCTeD100 As Boolean
Public StatusPeriodo As Boolean
Public UsarPeriodo As Boolean


Public DesconsiderarAbatimento As Boolean
Public DesconsiderarPISCOFINS As Boolean
Public SomarICMSSTProdutos As Boolean
Public InterromperProcesso As Boolean
Public SomarIPIProdutos As Boolean
Public chNFSemValidade As Boolean
Public AtualizarM400M800 As Boolean


'Links dos tutoriais em vídeo
Public Const AcessarPlataforma As String = "https://controldocs.club.hotmart.com/lesson/YOmwg3djed/boas-vindas-controldocs"
Public Const videoAutenticarUsuario As String = "https://controldocs.club.hotmart.com/lesson/R4jZyNYV4a/ativando-a-sua-automacao-fiscal"
Public Const videoCadastrarContribuinte As String = "https://controldocs.club.hotmart.com/lesson/k7QxdAWzey/cadastrando-o-contribuinte"
Public Const videoTutorial As String = "https://controldocs.club.hotmart.com/lesson/k7QxdAWqey/visao-geral-de-recursos-do-controldocs"
Public Const AcessarClub As String = "https://club.escoladaautomacaofiscal.com.br"

'---------------------------------------------
'#Variáveis
'---------------------------------------------
' Arquivo
Public ARQUIVO As String


'Byte
Public dias As Byte
Public RegimePISCOFINS As Byte


'Long
Public a As Long
Public Trava As Long
Public DocsSemValidade As Long


'Double
Public Comeco As Double


'Faixa Personalizada
Public Rib As IRibbonUI


'Dicionarios
Public dicLayoutContribuicoes As New Dictionary
Public dicInconsistenciasIgnoradas As New Dictionary
Public dicHierarquiaSPEDFiscal As New Dictionary
Public dicMapaChavesSPEDFiscal As New Dictionary
Public dicLayoutFiscal As New Dictionary
Public dicChavesRegistroSPED As New Dictionary
Public dicEnderecosXML As New Dictionary
Public dicChavesNivel As New Dictionary
Public dicTabelaCFOP As New Dictionary
Public dicTabelaCEST As New Dictionary
Public dicRegistros As New Dictionary
Public dicTabelaNCM As New Dictionary
Public ListaChaves As New Dictionary
Public StatusSPED As New Dictionary
Public dicTitulos As New Dictionary
Public dicFilhos As New Dictionary
Public dicNomes As New Dictionary
Public dicPais As New Dictionary


'ArrayLists
Public arrCamposIgnorar As New ArrayList
Public arrEnumeracoesSPEDFiscal As New ArrayList
Public arrEnumeracoesSPEDContribuicoes As New ArrayList


'Dados do Contribuinte
Public CNPJBase As String
Public UFContribuinte As String
Public CNPJContribuinte As String
Public InscContribuinte As String
Public RazaoContribuinte As String
Public PeriodoEspecifico As String
Public PeriodoImportacao As String
Public PeriodoInventario As String


'Classes Gerais
Public Util As New clsUtilitarios
Public fnXML As New clsFuncoesXML
Public fnNFSe As New clsFuncoesNFSe
Public fnCSV As New clsFuncoesCSV
Public fnSPED As New clsFuncoesSPED
Public SefazBA As New clsSefazBA
Public Erro As New clsTratamentoErros
Public Cripto As New clsCriptografia_MD5
Public fnSeguranca As New clsFuncoesSeguranca

'Public PisCofins As New clsAssistentePISCOFINS
Public fnExcel As New clsFuncoesExcel
Public valXML As New clsValidacoesXML
Public RegrasFiscais As New clsRegrasFiscais
Public ValidacoesSPED As New clsValidacoesSPED
Public RegrasCadastrais As New clsRegrasCadastrais
Public Assistente As New clsAssistentesInteligentes
Public RegrasTributarias As New clsRegrasTributarias
Public Correlacionamentos As New clsCorrelacionamentoSPEDXML
Public AssTributario As New AssistenteTributario
Public tribICMS As New AssistenteTributarioICMS
Public tribPISCOFINS As New AssistenteTributarioPISCOFINS
Public tribIPI As New AssistenteTributarioIPI
Public Oportunidades As New AssistOportunidadesFiscais
Public DivergenciasNotas As New AssistenteDivergenciasNotas
Public DivergenciasProd As New AssistenteDivergenciasProdutos
Public impTributario As New ImportadorTributario
Public impTributarioNCM As New ImportadorTributarioNCM
Public AnalistaICMS As New AnalistaApuracaoICMS
Public AnalistaPISCOFINS As New AnalistaApuracaoPISCOFINS
Public Estoque As New AssistenteEstoque
Public Inventario As New AssistenteInventario

'Assistentes Inteligentes
Public Otimizacoes As New AssistentesOtimizacoesFiscais

'Classes do Bloco 0
Public r0000 As New cls0000
Public r0100 As New cls0100
Public r0150 As New cls0150
Public r0200 As New cls0200

'Classes do Bloco A
Public rA100 As New clsA100
Public rA170 As New clsA170


'Classes do Bloco C
Public rC100 As New clsC100
Public rC170 As New clsC170
Public rC175Contr As New clsC175Contrib
Public rC181 As New clsC181
Public rC185 As New clsC185
Public rC190 As New clsC190
Public rC850 As New clsC850


'Classes do Bloco D
Public rD100 As New clsD100
Public rD190 As New clsD190

'Classes do Bloco E
Public rE111 As New clsE111
Public rE113 As New clsE113


'Classes do Bloco M
Public rM400 As New clsM400

'Classes do Bloco K
Public rK200 As New clsK200


'Classes do Bloco 1
Public r1010 As New cls1010
Public r1400 As New cls1400


'Variáveis Personalizadas
Public DadosDoce As DadosNotasFiscais
Public DadosCTe As DadosConhecimentos


'Data
Public Inicio As Date
Public Vencimento As Date
Public UltimaConsulta As Date


'Texto
Public MyTag As String
Public EmailAssinante As String
Public IngorarCHV_PAI As String
Public VersaoFiscal As String
Public VersaoContribuicoes As String


'Variáveis personalizadas EFD-ICMS/IPI
Public regEFD As RegistrosEFD
Public RelDiverg As CamposrelInteligenteDivergencias
