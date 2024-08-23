using SAPTests.Helpers;
using SAPTests.Helpers;
using Starline;
using System.Reflection;

namespace SAPTests.MaxCompra.Administracao.Compras
{
    public class TestHandler
    {
        private Dictionary<string, Action<Dictionary<string, string>>> testMethods;

        public TestHandler(Dictionary<string, Action<Dictionary<string, string>>> methods)
        {
            testMethods = methods ?? throw new ArgumentNullException(nameof(methods), "Dictionary of methods cannot be null.");
        }

        public static string GetCurrentMethodName()
        {
            MethodBase method = new System.Diagnostics.StackTrace().GetFrame(1).GetMethod();
            return method.Name;
        }

        public static string SetExecutionControl(string dataFilePath, string sheet, int testId, string queryName)
        {
            string executionType;
            try
            {
                executionType = Environment.GetEnvironmentVariable("executionType").ToString();
            }
            catch
            {
                executionType = "2";
                ExcelHelper.UpdateTestResult(dataFilePath, sheet, testId, "");
            }

            if (Global.firstRun)
            {
                if (executionType == "1")
                {
                    ExcelHelper.ResetTestResults(dataFilePath);
                }
                Global.firstRun = false;
            }

            Global.dataFetch = new DataFetch(ConnType: "Excel", ConnXLS: dataFilePath);
            Global.dataFetch.NewQuery(
                QueryName: queryName,
                QueryText: $"SELECT * FROM [{sheet}$] WHERE testId = {testId}"
            );
            string result = Global.dataFetch.GetValue("RESULT", queryName);
            return result;
        }

        public static void StartTest(DataFetch dataFetch, string queryName)
        {
            string scenarioName = dataFetch.GetValue("SCENARIONAME", queryName);
            string testName = dataFetch.GetValue("TESTNAME", queryName);
            string testType = dataFetch.GetValue("TESTTYPE", queryName);
            string analystName = dataFetch.GetValue("ANALYSTNAME", queryName);
            string testDesc = dataFetch.GetValue("TESTDESC", queryName);
            int reportID;
            try
            {
                reportID = int.Parse(Environment.GetEnvironmentVariable("reportID"));
                Console.WriteLine(reportID);
            }
            catch
            {
                reportID = int.Parse(dataFetch.GetValue("REPORTID", queryName));
            }
            Global.processTest.StartTest(Global.customerName, Global.suiteName, scenarioName, testName, testType, analystName, testDesc, reportID);
        }

        public static void DoTest(DataFetch dataFetch, string queryName)
        {
            string preCondition = dataFetch.GetValue("PRECONDITION", queryName);
            string postCondition = dataFetch.GetValue("POSTCONDITION", queryName);
            string inputData = dataFetch.GetValue("INPUTDATA", queryName);
            Global.processTest.DoTest(preCondition, postCondition, inputData);
        }

        public static void EndTest(DataFetch dataFetch, string queryName, bool closeWindow = true)
        {
            int reportID = int.Parse(dataFetch.GetValue("REPORTID", queryName));
            Global.processTest.EndTest(reportID, queryName);
            if (closeWindow == true)
            {
                WindowHandler.CloseWindow();
            }
        }

        public static void DefineSteps(string testName, int qtdLojas = 1)
        {
            switch (testName)
            {
                //App Login Screen
                case "RealizarLogin":
                    Global.processTest.DoStep("Abrir app");
                    Global.processTest.DoStep("Realizar login do analista");
                    Global.processTest.DoStep("Validar tela principal exibida");
                    break;

                // Gerenciador de Compras
                // Steps shortcuts
                case "Abrir Gerenciador de Compras":
                    Global.processTest.DoStep("Abrir menu Administração");
                    Global.processTest.DoStep("Abrir menu Compras");
                    Global.processTest.DoStep("Abrir menu Gerenciador de Compras");
                    break;

                case "Preencher fornecedor, categoria, abastecimento, recebimento e checkboxes":
                    Global.processTest.DoStep("Preencher fornecedor");
                    Global.processTest.DoStep("Selecionar categoria");
                    Global.processTest.DoStep("Preencher dias abastecimento");
                    Global.processTest.DoStep("Preencher Recebimento Em");
                    Global.processTest.DoStep("Habilitar checkbox Sugestão Compras");
                    break;
                case "Criar capa lote":
                    DefineSteps("RealizarLogin");
                    DefineSteps("Abrir Gerenciador de Compras");
                    DefineSteps("Preencher fornecedor, categoria, abastecimento, recebimento e checkboxes");
                    break;
                case "Adicionar lojas":
                    Global.processTest.DoStep("Adicionar lojas");
                    Global.processTest.DoStep("Confirmar janela Seleção de Empresas do Lote");
                    break;
                case "Incluir lote":
                    Global.processTest.DoStep("Incluir lote de compra");
                    Global.processTest.DoStep("Confirmar janela Filtros para Seleção de Produtos");
                    Global.processTest.DoStep("Maximizar janela");
                    break;
                case "Incluir lote com produtos inativos":
                    Global.processTest.DoStep("Incluir lote de compra");
                    Global.processTest.DoStep("Confirmar janela Filtros para Seleção de Produtos");
                    Global.processTest.DoStep("Confirmar janela Produtos Inativos");
                    Global.processTest.DoStep("Maximizar janela");
                    break;
                case "Incluir lote com tributação":
                    Global.processTest.DoStep("Incluir lote de compra");
                    Global.processTest.DoStep("Confirmar janela Filtros para Seleção de Produtos");
                    Global.processTest.DoStep("Confirmar janela Tributação");
                    Global.processTest.DoStep("Maximizar janela");
                    break;
                case "Incluir lote com produtos inativos e tributação":
                    Global.processTest.DoStep("Incluir lote de compra");
                    Global.processTest.DoStep("Confirmar janela Filtros para Seleção de Produtos");
                    Global.processTest.DoStep("Confirmar janela Produtos Inativos");
                    Global.processTest.DoStep("Confirmar janela Tributação");
                    Global.processTest.DoStep("Maximizar janela");
                    break;
                case "Gerar pedidos":
                    Global.processTest.DoStep("Gerar Pedidos");
                    Global.processTest.DoStep("Confirmar janela Consulta Lote de Compra");
                    break;
                case "Criar e incluir lote loja":
                    DefineSteps("Criar capa lote");
                    DefineSteps("Incluir lote com tributação");
                    break;
                case "Criar e incluir lote cd":
                    DefineSteps("Criar capa lote");
                    DefineSteps("Adicionar lojas");
                    Global.processTest.DoStep("Habilitar checkbox Incorporar Sugestão CD");
                    DefineSteps("Incluir lote com produtos inativos");
                    break;
                // Test Cases
                case "CriarLoteLoja":
                    DefineSteps("Criar e incluir lote loja");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar quantidade de compra dos produtos");
                    DefineSteps("Gerar pedidos");
                    break;
                case "CriarLoteCD":
                    DefineSteps("Criar e incluir lote cd");
                    Global.processTest.DoStep("Validar duplo click no campo QTD sugerida");
                    Global.processTest.DoStep("Edição do campo QtdeCompra no grid de produtos para lotes com Incorpora CD");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar quantidade de compra dos produtos");
                    DefineSteps("Gerar pedidos");
                    break;
                case "CriarLoteLojaBonificacao":
                    DefineSteps("Criar capa lote");
                    Global.processTest.DoStep("Preencher Limite Recebimento");
                    Global.processTest.DoStep("Alterar tipo do pedido");
                    Global.processTest.DoStep("Alterar tipo do acordo");
                    DefineSteps("Incluir lote com tributação");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar quantidade de compra dos produtos");
                    Global.processTest.DoStep("Gerar Pedidos");
                    Global.processTest.DoStep("Confirmar janela Manutenção de Acordos Promocionais");
                    break;
                case "CriarLoteCDBonificacao":
                    DefineSteps("Criar capa lote");
                    DefineSteps("Adicionar lojas");
                    Global.processTest.DoStep("Habilitar checkbox Incorporar Sugestão CD");
                    Global.processTest.DoStep("Preencher Limite Recebimento");
                    Global.processTest.DoStep("Alterar tipo do pedido");
                    Global.processTest.DoStep("Alterar tipo do acordo");
                    DefineSteps("Incluir lote com produtos inativos");
                    Global.processTest.DoStep("Validar duplo click no campo QTD sugerida");
                    Global.processTest.DoStep("Edição do campo QtdeCompra no grid de produtos para lotes com Incorpora CD");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar quantidade de compra dos produtos");
                    Global.processTest.DoStep("Gerar Pedidos");
                    Global.processTest.DoStep("Confirmar janela Manutenção de Acordos Promocionais");
                    break;

                case "CriarLoteFLVCompleto":
                    DefineSteps("CriarLoteFLVComprador");
                    for (int loja = 1; loja <= qtdLojas; loja++)
                    {
                        DefineSteps("PreencherLoteFLVChefeSessao");
                    }
                    DefineSteps("FinalizarLoteFLVComprador");
                    break;

                case "CriarLoteFLVComprador":
                    DefineSteps("Criar capa lote");
                    DefineSteps("Adicionar lojas");
                    Global.processTest.DoStep("Habilitar checkbox Restringe Empresa Loja");
                    DefineSteps("Incluir lote com produtos inativos e tributação");
                    Global.processTest.DoStep("Resgatar id do lote");
                    Global.processTest.DoStep("Fechar app");
                    break;

                case "PreencherLoteFLVChefeSessao":
                    DefineSteps("RealizarLogin");
                    DefineSteps("Abrir Gerenciador de Compras");
                    Global.processTest.DoStep("Abrir lote de compras");
                    Global.processTest.DoStep("Maximizar janela");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar quantidade de compra dos produtos");
                    Global.processTest.DoStep("Fechar app");
                    break;

                case "FinalizarLoteFLVComprador":
                    DefineSteps("RealizarLogin");
                    DefineSteps("Abrir Gerenciador de Compras");
                    Global.processTest.DoStep("Abrir lote de compras");
                    Global.processTest.DoStep("Maximizar janela");
                    Global.processTest.DoStep("Validar quantidade de compra dos produtos");
                    DefineSteps("Gerar pedidos");
                    break;

                case "ValidarAlteracaoPrazoPagamentoLoja":
                    DefineSteps("Criar e incluir lote loja");
                    Global.processTest.DoStep("Navegar para aba Parâmetros do Lote");
                    Global.processTest.DoStep("Alterar prazo pagamento");
                    Global.processTest.DoStep("Alterar prazo pagamento");
                    break;
                case "ValidarAlteracaoPrazoPagamentoCD":
                    DefineSteps("Criar e incluir lote cd");
                    Global.processTest.DoStep("Navegar para aba Parâmetros do Lote");
                    Global.processTest.DoStep("Alterar prazo pagamento");
                    Global.processTest.DoStep("Alterar prazo pagamento");
                    break;

                case "InserirItensLoteEmBrancoLoja":
                    DefineSteps("Criar capa lote");
                    Global.processTest.DoStep("Incluir lote de compra");
                    Global.processTest.DoStep("Confirmar janela Filtros para Seleção de Produtos");
                    Global.processTest.DoStep("Maximizar janela");
                    Global.processTest.DoStep("Alterar parâmetro de pesquisa");
                    Global.processTest.DoStep("Selecionar produtos na pesquisa");
                    Global.processTest.DoStep("Confirmar janela Pesquisa de Produtos");
                    Global.processTest.DoStep("Calcular tempo de execução");
                    break;

                case "InserirItensLoteEmBrancoCD":
                    DefineSteps("Criar capa lote");
                    DefineSteps("Adicionar lojas");
                    Global.processTest.DoStep("Habilitar checkbox Incorporar Sugestão CD");
                    Global.processTest.DoStep("Incluir lote de compra");
                    Global.processTest.DoStep("Confirmar janela Filtros para Seleção de Produtos");
                    Global.processTest.DoStep("Maximizar janela");
                    Global.processTest.DoStep("Alterar parâmetro de pesquisa");
                    Global.processTest.DoStep("Selecionar produtos na pesquisa");
                    Global.processTest.DoStep("Confirmar janela Pesquisa de Produtos");
                    Global.processTest.DoStep("Calcular tempo de execução");
                    break;

                case "ValidarBotaoAcataSugeridoLoja":
                    DefineSteps("Criar e incluir lote loja");
                    Global.processTest.DoStep("Clicar botão Acata Sugerido");
                    Global.processTest.DoStep("Validar Acata Sugerido");
                    break;

                case "ValidarBotaoAcataSugeridoCD":
                    DefineSteps("Criar e incluir lote cd");
                    Global.processTest.DoStep("Clicar botão Acata Sugerido");
                    Global.processTest.DoStep("Validar Acata Sugerido");
                    break;

                case "GerarLotePrazoUnicoLoja":
                case "ValidarPrazoVencimentoFixoGeracaoPedidoLoja":
                    DefineSteps("Criar capa lote");
                    Global.processTest.DoStep("Alterar prazo pagamento");
                    DefineSteps("Incluir lote com tributação");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar quantidade de compra dos produtos");
                    Global.processTest.DoStep("Gerar Pedidos");
                    //Global.processTest.DoStep("Resgatar ID do pedido");
                    Global.processTest.DoStep("Confirmar janela Consulta Lote de Compra");
                    Global.processTest.DoStep("Navegar para aba Parâmetros do Lote");
                    Global.processTest.DoStep("Abrir pedido gerado");
                    Global.processTest.DoStep("Validar Prazo de Pagamento");
                    break;

                case "GerarLotePrazoUnicoCD":
                case "ValidarPrazoVencimentoFixoGeracaoPedidoCD":
                    DefineSteps("Criar capa lote");
                    DefineSteps("Adicionar lojas");
                    Global.processTest.DoStep("Habilitar checkbox Incorporar Sugestão CD");
                    Global.processTest.DoStep("Alterar prazo pagamento");
                    DefineSteps("Incluir lote com produtos inativos");
                    Global.processTest.DoStep("Validar duplo click no campo QTD sugerida");
                    Global.processTest.DoStep("Edição do campo QtdeCompra no grid de produtos para lotes com Incorpora CD");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar quantidade de compra dos produtos");
                    DefineSteps("Gerar pedidos");
                    //Global.processTest.DoStep("Resgatar ID do pedido");
                    Global.processTest.DoStep("Confirmar janela Consulta Lote de Compra");
                    Global.processTest.DoStep("Navegar para aba Parâmetros do Lote");
                    Global.processTest.DoStep("Abrir pedido gerado");
                    Global.processTest.DoStep("Validar Prazo de Pagamento");
                    break;

                case "ZerarQuantidadeCompraLoja":
                    DefineSteps("Criar e incluir lote loja");
                    Global.processTest.DoStep("Clicar botão Acata Sugerido");
                    Global.processTest.DoStep("Validar Acata Sugerido");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar quantidade de compra do produto");
                    break;

                case "ZerarQuantidadeCompraCD":
                    DefineSteps("Criar e incluir lote cd");
                    Global.processTest.DoStep("Validar duplo click no campo QTD sugerida");
                    Global.processTest.DoStep("Edição do campo QtdeCompra no grid de produtos para lotes com Incorpora CD");
                    Global.processTest.DoStep("Clicar botão Acata Sugerido");
                    Global.processTest.DoStep("Validar Acata Sugerido");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar quantidade de compra do produto");
                    break;

                case "DiminuirQuantidadeCompraLoja":
                    DefineSteps("Criar e incluir lote loja");
                    Global.processTest.DoStep("Clicar botão Acata Sugerido");
                    Global.processTest.DoStep("Validar Acata Sugerido");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar quantidade de compra do produto");
                    break;

                case "DiminuirQuantidadeCompraCD":
                    DefineSteps("Criar e incluir lote cd");
                    Global.processTest.DoStep("Validar duplo click no campo QTD sugerida");
                    Global.processTest.DoStep("Edição do campo QtdeCompra no grid de produtos para lotes com Incorpora CD");
                    Global.processTest.DoStep("Clicar botão Acata Sugerido");
                    Global.processTest.DoStep("Validar Acata Sugerido");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar quantidade de compra do produto");
                    break;

                case "AumentarQuantidadeCompraLoja":
                    DefineSteps("Criar e incluir lote loja");
                    Global.processTest.DoStep("Clicar botão Acata Sugerido");
                    Global.processTest.DoStep("Validar Acata Sugerido");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar quantidade de compra do produto");
                    break;

                case "AumentarQuantidadeCompraCD":
                    DefineSteps("Criar e incluir lote cd");
                    Global.processTest.DoStep("Validar duplo click no campo QTD sugerida");
                    Global.processTest.DoStep("Edição do campo QtdeCompra no grid de produtos para lotes com Incorpora CD");
                    Global.processTest.DoStep("Clicar botão Acata Sugerido");
                    Global.processTest.DoStep("Validar Acata Sugerido");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar quantidade de compra do produto");
                    break;

                case "ValidarConsistenciaDataRecebimentoLoteExistenteLoja":
                    DefineSteps("Criar e incluir lote loja");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar quantidade de compra dos produtos");
                    Global.processTest.DoStep("Navegar para aba Parâmetros do Lote");
                    Global.processTest.DoStep("Preencher Recebimento Em");
                    Global.processTest.DoStep("Navegar para aba Produtos");
                    Global.processTest.DoStep("Tentar gerar Pedidos");
                    break;

                case "ValidarConsistenciaDataRecebimentoLoteExistenteCD":
                    DefineSteps("Criar e incluir lote cd");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar quantidade de compra dos produtos");
                    Global.processTest.DoStep("Navegar para aba Parâmetros do Lote");
                    Global.processTest.DoStep("Preencher Recebimento Em");
                    Global.processTest.DoStep("Navegar para aba Produtos");
                    Global.processTest.DoStep("Tentar gerar Pedidos");
                    break;

                case "ValidarQuantidadeTotalGerarPedidoLoja":
                    DefineSteps("Criar e incluir lote loja");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar quantidade total ao gerar pedido");
                    Global.processTest.DoStep("Confirmar janela Opções de geração do(s) pedido(s)");
                    Global.processTest.DoStep("Confirmar janela Consulta Lote de Compra");
                    break;

                case "ValidarQuantidadeTotalGerarPedidoCD":
                    DefineSteps("Criar e incluir lote cd");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar quantidade total ao gerar pedido");
                    Global.processTest.DoStep("Confirmar janela Opções de geração do(s) pedido(s)");
                    Global.processTest.DoStep("Confirmar janela Consulta Lote de Compra");
                    break;

                case "ValidarEmbalagemGridsLoja":
                    DefineSteps("Criar capa lote");
                    DefineSteps("Adicionar lojas");
                    DefineSteps("Incluir lote com produtos inativos e tributação");
                    Global.processTest.DoStep("Validar embalagens nos grids");
                    break;

                case "ValidarEmbalagemGridsCD":
                    DefineSteps("Criar e incluir lote cd");
                    Global.processTest.DoStep("Validar embalagens nos grids");
                    break;

                case "ValidarEstoqueDiasLoja":
                    DefineSteps("Criar e incluir lote loja");
                    Global.processTest.DoStep("Validar estoque dias na Grid Produtos");
                    Global.processTest.DoStep("Validar estoque dias na Consulta Produtos");
                    break;
                case "ValidarEstoqueDiasCD":
                    DefineSteps("Criar e incluir lote cd");
                    Global.processTest.DoStep("Validar estoque dias na Grid Produtos");
                    Global.processTest.DoStep("Validar estoque dias na Consulta Produtos");
                    break;

                case "ValidarObservacoesLoteLoja":
                    DefineSteps("Criar e incluir lote loja");
                    Global.processTest.DoStep("Preencher observações do lote");
                    Global.processTest.DoStep("Validar observações do lote");
                    break;

                case "ValidarObservacoesLoteCD":
                    DefineSteps("Criar e incluir lote cd");
                    Global.processTest.DoStep("Preencher observações do lote");
                    Global.processTest.DoStep("Validar observações do lote");
                    break;

                case "GerarLoteComItemDuplaProcedenciaLoja":
                    DefineSteps("Criar capa lote");
                    DefineSteps("Incluir lote");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar tipo abastecimento");
                    break;

                case "GerarLoteComItemDuplaProcedenciaCD":
                    DefineSteps("Criar e incluir lote cd");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar tipo abastecimento");
                    break;

                case "GerarLoteComDoisFornecedoresFLVCompleto":
                    DefineSteps("GerarLoteComDoisFornecedoresFLVCompradorCriarCapa");
                    DefineSteps("GerarLoteComDoisFornecedoresFLVChefeSessaoPreencher");
                    DefineSteps("GerarLoteComDoisFornecedoresCompradorFinalizar");
                    break;

                case "GerarLoteComDoisFornecedoresFLVCompradorCriarCapa":
                    DefineSteps("RealizarLogin");
                    DefineSteps("Abrir Gerenciador de Compras");
                    Global.processTest.DoStep("Preencher fornecedor");
                    Global.processTest.DoStep("Preencher fornecedor");
                    Global.processTest.DoStep("Selecionar categoria");
                    Global.processTest.DoStep("Preencher dias abastecimento");
                    Global.processTest.DoStep("Preencher Recebimento Em");
                    Global.processTest.DoStep("Habilitar checkbox Sugestão Compras");
                    Global.processTest.DoStep("Habilitar checkbox Restringe Empresa Loja");
                    DefineSteps("Incluir lote com tributação");
                    Global.processTest.DoStep("Resgatar id do lote");
                    Global.processTest.DoStep("Validar fornecedores");
                    break;

                case "GerarLoteComDoisFornecedoresFLVChefeSessaoPreencher":
                    DefineSteps("RealizarLogin");
                    DefineSteps("Abrir Gerenciador de Compras");
                    Global.processTest.DoStep("Abrir lote de compras");
                    Global.processTest.DoStep("Maximizar janela");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos");
                    Global.processTest.DoStep("Validar quantidade de compra dos produtos");
                    break;     
                
                case "GerarLoteComDoisFornecedoresCompradorFinalizar":
                    DefineSteps("RealizarLogin");
                    DefineSteps("Abrir Gerenciador de Compras");
                    Global.processTest.DoStep("Abrir lote de compras");
                    Global.processTest.DoStep("Maximizar janela");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos");
                    Global.processTest.DoStep("Alterar fornecedor na grid de produtos");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos");
                    DefineSteps("Gerar pedidos");
                    Global.processTest.DoStep("Validar fornecedores");
                    Global.processTest.DoStep("Confirmar janela Consulta Lote de Compra");
                    break;

                case "ValidarBloqueioSessaoSimulaneaMesmoLoteCompleto":
                    DefineSteps("ValidarBloqueioSessaoSimulaneaMesmoLoteSessaoUm");
                    DefineSteps("ValidarBloqueioSessaoSimulaneaMesmoLoteSessaoDois");
                    break;            
                
                case "ValidarBloqueioSessaoSimulaneaMesmoLoteSessaoUm":
                    DefineSteps("Criar e incluir lote loja");
                    Global.processTest.DoStep("Resgatar id do lote");
                    break;
                           
                case "ValidarBloqueioSessaoSimulaneaMesmoLoteSessaoDois":
                    DefineSteps("RealizarLogin");
                    DefineSteps("Abrir Gerenciador de Compras");
                    Global.processTest.DoStep("Abrir lote de compras");
                    Global.processTest.DoStep("Fechar app");
                    Global.processTest.DoStep("Fechar app");
                    break;

                case "ValidarPendenciaReceberAposCancelarEClicarRefresh":
                    DefineSteps("Criar e incluir lote loja");
                    Global.processTest.DoStep("Cancelar pendencia a receber de produto");
                    Global.processTest.DoStep("Validar pendencia receber, apos refresh");
                    break;

                default:
                    throw new Exception($"{testName} has no steps definition.");
            }
        }
    }
}