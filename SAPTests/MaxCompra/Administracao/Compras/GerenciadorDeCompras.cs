using SAPTests.Helpers;
using SAPTests.MaxCompra.PageObjects.Administracao.Compras;
using SAPTests.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;
using System.Collections.ObjectModel;
using static SAPTests.Helpers.ElementHandler;

namespace SAPTests.MaxCompra.Administracao.Compras
{
    [TestClass]
    public class GerenciadorDeComprasTests : MaxCompraTests
    {
        private static GerenciadorDeComprasPageObject pageObject;
        private string idLote;
        private string idPedido;
        private string qtdeCompraEditValue;
        private int qtQtdeCompraTotal;
        private string embalagemProductGridValue;
        private double estoqueDiasExpected;
        private string pendReceberAntes;
        protected string sheet = "GerenciadorDeCompras";

        private void GetLoteId()
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep("Resgatar id do lote", logMsg: $"Resgatar id do lote",
                paramName: "Global.app", paramValue: Global.app);
            try
            {
                idLote = pageObject.GetIdLote();
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Resgate do id {idLote} com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "Erro ao tentar resgatar id do lote");
                throw;
            }

        }

        private void CloseApp()
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep("Fechar app", logMsg: $"Fechar app {Global.app}",
                paramName: "Global.app", paramValue: Global.app);
            try
            {
                WindowHandler.CloseWindow();
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "App fechado com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "Erro ao tentar fechar app");
                throw;
            }
        }

        private void OpenGerenciadorDeCompras()
        {
            string printFileName;

            string menuName = "Administração";
            int lgsID = Global.processTest.StartStep($"Abrir menu {menuName}", logMsg: $"Abrir menu {menuName}",
                paramName: "menuName", paramValue: menuName);
            try
            {
                OpenMenu(menuName);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"menu {menuName} aberto com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"Erro na abertura do menu {menuName}");
                throw;
            }

            menuName = "Compras";
            lgsID = Global.processTest.StartStep($"Abrir menu {menuName}", logMsg: $"Abrir menu {menuName}", paramName: "menuName",
                paramValue: menuName);
            try
            {
                OpenMenu(menuName);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"menu {menuName} aberto com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"Erro na abertura do menu {menuName}");
                throw;
            }

            menuName = "Gerenciador de Compras";
            lgsID = Global.processTest.StartStep($"Abrir menu {menuName}", logMsg: $"Abrir menu {menuName}", paramName: "menuName",
                paramValue: menuName);
            try
            {
                OpenMenu(menuName);
                WaitSeconds(1);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"menu {menuName} aberto com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"Erro na abertura do menu {menuName}");
                throw;
            }
            pageObject = new GerenciadorDeComprasPageObject();
        }

        private void SetSupplier(string codFornecedor, bool setMain = true)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep("Preencher fornecedor", logMsg: $"Preencher fornecedor {codFornecedor}",
                paramName: "fornecedorId", paramValue: codFornecedor);
            try
            {
                pageObject.FillFornecedor(codFornecedor, setMain);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Fornecedor preenchido com sucesso");
                KeyPresser.PressKey("F10");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro no preenchimento do fornecedor");
                throw;
            }
        }

        private void MaximizeWindow()
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep("Maximizar janela",
                logMsg: $"Tentando maximizar janela",
                paramName: "", paramValue: "");
            try
            {
                WindowHandler.MaximizeWindow();
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName,
                    logMsg: "Janela maximizada com sucesso.");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                    logMsg: "Erro ao tentar maximizar janela.");
                throw;
            }
        }

        private void AddLojas(List<string> lojas, string divisao, int qtdLojas)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep("Adicionar lojas", logMsg: $"Adicionar {lojas.Count} lojas", paramName: "lojas",
                paramValue: $"{lojas}");
            try
            {
                pageObject.OpenSelecaoDeLojas();
                WaitSeconds(2);
                pageObject.RemoveDivisoes();
                pageObject.AddDivisao(divisao);
                pageObject.RemoveLojas();
                pageObject.AddLojasPorNome(lojas);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Adição de {qtdLojas} lojas com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"Erro na adição de {qtdLojas} lojas");
                throw;
            }
        }

        private void GeneratePedidos(string tipoLote)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Gerar Pedidos", logMsg: $"Tentando Gerar Pedidos", paramName: "", paramValue: "");
            try
            {
                pageObject.ClickGerarPedidos();
                if (tipoLote == "cd")
                {
                    pageObject.ExitWindow("Atenção");
                }
                if (tipoLote == "flv")
                {
                    pageObject.ExitWindow("Atenção");
                    WindowsElement warning = FindElementByName("Atenção");
                    if (warning != null)
                    {
                        pageObject.ExitWindow("Atenção");
                    }
                }
                pageObject.ExitWindow("Opções de geração do(s) pedido(s)");
                WindowsElement warn = FindElementByXPathPartialName("com sucesso.");
                Assert.IsNotNull(warn);
                idPedido = pageObject.GetIdPedido();
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Pedido {idPedido} gerado com sucesso");
                // Confirma dialog do pedido criado
                KeyPresser.PressKey("RETURN");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                    logMsg: $"Erro ao tentar gerar Pedidos");
                throw;
            }
        }

        private void SholdNotGeneratePedidos(string tipoLote)
        {
            string printFileName;
            int lgsID;
            switch (tipoLote)
            {
                case "loja":
                    lgsID = Global.processTest.StartStep($"Tentar gerar Pedidos", logMsg: $"Tentando Gerar Pedidos com data anterior à hoje.",
                    paramName: "", paramValue: "");
                    try
                    {
                        pageObject.ClickGerarPedidos();
                        WindowsElement secondAlert = FindElementByName("Data de recebimento não pode ser anterior a data atual.");
                        Assert.IsNotNull(secondAlert);
                        printFileName = Global.processTest.CaptureWholeScreen();
                        Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Pedidos não gerados por restrição de data.");
                    }
                    catch
                    {
                        printFileName = Global.processTest.CaptureWholeScreen();
                        Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                            logMsg: $"Erro ao validar restrição de geração de pedido.");
                        throw;
                    }
                    break;
                case "cd":
                    lgsID = Global.processTest.StartStep($"Tentar gerar Pedidos", logMsg: $"Tentando Gerar Pedidos com data anterior à hoje.",
                    paramName: "", paramValue: "");
                    try
                    {
                        pageObject.ClickGerarPedidos();
                        WindowsElement firstAlert = FindElementByXPathPartialName("Ainda há itens inativos para a(s) Loja(s)");
                        Assert.IsNotNull(firstAlert);

                        ReadOnlyCollection<WindowsElement> buttons = FindElementsByClassName("Button");
                        WindowsElement exitButton = buttons[0];
                        exitButton.Click();

                        WindowsElement secondAlert = FindElementByName("Data de recebimento não pode ser anterior a data atual.");
                        Assert.IsNotNull(secondAlert);
                        printFileName = Global.processTest.CaptureWholeScreen();
                        Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Pedidos não gerados por restrição de data.");
                    }
                    catch
                    {
                        printFileName = Global.processTest.CaptureWholeScreen();
                        Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                            logMsg: $"Erro ao validar restrição de geração de pedido.");
                        throw;
                    }
                    break;
                case "flv":
                    lgsID = Global.processTest.StartStep($"Tentar gerar Pedidos", logMsg: $"Tentando Gerar Pedidos com data anterior à hoje.",
                    paramName: "", paramValue: "");
                    try
                    {
                        pageObject.ClickGerarPedidos();
                        WindowsElement alert = FindElementByName("Data de recebimento não pode ser anterior a data atual.");
                        Assert.IsNotNull(alert);
                        printFileName = Global.processTest.CaptureWholeScreen();
                        Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Pedidos não gerados por restrição de data.");
                    }
                    catch
                    {
                        printFileName = Global.processTest.CaptureWholeScreen();
                        Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                            logMsg: $"Erro ao validar restrição de geração de pedido.");
                        throw;
                    }
                    break;
                default:
                    throw new Exception("Erro ao tentar gerar pedidos!");
            }
            KeyPresser.PressKey("RETURN");
        }

        private void FillProdutos(int qtdProdutos, int qtdeCompra, int qtdLojas, string tipoLote, int productIndex = 1, bool selectProduct = true)
        {
            string printFileName;

            int lgsID = Global.processTest.StartStep($"Preencher quantidade de compra dos produtos",
            logMsg: $"Tentando preencher quantidade de compra do(s) produto(s)",
            paramName: "tipoLote, qtdProdutos, qtdeCompra, qtdLojas", paramValue: $"{tipoLote}, {qtdProdutos}, {qtdeCompra}, {qtdLojas}");
            try
            {
                pageObject.FillQtdeCompra(qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, productIndex: productIndex, selectProduct: selectProduct);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg:
                    $"Preenchido: {qtdProdutos} produto(s) com qtdeCompra: {qtdeCompra}");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                    logMsg: $"Erro ao tentar preencher quantidade de compra do(s) produto(s)");
                throw;
            }
        }

        private void ValidateQtdeCompraTotal(int qtdProdutos, int qtdeCompra, int qtdLojas, string tipoLote, string tipoPedido)
        {
            string printFileName;
            if (tipoLote == "cd")
            {
                int lgsID = Global.processTest.StartStep($"Validar quantidade de compra dos produtos",
             logMsg: $"Tentando validar quantidade de compra dos produtos",
             paramName: "qtdProdutos, qtdeCompra", paramValue: $"{qtdProdutos}, {qtdeCompra}");

                qtQtdeCompraTotal = pageObject.GetQtdeCompraTotal();
                bool totalIsValid = pageObject.ValidateQtdeComprasValue(qtQtdeCompraTotal, qtdProdutos, qtdeCompra, qtdLojas, tipoLote);

                if (!totalIsValid)
                {
                    printFileName = Global.processTest.CaptureWholeScreen();
                    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                        logMsg: $"Erro, esperado: {qtdProdutos * qtdeCompra}, atual: {qtQtdeCompraTotal}");
                    throw new Exception($"Erro, esperado: {qtdProdutos * qtdeCompra}, atual: {qtQtdeCompraTotal}");
                }

                WaitSeconds(3);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Sucesso, esperado: {qtdProdutos * qtdeCompra}, atual: {qtQtdeCompraTotal}");
            }
            else
            {
                int lgsID = Global.processTest.StartStep($"Validar quantidade de compra dos produtos",
             logMsg: $"Tentando validar quantidade de compra dos produtos",
             paramName: "qtdProdutos, qtdeCompra, qtdLojas", paramValue: $"{qtdProdutos}, {qtdeCompra}, {qtdLojas}");

                qtQtdeCompraTotal = pageObject.GetQtdeCompraTotal();
                bool totalIsValid = pageObject.ValidateQtdeComprasValue(qtQtdeCompraTotal, qtdProdutos, qtdeCompra, qtdLojas, tipoLote);

                if (!totalIsValid)
                {
                    printFileName = Global.processTest.CaptureWholeScreen();
                    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                        logMsg: $"Erro, esperado: {qtdProdutos * qtdeCompra * qtdLojas}, atual: {qtQtdeCompraTotal}");
                    throw new Exception($"Erro, esperado: {qtdProdutos * qtdeCompra * qtdLojas}, atual: {qtQtdeCompraTotal}");
                }
                WaitSeconds(3);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Sucesso, esperado: {qtdProdutos * qtdeCompra * qtdLojas}, atual: {qtQtdeCompraTotal}");
            }
        }

        private void ValidateQtdeTotalGerarPedido(string tipoLote)
        {
            string printFileName;
            int qtdTotalOpcoesGeracaoPedido;
            {
                int lgsID = Global.processTest.StartStep($"Validar quantidade total ao gerar pedido",
                 logMsg: $"Tentando validar quantidade total ao gerar pedido",
                 paramName: "qtQtdeCompraTotal", paramValue: $"{qtQtdeCompraTotal}");
                try
                {
                    pageObject.ClickGerarPedidos();
                    if (tipoLote == "cd")
                    {
                        pageObject.ExitWindow("Atenção");
                    }
                    if (tipoLote == "flv")
                    {
                        pageObject.ExitWindow("Atenção");
                        pageObject.ExitWindow("Atenção");
                    }

                    FindElementByName("Opções de geração do(s) pedido(s)");
                    qtdTotalOpcoesGeracaoPedido = pageObject.GetQtdTotalOpcoesGeracaoPedido(tipoLote);
                    Assert.AreEqual(qtQtdeCompraTotal, qtdTotalOpcoesGeracaoPedido);

                    printFileName = Global.processTest.CaptureWholeScreen();
                    Global.processTest.EndStep(lgsID, printPath: printFileName,
                        logMsg: $"Quantidade total validada, esperado: {qtQtdeCompraTotal}, atual: {qtdTotalOpcoesGeracaoPedido}");
                }
                catch
                {
                    FindElementByName("Opções de geração do(s) pedido(s)");
                    qtdTotalOpcoesGeracaoPedido = pageObject.GetQtdTotalOpcoesGeracaoPedido(tipoLote);

                    printFileName = Global.processTest.CaptureWholeScreen();
                    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                        logMsg: $"Erro, esperado: {qtQtdeCompraTotal}, atual: {qtdTotalOpcoesGeracaoPedido}");
                    throw;
                }
            }
        }

        private void ValidateQtdeCompraUmProduto(int qtdeCompra, int qtdLojas, string tipoLote, string codProduto)
        {
            string printFileName;
            {
                int lgsID = Global.processTest.StartStep($"Validar quantidade de compra do produto",
             logMsg: $"Tentando validar quantidade de compra do produto",
             paramName: "qtdeCompra", paramValue: $"{qtdeCompra}");

                int total = pageObject.GetQtdeCompraUmProduto(codProduto);
                bool totalIsValid = pageObject.ValidateQtdeComprasValue(total, 1, qtdeCompra, qtdLojas, tipoLote);

                if (!totalIsValid)
                {
                    printFileName = Global.processTest.CaptureWholeScreen();
                    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                        logMsg: $"Erro, esperado: {1 * qtdeCompra}, atual: {total}");
                    throw new Exception($"Erro, esperado: {1 * qtdeCompra}, atual: {total}");
                }

                WaitSeconds(3);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Sucesso, esperado: {1 * qtdeCompra}, atual: {total}");
            }
        }

        private void ClickOnIncluirLote()
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep("Incluir lote de compra", logMsg: $"Clicar no botão Incluir lote", paramName: "", paramValue: "");
            try
            {
                pageObject.ClickOnIncluirLote();
                WindowsElement button = FindElementByName("Gera Pedidos");
                Assert.IsNotNull(button);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Clique botão Incluir lote efetuado com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "Erro ao clicar no botão Incluir lote");
                throw;
            }
        }

        private void EnableCheckbox(string feature, string paramName = "", string paramValue = "")
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Habilitar checkbox {feature}", logMsg: $"Habilitar {feature}",
                            paramName: paramName, paramValue: paramValue);
            try
            {
                switch (feature)
                {
                    case "Incorporar Sugestão CD":
                        pageObject.EnableCheckbox(feature);
                        pageObject.SetCD(paramValue);
                        break;
                    case "Sugestão Compras":
                    case "Restringe Empresa Loja":
                        pageObject.EnableCheckbox(feature);
                        break;
                    default:
                        throw new ArgumentException($"Unsupported feature: {feature}");
                }

                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Habilitação {feature} com sucesso");
            }
            catch (Exception ex)
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                     logMsg: $"erro na habilitação do {feature}: {ex.Message}");
                throw;
            }
        }

        private void FillAbastecimentoDias(string diasAbastecimento)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep("Preencher dias abastecimento", logMsg: $"Preencher dias abastecimento com {diasAbastecimento}",
                paramName: "diasAbastecimento", paramValue: diasAbastecimento);
            try
            {
                pageObject.FillAbastecimentoDias(diasAbastecimento);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Preenchimento dias abastecimento com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro no preenchimento de dias abastecimento");
                throw;
            }
        }

        private void SelectCategoria(string categoria)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep("Selecionar categoria", logMsg: $"Selecionar categoria {categoria}",
                paramName: "categoria", paramValue: categoria);
            try
            {
                pageObject.SelectCategoria(categoria);
                WaitSeconds(3);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Seleção da categoria com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro na seleção da categoria");
                throw;
            }
        }

        private void ConfirmWindow(string windowName, int buttonIndex = 0)
        {
            string printFileName;
            int lgsID;

            switch (windowName)
            {
                case "Seleção de Empresas do Lote":
                case "Filtros para Seleção de Produtos":
                case "Produtos Inativos":
                case "Tributação":
                case "Atenção":
                case "Opções de geração do(s) pedido(s)":
                case "Consulta Lote de Compra":

                    lgsID = Global.processTest.StartStep($"Confirmar janela {windowName}", logMsg: $"Tentando confirmar janela {windowName}",
                        paramName: "windowName", paramValue: windowName);
                    try
                    {
                        WaitSeconds(5);
                        printFileName = Global.processTest.CaptureWholeScreen();
                        pageObject.ExitWindow(windowName, buttonIndex);
                        WaitSeconds(5);
                        Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Confirmação janela {windowName} com sucesso");
                    }
                    catch
                    {
                        printFileName = Global.processTest.CaptureWholeScreen();
                        Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"Erro na confirmação janela {windowName}");
                        throw;
                    }
                    break;
                case "Manutenção de Acordos Promocionais":
                    lgsID = Global.processTest.StartStep($"Confirmar janela {windowName}",
                        logMsg: $"Tentando confirmar janela {windowName}",
                        paramName: "windowName", paramValue: windowName);
                    try
                    {
                        WindowsElement window = FindElementByName(windowName);
                        Assert.IsNotNull(window);
                        printFileName = Global.processTest.CaptureWholeScreen();
                        pageObject.ExitWindow(windowName, buttonIndex);
                        WaitSeconds(5);
                        Global.processTest.EndStep(lgsID, printPath: printFileName,
                            logMsg: $"Confirmação janela {windowName} com sucesso");
                    }
                    catch
                    {
                        printFileName = Global.processTest.CaptureWholeScreen();
                        Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                            logMsg: $"Erro na confirmação janela {windowName}");
                        throw;
                    }
                    break;
                default:
                    throw new Exception($"Window {windowName} not found.");
            }
        }

        private ExecutionTimer ConfirmWindowWithTimer(string windowName, int buttonIndex = 0)
        {
            string printFileName;
            int lgsID;

            switch (windowName)
            {
                case "Pesquisa de Produtos":
                    ExecutionTimer timer = new ExecutionTimer();
                    lgsID = Global.processTest.StartStep($"Confirmar janela {windowName}", logMsg: $"Tentando confirmar janela {windowName}",
                        paramName: "windowName", paramValue: windowName);
                    try
                    {
                        timer = pageObject.ExitWindowWithTimer(timer, windowName, buttonIndex);
                        printFileName = Global.processTest.CaptureWholeScreen();
                        Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Confirmação janela {windowName} com sucesso");
                        return timer;
                    }
                    catch
                    {
                        printFileName = Global.processTest.CaptureWholeScreen();
                        Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"Erro na confirmação janela {windowName}");
                        throw;
                    }
                default:
                    throw new Exception($"Window {windowName} not found.");
            }
        }

        private void OpenLote(string idLote)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Abrir lote de compras", logMsg: $"Abrir lote de compras",
                paramName: "idLote", paramValue: idLote);
            try
            {
                pageObject.OpenLote(idLote);
                string message = "Não será possível visualizar este lote pois ele está aberto em outra sessão.";
                WindowsElement blockWarn = FindElementByXPathPartialName(message);
                if (blockWarn != null)
                {
                    printFileName = Global.processTest.CaptureWholeScreen();
                    Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg:
                        $"Lote de compras não aberto pois está aberto em outra sessão");
                    blockWarn.Click();
                    KeyPresser.PressKey("RETURN");
                }
                else
                {
                    string idOpenned = pageObject.GetIdLote();
                    Assert.AreEqual(idLote, idOpenned);
                    WaitSeconds(3);
                    printFileName = Global.processTest.CaptureWholeScreen();
                    Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg:
                        $"Lote de compras aberto com sucesso");

                }
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"Erro ao tentar abrir lote de compras: {idLote}");
                throw;
            }
        }

        private void ValidateDoubleClickOnQtdSugerida()
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Validar duplo click no campo QTD sugerida",
                logMsg: $"Duplo click no campo QTD sugerida para lotes com incorpora cd",
                paramName: "", paramValue: "");
            try
            {
                pageObject.DoubleClickOnQtdSugerida();
                printFileName = Global.processTest.CaptureWholeScreen();
                KeyPresser.PressKey("RETURN");
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Duplo click no campo QTD sugerida exibe aviso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao tentar duplo click no campo QTD sugerida para lotes com incorpora cd");
                throw;
            }
        }

        private void ValidateProductsGridEdit(int qtdeCompra)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Edição do campo QtdeCompra no grid de produtos para lotes com Incorpora CD",
                logMsg: $"Campo QtdeCompra no grid de produtos é editável",
                paramName: "qtdeCompra", paramValue: qtdeCompra.ToString());
            try
            {
                pageObject.FillProductsGrid(qtdeCompra);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Edição do campo QtdeCompra permitido no grid de produtos");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao tentar editar o campo QtdeCompra, no grid de produtos, para lotes com incorpora cd");
                throw;
            }
        }

        private void UpdateTipoPedido(string tipoPedido)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Alterar tipo do pedido",
                logMsg: $"Tentando alterar o tipo de pedido para {tipoPedido}",
                paramName: "tipoLote", paramValue: tipoPedido);
            try
            {
                pageObject.UpdateTipoPedido(tipoPedido);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Alteração do tipo do pedido com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao tentar alterar tipo do pedido");
                throw;
            }
        }

        private void FillLimiteRecebimento(string data)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Preencher Limite Recebimento",
                logMsg: $"Preencher Limite Recebimento com {data}",
                paramName: "dataAtual", paramValue: data);
            try
            {
                pageObject.FillLimiteRecebimento(data);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Campo Limite Recebimento preenchido");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao tentar preencher campo Limite Recebimento");
                throw;
            }
        }

        private void FillRecebimentoEm(string data)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Preencher Recebimento Em",
                logMsg: $"Preencher Recebimento Em com {data}",
                paramName: "data", paramValue: data);
            try
            {
                pageObject.FillRecebimentoEm(data);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Campo Recebimento Em preenchido");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao tentar preencher campo Recebimento Em");
                throw;
            }
        }

        private void UpdateTipoAcordo(string tipoAcordo)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Alterar tipo do acordo",
                logMsg: $"Alterar tipo acordo para {tipoAcordo}",
                paramName: "tipoAcordo", paramValue: tipoAcordo);
            try
            {
                pageObject.UpdateTipoAcordo(tipoAcordo);
                ReadOnlyCollection<WindowsElement> edits = FindElementsByClassName("Edit");
                string tipoAcordoIdEditValue = edits[41].GetAttribute("Value.Value");
                Assert.AreEqual(tipoAcordoIdEditValue, "8");
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Tipo acordo alterado.");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao tentar alterar tipo acordo.");
                throw;
            }
        }

        private void ValidatePrazoPagamento(string tipoPrazoPagamento, string prazoPagamento, string tipoLote)
        {
            string printFileName;
            if (tipoPrazoPagamento == "Prazo Fixo")
            {
                int lgsID = Global.processTest.StartStep($"Validar Prazo de Pagamento",
                logMsg: $"Validar Prazo de Pagamento: {prazoPagamento}",
                paramName: "prazoPagamento", paramValue: prazoPagamento);
                try
                {
                    bool valid = pageObject.ValidateVencimentoFixo(prazoPagamento, tipoLote);
                    Assert.IsTrue(valid);
                    printFileName = Global.processTest.CaptureWholeScreen();
                    Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Prazo de Pagamento validado: {prazoPagamento}.");
                }
                catch
                {
                    printFileName = Global.processTest.CaptureWholeScreen();
                    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                        $"Erro ao tentar validar Prazo de Pagamento: {prazoPagamento}.");
                    throw;
                }
            }
            else
            {
                int lgsID = Global.processTest.StartStep($"Validar Prazo de Pagamento",
                logMsg: $"Validar Prazo de Pagamento: {prazoPagamento}",
                paramName: "prazoPagamento", paramValue: prazoPagamento);
                try
                {
                    bool valid = pageObject.ValidatePrazoPagamento(prazoPagamento, tipoLote);
                    Assert.IsTrue(valid);
                    printFileName = Global.processTest.CaptureWholeScreen();
                    Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Prazo de Pagamento validado: {prazoPagamento}.");
                }
                catch
                {
                    printFileName = Global.processTest.CaptureWholeScreen();
                    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                        $"Erro ao tentar validar Prazo de Pagamento: {prazoPagamento}.");
                    throw;
                }
            }
        }

        private void ClickButtonAcataSugerido()
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Clicar botão Acata Sugerido",
                logMsg: $"Clicar botão Acata Sugerido",
                paramName: "", paramValue: "");
            try
            {
                pageObject.ClickButtonAcataSugerido();
                FindElementByName("OK");
                printFileName = Global.processTest.CaptureWholeScreen();
                WindowsElement button = FindElementByName("OK");
                button.Click();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg:
                    $"Botão Acata Sugerido funcional.");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao tentar clicar no botão Acata Sugerido.");
                throw;
            }
        }

        private void ValidateAcataSugerido(string tipoLote, string codProduto)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Validar Acata Sugerido",
                logMsg: $"Validar Acata Sugerido para {tipoLote}",
                paramName: "tipoLote", paramValue: tipoLote);
            string productFound = "";
            try
            {
                //pageObject.SelectFirstItemWithQtdeCompraFilled();
                productFound = pageObject.SelectItemOnProductGridByCode(codProduto, clickColumn: true, X: 1080);//, Y: 270);
                Assert.AreEqual(codProduto, productFound);

                qtdeCompraEditValue = pageObject.GetEditValueAfterAcataSugerido();
                string qtdeSugerida = pageObject.GetQtdeSugeridaValue(tipoLote);
                Assert.AreEqual(qtdeCompraEditValue, qtdeSugerida);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg:
                    $"Acata Sugerido refletiu corretamente, valor esperado: {qtdeSugerida}, valor encontrado: {qtdeCompraEditValue}.");
            }
            catch
            {
                if (productFound != "")
                {
                    string qtdeCompraEditValue = pageObject.GetEditValueAfterAcataSugerido();
                    string qtdeSugerida = pageObject.GetQtdeSugeridaValue(tipoLote);
                    printFileName = Global.processTest.CaptureWholeScreen();
                    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                        $"Erro ao validar Acata Sugerido, valor esperado: {qtdeSugerida}, valor encontrado: {qtdeCompraEditValue}.");
                    throw;
                }
                else
                {
                    printFileName = Global.processTest.CaptureWholeScreen();
                    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                        $"Erro ao encontrar o produto ID: {codProduto}.");
                    throw;
                }
            }
        }

        private void AlterSearchParameters(string searchParameter)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Alterar parâmetro de pesquisa",
                logMsg: $"Alterar parâmetro de pesquisa para ${searchParameter}",
                paramName: "searchParameter", paramValue: searchParameter);
            try
            {
                pageObject.AlterSearchParameters(searchParameter);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg:
                    $"Alteração de parâmetro de pesquisa para {searchParameter} refletiu corretamente.");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao alterar parâmetro de pesquisa para ${searchParameter}.");
                throw;
            }
        }

        private void SelectProducts(int qtdProdutos)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Selecionar produtos na pesquisa",
                logMsg: $"Seleção de {qtdProdutos} produtos na pesquisa.",
                paramName: "qtdProdutos", paramValue: qtdProdutos.ToString());
            try
            {
                pageObject.SelectProducts(qtdProdutos);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg:
                    $"Seleção de {qtdProdutos} produtos na pesquisa refletiu corretamente.");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao selecionar {qtdProdutos} produtos na pesquisa.");
                throw;
            }
        }

        private void LogExecutionTime(string scenario, ExecutionTimer timer)
        {
            string executionTime = "";
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Calcular tempo de execução",
                logMsg: $"Calculando tempo de execução para {scenario}.",
                paramName: "", paramValue: "");
            try
            {
                executionTime = pageObject.LogExecutionTime(scenario, timer);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg:
                    $"Tempo de execução para {scenario}: {executionTime}");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao calcular tempo de execução para {scenario}, valor encontrado: {executionTime}.");
                throw;
            }
        }

        private void UpdatePrazoPagamento(string tipoPrazoPagamento, string prazoPagamento)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Alterar prazo pagamento",
                logMsg: $"Alterar prazo pagamento para {tipoPrazoPagamento}.",
                paramName: "tipoPrazoPagamento, prazoPagamento", paramValue: $"{tipoPrazoPagamento}, {prazoPagamento}");
            try
            {
                pageObject.UpdatePrazoPagamento(tipoPrazoPagamento, prazoPagamento);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg:
                    $"Sucesso ao alterar prazo pagamento para {tipoPrazoPagamento}, {prazoPagamento}.");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao alterar prazo pagamento para {tipoPrazoPagamento}, {prazoPagamento}.");
                throw;
            }
        }

        private void GetPedidoID()
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Resgatar ID do pedido",
                logMsg: $"Resgatar ID do pedido.",
                paramName: "idPedido", paramValue: idPedido);
            try
            {
                idPedido = pageObject.GetIdPedido();
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg:
                    $"Sucesso ao resgatar ID do pedido: {idPedido}.");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao resgatar ID do pedido: {idPedido}.");
                throw;
            }
        }

        private void GoToTabParametrosDoLote()
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Navegar para aba Parâmetros do Lote",
                logMsg: $"Navegar para aba Parâmetros do Lote.",
                paramName: "", paramValue: "");
            try
            {
                pageObject.GoToTabParametrosDoLote();
                WindowsElement empresasButton = FindElementByName("Empresas");
                Assert.IsNotNull(empresasButton);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg:
                    $"Sucesso ao navegar para aba Parâmetros do Lote.");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao navegar para aba Parâmetros do Lote.");
                throw;
            }
        }

        private void ValidateSuppliers(List<string> suppliers, string scenario)
        {
            string fornecedores = StringHandler.ConvertListToString(suppliers);
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Validar fornecedores",
                logMsg: $"Tentando validar fornecedores: {suppliers} em {scenario}.",
                paramName: "suppliers", paramValue: $"{fornecedores}");

            try
            {
                int suppliersCount = pageObject.GetSuppliersCount(scenario, suppliers);
                Assert.AreEqual(2, suppliersCount);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg:
                    $"Sucesso ao validar fornecedores: {fornecedores} em {scenario}.");
                KeyPresser.PressKey("ESCAPE");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao validar fornecedores: {fornecedores} em {scenario}.");
                throw;
            }
        }

        private void GoToTabProdutos()
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Navegar para aba Produtos",
                logMsg: $"Navegar para aba Produtos.",
                paramName: "", paramValue: "");
            try
            {
                pageObject.GoToTabProdutos();
                WindowsElement button = FindElementByName("Gera Pedidos");
                Assert.IsNotNull(button);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg:
                    $"Sucesso ao navegar para aba Produtos.");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao navegar para aba Produtos");
                throw;
            }
        }

        private void OpenPedido()
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Abrir pedido gerado",
                logMsg: $"Tentando abrir pedido gerado {idPedido}.",
                paramName: "", paramValue: "");
            try
            {
                pageObject.OpenPedido(idPedido);
                FindElementByName("Pedido de Compras / Transferências de Suprimentos");
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg:
                    $"Sucesso ao tentar abrir pedido {idPedido}.");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao tentar abrir pedido {idPedido}.");
                throw;
            }
        }

        private void ValidateEmbalagens(string tipoLote)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Validar embalagens nos grids",
                logMsg: $"Tentando validar embalagens nos grids.",
                paramName: "", paramValue: "");

            WaitSeconds(5);
            embalagemProductGridValue = pageObject.GetEmbalagemProductGridValue(tipoLote);
            string embalagemEmpresasGridValue = pageObject.GetEmbalagemLojasGridValue(tipoLote);

            try
            {
                Assert.AreEqual(embalagemProductGridValue, embalagemEmpresasGridValue);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg:
                    $"Sucesso ao validar embalagens nos grids de produto ({embalagemProductGridValue}) e empresa ({embalagemEmpresasGridValue}).");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao validar embalagens nos grids de produto ({embalagemProductGridValue}) e empresa ({embalagemEmpresasGridValue}).");
                throw;
            }
        }

        private void ValidateEstoqueDias(string windowName, string tipoLote, List<string> lojas)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Validar estoque dias na {windowName}",
                logMsg: $"Tentando validar estoque dias na {windowName}.",
                paramName: "windowName, tipoLote", paramValue: $"{windowName}, {tipoLote}");
            if (windowName == "Grid Produtos")
            {
                pageObject.SelectFirstItemWithEstoqueFilled(tipoLote);
                embalagemProductGridValue = pageObject.GetEmbalagemProductGridValue(tipoLote);
                WaitSeconds(2);
                double estoqueCD;
                if (tipoLote == "loja")
                {
                    estoqueCD = 0;
                }
                else
                {
                    estoqueCD = pageObject.GeteEtqQtdCentral();
                }
                estoqueDiasExpected = pageObject.CalculateEstoqueDiasExpected(estoqueCD);
                double estoqueDiasActual = pageObject.GetEstqDias();
                double estoqueQtd = pageObject.GetEstoqueQtd();
                double mediaVendaDia = pageObject.GetMedVDiaPeriodo();
                try
                {
                    Assert.AreEqual(estoqueDiasExpected, estoqueDiasActual);
                    printFileName = Global.processTest.CaptureWholeScreen();
                    Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg:
                        $"Fórmula: " +
                        $"estoqueDias({estoqueDiasActual}) = (estoqueQtd({estoqueQtd}) + estoqueCD({estoqueCD})) / mediaVendaDia({mediaVendaDia}) | " +
                        $"esperado: {estoqueDiasExpected}, " +
                        $"encontrado: {estoqueDiasActual}");
                }
                catch
                {
                    printFileName = Global.processTest.CaptureWholeScreen();
                    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                        $"Erro ao validar estoque dias, " +
                        $"esperado: {estoqueDiasExpected}, " +
                        $"encontrado: {estoqueDiasActual})");
                    throw;
                }
                WaitSeconds(2);
            }

            else if (windowName == "Consulta Produtos")
            {
                double emDiasActual = pageObject.GetEmDiasValue(lojas, embalagemProductGridValue);
                double emDiasExpected = Math.Round(estoqueDiasExpected, 1);
                try
                {
                    Assert.AreEqual(emDiasActual, emDiasExpected);
                    printFileName = Global.processTest.CaptureWholeScreen();
                    Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg:
                        $"Sucesso, " +
                        $"esperado: {emDiasExpected}, " +
                        $"encontrado: {emDiasActual}");
                }
                catch
                {
                    printFileName = Global.processTest.CaptureWholeScreen();
                    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                        $"Erro, " +
                        $"esperado: {emDiasExpected}, " +
                        $"encontrado: {emDiasActual}");
                    throw;
                }
            }
        }

        private void ValidateTipoAbastecimento(int productIndex, string codProduto, string tipoLote)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Validar tipo abastecimento",
                logMsg: $"Tentando validar tipo abastecimento.",
                paramName: "productIndex, tipoLote", paramValue: $"{productIndex}, {tipoLote}");
            pageObject.SelectItemOnProductGridByIndex(productIndex);
            pageObject.OpenConsultaProdutos();
            List<string> formasAbastecimento = pageObject.GetFormasAbastecimento(tipoLote);
            string expected = "3 - Entregue CD/Entrada Direto Loja";
            try
            {
                string actual = formasAbastecimento[0];
                Assert.AreEqual(expected, actual);
                if (tipoLote == "cd")
                {
                    actual = formasAbastecimento[1];
                    Assert.AreEqual(expected, actual);
                }
                string result = string.Join(", ", formasAbastecimento);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg:
                    $"Sucesso ao validar tipo abastecimento do produto {codProduto}, " +
                    $"Esperado: {expected}, " +
                    $"Encontrado: {result}.");
            }
            catch
            {
                string result = string.Join(", ", formasAbastecimento);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao validar tipo abastecimento do produto {codProduto}, " +
                    $"Esperado: {expected}, " +
                    $"Encontrado: {result}.");
                throw;
            }
        }

        private void ChangeSupplier()
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Alterar fornecedor na grid de produtos",
                logMsg: $"Tentando alterar fornecedor na grid de produtos.",
                paramName: "", paramValue: "");
            try
            {
                WaitSeconds(15);
                pageObject.ChangeSupplier();
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg:
                    $"Sucesso ao alterar fornecedor na grid de produtos.");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao alterar fornecedor na grid de produtos.");
                throw;
            }
        }

        private void CancelPendReceber(string codProduto)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Cancelar pendencia a receber de produto",
                logMsg: $"Tentando cancelar pendencia a receber de produto",
                paramName: "", paramValue: "");
            try
            {
                pendReceberAntes = pageObject.GetPendReceberValue(codProduto);
                pageObject.CancelPendReceber();
                WindowsElement warn = FindElementByName("Item cancelados. Por favor, atualize as informações de pendência clicando no botão Refresh");
                Assert.IsNotNull(warn);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg:
                    $"Sucesso ao cancelar pendencia a receber de produto");

                WindowsElement exitButton = FindElementByClassAndName("Button", "OK");
                exitButton.Click();
                KeyPresser.PressKey("F10");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                string errorMsg = $"Erro ao cancelar pendencia a receber de produto";
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: errorMsg);
                throw new Exception(errorMsg);
            }
        }

        private void ValidatePendReceber(string codProduto)
        {
            string printFileName;
            string expected = "0";
            int lgsID = Global.processTest.StartStep($"Validar pendencia receber, apos refresh",
                logMsg: $"Tentando validar pendencia receber, apos refresh",
                paramName: "expected", paramValue: $"{expected}");

            pageObject.ClickOnRefreshButton();
            string actual = pageObject.GetPendReceberValue(codProduto);

            try
            {
                Assert.AreEqual(expected, actual);
                printFileName = Global.processTest.CaptureWholeScreen();
                string sucessMsg =
                    $"Sucesso ao validar pendencia receber, apos refresh: " +
                    $"Esperado: {expected}, " +
                    $"Encontrado: {actual}";
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: sucessMsg);
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                string errorMsg = $"Erro ao validar pendencia receber, apos refresh: " +
                    $"Esperado: {expected}, " +
                    $"Encontrado: {actual}";
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: errorMsg);
                throw new Exception(errorMsg);
            }
        }

        private void SelectProduct(string codProduto)
        {
            string printFileName;
            string expected = codProduto;
            int lgsID = Global.processTest.StartStep($"Selecionar produto",
                logMsg: $"Tentando selecionar produto",
                paramName: "codProduto", paramValue: $"{expected}");

            string productFound = pageObject.SelectItemOnProductGridByCode(expected);
            try
            {
                Assert.AreEqual(expected, productFound);
                printFileName = Global.processTest.CaptureWholeScreen();
                string sucessMsg =
                    $"Sucesso ao selecionar produto: " +
                    $"Esperado: {expected}, " +
                    $"Encontrado: {productFound}";
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: sucessMsg);
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                string errorMsg = $"Erro ao selecionar produto: " +
                    $"Esperado: {expected}, " +
                    $"Encontrado: {productFound}";
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: errorMsg);
                throw new Exception(errorMsg);
            }
        }

        //Test Cases
        [TestMethod, TestCategory("done"), Priority(1)]
        public void CriarLoteLoja()
        {
            int testId = 2;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                int qtdeCompra = int.Parse(Global.dataFetch.GetValue("QTDECOMPRA", queryName));
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string tipoPedido = Global.dataFetch.GetValue("TIPOPEDIDO", queryName);
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Tributação");
                    MaximizeWindow();
                    FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote);
                    ValidateQtdeCompraTotal(tipoPedido: tipoPedido, qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas);
                    GeneratePedidos(tipoLote);
                    ConfirmWindow("Consulta Lote de Compra", 2);

                    ExcelHelper.UpdateTestResult(dataFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done"), Priority(2)]
        public void CriarLoteCD()
        {
            int testId = 3;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                List<string> lojas = StringHandler.ParseStringToList(Global.dataFetch.GetValue("LOJAS", queryName));
                string divisao = Global.dataFetch.GetValue("DIVISAO", queryName);
                string cdNome = Global.dataFetch.GetValue("CDNOME", queryName);
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                int qtdeCompra = int.Parse(Global.dataFetch.GetValue("QTDECOMPRA", queryName));
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string tipoPedido = Global.dataFetch.GetValue("TIPOPEDIDO", queryName);
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    AddLojas(lojas, divisao, qtdLojas);
                    ConfirmWindow("Seleção de Empresas do Lote", 1);
                    EnableCheckbox(feature: "Incorporar Sugestão CD", paramName: "cdNome", paramValue: cdNome);
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Produtos Inativos");
                    MaximizeWindow();
                    ValidateDoubleClickOnQtdSugerida();
                    ValidateProductsGridEdit(qtdeCompra);
                    FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote);
                    ValidateQtdeCompraTotal(tipoPedido: tipoPedido, qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas);
                    GeneratePedidos(tipoLote);
                    ConfirmWindow("Consulta Lote de Compra", 2);

                    ExcelHelper.UpdateTestResult(dataFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        public void CriarLoteFLVComprador()
        {
            int testId = 4;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                List<string> lojas = StringHandler.ParseStringToList(Global.dataFetch.GetValue("LOJAS", queryName));
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                string divisao = Global.dataFetch.GetValue("DIVISAO", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    AddLojas(lojas, divisao, qtdLojas);
                    ConfirmWindow("Seleção de Empresas do Lote", 1);
                    EnableCheckbox("Restringe Empresa Loja");
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Produtos Inativos");
                    ConfirmWindow("Tributação");
                    MaximizeWindow();
                    GetLoteId();
                    CloseApp();

                    ExcelHelper.UpdateTestResult(dataFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName, false);
                }
            }
        }

        public void FinalizarLoteFLVComprador()
        {
            int testId = 8;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                int qtdeCompra = int.Parse(Global.dataFetch.GetValue("QTDECOMPRA", queryName));
                string tipoPedido = Global.dataFetch.GetValue("TIPOPEDIDO", queryName);

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    OpenLote(idLote);
                    MaximizeWindow();
                    ValidateQtdeCompraTotal(tipoPedido: tipoPedido, qtdProdutos: 1, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: 1);
                    GeneratePedidos(tipoLote);
                    ConfirmWindow("Consulta Lote de Compra", 2);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        public void PreencherLoteFLVChefeSessao()
        {
            List<int> testIds = [5, 6, 7];
            for (int i = 0; i < testIds.Count; i++)
            {
                int testId = testIds[i];
                string queryName = TestHandler.GetCurrentMethodName();
                if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
                {
                    List<string> lojas = StringHandler.ParseStringToList(Global.dataFetch.GetValue("LOJAS", queryName));
                    int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                    int qtdeCompra = int.Parse(Global.dataFetch.GetValue("QTDECOMPRA", queryName));
                    int qtdLojas = lojas.Count;
                    string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                    string tipoPedido = Global.dataFetch.GetValue("TIPOPEDIDO", queryName);

                    TestHandler.StartTest(Global.dataFetch, queryName);
                    try
                    {
                        TestHandler.DoTest(Global.dataFetch, queryName);
                        TestHandler.DefineSteps(queryName);

                        Login(Global.dataFetch, queryName);
                        OpenGerenciadorDeCompras();
                        OpenLote(idLote);
                        MaximizeWindow();
                        FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote);
                        ValidateQtdeCompraTotal(tipoPedido: tipoPedido, qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas);
                        CloseApp();

                        ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                    }
                    catch (Exception ex)
                    {
                        ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                        throw;
                    }
                    finally
                    {
                        TestHandler.EndTest(Global.dataFetch, queryName, false);
                    }
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void CriarLoteFLVCompleto()
        {
            CriarLoteFLVComprador();
            PreencherLoteFLVChefeSessao();
            FinalizarLoteFLVComprador();
        }

        [TestMethod, TestCategory("done")]
        public void CriarLoteLojaBonificacao()
        {
            int testId = 9;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                int qtdeCompra = int.Parse(Global.dataFetch.GetValue("QTDECOMPRA", queryName));
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string tipoPedido = Global.dataFetch.GetValue("TIPOPEDIDO", queryName);
                string tipoAcordo = Global.dataFetch.GetValue("TIPOACORDO", queryName);

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    FillLimiteRecebimento(hoje);
                    UpdateTipoPedido(tipoPedido);
                    UpdateTipoAcordo(tipoAcordo);
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Tributação");
                    MaximizeWindow();
                    FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote);
                    ValidateQtdeCompraTotal(tipoPedido: tipoPedido, qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas);
                    GeneratePedidos(tipoLote);
                    ConfirmWindow("Manutenção de Acordos Promocionais");

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void CriarLoteCDBonificacao()
        {
            int testId = 10;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");
                List<string> lojas = StringHandler.ParseStringToList(Global.dataFetch.GetValue("LOJAS", queryName));
                string divisao = Global.dataFetch.GetValue("DIVISAO", queryName);
                string cdNome = Global.dataFetch.GetValue("CDNOME", queryName);
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                int qtdeCompra = int.Parse(Global.dataFetch.GetValue("QTDECOMPRA", queryName));
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string tipoPedido = Global.dataFetch.GetValue("TIPOPEDIDO", queryName);
                string tipoAcordo = Global.dataFetch.GetValue("TIPOACORDO", queryName);
                int productIndex = int.Parse(Global.dataFetch.GetValue("PRODUCTINDEX", queryName));

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    AddLojas(lojas, divisao, qtdLojas);
                    ConfirmWindow("Seleção de Empresas do Lote", 1);
                    EnableCheckbox(feature: "Incorporar Sugestão CD", paramName: "cdNome", paramValue: cdNome);
                    FillLimiteRecebimento(hoje);
                    UpdateTipoPedido(tipoPedido);
                    UpdateTipoAcordo(tipoAcordo);
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Produtos Inativos");
                    MaximizeWindow();
                    ValidateDoubleClickOnQtdSugerida();
                    ValidateProductsGridEdit(qtdeCompra);
                    FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote, productIndex: productIndex);
                    ValidateQtdeCompraTotal(tipoPedido: tipoPedido, qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas);
                    GeneratePedidos(tipoLote);
                    ConfirmWindow("Manutenção de Acordos Promocionais");

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void ValidarAlteracaoPrazoPagamentoLoja()
        {
            int testId = 11;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                List<string> tipoPrazoPagamento = StringHandler.ParseStringToList(Global.dataFetch.GetValue("TIPOPRAZOPAGAMENTO", queryName));
                List<string> prazoPagamento = StringHandler.ParseStringToList(Global.dataFetch.GetValue("PRAZOPAGAMENTO", queryName));
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Tributação");
                    MaximizeWindow();
                    GoToTabParametrosDoLote();
                    UpdatePrazoPagamento(tipoPrazoPagamento[1], prazoPagamento[0]);
                    UpdatePrazoPagamento(tipoPrazoPagamento[2], prazoPagamento[1]);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void ValidarAlteracaoPrazoPagamentoCD()
        {
            int testId = 12;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                List<string> lojas = StringHandler.ParseStringToList(Global.dataFetch.GetValue("LOJAS", queryName));
                List<string> tipoPrazoPagamento = StringHandler.ParseStringToList(Global.dataFetch.GetValue("TIPOPRAZOPAGAMENTO", queryName));
                List<string> prazoPagamento = StringHandler.ParseStringToList(Global.dataFetch.GetValue("PRAZOPAGAMENTO", queryName));
                string divisao = Global.dataFetch.GetValue("DIVISAO", queryName);
                string cdNome = Global.dataFetch.GetValue("CDNOME", queryName);
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    AddLojas(lojas, divisao, qtdLojas);
                    ConfirmWindow("Seleção de Empresas do Lote", 1);
                    EnableCheckbox(feature: "Incorporar Sugestão CD", paramName: "cdNome", paramValue: cdNome);
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Produtos Inativos");
                    MaximizeWindow();
                    GoToTabParametrosDoLote();
                    UpdatePrazoPagamento(tipoPrazoPagamento[1], prazoPagamento[0]);
                    UpdatePrazoPagamento(tipoPrazoPagamento[2], prazoPagamento[1]);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void InserirItensLoteEmBrancoLoja()
        {
            int testId = 13;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 7);
                    MaximizeWindow();
                    AlterSearchParameters("Pesq. Produto");
                    SelectProducts(qtdProdutos);
                    ExecutionTimer timer = ConfirmWindowWithTimer("Pesquisa de Produtos", 1);
                    LogExecutionTime($"Inserir {qtdProdutos} itens", timer);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void InserirItensLoteEmBrancoCD()
        {
            int testId = 14;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                List<string> lojas = StringHandler.ParseStringToList(Global.dataFetch.GetValue("LOJAS", queryName));
                string divisao = Global.dataFetch.GetValue("DIVISAO", queryName);
                string cdNome = Global.dataFetch.GetValue("CDNOME", queryName);
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    AddLojas(lojas, divisao, qtdLojas);
                    ConfirmWindow("Seleção de Empresas do Lote", 1);
                    EnableCheckbox(feature: "Incorporar Sugestão CD", paramName: "cdNome", paramValue: cdNome);
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 7);
                    MaximizeWindow();
                    AlterSearchParameters("Pesq. Produto");
                    SelectProducts(qtdProdutos);
                    ExecutionTimer timer = ConfirmWindowWithTimer("Pesquisa de Produtos", 1);
                    LogExecutionTime($"Inserir {qtdProdutos} itens", timer);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void GerarLotePrazoUnicoLoja()
        {
            int testId = 17;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string tipoPrazoPagamento = Global.dataFetch.GetValue("TIPOPRAZOPAGAMENTO", queryName);
                string prazoPagamento = Global.dataFetch.GetValue("PRAZOPAGAMENTO", queryName);
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                int qtdeCompra = int.Parse(Global.dataFetch.GetValue("QTDECOMPRA", queryName));
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string tipoPedido = Global.dataFetch.GetValue("TIPOPEDIDO", queryName);
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    UpdatePrazoPagamento(tipoPrazoPagamento, prazoPagamento);
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Tributação");
                    MaximizeWindow();
                    FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote);
                    ValidateQtdeCompraTotal(tipoPedido: tipoPedido, qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas);
                    GeneratePedidos(tipoLote);
                    //GetPedidoID();
                    ConfirmWindow("Consulta Lote de Compra", 2);
                    GoToTabParametrosDoLote();
                    OpenPedido();
                    ValidatePrazoPagamento(tipoPrazoPagamento, prazoPagamento, tipoLote);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void GerarLotePrazoUnicoCD()
        {
            int testId = 18;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                List<string> lojas = StringHandler.ParseStringToList(Global.dataFetch.GetValue("LOJAS", queryName));
                string divisao = Global.dataFetch.GetValue("DIVISAO", queryName);
                string cdNome = Global.dataFetch.GetValue("CDNOME", queryName);
                string tipoPrazoPagamento = Global.dataFetch.GetValue("TIPOPRAZOPAGAMENTO", queryName);
                string prazoPagamento = Global.dataFetch.GetValue("PRAZOPAGAMENTO", queryName);
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                int qtdeCompra = int.Parse(Global.dataFetch.GetValue("QTDECOMPRA", queryName));
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string tipoPedido = Global.dataFetch.GetValue("TIPOPEDIDO", queryName);
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    AddLojas(lojas, divisao, qtdLojas);
                    ConfirmWindow("Seleção de Empresas do Lote", 1);
                    EnableCheckbox(feature: "Incorporar Sugestão CD", paramName: "cdNome", paramValue: cdNome);
                    UpdatePrazoPagamento(tipoPrazoPagamento, prazoPagamento);
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Produtos Inativos");
                    MaximizeWindow();
                    ValidateDoubleClickOnQtdSugerida();
                    ValidateProductsGridEdit(qtdeCompra);
                    FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote);
                    ValidateQtdeCompraTotal(tipoPedido: tipoPedido, qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas);
                    GeneratePedidos(tipoLote);
                    //GetPedidoID();
                    ConfirmWindow("Consulta Lote de Compra", 2);
                    GoToTabParametrosDoLote();
                    OpenPedido();
                    ValidatePrazoPagamento(tipoPrazoPagamento, prazoPagamento, tipoLote);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void ZerarQuantidadeCompraLoja()
        {
            int testId = 19;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                int qtdeCompra = int.Parse(Global.dataFetch.GetValue("QTDECOMPRA", queryName));
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string codProduto = Global.dataFetch.GetValue("CODPRODUTO", queryName);
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Tributação");
                    MaximizeWindow();
                    ClickButtonAcataSugerido();
                    ValidateAcataSugerido(tipoLote, codProduto);
                    FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote, selectProduct: false);
                    ValidateQtdeCompraUmProduto(qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas, codProduto: codProduto);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void ZerarQuantidadeCompraCD()
        {
            int testId = 20;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                List<string> lojas = StringHandler.ParseStringToList(Global.dataFetch.GetValue("LOJAS", queryName));
                string divisao = Global.dataFetch.GetValue("DIVISAO", queryName);
                string cdNome = Global.dataFetch.GetValue("CDNOME", queryName);
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                int qtdeCompra = int.Parse(Global.dataFetch.GetValue("QTDECOMPRA", queryName));
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string codProduto = Global.dataFetch.GetValue("CODPRODUTO", queryName);
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    AddLojas(lojas, divisao, qtdLojas);
                    ConfirmWindow("Seleção de Empresas do Lote", 1);
                    EnableCheckbox(feature: "Incorporar Sugestão CD", paramName: "cdNome", paramValue: cdNome);
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Produtos Inativos");
                    MaximizeWindow();
                    ValidateDoubleClickOnQtdSugerida();
                    ValidateProductsGridEdit(qtdeCompra);
                    ClickButtonAcataSugerido();
                    ValidateAcataSugerido(tipoLote, codProduto);
                    FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote, selectProduct: false);
                    ValidateQtdeCompraUmProduto(qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas, codProduto: codProduto);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void DiminuirQuantidadeCompraLoja()
        {
            int testId = 21;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string codProduto = Global.dataFetch.GetValue("CODPRODUTO", queryName);
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Tributação");
                    MaximizeWindow();
                    ClickButtonAcataSugerido();
                    ValidateAcataSugerido(tipoLote, codProduto);
                    int newValue = int.Parse(qtdeCompraEditValue) - 1;
                    FillProdutos(qtdProdutos, newValue, qtdLojas, tipoLote, selectProduct: false);
                    ValidateQtdeCompraUmProduto(qtdeCompra: newValue, tipoLote: tipoLote, qtdLojas: qtdLojas, codProduto: codProduto);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void DiminuirQuantidadeCompraCD()
        {
            int testId = 22;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                List<string> lojas = StringHandler.ParseStringToList(Global.dataFetch.GetValue("LOJAS", queryName));
                string divisao = Global.dataFetch.GetValue("DIVISAO", queryName);
                string cdNome = Global.dataFetch.GetValue("CDNOME", queryName);
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                int qtdeCompra = int.Parse(Global.dataFetch.GetValue("QTDECOMPRA", queryName));
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string codProduto = Global.dataFetch.GetValue("CODPRODUTO", queryName);
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    AddLojas(lojas, divisao, qtdLojas);
                    ConfirmWindow("Seleção de Empresas do Lote", 1);
                    EnableCheckbox(feature: "Incorporar Sugestão CD", paramName: "cdNome", paramValue: cdNome);
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Produtos Inativos");
                    MaximizeWindow();
                    ValidateDoubleClickOnQtdSugerida();
                    ValidateProductsGridEdit(qtdeCompra);
                    ClickButtonAcataSugerido();
                    ValidateAcataSugerido(tipoLote, codProduto);
                    int newValue = int.Parse(qtdeCompraEditValue) - 1;
                    FillProdutos(qtdProdutos, newValue, qtdLojas, tipoLote, selectProduct: false);
                    ValidateQtdeCompraUmProduto(qtdeCompra: newValue, tipoLote: tipoLote, qtdLojas: qtdLojas, codProduto: codProduto);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void AumentarQuantidadeCompraLoja()
        {
            int testId = 23;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string codProduto = Global.dataFetch.GetValue("CODPRODUTO", queryName);
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Tributação");
                    MaximizeWindow();
                    ClickButtonAcataSugerido();
                    ValidateAcataSugerido(tipoLote, codProduto);
                    int newValue = int.Parse(qtdeCompraEditValue) + 1;
                    FillProdutos(qtdProdutos, newValue, qtdLojas, tipoLote, selectProduct: false);
                    ValidateQtdeCompraUmProduto(qtdeCompra: newValue, tipoLote: tipoLote, qtdLojas: qtdLojas, codProduto: codProduto);
                    //newValue = newValue + 1;
                    //FillProdutos(qtdProdutos, newValue, qtdLojas, tipoLote);
                    //ValidateQtdeCompraUmProduto(qtdeCompra: newValue, tipoLote: tipoLote, qtdLojas: qtdLojas);
                    //newValue = 0;
                    //FillProdutos(qtdProdutos, newValue, qtdLojas, tipoLote);
                    //ValidateQtdeCompraUmProduto(qtdeCompra: newValue, tipoLote: tipoLote, qtdLojas: qtdLojas);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void AumentarQuantidadeCompraCD()
        {
            int testId = 24;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                List<string> lojas = StringHandler.ParseStringToList(Global.dataFetch.GetValue("LOJAS", queryName));
                string divisao = Global.dataFetch.GetValue("DIVISAO", queryName);
                string cdNome = Global.dataFetch.GetValue("CDNOME", queryName);
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                int qtdeCompra = int.Parse(Global.dataFetch.GetValue("QTDECOMPRA", queryName));
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string codProduto = Global.dataFetch.GetValue("CODPRODUTO", queryName);
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    AddLojas(lojas, divisao, qtdLojas);
                    ConfirmWindow("Seleção de Empresas do Lote", 1);
                    EnableCheckbox(feature: "Incorporar Sugestão CD", paramName: "cdNome", paramValue: cdNome);
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Produtos Inativos");
                    MaximizeWindow();
                    ValidateDoubleClickOnQtdSugerida();
                    ValidateProductsGridEdit(qtdeCompra);
                    ClickButtonAcataSugerido();
                    ValidateAcataSugerido(tipoLote, codProduto);
                    int newValue = int.Parse(qtdeCompraEditValue) + 1;
                    FillProdutos(qtdProdutos, newValue, qtdLojas, tipoLote, selectProduct: false);
                    ValidateQtdeCompraUmProduto(qtdeCompra: newValue, tipoLote: tipoLote, qtdLojas: qtdLojas, codProduto: codProduto);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void ValidarConsistenciaDataRecebimentoLoteExistenteLoja()
        {
            int testId = 25;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                int qtdeCompra = int.Parse(Global.dataFetch.GetValue("QTDECOMPRA", queryName));
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string tipoPedido = Global.dataFetch.GetValue("TIPOPEDIDO", queryName);
                string ontem = DateHandler.GetYesterdaysDate().ToString("ddMMyyyy");
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);
                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Tributação");
                    MaximizeWindow();
                    FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote);
                    ValidateQtdeCompraTotal(tipoPedido: tipoPedido, qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas);
                    GoToTabParametrosDoLote();
                    FillRecebimentoEm(ontem);
                    GoToTabProdutos();
                    SholdNotGeneratePedidos(tipoLote);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void ValidarConsistenciaDataRecebimentoLoteExistenteCD()
        {
            int testId = 26;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                int qtdeCompra = int.Parse(Global.dataFetch.GetValue("QTDECOMPRA", queryName));
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string tipoPedido = Global.dataFetch.GetValue("TIPOPEDIDO", queryName);
                string ontem = DateHandler.GetYesterdaysDate().ToString("ddMMyyyy");
                List<string> lojas = StringHandler.ParseStringToList(Global.dataFetch.GetValue("LOJAS", queryName));
                string divisao = Global.dataFetch.GetValue("DIVISAO", queryName);
                string cdNome = Global.dataFetch.GetValue("CDNOME", queryName);
                idLote = Global.dataFetch.GetValue("IDLOTE", queryName);
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    AddLojas(lojas, divisao, qtdLojas);
                    ConfirmWindow("Seleção de Empresas do Lote", 1);
                    EnableCheckbox(feature: "Incorporar Sugestão CD", paramName: "cdNome", paramValue: cdNome);
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Produtos Inativos");
                    MaximizeWindow();
                    FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote);
                    ValidateQtdeCompraTotal(tipoPedido: tipoPedido, qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas);
                    GoToTabParametrosDoLote();
                    FillRecebimentoEm(ontem);
                    GoToTabProdutos();
                    SholdNotGeneratePedidos(tipoLote);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void ValidarQuantidadeTotalGerarPedidoLoja()
        {
            int testId = 27;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                int qtdeCompra = int.Parse(Global.dataFetch.GetValue("QTDECOMPRA", queryName));
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string tipoPedido = Global.dataFetch.GetValue("TIPOPEDIDO", queryName);
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Tributação");
                    MaximizeWindow();
                    FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote);
                    ValidateQtdeCompraTotal(tipoPedido: tipoPedido, qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas);
                    ValidateQtdeTotalGerarPedido(tipoLote);
                    ConfirmWindow("Opções de geração do(s) pedido(s)");
                    WindowsElement warn = FindElementByXPathPartialName("com sucesso.");
                    Assert.IsNotNull(warn);
                    idPedido = pageObject.GetIdPedido();
                    KeyPresser.PressKey("RETURN");
                    ConfirmWindow("Consulta Lote de Compra");

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void ValidarQuantidadeTotalGerarPedidoCD()
        {
            int testId = 28;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                List<string> lojas = StringHandler.ParseStringToList(Global.dataFetch.GetValue("LOJAS", queryName));
                string divisao = Global.dataFetch.GetValue("DIVISAO", queryName);
                string cdNome = Global.dataFetch.GetValue("CDNOME", queryName);
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                int qtdeCompra = int.Parse(Global.dataFetch.GetValue("QTDECOMPRA", queryName));
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string tipoPedido = Global.dataFetch.GetValue("TIPOPEDIDO", queryName);
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    AddLojas(lojas, divisao, qtdLojas);
                    ConfirmWindow("Seleção de Empresas do Lote", 1);
                    EnableCheckbox(feature: "Incorporar Sugestão CD", paramName: "cdNome", paramValue: cdNome);
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Produtos Inativos");
                    MaximizeWindow();
                    FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote);
                    ValidateQtdeCompraTotal(tipoPedido: tipoPedido, qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas);
                    ValidateQtdeTotalGerarPedido(tipoLote);
                    ConfirmWindow("Opções de geração do(s) pedido(s)");
                    WindowsElement warn = FindElementByXPathPartialName("com sucesso.");
                    Assert.IsNotNull(warn);
                    idPedido = pageObject.GetIdPedido();
                    KeyPresser.PressKey("RETURN");
                    ConfirmWindow("Consulta Lote de Compra");

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void ValidarEmbalagemGridsLoja()
        {
            int testId = 29;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                List<string> lojas = StringHandler.ParseStringToList(Global.dataFetch.GetValue("LOJAS", queryName));
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                string divisao = Global.dataFetch.GetValue("DIVISAO", queryName);
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    AddLojas(lojas, divisao, qtdLojas);
                    ConfirmWindow("Seleção de Empresas do Lote", 1);
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Produtos Inativos");
                    ConfirmWindow("Tributação");
                    MaximizeWindow();
                    ValidateEmbalagens(tipoLote);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void ValidarEmbalagemGridsCD()
        {
            int testId = 30;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                List<string> lojas = StringHandler.ParseStringToList(Global.dataFetch.GetValue("LOJAS", queryName));
                string divisao = Global.dataFetch.GetValue("DIVISAO", queryName);
                string cdNome = Global.dataFetch.GetValue("CDNOME", queryName);
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    AddLojas(lojas, divisao, qtdLojas);
                    ConfirmWindow("Seleção de Empresas do Lote", 1);
                    EnableCheckbox(feature: "Incorporar Sugestão CD", paramName: "cdNome", paramValue: cdNome);
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Produtos Inativos");
                    MaximizeWindow();
                    ValidateEmbalagens(tipoLote);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void ValidarPrazoVencimentoFixoGeracaoPedidoLoja()
        {
            int testId = 31;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string tipoPrazoPagamento = Global.dataFetch.GetValue("TIPOPRAZOPAGAMENTO", queryName);
                string prazoPagamento = Global.dataFetch.GetValue("PRAZOPAGAMENTO", queryName);
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                int qtdeCompra = int.Parse(Global.dataFetch.GetValue("QTDECOMPRA", queryName));
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string tipoPedido = Global.dataFetch.GetValue("TIPOPEDIDO", queryName);
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    UpdatePrazoPagamento(tipoPrazoPagamento, prazoPagamento);
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Tributação");
                    MaximizeWindow();
                    FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote);
                    ValidateQtdeCompraTotal(tipoPedido: tipoPedido, qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas);
                    GeneratePedidos(tipoLote);
                    //GetPedidoID();
                    ConfirmWindow("Consulta Lote de Compra", 2);
                    GoToTabParametrosDoLote();
                    OpenPedido();
                    ValidatePrazoPagamento(tipoPrazoPagamento, prazoPagamento, tipoLote);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void ValidarPrazoVencimentoFixoGeracaoPedidoCD()
        {
            int testId = 32;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                List<string> tipoPrazoPagamento = StringHandler.ParseStringToList(Global.dataFetch.GetValue("TIPOPRAZOPAGAMENTO", queryName));
                string prazoPagamento = Global.dataFetch.GetValue("PRAZOPAGAMENTO", queryName);
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                int qtdeCompra = int.Parse(Global.dataFetch.GetValue("QTDECOMPRA", queryName));
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string tipoPedido = Global.dataFetch.GetValue("TIPOPEDIDO", queryName);
                List<string> lojas = StringHandler.ParseStringToList(Global.dataFetch.GetValue("LOJAS", queryName));
                string divisao = Global.dataFetch.GetValue("DIVISAO", queryName);
                string cdNome = Global.dataFetch.GetValue("CDNOME", queryName);
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    AddLojas(lojas, divisao, qtdLojas);
                    ConfirmWindow("Seleção de Empresas do Lote", 1);
                    EnableCheckbox(feature: "Incorporar Sugestão CD", paramName: "cdNome", paramValue: cdNome);
                    UpdatePrazoPagamento(tipoPrazoPagamento[0], prazoPagamento);
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Produtos Inativos");
                    MaximizeWindow();
                    ValidateDoubleClickOnQtdSugerida();
                    ValidateProductsGridEdit(qtdeCompra);
                    FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote);
                    ValidateQtdeCompraTotal(tipoPedido: tipoPedido, qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas);
                    GeneratePedidos(tipoLote);
                    //GetPedidoID();
                    ConfirmWindow("Consulta Lote de Compra", 2);
                    GoToTabParametrosDoLote();
                    OpenPedido();
                    ValidatePrazoPagamento(tipoPrazoPagamento[0], prazoPagamento, tipoLote);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void ValidarEstoqueDiasLoja()
        {
            int testId = 33;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                List<string> lojas = StringHandler.ParseStringToList(Global.dataFetch.GetValue("LOJAS", queryName));
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Tributação");
                    MaximizeWindow();
                    ValidateEstoqueDias("Grid Produtos", tipoLote, lojas);
                    ValidateEstoqueDias("Consulta Produtos", tipoLote, lojas);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void ValidarEstoqueDiasCD()
        {
            int testId = 34;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                List<string> lojas = StringHandler.ParseStringToList(Global.dataFetch.GetValue("LOJAS", queryName));
                string divisao = Global.dataFetch.GetValue("DIVISAO", queryName);
                string cdNome = Global.dataFetch.GetValue("CDNOME", queryName);
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    AddLojas(lojas, divisao, qtdLojas);
                    ConfirmWindow("Seleção de Empresas do Lote", 1);
                    EnableCheckbox(feature: "Incorporar Sugestão CD", paramName: "cdNome", paramValue: cdNome);
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Produtos Inativos");
                    MaximizeWindow();
                    ValidateEstoqueDias("Grid Produtos", tipoLote, lojas);
                    ValidateEstoqueDias("Consulta Produtos", tipoLote, lojas);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void GerarLoteComItemDuplaProcedenciaLoja()
        {
            int testId = 35;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string tipoPedido = Global.dataFetch.GetValue("TIPOPEDIDO", queryName);
                string codProduto = Global.dataFetch.GetValue("CODPRODUTO", queryName);
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                int qtdeCompra = int.Parse(Global.dataFetch.GetValue("QTDECOMPRA", queryName));
                int productIndex = int.Parse(Global.dataFetch.GetValue("PRODUCTINDEX", queryName));
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    //ConfirmWindow("Produtos Inativos");
                    MaximizeWindow();
                    FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote, productIndex: productIndex);
                    ValidateQtdeCompraTotal(tipoPedido: tipoPedido, qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas);
                    ValidateTipoAbastecimento(productIndex, codProduto, tipoLote);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void GerarLoteComItemDuplaProcedenciaCD()
        {
            int testId = 36;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                List<string> lojas = StringHandler.ParseStringToList(Global.dataFetch.GetValue("LOJAS", queryName));
                string divisao = Global.dataFetch.GetValue("DIVISAO", queryName);
                string cdNome = Global.dataFetch.GetValue("CDNOME", queryName);
                string tipoPedido = Global.dataFetch.GetValue("TIPOPEDIDO", queryName);
                string codProduto = Global.dataFetch.GetValue("CODPRODUTO", queryName);
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                int qtdeCompra = int.Parse(Global.dataFetch.GetValue("QTDECOMPRA", queryName));
                int productIndex = int.Parse(Global.dataFetch.GetValue("PRODUCTINDEX", queryName));
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    AddLojas(lojas, divisao, qtdLojas);
                    ConfirmWindow("Seleção de Empresas do Lote", 1);
                    EnableCheckbox(feature: "Incorporar Sugestão CD", paramName: "cdNome", paramValue: cdNome);
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Produtos Inativos");
                    MaximizeWindow();
                    FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote, productIndex: productIndex);
                    ValidateQtdeCompraTotal(tipoPedido: tipoPedido, qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas);
                    ValidateTipoAbastecimento(productIndex, codProduto, tipoLote);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void ValidarBloqueioSessaoSimulaneaMesmoLoteCompleto()
        {
            ValidarBloqueioSessaoSimulaneaMesmoLoteSessaoUm();
            ValidarBloqueioSessaoSimulaneaMesmoLoteSessaoDois();
        }

        public void ValidarBloqueioSessaoSimulaneaMesmoLoteSessaoUm()
        {
            int testId = 40;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Tributação");
                    MaximizeWindow();
                    GetLoteId();

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName, false);
                }
            }
        }
           
        public void ValidarBloqueioSessaoSimulaneaMesmoLoteSessaoDois()
        {
            int testId = 41;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    OpenLote(idLote);
                    CloseApp();
                    SetAppSession();
                    CloseApp();

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void ValidarPendenciaReceberAposCancelarEClicarRefresh()
        {
            int testId = 42;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string codFornecedor = Global.dataFetch.GetValue("CODFORNECEDOR", queryName);
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int productIndex = int.Parse(Global.dataFetch.GetValue("PRODUCTINDEX", queryName));
                string codProduto = Global.dataFetch.GetValue("CODPRODUTO", queryName);
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName);
                    OpenGerenciadorDeCompras();
                    SetSupplier(codFornecedor);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    ClickOnIncluirLote();
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Tributação");
                    MaximizeWindow();
                    CancelPendReceber(codProduto);
                    ValidatePendReceber(codProduto);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        [TestMethod, TestCategory("done")]
        public void GerarLoteComDoisFornecedoresFLVCompleto()
        {
            GerarLoteComDoisFornecedoresFLVCompradorCriarCapa();
            GerarLoteComDoisFornecedoresFLVChefeSessaoPreencher();
            GerarLoteComDoisFornecedoresCompradorFinalizar();
        }

        public void GerarLoteComDoisFornecedoresFLVCompradorCriarCapa()
        {
            int testId = 37;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string tipoPedido = Global.dataFetch.GetValue("TIPOPEDIDO", queryName);
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                int qtdeCompra = int.Parse(Global.dataFetch.GetValue("QTDECOMPRA", queryName));
                int productIndex = int.Parse(Global.dataFetch.GetValue("PRODUCTINDEX", queryName));
                List<string> suppliers = StringHandler.ParseStringToList(Global.dataFetch.GetValue("CODFORNECEDOR", queryName));
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    //Criar capa lote | Comprador
                    Login(Global.dataFetch, queryName, 0);
                    OpenGerenciadorDeCompras();
                    SetSupplier(suppliers[0], true);
                    SetSupplier(suppliers[1], false);
                    SelectCategoria(categoria);
                    FillAbastecimentoDias(diasAbastecimento);
                    FillRecebimentoEm(hoje);
                    EnableCheckbox("Sugestão Compras");
                    EnableCheckbox("Restringe Empresa Loja");
                    ClickOnIncluirLote();
                    ValidateSuppliers(suppliers, "Filtros para Seleção de Produtos");
                    ConfirmWindow("Filtros para Seleção de Produtos", 6);
                    ConfirmWindow("Tributação");
                    MaximizeWindow();
                    GetLoteId();
                    ValidateSuppliers(suppliers, "Grid de Produtos");

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        public void GerarLoteComDoisFornecedoresFLVChefeSessaoPreencher()
        {
            int testId = 38;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string tipoPedido = Global.dataFetch.GetValue("TIPOPEDIDO", queryName);
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                int qtdeCompra = int.Parse(Global.dataFetch.GetValue("QTDECOMPRA", queryName));
                int productIndex = int.Parse(Global.dataFetch.GetValue("PRODUCTINDEX", queryName));
                List<string> suppliers = StringHandler.ParseStringToList(Global.dataFetch.GetValue("CODFORNECEDOR", queryName));
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    //Preencher lote | Chefe de sessão
                    Login(Global.dataFetch, queryName, 1);
                    OpenGerenciadorDeCompras();
                    OpenLote(idLote);
                    MaximizeWindow();
                    FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote, productIndex: productIndex);
                    ValidateQtdeCompraTotal(tipoPedido: tipoPedido, qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }

        public void GerarLoteComDoisFornecedoresCompradorFinalizar()
        {
            int testId = 39;
            string queryName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testId, queryName) == "")
            {
                string categoria = Global.dataFetch.GetValue("CATEGORIA", queryName);
                string diasAbastecimento = Global.dataFetch.GetValue("DIASABASTECIMENTO", queryName);
                int qtdLojas = int.Parse(Global.dataFetch.GetValue("QTDLOJAS", queryName));
                string tipoLote = Global.dataFetch.GetValue("TIPOLOTE", queryName);
                string tipoPedido = Global.dataFetch.GetValue("TIPOPEDIDO", queryName);
                int qtdProdutos = int.Parse(Global.dataFetch.GetValue("QTDPRODUTOS", queryName));
                int qtdeCompra = int.Parse(Global.dataFetch.GetValue("QTDECOMPRA", queryName));
                int productIndex = int.Parse(Global.dataFetch.GetValue("PRODUCTINDEX", queryName));
                List<string> suppliers = StringHandler.ParseStringToList(Global.dataFetch.GetValue("CODFORNECEDOR", queryName));
                string hoje = DateHandler.GetTodaysDate().ToString("ddMMyyyy");

                TestHandler.StartTest(Global.dataFetch, queryName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, queryName);
                    TestHandler.DefineSteps(queryName);

                    Login(Global.dataFetch, queryName, 0);
                    OpenGerenciadorDeCompras();
                    OpenLote(idLote);
                    MaximizeWindow();
                    int qtdeCompraFirstSupplier = qtdeCompra / 2 - 1;
                    FillProdutos(qtdProdutos, qtdeCompraFirstSupplier, qtdLojas, tipoLote, productIndex: productIndex);
                    ChangeSupplier();
                    int qtdeCompraSecondSupplier = qtdeCompra / 2 + 1;
                    FillProdutos(qtdProdutos, qtdeCompraSecondSupplier, qtdLojas, tipoLote, productIndex: productIndex);
                    GeneratePedidos(tipoLote);
                    ValidateSuppliers(suppliers, "Consulta Lote de Compra");
                    ConfirmWindow("Consulta Lote de Compra", 2);

                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataSetFilePath, sheet, testId, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, queryName);
                }
            }
        }
    }
}