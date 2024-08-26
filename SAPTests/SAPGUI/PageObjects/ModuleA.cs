using SAPTests.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using System.Collections.ObjectModel;
using static SAPTests.Helpers.ElementHandler;

namespace SAPTests.PageObjects
{
    public class ModuleAPageObject
    {
        private WindowsElement window;
        private AppiumWebElement pane;

        public ModuleAPageObject()
        {
            window = FindElementByName("Gerenciador de Compras");
            pane = window.FindElementByClassName("Centura:Form");
        }

        public void ExitWindow(string windowName, int buttonIndex = 0)
        {
            switch (windowName)
            {
                case "Seleção de Empresas do Lote":
                case "Filtros para Seleção de Produtos":
                case "Pesquisa de Lotes de Compra":
                case "Atenção":
                    ConfirmWindow(windowName, buttonIndex);
                    break;
                case "Produtos Inativos":
                case "Tributação":
                    ConfirmWindow(windowName, 0, 2000);
                    FindElementByName("Restringe Empresa Loja");
                    break;
                case "Manutenção de Acordos Promocionais":
                    ConfirmWindow(windowName, 0, 3000);
                    break;
                case "Opções de geração do(s) pedido(s)":
                    ConfirmWindow(windowName, 1); // Atualiza pedido
                    ConfirmWindow(windowName, 0); // Confirma validação
                    ExitWindow("Atenção");
                    break;
                case "Consulta Lote de Compra":
                    ConfirmWindow(windowName, buttonIndex, 3000);
                    break;
                default:
                    throw new Exception($"Window {windowName} not found.");
            }
        }

        public ExecutionTimer ExitWindowWithTimer(ExecutionTimer timer, string windowName, int buttonIndex = 0)
        {
            switch (windowName)
            {
                case "Pesquisa de Produtos":
                    ConfirmWindow(windowName, buttonIndex);
                    timer.Start();
                    return timer;
                default:
                    throw new Exception($"Window {windowName} not found.");
            }
        }

        public void FillFornecedor(string codFornecedor, bool setMain = true)
        {
            string buttonName = "Fornecedor";
            WindowsElement fornecedorButton = FindElementByClassAndName("Button", buttonName);
            fornecedorButton.Click();

            BoundingRectangle sequenciaVsField = new BoundingRectangle(220, 158, 286, 178);
            MouseHandler.Click(sequenciaVsField);
            WinAppDriver.FillField(codFornecedor);
            KeyPresser.PressKey("F8");
            MouseHandler.DoubleClick(125, 255);

            if (setMain == true)
            {
                BoundingRectangle principalCheckBox = new BoundingRectangle(579, 462, 592, 475);
                MouseHandler.Click(principalCheckBox);
            }
        }

        public void OpenSelecaoDeLojas()
        {
            AppiumWebElement empresasButton = pane.FindElementByName("Empresas");
            empresasButton.Click();
        }

        public void RemoveDivisoes()
        {
            BoundingRectangle removeDivisoesButton = new BoundingRectangle(247, 270, 271, 295);
            MouseHandler.Click(removeDivisoesButton);
        }

        public void AddDivisao(string divisao)
        {
            var divisaoListItem = FindElementByName(divisao);
            divisaoListItem.Click();

            BoundingRectangle addDivisaoButton = new BoundingRectangle(247, 195, 271, 220);
            MouseHandler.Click(addDivisaoButton);
        }

        public void RemoveLojas()
        {
            BoundingRectangle removeLojasButton = new BoundingRectangle(247, 497, 271, 522);
            MouseHandler.Click(removeLojasButton);
        }

        public void AddLojasPorQuantidade(int qtdLojas)
        {
            AppiumWebElement empresasButton = pane.FindElementByName("Empresas");
            empresasButton.Click();

            string windowName = "Seleção de Empresas do Lote";
            try
            {
                FindElementByName(windowName);

                // Adds first X items from the list of Lojas
                BoundingRectangle empresasFirstItem = new BoundingRectangle(80, 391, 207, 404);
                for (int i = 0; i < qtdLojas; i++)
                {
                    MouseHandler.DoubleClickOn(empresasFirstItem);
                }
            }
            catch
            {
                throw new Exception($"Erro ao tentar adicionar ${qtdLojas} lojas na janela {windowName}");
            }
        }

        public void AddLojasPorNome(List<string> lojas)
        {
            BoundingRectangle addLojasButton = new BoundingRectangle(247, 422, 271, 447);
            string windowName = "Seleção de Empresas do Lote";
            try
            {
                lojas.ForEach(loja =>
                {
                    WindowsElement lojaListItem = FindElementByName(loja);
                    lojaListItem.Click();
                });
                MouseHandler.Click(addLojasButton);
            }
            catch
            {
                throw new Exception($"Erro ao tentar adicionar {lojas.Count} lojas na janela {windowName}");
            }
        }

        public void SelectCategoria(string categoria)
        {
            BoundingRectangle categoriaComboBox = new BoundingRectangle(377, 201, 563, 222);
            MouseHandler.Click(categoriaComboBox);
            WinAppDriver.FillField(categoria);
            KeyPresser.PressKey("RETURN");
        }

        public void FillAbastecimentoDias(string dias)
        {
            BoundingRectangle abastecimentoDiasField = new BoundingRectangle(431, 338, 460, 357);
            MouseHandler.DoubleClickOn(abastecimentoDiasField);
            WinAppDriver.FillField(dias);
        }

        public void EnableCheckbox(string feature)
        {
            switch (feature)
            {
                case "Sugestão Compras":
                    string className = "Centura:GPCheck";
                    //CD
                    List<string> checkboxNames = [
                        "Considera Saldo Pend Receber",
                        "Considera Saldo Pend Expedir",
                        "Considera Qtde a Comprar Lote",
                    ];
                    foreach (string name in checkboxNames)
                    {
                        WindowsElement checkbox = FindElementByClassAndName(className, name);
                        if (VerifyCheckBoxIsOn(checkbox))
                            continue;
                        else
                            checkbox.Click();
                    }

                    //Loja
                    List<string> checkBoxIds = ["4217", "4218"];
                    checkBoxIds.ForEach(id =>
                    {
                        WindowsElement checkBox = FindElementByAccessibilityId(id);
                        if (!VerifyCheckBoxIsOn(checkBox))
                            checkBox.Click();
                    });
                    break;
                case "Restringe Empresa Loja":
                    WindowsElement checkboxRestringe = FindElementByName(feature);
                    if (!VerifyCheckBoxIsOn(checkboxRestringe))
                        checkboxRestringe.Click();
                    break;
                case "Incorporar Sugestão CD":
                    string checkboxName = "Incorporar Sugestão";
                    WindowsElement checkboxIncorporar = FindElementByName(checkboxName);
                    if (!VerifyCheckBoxIsOn(checkboxIncorporar))
                        checkboxIncorporar.Click();
                    break;
                default:
                    throw new ArgumentException($"Unsupported feature: {feature}");
            }
        }

        public void SetCD(string cdNome)
        {
            BoundingRectangle cdPrincipalCombobox = new BoundingRectangle(359, 645, 551, 666);
            MouseHandler.Click(cdPrincipalCombobox);
            WindowsElement chosenCd = FindElementByName(cdNome);
            chosenCd.Click();
        }

        public void ClickOnIncluirLote()
        {
            ReadOnlyCollection<WindowsElement> buttons = FindElementsByClassName("Button");
            WindowsElement incluir = buttons[10];
            incluir.Click();
        }

        public void FillQtdeCompra(int qtdProdutos, int qtdeCompra, string tipoLote, int productIndex = 1, bool selectProduct = true)
        {
            string gridClassName = "Centura:ChildTable";
            WindowsElement grid = FindElementByClassName(gridClassName);
            if (grid == null)
                throw new Exception("Grid not found");
            switch (tipoLote)
            {
                case "loja":
                    {
                        if (selectProduct)
                        {
                            if (productIndex != 1)
                            {
                                SelectItemOnProductGridByIndex(productIndex);
                                WinAppDriver.WaitSeconds(8);
                            }
                        }
                        BoundingRectangle qtdeCompraFirstLoja = new BoundingRectangle(461, 422, 513, 435);
                        MouseHandler.Click(qtdeCompraFirstLoja);
                        WinAppDriver.WaitSeconds(3);
                        for (int i = 0; i < qtdProdutos; i++)
                        {
                            WinAppDriver.FillField(qtdeCompra.ToString());
                            KeyPresser.PressKey("RETURN");
                        }
                    }
                    break;
                case "cd":
                    {
                        if (selectProduct)
                        {
                            SelectItemOnProductGridByIndex(2);
                        }
                        WinAppDriver.WaitSeconds(3);
                        BoundingRectangle qtdeCompraCD = new BoundingRectangle(540, 404, 592, 417);
                        MouseHandler.Click(qtdeCompraCD);
                        WinAppDriver.WaitSeconds(3);
                        for (int i = 0; i < qtdProdutos; i++)
                        {
                            WinAppDriver.FillField(qtdeCompra.ToString());
                            KeyPresser.PressKey("RETURN");
                        }
                        WindowsElement warning = FindElementByName("Atenção");
                        if (warning != null)
                        {
                            ExitWindow("Atenção");
                        }
                        break;

                    }
                case "flv":
                    {
                        BoundingRectangle qtdeCompraFirstLoja = new BoundingRectangle(461, 422, 513, 435);
                        MouseHandler.Click(qtdeCompraFirstLoja);
                        for (int i = 0; i < qtdProdutos; i++)
                        {
                            WinAppDriver.FillField(qtdeCompra.ToString());
                            KeyPresser.PressKey("RETURN");
                        }
                    }
                    break;
            }
        }

        public bool ValidateQtdeComprasValue(int total, int qtdProdutos, int qtdeCompra, int qtdLojas, string tipoLote)
        {
            if (tipoLote == "cd")
            {
                return total == qtdProdutos * qtdeCompra;
            }
            else
            {
                return total == qtdProdutos * qtdeCompra * qtdLojas;
            }
        }

        public void ClickGerarPedidos()
        {
            AppiumWebElement geraPedidosButton = FindElementByName("Gera Pedidos");
            geraPedidosButton.Click();
        }

        public void UpdateLote()
        {
            BoundingRectangle refreshButton = new BoundingRectangle(179, 78, 207, 106);
            MouseHandler.Click(refreshButton);
            ConfirmWindow("Atenção");
        }

        public void OpenLote(string idLote)
        {
            BoundingRectangle pesquisarButton = new BoundingRectangle(151, 78, 179, 106);
            MouseHandler.Click(pesquisarButton);

            BoundingRectangle sequenciaqLoteField = new BoundingRectangle(114, 137, 165, 157);
            MouseHandler.Click(sequenciaqLoteField);
            WinAppDriver.FillField(idLote);

            string windowName = "Pesquisa de Lotes de Compra";
            ExitWindow(windowName, 2);
        }

        public int GetValueFromEdit()
        {
            string className = "Edit";
            ReadOnlyCollection<WindowsElement> editElements = FindElementsByClassName(className);
            //System.Console.WriteLine($@"editElements found: {editElements.Count}");
            WindowsElement edit = editElements[9];
            string editValue = edit.GetAttribute("Value.Value");
            return int.Parse(editValue);
        }

        public int GetQtdeCompraTotal()
        {
            WinAppDriver.WaitSeconds(3);
            ResetColumnLeft();
            ClickOnQtdeCompraTotalProductGrid();
            return GetValueFromEdit();
        }

        public void SelectProductGrid()
        {
            WindowsElement productGrid = FindElementByClassName("Centura:Pict");
            productGrid.Click();
        }

        public void ResetColumnLeft()
        {
            SelectProductGrid();
            BoundingRectangle columnLeftButton = new BoundingRectangle(9, 337, 26, 354);
            MouseHandler.RightClick(columnLeftButton);
            KeyPresser.PressKey("DOWN");
            KeyPresser.PressKey("DOWN");
            KeyPresser.PressKey("RETURN");
        }

        public int GetQtdeCompraUmProduto(string codProduto)
        {
            WinAppDriver.WaitSeconds(5);
            SelectItemOnProductGridByCode(codProduto, clickColumn: true, X: 1085);//, Y: 270);
            //ClickOnQtdeCompraFirstItemWithQtdeCompraFilled();
            return GetValueFromEdit();
        }

        public void ClickOnQtdeCompraTotalProductGrid()
        {
            MouseHandler.Click(1080, 327);
        }

        public void ClickOnQtdeCompraFirstItemWithQtdeCompraFilled()
        {
            MouseHandler.Click(1085, 270);
        }

        public string GetIdLote()
        {
            ReadOnlyCollection<WindowsElement> editList = FindElementsByClassName("Edit");
            return editList[5].GetAttribute("Value.Value");
        }

        public void DoubleClickOnQtdSugerida()
        {
            WinAppDriver.WaitSeconds(7);
            MouseHandler.DoubleClick(965, 410);
        }

        public void FillProductsGrid(int qtdeCompra)
        {
            BoundingRectangle qtdeCompraSegundoProduto = new BoundingRectangle(1047, 156, 1099, 169);
            MouseHandler.Click(qtdeCompraSegundoProduto);
            WinAppDriver.WaitSeconds(3);
            WinAppDriver.FillField(qtdeCompra.ToString());
        }

        public void UpdateTipoPedido(string tipoPedido)
        {
            BoundingRectangle tipoPedidoComboBox = new BoundingRectangle(299, 203, 316, 220);
            MouseHandler.Click(tipoPedidoComboBox);
            WindowsElement listItem = FindElementByName(tipoPedido);
            listItem.Click();
            KeyPresser.PressKey("RETURN");
        }

        public void UpdateTipoAcordo(string tipoAcordo)
        {
            switch (tipoAcordo)
            {
                case "DIFERENCA DE PRECO":
                    WindowsElement tipoAcordoButton = FindElementByName("Tipo Acordo");
                    tipoAcordoButton.Click();

                    int qtdClicks = 7;
                    for (int i = 0; i < qtdClicks; i++)
                    {
                        KeyPresser.PressKey("DOWN");
                    }
                    KeyPresser.PressKey("RETURN");
                    break;
                default:
                    throw new Exception($"{tipoAcordo} não encontrado.");
            }
        }

        public void FillLimiteRecebimento(string data)
        {
            BoundingRectangle limiteRecebimentoEdit = new BoundingRectangle(125, 506, 195, 527);
            MouseHandler.Click(limiteRecebimentoEdit);
            WinAppDriver.FillField(data);
        }

        public void FillRecebimentoEm(string data)
        {
            ReadOnlyCollection<WindowsElement> edits = FindElementsByClassName("Edit");
            WindowsElement recebimentoEm = edits[12];
            recebimentoEm.Click();
            KeyPresser.PressKeys("HOME");
            int qtdClicks = 8;
            for (int i = 0; i < qtdClicks; i++)
            {
                KeyPresser.PressKey("DELETE");
            }
            WinAppDriver.FillField(data);
        }

        public void GoToTabParametrosDoLote()
        {
            MouseHandler.Click(25, 90);
        }

        public void GoToTabProdutos()
        {
            MouseHandler.Click(170, 90);
        }

        public bool ValidatePrazoPagamento(string prazoPagamento, string tipoLote)
        {
            if (tipoLote == "cd")
            {
                KeyPresser.PressKey("RETURN");
            }
            ReadOnlyCollection<WindowsElement> edits = FindElementsByClassName("Edit");
            string prazoPagamentoEditValue = edits[8].GetAttribute("Value.Value");
            return prazoPagamento == prazoPagamentoEditValue;
        }

        public bool ValidateVencimentoFixo(string prazoPagamento, string tipoLote)
        {
            if (tipoLote == "cd")
            {
                KeyPresser.PressKey("RETURN");
            }
            ReadOnlyCollection<WindowsElement> edits = FindElementsByClassName("Edit");
            string vencimentoFixoEditValue = edits[10].GetAttribute("Value.Value");
            return prazoPagamento == vencimentoFixoEditValue;
        }

        public void UpdatePrazoPagamento(string tipoPrazoPagamento, string prazoPagamento)
        {
            switch (tipoPrazoPagamento)
            {
                case "Prazo Único":
                    WindowsElement radioButtonPrazoPagamento = FindElementByName(tipoPrazoPagamento);
                    radioButtonPrazoPagamento.Click();
                    KeyPresser.PressKey("TAB");
                    WinAppDriver.FillField(prazoPagamento);
                    break;
                case "Prazo Fixo":
                    radioButtonPrazoPagamento = FindElementByName(tipoPrazoPagamento);
                    radioButtonPrazoPagamento.Click();
                    KeyPresser.PressKey("TAB");
                    WinAppDriver.FillField(prazoPagamento);
                    break;
                default:
                    throw new Exception($"{tipoPrazoPagamento} não encontrado.");
            }
        }

        public void ClickButtonAcataSugerido()
        {
            WindowsElement button = FindElementByName("Acata Sugerido");
            button.Click();
            ExitWindow("Atenção");
        }

        public void SelectFirstItemWithQtdeCompraFilled()
        {
            WindowsElement button = FindElementByName("OK");
            button.Click();
            ClickOnProductGridLineDown(11);
            MouseHandler.Click(1080, 290);
        }

        public string GetEditValueAfterAcataSugerido()
        {
            return GetValueFromEdit().ToString();
        }

        public string GetQtdeSugeridaValue(string tipoLote)
        {
            int qtdCarac;
            string qtdeSugeridaValue;
            if (tipoLote == "loja")
            {
                qtdeSugeridaValue = OCRScanner.ExtractText(853, 422, 54, 13);
            }
            else
            {
                qtdeSugeridaValue = OCRScanner.ExtractText(934, 404, 52, 13);
            }
            return StringHandler.RemoveCommasAndTrailingZeros(qtdeSugeridaValue);
        }

        public void AlterSearchParameters(string searchParameter)
        {
            BoundingRectangle searchComboBox = new BoundingRectangle(157, 694, 174, 711);
            MouseHandler.Click(searchComboBox);
            WindowsElement listItem = FindElementByName(searchParameter);
            listItem.Click();
        }

        public void SelectProducts(int qtdProdutos)
        {
            ReadOnlyCollection<WindowsElement> buttons = FindElementsByClassName("Button");
            WindowsElement newLineBtn = buttons[3];
            newLineBtn.Click();

            buttons = FindElementsByClassName("Button");
            WindowsElement searchBtn = buttons[2];
            searchBtn.Click();

            WindowsElement middleElement = FindElementByName("Open");
            middleElement.Click();

            KeyPresser.PressKeys("TAB");
            KeyPresser.PressKeys("DOWN");
            KeyPresser.PressKeys("UP");

            for (int i = 1; i < qtdProdutos; i++)
            {
                KeyPresser.PressKeys("SHIFT", "DOWN");
            }
            WinAppDriver.WaitSeconds(3);
        }

        public string LogExecutionTime(string scenario, ExecutionTimer timer)
        {
            TimeSpan executionTime;
            switch (scenario)
            {
                case "Inserir 10 itens":
                    string expectedValue = "10";
                    string actualValue = GetQtdProdutosShownProductGrid();

                    while (actualValue != expectedValue)
                    {
                        WinAppDriver.WaitSeconds(5);
                        actualValue = GetQtdProdutosShownProductGrid();
                    }

                    executionTime = timer.Stop();
                    int minutes = executionTime.Minutes; // Minutes portion
                    int seconds = executionTime.Seconds; // Seconds portion
                    return $"{minutes}min e {seconds}seg";
                default:
                    throw new Exception($"{scenario} not found.");
            }
        }

        private string GetQtdProdutosShownProductGrid()
        {
            string actualValue = OCRScanner.ExtractText(82, 319, 105, 16);
            return actualValue.Substring(0, 2).Trim();
        }

        public string GetIdPedido()
        {
            WindowsElement dialog = FindElementByXPathPartialName("com sucesso.");
            string dialogText = dialog.Text;
            string pattern = @":\s*(.*)";
            string pedidoId = ExtractTextWithPattern(dialogText, pattern);
            return pedidoId;
        }

        public void OpenPedido(string idPedido)
        {
            WindowsElement pedidoGerado = FindElementByXPathPartialName(idPedido);
            MouseHandler.DoubleClick(pedidoGerado);
        }

        public int GetQtdTotalOpcoesGeracaoPedido(string tipoLote)
        {
            int qtdCarac;
            (int screenWidth, int screenHeight) = ScreenHandler.GetScreenResolution();
            string QtdTotalOpcoesGeracaoPedido;
            if (screenWidth == 1366 && screenHeight == 768)
            {
                QtdTotalOpcoesGeracaoPedido = OCRScanner.TextToDouble(525, 384, 56, 14).ToString();
            }
            else
            {
                QtdTotalOpcoesGeracaoPedido = OCRScanner.TextToDouble(820, 540, 37, 16).ToString();
            }
            if (tipoLote == "cd")
            {
                qtdCarac = 2;
            }
            else
            {
                qtdCarac = 1;
            }
            QtdTotalOpcoesGeracaoPedido = StringHandler.RemoveCommasAndTrailingZeros(QtdTotalOpcoesGeracaoPedido.Substring(0, qtdCarac));
            return int.Parse(QtdTotalOpcoesGeracaoPedido);
        }

        public void ClickOnProductGridLineDown(int clicks)
        {
            WindowsElement button = FindElementByName("Line down");
            for (int i = 0; i < clicks; i++)
            {
                button.Click();
            }
        }

        public void SelectItemOnProductGridByIndex(int productIndex)
        {
            int intialX = 10;
            int intialY = 144;
            int height = 18;
            try
            {
                if (productIndex == 1)
                {
                    MouseHandler.Click(intialX, intialY);
                }
                else
                {
                    MouseHandler.Click(intialX, intialY + height * (productIndex - 1));
                }
            }
            catch
            {
                WinAppDriver.WaitSeconds(5);
                if (productIndex == 1)
                {
                    MouseHandler.Click(intialX, intialY);
                }
                else
                {
                    MouseHandler.Click(intialX, intialY + height * (productIndex - 1));
                }
            }
        }

        public string SelectItemOnProductGridByCode(string codProduto, bool clickColumn = false, int X = 0)
        {
            WinAppDriver.WaitSeconds(5);
            SelectItemOnProductGridByIndex(1);
            WinAppDriver.WaitSeconds(5);
            WindowsElement scrollBar = FindElementByClassName("ScrollBar");
            string productFound = FindElementOnProductGrid(codProduto, 30, 138, 48, 13, 18, scrollBar, clickColumn: clickColumn, X: X);
            WinAppDriver.WaitSeconds(2);
            return productFound;
        }

        private string FindElementOnProductGrid(string expectedValue, int initialX, int initialY, int width, int height, int rowHeight, WindowsElement scrollBar, int maxAttempts = 9, int elementIndex = 9, bool clickColumn = false, int X = 0)
        {
            bool elementFound = false;
            int attempts = 0;

            //first 9 elements on grid
            string valueFound = OCRScanner.ExtractText(initialX, initialY, width, height, invertedColors: true);
            if (valueFound == expectedValue)
            {
                if (clickColumn)
                {
                    WinAppDriver.WaitSeconds(5);
                    MouseHandler.Click(X, initialY + height / 2);
                }
                return valueFound;
            }
            attempts++;

            while (!elementFound && attempts < maxAttempts)
            {
                initialY += rowHeight;
                valueFound = OCRScanner.ExtractText(initialX, initialY, width, height, invertedColors: false);

                if (valueFound == expectedValue)
                {
                    if (clickColumn)
                    {
                        WinAppDriver.WaitSeconds(5);
                        MouseHandler.Click(X, initialY + height / 2);
                    }
                    return valueFound;
                }
                attempts++;
            }

            //10th element and so on
            bool canScrollDown = CanScrollDown(scrollBar);
            //SelectItemOnProductGridByIndex(elementIndex);
            while (canScrollDown && !elementFound)
            {
                SelectItemOnProductGridByIndex(elementIndex);
                KeyPresser.PressKey("DOWN");
                WinAppDriver.WaitSeconds(1);
                //valueFound on 9th position
                valueFound = OCRScanner.ExtractText(initialX, 282, width, height, invertedColors: true);

                if (valueFound == expectedValue)
                {
                    if (clickColumn)
                    {
                        WinAppDriver.WaitSeconds(5);
                        MouseHandler.Click(X, initialY + height / 2);
                    }
                    return valueFound;
                }
                else
                {
                    canScrollDown = CanScrollDown(scrollBar);
                }
            }
            return $"{expectedValue} not found after searching entire grid.";
        }

        public string GetEmbalagemProductGridValue(string tipoLote)
        {
            SelectProductGrid();
            if (tipoLote == "cd")
            {
                SelectItemOnProductGridByIndex(1);
                ResetColumnLeft();
            }
            string fistItemEmbalagem = OCRScanner.ExtractText(349, 138, 62, 12, invertedColors: true);
            return fistItemEmbalagem.ToUpper();
        }

        public string GetEmbalagemLojasGridValue(string tipoLote)
        {
            // Dado um valor de embalagem do grid de produtos (GetEmbalagemProductGridValue)
            // Quando a primeira loja refletir o mesmo valor
            // E as lojas na sequência forem da mesma diisão
            // Então podemos assumir que o valor da embalagem será o mesmo para elas

            if (tipoLote == "cd")
            {
                string value = OCRScanner.ExtractText(248, 404, 39, 16);
                return value.ToUpper();
            }
            else
            {
                //primeira loja da grid
                MouseHandler.Click(37, 427);
                string actual = OCRScanner.ExtractText(175, 422, 42, 13, invertedColors: true);
                return actual.ToUpper();
            }
        }

        public double GetEstoqueQtd()
        {
            return OCRScanner.TextToDouble(451, 136, 57, 16, invertColors: true);
        }

        public double GetEstqQtdCentral()
        {
            return OCRScanner.TextToDouble(510, 135, 58, 19, invertColors: true);
        }

        public double GetMedVDiaPeriodo()
        {
            return OCRScanner.TextToDouble(1101, 138, 58, 14, invertColors: true);
        }

        public double GetEstqDias()
        {
            return OCRScanner.TextToDouble(641, 137, 38, 15, invertColors: true);
        }

        public void SelectFirstItemWithEstoqueFilled(string tipoLote)
        {
            if (tipoLote == "loja")
            {
                SelectItemOnProductGridByIndex(1);
                KeyPresser.PressKey("PageDown", 5);
            }
            else
            {
                SelectItemOnProductGridByIndex(1);
                KeyPresser.PressKey("PageDown");
                ClickOnProductGridLineDown(6);
                SelectItemOnProductGridByIndex(1);
            }
        }

        public double CalculateEstoqueDiasExpected(double estoqueCD)
        {
            double estoqueQtd = GetEstoqueQtd();
            double mediaVendaDia = GetMedVDiaPeriodo();
            double estoqueDiasExpected = (estoqueQtd + estoqueCD) / mediaVendaDia;
            estoqueDiasExpected = Math.Round(estoqueDiasExpected, 3);
            return estoqueDiasExpected;
        }

        public void OpenConsultaProdutos()
        {
            ReadOnlyCollection<WindowsElement> buttons = FindElementsByClassName("Button");
            WindowsElement button = buttons[6];
            button.Click();
        }

        public double GetEmDiasValue(List<string> lojas, string embalagem)
        {
            OpenConsultaProdutos();
            FilterEmpresasConsultaProdutos(lojas);
            FilterEmbalagem(embalagem);
            FilterEmpresasPromocao("Exceto Promoção");
            WinAppDriver.WaitSeconds(2);
            double value = OCRScanner.TextToDouble(928, 274, 56, 15, invertColors: false);
            return value;
        }

        private void FilterEmpresasPromocao(string listItem)
        {
            MouseHandler.Click(new BoundingRectangle(808, 108, 825, 125));
            WindowsElement option = FindElementByName(listItem);
            option.Click();
            KeyPresser.PressKey("TAB");
        }

        private void FilterEmbalagem(string embalagem)
        {
            MouseHandler.Click(new BoundingRectangle(700, 108, 717, 125));
            embalagem = embalagem.Replace(" ", "  ").ToUpper();
            WindowsElement listItem = FindElementByName(embalagem);
            listItem.Click();
        }

        private void FilterEmpresasConsultaProdutos(List<string> lojas)
        {
            MouseHandler.Click(new BoundingRectangle(519, 653, 603, 675));
            RemoveLojas();
            AddLojasPorNome(lojas);
            ExitWindow("Seleção de Empresas do Lote");
        }

        public double GeteEtqQtdCentral()
        {
            double value = OCRScanner.TextToDouble(512, 137, 54, 15, invertColors: true);
            return value;
        }

        public List<string> GetFormasAbastecimento(string tipoLote)
        {
            List<string> formasAbastecimento = [];

            ClickOnPButton();

            WinAppDriver.WaitSeconds(3);
            GoToTabLogisticaAbastecimento();

            MouseHandler.RightClick(new BoundingRectangle(753, 297, 770, 314));
            KeyPresser.PressKey("DOWN", 3);
            KeyPresser.PressKey("RETURN");

            MouseHandler.Click(50, 250);
            string formasAbastecimentoLoja = OCRScanner.ExtractText(486, 244, 218, 15, invertedColors: true);
            formasAbastecimento.Add(formasAbastecimentoLoja);

            if (tipoLote == "cd")
            {
                MouseHandler.RightClick(new BoundingRectangle(770, 280, 787, 297));
                KeyPresser.PressKey("DOWN", 3);
                KeyPresser.PressKey("RETURN");

                BoundingRectangle upButton = new BoundingRectangle(770, 194, 787, 211);
                MouseHandler.Click(upButton, 8);

                MouseHandler.Click(50, 250);
                string formaAbastecimentoCD = OCRScanner.ExtractText(486, 244, 218, 15, invertedColors: true);
                formasAbastecimento.Add(formaAbastecimentoCD);
            }
            return formasAbastecimento;
        }

        private void ClickOnPButton()
        {
            MouseHandler.Click(new BoundingRectangle(363, 106, 380, 126));
        }

        private void GoToTabLogisticaAbastecimento()
        {
            MouseHandler.Click(465, 160);
        }

        public int GetSuppliersCount(string scenario, List<string> suppliers)
        {
            switch (scenario)
            {
                case "Filtros para Seleção de Produtos":
                    {
                        BoundingRectangle listBoxRect = new BoundingRectangle(27, 163, 274, 273);
                        MouseHandler.Click(listBoxRect);
                        WindowsElement listBox = FindElementByClassName("ListBox");
                        ReadOnlyCollection<AppiumWebElement> listItemElements = listBox.FindElementsByXPath(".//*[@LocalizedControlType='list item']");
                        return listItemElements.Count;
                    }
                case "Grid de Produtos":
                    {
                        OpenSupplierDropdown();
                        WindowsElement comboListBox = FindElementByClassName("ComboLBox");
                        ReadOnlyCollection<AppiumWebElement> listItemElements = comboListBox.FindElementsByXPath(".//*[@LocalizedControlType='list item']");
                        return listItemElements.Count;
                    }
                case "Consulta Lote de Compra":
                    {
                        WinAppDriver.WaitSeconds(7);
                        BoundingRectangle rightEdgeButton = new BoundingRectangle(726, 314, 743, 331);
                        MouseHandler.Click(rightEdgeButton, 3);
                        try
                        {
                            List<string> suppliersFound = [];
                            string firstSupplier = OCRScanner.ExtractText(270, 212, 38, 14, invertedColors: true);
                            suppliersFound.Add(firstSupplier);
                            string secondSupplier = OCRScanner.ExtractText(270, 230, 38, 14, invertedColors: false);
                            suppliersFound.Add(secondSupplier);
                            Assert.IsTrue(suppliersFound.All(suppliers.Contains));
                            return 2;
                        }
                        catch
                        {
                            throw new Exception($"Fornecedores inseridos não encontrados na janela {scenario}");
                        }
                    }
                default:
                    return 0;
            }
        }

        public void ChangeSupplier()
        {
            OpenSupplierDropdown();
            KeyPresser.PressKey("UP");
        }

        public void OpenSupplierDropdown()
        {
            BoundingRectangle dropdown = new BoundingRectangle(799, 81, 816, 98);
            MouseHandler.Click(dropdown);
        }

        public string GetPendReceberValue(string codProduto)
        {
            //int pageDown = 5;
            //KeyPresser.PressKey("PageDown", pageDown);
            //WinAppDriver.WaitSeconds(15);
            //SelectItemOnProductGridByIndex(productIndex);
            SelectItemOnProductGridByCode(codProduto);
            WinAppDriver.WaitSeconds(10);
            return OCRScanner.ExtractText(666, 422, 51, 13, invertedColors: false);
        }

        public void CancelPendReceber()
        {
            OpenPendReceberGridLojas();
            SelectFirstPendReceber();

            string buttonName = "Cancela Item a Receber";
            WindowsElement cancelButton = FindElementByClassAndName("Button", buttonName);
            cancelButton.Click();

            string dialogName = "Cancelamento de itens a receber";
            if (dialogName != null)
            {
                buttonName = "Yes";
                WindowsElement confirmButton = FindElementByClassAndName("Button", buttonName);
                confirmButton.Click();

            }
            else
            {
                throw new Exception($"Erro ao tentar cancelar item a receber.");
            }
        }

        public void OpenPendReceberGridLojas()
        {
            MouseHandler.DoubleClick(700, 430);
        }

        public void SelectFirstPendReceber()
        {
            MouseHandler.Click(60, 185);
        }

        public void ClickOnRefreshButton()
        {
            ReadOnlyCollection<WindowsElement> buttons = FindElementsByClassName("Button");
            WindowsElement refreshButton = buttons[7];
            refreshButton.Click();

            buttons = FindElementsByClassName("Button");
            WindowsElement confirmButton = buttons[0];
            confirmButton.Click();
            WinAppDriver.WaitSeconds(30);
        }
    }
}