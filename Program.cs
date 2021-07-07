using DesafioRPA.Helper;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.IO;

namespace DesafioRPA
{
    class Program
    {
        static void Main(string[] args)
        {
            if (File.Exists(HelperData.fileNameLog))
                File.Delete(HelperData.fileNameLog);
            if (File.Exists(HelperData.fileNameCepRequest))
                File.Delete(HelperData.fileNameCepRequest);
            
            HelperData.CreateResultFile();
            SearchCep(HelperData.ExtractData());
        }

        public static void SearchCep(List<(int, int)> listCepRange)
        {
            try
            {
                HelperData.Log(HelperData.fileNameLog, $"Inicia processo de busca");
                listCepRange.ForEach(listCep =>
                {
                    HelperData.Log(HelperData.fileNameLog, $"Inicia busca pela faixa de CEP {listCep.Item1} - {listCep.Item2}");

                    var rangeCepInicial = listCep.Item1.ToString();
                    rangeCepInicial = rangeCepInicial.PadLeft(8, '0');
                    var inicio = rangeCepInicial.EndsWith('0') ? 0 : 1;
                    rangeCepInicial = rangeCepInicial.Substring(0, rangeCepInicial.Length - 3);

                    var rangeCepFinal = listCep.Item2.ToString();
                    rangeCepFinal = rangeCepFinal.PadLeft(8, '0');
                    rangeCepFinal = rangeCepFinal.Substring(0, rangeCepFinal.Length - 3);

                    ChromeOptions opt = new ChromeOptions();
                    opt.AddArguments("--headless");
                    opt.AddArguments("--disable-gpu");
                    using (var driver = new ChromeDriver(opt))
                    {
                        //Homepage do correios 
                        driver.Navigate().GoToUrl("https://buscacepinter.correios.com.br/app/endereco/index.php");

                        var faixaInicial = Convert.ToInt32(rangeCepInicial);
                        var faixaFinal = Convert.ToInt32(rangeCepFinal);
                        for (int faixa = faixaInicial; faixa <= faixaFinal; faixa++)
                        {
                            for (int i = inicio; i <= 999; i++)
                            {
                                var sufixo = i.ToString().PadLeft(3, '0');
                                var cep = faixa + sufixo;

                                //Pega os elementos necessarios para a pesquisa
                                if (driver.PageSource.Contains("503 Service Unavailable")) 
                                {
                                    var tentativaDeReconexao = 3;
                                    var count = 0;
                                    while ((count != tentativaDeReconexao) && driver.PageSource.Contains("503 Service Unavailable")) {
                                        driver.Navigate().GoToUrl("https://buscacepinter.correios.com.br/app/endereco/index.php");
                                        count++;
                                    }

                                    if (driver.PageSource.Contains("503 Service Unavailable"))
                                    {
                                        HelperData.Log(HelperData.fileNameLog, "Servidor do correio indisponível");
                                        return;
                                    }
                                }
                                var address = driver.FindElementById("endereco");
                                var searchButton = driver.FindElementById("btn_pesquisar");

                                address.SendKeys(cep); ;
                                searchButton.Click();

                                #region Busca as informações encontradas
                                var table = driver.FindElementByXPath("//table[@id='resultado-DNEC']//tbody");
                                var rows = table.FindElements(By.TagName("tr"));
                                var isSaved = false;
                                foreach (var row in rows)
                                {
                                    var rowTds = row.FindElements(By.TagName("td"));
                                    if (rowTds.Count > 0)
                                    {
                                        InfoLocation infoLocation = new InfoLocation();
                                        infoLocation.Logradouro = rowTds[0].Text;
                                        infoLocation.Bairro = rowTds[1].Text;
                                        infoLocation.LocalidadeUF = rowTds[2].Text;
                                        infoLocation.CEP = rowTds[3].Text;
                                        infoLocation.Data = DateTime.Now;
                                        HelperData.WriteResultFile(infoLocation);
                                        isSaved = true;
                                    }
                                }
                                #endregion
                            
                                HelperData.Log(HelperData.fileNameCepRequest, $"{cep}{(isSaved?"- Encontrado":"")}");

                                var backButton = driver.FindElementByXPath("//div[@id='retornar']//div//div//div//button[@id='btn_voltar']");
                                if (backButton.Selected)
                                    backButton.Click();
                                else
                                    driver.Navigate().GoToUrl("https://buscacepinter.correios.com.br/app/endereco/index.php");

                            }
                        }
                    }
                    HelperData.Log(HelperData.fileNameLog, $"Finaliza busca pela faixa de CEP  {listCep.Item1} - {listCep.Item2}");
                });
                HelperData.Log(HelperData.fileNameLog, $"Finaliza processo de busca");
            }
            catch (InvalidOperationException ex)
            {
                HelperData.WriteExceptionLog(ex.GetType().FullName, ex.Message);
            }
        }
    }
}
