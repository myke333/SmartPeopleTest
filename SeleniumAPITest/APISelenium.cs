using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Chrome;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System.Collections.ObjectModel;
using OpenQA.Selenium.Interactions;
using System.Windows.Forms;
using System.Diagnostics;
using System.Drawing;
using Keys = OpenQA.Selenium.Keys;
using Keyboard = System.Windows.Forms.SendKeys;

namespace APITest
{
    public class APISelenium
    {
        ChromeDriver driver;
        APIFuncionesVitales FV = new APIFuncionesVitales();
        String mainWin;
        String auxModalWin;

        public APISelenium()
        {
        }
        public string modeSelection() {
            return "normal";
        }
        public void LoginApps(string App, string User, string Password, string UrlApp, string file) //LOGIN SELFSERVICE Y RL
        {
            var options = new ChromeOptions();           
            if (modeSelection() == "normal") {
                
                options.AddArgument("-no-sandbox");
                driver = new ChromeDriver(@"C:\deployment\", options, TimeSpan.FromSeconds(240));
                SetImplicitTimeoutSeconds(240);
                driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(240);
                driver.Manage().Window.Maximize();
                driver.Navigate().GoToUrl(UrlApp);

               
            }
            else if(modeSelection()=="headless") {
                options.AddArguments(new List<string>() {
                                "--headless",
                                "--disable-gpu",
                                "--disable-software-rasterizer",
                                "--log-level=3",
                                "--window-size=1600x900"
                                });

                //options.AddArguments("--disable-gpu");
                //options.AddArguments("--disable-software-rasterizer");
                //options.AddArguments("--ignore-gpu-blacklist");
                //options.AddArguments("--use-angle=gl");
                driver = new ChromeDriver(@"C:\deployment\", options, TimeSpan.FromSeconds(240));
                SetImplicitTimeoutSeconds(240);
                driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(240);

                //driver.Manage().Window.Maximize();
                driver.Navigate().GoToUrl(UrlApp);

                Thread.Sleep(500);
                
            }
            Thread.Sleep(1500);
            Screenshot("Abrir programa", true, file);


            if (App.ToLower() == "selfservice" || App.ToLower() == "smartpeople")
            {
                SendKeys("//input[contains(@name,'txtCodUsua')]", User);
                Thread.Sleep(2000);
                Click("//input[contains(@name,'txtPasUsua')]");
                Thread.Sleep(2000);
                SendKeys("//input[contains(@name,'txtPasUsua')]", Keys.Enter);
                Thread.Sleep(2000);
                SendKeys("//input[contains(@name,'txtPasUsua')]", Password);
                Thread.Sleep(2000);
                SendKeys("//input[contains(@name,'txtPasUsua')]", Password);
                Thread.Sleep(2000);
                Keyboard.SendWait("{TAB}");
                Thread.Sleep(5000);
                Click("//*[@id='btnIngresar']");
                Thread.Sleep(2000);

            }
            else if (App.ToLower() == "reclutamiento")
            {
                //Clic a ventana emergente

                Click("//div[@id='dlogin']");
                Thread.Sleep(2000);
                SendKeys("//input[@id='txtMail']", User);
                Thread.Sleep(2000);
                Thread.Sleep(500);
                Click("//input[@id='txtPassword']");
                Thread.Sleep(2000);
                SendKeys("//input[@id='txtPassword']", Password);
                Thread.Sleep(2000);
                driver.FindElement(By.XPath("//input[contains(@id,'btnIngresar')]")).Click();
            }
            Thread.Sleep(2000);
            Screenshot("Login", true, file);

           
        }
        public void Screenshot(string maestro, bool bandera, string file)
        {

            //Image MyImage = null;
            //string UrlImage = @"C:\Reportes\" + maestro + ".bmp";
            //MyImage = ((ITakesScreenshot)session).GetScreenshot();
            Screenshot image = ((ITakesScreenshot)driver).GetScreenshot();
            string name = maestro;
            string path = @"C:\Reportes\" + maestro + ".bmp";
            Image imgSource;
            image.SaveAsFile(string.Format("C:\\Reportes\\{0}.bmp", name), ScreenshotImageFormat.Bmp);
            //image.Save(path, System.Drawing.Imaging.ImageFormat.Bmp);
            APIFuncionesVitales.InsertAPicture(file, path, maestro, bandera);
        }

        public void Screenshotv2(string maestro, bool bandera, string file)
        {

            //Image MyImage = null;
            //string UrlImage = @"C:\Reportes\" + maestro + ".bmp";
            //MyImage = ((ITakesScreenshot)session).GetScreenshot();
            Screenshot image = ((ITakesScreenshot)driver).GetScreenshot();
            string name = maestro;
            string path = @"C:\Reportes\" + maestro + ".bmp";
            Image imgSource;
            image.SaveAsFile(string.Format("C:\\Reportes\\{0}.bmp", name), ScreenshotImageFormat.Bmp);
            //image.Save(path, System.Drawing.Imaging.ImageFormat.Bmp);
            APIFuncionesVitales.InsertAPicture(file, path, maestro, bandera);
        }


        public void Click(string control) //DA CLICK SOBRE EL CONTROL ENVIADO
        {
            driver.FindElement(By.XPath(control)).Click();
        }

        public void SendKeys(string control, string value) //ESCRIBE EN UNA CAJA DE TEXTO
        {
            //driver.FindElement(By.XPath(control)).Click();
            driver.FindElement(By.XPath(control)).Clear();
            driver.FindElement(By.XPath(control)).SendKeys(value);
        }

        public void SendKeys1(string control, string value) //ESCRIBE EN UNA CAJA DE TEXTO
        {
            //driver.FindElement(By.XPath(control)).Click();
            //driver.FindElement(By.XPath(control)).Clear();
            driver.FindElement(By.XPath(control)).SendKeys(value);
        }

        public void Scroll(string control) //HACE SCROLL HASTA QUE ENCUENTE EL CONTROL
        {
            Thread.Sleep(2000);
            IWebElement endScroll = driver.FindElement(By.XPath(control));
            IJavaScriptExecutor js = driver as IJavaScriptExecutor;
            js.ExecuteScript("arguments[0].scrollIntoView(true);", endScroll);
            Thread.Sleep(2000);
        }

        public void ScrollTo(string x, string y)
        {
            var js = String.Format($"window.scrollTo({x}, {y})");
            IJavaScriptExecutor scroll = driver as IJavaScriptExecutor;
            scroll.ExecuteScript(js);
        }

        public void SelectElementByName(string control, string value) //SELECCIONA EL NOMBRE DE UNA LISTA DESPLEGABLE
        {
            SelectElement select = new SelectElement(driver.FindElement(By.XPath(control)));
            select.SelectByText(value);
        }

        public void ChangeMainWindow() //CAMBIA A LA VENTANA PRINCIPAL
        {
            driver.SwitchTo().Window(mainWin);
        }

        public void ChangeAuxWindow() //CAMBIA A UNA VENTANA AUXILIAR SI LA HAY
        {
            ReadOnlyCollection<String> windowHandles = driver.WindowHandles;
            mainWin = windowHandles[0];
            auxModalWin = windowHandles[windowHandles.Count - 1];
        }

        public void ChangeWindow(String window)
        {
            driver.SwitchTo().Window(window);
        }

        public void MaximizeWindow()
        {
            driver.Manage().Window.Maximize();
        }

        public int CountWindow()
        {
            ReadOnlyCollection<String> windowHandles = driver.WindowHandles;
            return windowHandles.Count();
        }

        public String MainWindow()
        {
            ReadOnlyCollection<String> windowHandles = driver.WindowHandles;
            return windowHandles[0];
        }

        public String PopupWindow()
        {
            ReadOnlyCollection<String> windowHandles = driver.WindowHandles;
            return windowHandles[windowHandles.Count - 1];
        }

        public bool IsPresent(string control) //DEVUELVE BOOLEANO SI ESTA PRESENTE EL CONTROL A BUSCAR
        {
            return driver.FindElement(By.XPath(control)).Displayed;
        }

        public string GetText(string control) //OBTIENE EL TEXTO SEGUN EL CONTROL INDICADO
        {
            return driver.FindElement(By.XPath(control)).Text;
        }

        public bool ExistControl(string control) //INTENTA BUSCAR EL CONTROL ESPECIFICADO
        {
            SetImplicitTimeoutSeconds(7);
            List<IWebElement> Exist = new List<IWebElement>();
            Exist.AddRange(driver.FindElements(By.XPath(control)));
            if (Exist.Count > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public int CountControl(string control) //INTENTA BUSCAR EL CONTROL ESPECIFICADO
        {
            SetImplicitTimeoutSeconds(5);
            List<IWebElement> Exist = new List<IWebElement>();
            Exist.AddRange(driver.FindElements(By.XPath(control)));
            if (Exist.Count > 0)
            {
                return Exist.Count;
            }
            else
            {
                return 0;
            }
        }

        public void SetImplicitTimeoutSeconds(int seconds) //SETEA UN TIEMPO IMPLICITO DE ESPERA EN SEGUNDOS
        {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(seconds);
        }

        public void Clear(string control) //LIMPIA EL CAMPO ENVIADO COMO PARAMETRO
        {
            driver.FindElement(By.XPath(control)).Clear();
        }

        public void Close()
        {
            driver.Close();
        }

        public void Dispose()
        {
            driver.Dispose();
        }

        public void Quit()
        {
            driver.Quit();
        }

        public void AcceptAlert()
        {

            //driver.SwitchTo().Alert().Accept();
            Thread.Sleep(5000);
            Keyboard.SendWait("{ENTER}");
            Thread.Sleep(5000);

        }

        public ChromeDriver returnDriver()
        {
            return driver;
        }

        public string GetTextFromTextBox(string control)
        {
            return driver.FindElement(By.XPath(control)).GetAttribute("value");
        }

        public void Enter(string control)
        {
            driver.FindElement(By.XPath(control)).SendKeys(Keys.Enter);
        }

        public void Tab(string control)
        {
            driver.FindElement(By.XPath(control)).SendKeys(Keys.Tab);
        }

        public void ActiveElement()
        {
            driver.SwitchTo().ActiveElement();
        }

        public string Title()
        {
            string Titulo = driver.Title;
            return Titulo;
        }

        public string Subtitulo(string control)
        {
            IWebElement subtittle = driver.FindElement(By.Id(control));
            string Subtittle = subtittle.GetAttribute("data-original-title");
            return Subtittle;
        }

        public string Emergente(string control)
        {
            IWebElement subtittle = driver.FindElement(By.XPath(control));
            string Subtittle = subtittle.GetAttribute("data-original-title");
            return Subtittle;
        }

        public string EmergenteBotones(string control)
        {
            IWebElement icono = driver.FindElement(By.Id(control));
            string tittleIcono = icono.GetAttribute("title");
            return tittleIcono;
        }

        public string CamposEmergentes(string control, string Nombre, string file)
        {
            Click("//div[@id='ctl00_pBotones']/div");
            Thread.Sleep(500);
            Click(control);
            Thread.Sleep(1000);
            Screenshotv2(Nombre, true, file);
            Thread.Sleep(100);
            string var = Emergente(control);
            Thread.Sleep(100);
            return var;
        }

        public void ValTabs(string file)
        {

            Click("//div[@id='ctl00_pBotones']/div");
            List<IWebElement> elementList = new List<IWebElement>();
            Thread.Sleep(800);
            elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
                                            
            if (elementList.Count > 0)
            {
                elementList[0].Click();
                Thread.Sleep(1500);
                for (int i = 0; i < elementList.Count; i++)
                {
                    try
                    {
                        var tab = elementList[i];
                        tab.SendKeys(Keys.Tab);
                        Screenshot("TAB", true, file);

                        Thread.Sleep(100);
                    }
                    catch (Exception e)
                    {
                        var tab = elementList[i];
                        Tab("//*[contains(@name,'ctl00$ContenidoPagina$')]");
                        Screenshot("TAB", true, file);
                    }
                }
            }
        }

        public void ClockAut(string Time, string Horario, int numNthChild) //REALIZA EL RELOJ APLICANDO LA HORA INGRESADA POR EL MTM //numNthChild hace referencia al reloj, el primer reloj es el 5 y el segundo es el 6 y asi sucesivamente
        {
            string firstTime = null;
            string secondTime = null;
            //Debugger.Launch();
            //Confirmo el formato de la Hora
            int Tamano = Time.Length;
            if (Tamano > 2)
            {
                //Obtengo valores de las horas parametrizadas en una Lista
                char delimitador = ':';
                string[] horas = Time.Split(delimitador);
                firstTime = horas[0];
                secondTime = horas[1];
            }
            else
            {
                firstTime = Time;
            }
            //Cambio de Ventana
            ChangeAuxWindow();
            //Aplico la hora en el reloj
            switch (firstTime)
            {
                case "1":
                    Thread.Sleep(100);
                    driver.FindElement(By.CssSelector(".mdtimepicker:nth-child("+numNthChild+") .mdtp__hour_holder > .rotate-120 > span")).Click();
                    break;
                case "2":
                    Thread.Sleep(100);
                    driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__hour_holder > .rotate-150 > span")).Click();
                    break;
                case "3":
                    Thread.Sleep(100);
                    driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__hour_holder > .rotate-180 > span")).Click();
                    break;
                case "4":
                    Thread.Sleep(100);
                    driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__hour_holder > .rotate-210 > span")).Click();
                    break;
                case "5":
                    Thread.Sleep(100);
                    driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__hour_holder > .rotate-240 > span")).Click();
                    break;
                case "6":
                    Thread.Sleep(100);
                    driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__hour_holder > .rotate-270 > span")).Click();
                    break;
                case "7":
                    Thread.Sleep(100);
                    driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__hour_holder > .rotate-300 > span")).Click();
                    break;
                case "8":
                    Thread.Sleep(100);
                    driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__hour_holder > .rotate-330 > span")).Click();
                    break;
                case "9":
                    Thread.Sleep(100);
                    driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__hour_holder > .rotate-0 > span")).Click();
                    break;
                case "10":
                    Thread.Sleep(100);
                    driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__hour_holder > .rotate-30 > span")).Click();
                    break;
                case "11":
                    Thread.Sleep(100);
                    driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__hour_holder > .rotate-60 > span")).Click();
                    break;
                case "12":
                    Thread.Sleep(100);
                    driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__hour_holder > .rotate-90 > span")).Click();
                    break;
                default:
                    Thread.Sleep(100);
                    driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__hour_holder > .rotate-120 > span")).Click();
                    Debug.WriteLine("La hora ingresada se encuentra fuera del rango de horas naturales");
                    break;
            }

            if (secondTime != null)
            {
                switch (secondTime)
                {
                    case "05":
                        Thread.Sleep(100);
                        driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__minute_holder > .rotate-120 > span")).Click();
                        break;
                    case "10":
                        Thread.Sleep(100);
                        driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__minute_holder > .rotate-150 > span")).Click(); break;
                    case "15":
                        Thread.Sleep(100);
                        driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__minute_holder > .rotate-180 > span")).Click();
                        break;
                    case "20":
                        Thread.Sleep(100);
                        driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__minute_holder > .rotate-210 > span")).Click();
                        break;
                    case "25":
                        Thread.Sleep(100);
                        driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__minute_holder > .rotate-240 > span")).Click();
                        break;
                    case "30":
                        Thread.Sleep(100);
                        driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__minute_holder > .rotate-270 > span")).Click();
                        break;
                    case "35":
                        Thread.Sleep(100);
                        driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__minute_holder > .rotate-300 > span")).Click();
                        break;
                    case "40":
                        Thread.Sleep(100);
                        driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__minute_holder > .rotate-330 > span")).Click();
                        break;
                    case "45":
                        Thread.Sleep(100);
                        driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__minute_holder > .rotate-0 > span")).Click();
                        break;
                    case "50":
                        Thread.Sleep(100);
                        driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__minute_holder > .rotate-30 > span")).Click();
                        break;
                    case "55":
                        Thread.Sleep(100);
                        driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__minute_holder > .rotate-60 > span")).Click();
                        break;
                    case "00":
                        Thread.Sleep(100);
                        driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__minute_holder > .rotate-90 > span")).Click();
                        break;
                    default:
                        Thread.Sleep(100);
                        driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__minute_holder > .rotate-90 > span")).Click();
                        Debug.WriteLine("Los minutos ingresada se encuentra fuera del rango de minutos principales");
                        break;
                }
            }
            else
            {
                //Si el segundo dato es nulo seleccione 00
                Thread.Sleep(100);
                driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__minute_holder > .rotate-90 > span")).Click();
            }


            //Configuro el Horario
            if (Horario != null)
            {
                string horario = Horario.ToLower();
                switch (horario)
                {
                    case "am":
                        Thread.Sleep(100);
                        driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__am")).Click();
                        break;
                    case "pm":
                        Thread.Sleep(100);
                        driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__pm")).Click();
                        break;
                    default:
                        Thread.Sleep(100);
                        driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__pm")).Click();
                        Debug.WriteLine("El horario ingresado es diferente a los horarios existentes");
                        break;
                }
            }
            else
            {
                //Si Horario es nulo selecciona PM
                Thread.Sleep(100);
                driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .mdtp__pm")).Click();
            }

            //Aplico
            Thread.Sleep(300);
            driver.FindElement(By.CssSelector(".mdtimepicker:nth-child(" + numNthChild + ") .ok")).Click();

        }


    }
}
